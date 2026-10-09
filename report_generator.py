# -*- coding: utf-8 -*-
"""
DOE工艺验证报告自动生成器
基于GT工艺验证模板，自动填充数据和图表
"""
import os
import sys
import tempfile
import numpy as np
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

import parameter as P

# 模板所需的最小页数（封面/汇总表/验证信息/DOE参数/位置精度/可靠性/切片/附件）
# 各 slide 模块按此页序复制母版，模板页序变动会静默错位，故在此校验
EXPECTED_TEMPLATE_SLIDES = 7


def generate_charts(data, specs):
    """生成所有图表。data/specs 由调用方提供，避免重复解析 Excel。"""
    P.ensure_image_dirs()

    print("=" * 50)
    print("生成位置精度图...")
    import position

    groups_data, group_names, devices_per_group = (
        position.build_precision_inputs(data["位置精度"])
    )
    for group in group_names:
        devs = devices_per_group.get(group, [])
        out_path = os.path.join(P.POSITION_IMG_DIR, f"{group}_position.png")
        position.plot_single_profile(
            profile_name=group,
            profile_data=groups_data[group],
            devices=devs,
            out_path=out_path,
            specs=specs,
        )

    print("=" * 50)
    print("生成可靠性图...")
    import reliability
    all_data = {}
    for cond, (sheet_key, block_name) in P.RELIABILITY_MAP.items():
        if sheet_key not in data:
            continue
        sheet_data = data[sheet_key]
        if block_name not in sheet_data:
            continue
        block_data = sheet_data[block_name]
        all_data[cond] = {}
        for prof, content in block_data.items():
            all_data[cond][prof] = {}
            for dev_name in content:
                if dev_name == "SN":
                    continue
                if isinstance(content[dev_name], dict):
                    all_data[cond][prof][dev_name] = content[dev_name]

    all_devices = set()
    for cond_data in all_data.values():
        for prof_data in cond_data.values():
            for dev_name in prof_data:
                if dev_name != "SN":
                    all_devices.add(dev_name)

    for dev_name in sorted(all_devices):
        dim_name = None
        for cond_data in all_data.values():
            for prof_data in cond_data.values():
                dev_data = prof_data.get(dev_name, {})
                for k in dev_data:
                    if k not in ("SN", "统计"):
                        dim_name = k
                        break
                if dim_name:
                    break
            if dim_name:
                break
        if not dim_name:
            dim_name = "剪切力"
        out_path = os.path.join(P.RELIABILITY_IMG_DIR,
                                f"{dev_name}_reliability.png")
        reliability.plot_device(dev_name, dim_name, all_data, out_path, specs)


def load_data(data_file):
    """加载所有数据"""
    import data_reader
    return data_reader.main(data_file)


def set_text_box(shape, text, font_size=12, bold=False):
    """设置文本框内容（保留原有格式）"""
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    tf.word_wrap = True


def replace_in_shape(shape, old, new):
    """在文本框中替换文字，保留原有格式"""
    tf = shape.text_frame
    for para in tf.paragraphs:
        for run in para.runs:
            if old in run.text:
                run.text = run.text.replace(old, new)


def replace_by_line(shape, replacements):
    """按行替换文本，保留原有格式。replacements: [(old, new), ...]"""
    tf = shape.text_frame
    for para in tf.paragraphs:
        for run in para.runs:
            for old, new in replacements:
                if old in run.text:
                    run.text = run.text.replace(old, new)


def build_summary_data(data):
    """构建汇总表数据"""
    summary = []
    spec = data.get("物料信息", {}).get("spec", {})
    for dev_name, vals in spec.items():
        pos_spec = vals.get("位置", "N/A")
        shear_spec = vals.get("剪切力", "N/A")
        summary.append([dev_name, f"精度{pos_spec}",
                        f"剪切力{shear_spec}", "待验证", "待定"])
    return summary


def build_doe_params(data):
    """构建DOE参数表"""
    material_info = data.get("物料信息", {})
    doe_params = material_info.get("doe_params", [])
    if not doe_params:
        return []
    headers = ["Profile", "NO.", "Sample size", "Pre-heating Temp/℃",
               "Ramp rate\n℃/s", "Solder Temp\n℃", "Press Time\ns",
               "Solder Time\ns", "Cooling Rate\n℃/s", "Cooling Temp\n℃",
               "Solder Force\ngf"]
    rows = [headers]
    for param in doe_params:
        row = [param.get("Profile", ""), param.get("NO.", ""),
               param.get("Sample size", ""), param.get("Pre-heating Temp", ""),
               param.get("Ramp rate", ""), param.get("Solder Temp", ""),
               param.get("Press Time", ""), param.get("Solder Time", ""),
               param.get("Cooling Rate", ""), param.get("Cooling Temp", ""),
               param.get("Solder Force", "")]
        rows.append(row)
    return rows


def _parse_pos_bound(pos_str):
    """从 '±0.02' 解析精度SPEC数值"""
    import re
    m = re.search(r'[±+-]?([\d.]+)', str(pos_str))
    return float(m.group(1)) if m else None


def _fmt_void_spec(v):
    """空洞率格式化：单元格里已是百分数（0.6 = 0.6%），直接输出不换算"""
    if v is None or v == "":
        return "N/A"
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    return f"{f:g}"


def build_slide2_table_data(data):
    """构建第二页汇总表格数据
    精度结果 = 类似作图：各样本减均值后，标准差的最大值（排除angle）
    Die shear结果 = 所有可靠性条件、所有profile、所有样本剪切力原始数据的最小值
    结论 = 精度结果<=精度SPEC 且 Die shear结果>=剪切力SPEC 且 空洞率结果<=空洞率SPEC ? PASS : FAIL
    空洞率SPEC 来自物料信息 SPEC 表“空洞率”列，空洞率结果取“空洞率”sheet 各
    profile 实测最大值（最差）；结果为空时显示 "/" 且不参与判定
    """
    spec = data.get("物料信息", {}).get("spec", {})
    position_data = data.get("位置精度", {})
    rel_keys = ["可靠性-T0", "可靠性-T168", "可靠性-T500"]

    precision_results = {}
    for dev_name in spec.keys():
        profile_stds = []
        for group_name, group_data in position_data.items():
            if group_name == "统计":
                continue
            if dev_name not in group_data:
                continue
            dev_data = group_data[dev_name]
            for dim, raw in dev_data.items():
                if dim == "统计":
                    continue
                if str(dim).lower() == "angle":
                    continue
                if not isinstance(raw, list):
                    continue
                vals = [float(v) for v in raw if v is not None]
                if not vals:
                    continue
                m = np.mean(vals)
                centered = [v - m for v in vals]
                profile_stds.append(float(np.std(centered)))
        precision_results[dev_name] = max(profile_stds) if profile_stds else 0.0

    # 原逻辑：各组剪切力均值的最小值（已弃用）
    # shear_results = {}
    # for dev_name in spec.keys():
    #     group_means = []
    #     for key in rel_keys:
    #         if key not in data:
    #             continue
    #         for block_name, block_content in data[key].items():
    #             for prof_name, prof_content in block_content.items():
    #                 if dev_name not in prof_content:
    #                     continue
    #                 stats = prof_content[dev_name].get("统计", {})
    #                 m = stats.get("mean", {}).get("剪切力")
    #                 if isinstance(m, (int, float)):
    #                     group_means.append(m)
    #     shear_results[dev_name] = min(group_means) if group_means else 0.0

    # 新逻辑：所有可靠性条件、所有profile、所有样本原始剪切力的最小值
    shear_results = {}
    for dev_name in spec.keys():
        all_vals = []
        for key in rel_keys:
            if key not in data:
                continue
            for block_name, block_content in data[key].items():
                for prof_name, prof_content in block_content.items():
                    if dev_name not in prof_content:
                        continue
                    raw = prof_content[dev_name].get("剪切力", [])
                    if isinstance(raw, list):
                        for v in raw:
                            if isinstance(v, (int, float)) and not isinstance(v, bool):
                                all_vals.append(float(v))
                    elif isinstance(raw, (int, float)) and not isinstance(raw, bool):
                        all_vals.append(float(raw))
        shear_results[dev_name] = min(all_vals) if all_vals else 0.0

    # 空洞率结果：取该器件所有 profile 实测值的最大值（最差），“空洞率”sheet
    void_results = {}
    void_data = data.get("空洞率", {}) or {}
    for dev_name in spec.keys():
        vals = []
        for prof_vals in void_data.values():
            if not isinstance(prof_vals, dict):
                continue
            v = prof_vals.get(dev_name)
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                vals.append(float(v))
        void_results[dev_name] = max(vals) if vals else None

    rows = []
    for dev_name, spec_val in spec.items():
        pos_spec = spec_val.get("位置", "N/A")
        shear_spec = spec_val.get("剪切力", "N/A")
        prec_result = precision_results.get(dev_name, 0.0)
        shear_result = shear_results.get(dev_name, 0.0)

        # 结论：精度 + Die shear + 空洞率三项都满足SPEC -> PASS
        pos_bound = _parse_pos_bound(pos_spec)
        pass_pos = pos_bound is not None and prec_result <= pos_bound
        pass_shear = (isinstance(shear_spec, (int, float))
                      and shear_result >= shear_spec)
        void_spec_raw = spec_val.get("空洞率")
        void_result = void_results.get(dev_name)
        pass_void = None
        if (isinstance(void_spec_raw, (int, float))
                and not isinstance(void_spec_raw, bool)
                and void_result is not None):
            pass_void = void_result <= float(void_spec_raw)

        if not pass_pos or not pass_shear or pass_void is False:
            conclusion = "FAIL"
        else:
            # 空洞率结果为空时不参与判定（只看精度 + Die shear）
            conclusion = "PASS"

        shear_spec_s = f"{shear_spec:.2f}" if isinstance(shear_spec, (int, float)) else str(shear_spec)
        void_spec = _fmt_void_spec(spec_val.get("空洞率"))
        void_result_s = _fmt_void_spec(void_result) if void_result is not None else "/"
        rows.append([dev_name, str(pos_spec), f"{prec_result:.4f}",
                      shear_spec_s, f"{shear_result:.2f}",
                      void_spec, void_result_s, conclusion, ""])
    return rows


def set_cell_text(cell, text):
    """设置单元格文字，尽量保留原格式"""
    tf = cell.text_frame
    if tf.paragraphs and tf.paragraphs[0].runs:
        p0 = tf.paragraphs[0]
        p0.runs[0].text = str(text)
        for r in p0.runs[1:]:
            r.text = ""
        for p in tf.paragraphs[1:]:
            for r in p.runs:
                r.text = ""
    else:
        cell.text = str(text)


def _add_table_row(table):
    """复制最后一行来添加新行，保留格式"""
    import copy
    from pptx.oxml.ns import qn
    rows = table._tbl.findall(qn('a:tr'))
    if not rows:
        return
    new_tr = copy.deepcopy(rows[-1])
    table._tbl.append(new_tr)


def _remove_table_rows(table, keep):
    """删除第 keep 行之后的多余行（模板预置的数据行多于实际数据时用）"""
    from pptx.oxml.ns import qn
    trs = table._tbl.findall(qn('a:tr'))
    if keep < 1:
        keep = 1  # 至少保留表头
    for tr in trs[keep:]:
        table._tbl.remove(tr)


def _add_table_col(table):
    """复制最后一列来添加新列，保留格式，并均摊总宽度"""
    import copy
    from pptx.oxml.ns import qn
    from pptx.util import Emu
    tbl = table._tbl
    grid = tbl.find(qn('a:tblGrid'))
    if grid is None:
        return
    cols = grid.findall(qn('a:gridCol'))
    if not cols:
        return
    total_w = sum(int(c.get('w', 0)) for c in cols)
    new_col = copy.deepcopy(cols[-1])
    grid.append(new_col)
    cols = grid.findall(qn('a:gridCol'))
    # 均摊宽度，保持总宽不变
    n = len(cols)
    each = total_w // n if n else 0
    for c in cols:
        c.set('w', str(each))
    # 每行复制最后一个单元格
    for tr in tbl.findall(qn('a:tr')):
        tcs = tr.findall(qn('a:tc'))
        if tcs:
            tr.append(copy.deepcopy(tcs[-1]))


# 表格尺寸强制值：填数字(cm)则覆盖模板原尺寸；设为 None 则沿用模板原尺寸
SLIDE2_TABLE_WIDTH_CM = None
SLIDE2_TABLE_HEIGHT_CM = None
SLIDE2_TABLE_FONT = "Times New Roman"

SLIDE3_TABLE_WIDTH_CM = None
SLIDE3_TABLE_HEIGHT_CM = None
FONT_EA = "微软雅黑"
FONT_LATIN = "Times New Roman"
TABLE_FONT_SIZE_PT = 14


def _set_run_font(run, latin_font, ea_font=None, size_pt=None):
    """设置run字体：西文latin_font，中文ea_font（缺省同latin）"""
    if ea_font is None:
        ea_font = latin_font
    run.font.name = latin_font
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(qn('a:latin'))
    if latin is None:
        latin = rPr.makeelement(qn('a:latin'), {'typeface': latin_font})
        rPr.append(latin)
    else:
        latin.set('typeface', latin_font)
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {'typeface': ea_font})
        latin.addnext(ea)
    else:
        ea.set('typeface', ea_font)
    cs = rPr.find(qn('a:cs'))
    if cs is None:
        cs = rPr.makeelement(qn('a:cs'), {'typeface': latin_font})
        ea.addnext(cs)
    else:
        cs.set('typeface', latin_font)


def _fix_table_style(shape, latin_font, ea_font, width_cm, height_cm,
                     size_pt=None):
    """统一表格字体，并调整表格尺寸（宽×高，cm）。

    width_cm / height_cm 传 None 时沿用模板原尺寸（只重排行高列宽）。
    """
    if size_pt is None:
        size_pt = TABLE_FONT_SIZE_PT
    table = shape.table
    for row in table.rows:
        for cell in row.cells:
            tf = cell.text_frame
            for para in tf.paragraphs:
                for run in para.runs:
                    _set_run_font(run, latin_font, ea_font, size_pt)
            for latin in cell._tc.iter(qn('a:latin')):
                latin.set('typeface', latin_font)
            for ea in cell._tc.iter(qn('a:ea')):
                ea.set('typeface', ea_font)
            for cs in cell._tc.iter(qn('a:cs')):
                cs.set('typeface', latin_font)
            for def_rpr in cell._tc.iter(qn('a:defRPr')):
                def_rpr.set('sz', str(int(size_pt * 100)))
            for end_rpr in cell._tc.iter(qn('a:endParaRPr')):
                end_rpr.set('sz', str(int(size_pt * 100)))

    n_rows = len(table.rows)
    n_cols = len(table.columns)
    if n_rows and n_cols:
        target_w = Cm(width_cm) if width_cm else shape.width
        target_h = Cm(height_cm) if height_cm else shape.height
        cur_w = sum(c.width for c in table.columns) or 1
        acc_w = 0
        for i, col in enumerate(table.columns):
            if i == n_cols - 1:
                col.width = target_w - acc_w
            else:
                w = int(col.width * target_w / cur_w)
                col.width = w
                acc_w += w
        row_h = target_h // n_rows
        acc_h = 0
        for i, row in enumerate(table.rows):
            if i == n_rows - 1:
                row.height = target_h - acc_h
            else:
                row.height = row_h
                acc_h += row_h
        shape.width = target_w
        shape.height = target_h


def _fix_slide2_table_style(shape, font_name, width_cm, height_cm):
    """第二页：全表统一为font_name"""
    _fix_table_style(shape, font_name, font_name, width_cm, height_cm)


def build_slide3_material_rows(data):
    """构建第三页底部物料表数据行：[名称, 料号, 供应商, 批次]"""
    material_info = data.get("物料信息", {}).get("material_info", {})
    names = material_info.get("器件行") or []
    part_nos = material_info.get("料号") or []
    suppliers = material_info.get("供应商信息") or []
    batches = material_info.get("批次信息") or []
    if not isinstance(part_nos, list):
        part_nos = [part_nos]
    if not isinstance(suppliers, list):
        suppliers = [suppliers]
    if not isinstance(batches, list):
        batches = [batches]

    rows = []
    n = max(len(names), len(part_nos), len(suppliers), len(batches))
    for i in range(n):
        def pick(lst):
            v = lst[i] if i < len(lst) else None
            if v is None:
                return ""
            return str(v)
        rows.append([pick(names), pick(part_nos), pick(suppliers), pick(batches)])
    return rows


def fill_slide3(prs, data):
    """填充第三页底部物料信息表（表格5）
    行结构：[0]标题 [1]表头 [2..]数据行
    字体：中文微软雅黑 / 西文Times New Roman；尺寸沿用模板原表（可由常量覆盖）
    """
    slide3 = prs.slides[2]
    table_shapes = [s for s in slide3.shapes if s.has_table]
    if not table_shapes:
        return

    # 按表内容识别物料表（首行首格含“物料信息”），避免依赖表在页面中的顺序
    shape = None
    for s in table_shapes:
        t = s.table
        if len(t.rows) and len(t.columns) and \
                t.cell(0, 0).text.strip().startswith("物料信息"):
            shape = s
            break
    if shape is None:
        print("[WARN] 验证信息页未找到物料信息表")
        return
    table = shape.table
    rows = build_slide3_material_rows(data)

    headers = ["物料名称", "料号", "供应商", "批次"]
    for ci, h in enumerate(headers):
        if ci < len(table.columns):
            set_cell_text(table.cell(1, ci), h)

    # 数据行从 row2 开始，不够则加行
    need = len(rows)
    have = len(table.rows) - 2
    while have < need:
        _add_table_row(table)
        have += 1
    # 模板预置数据行多于实际数据时，删掉多余空行
    if have > need:
        _remove_table_rows(table, need + 2)

    for ri, row_data in enumerate(rows):
        row_idx = ri + 2
        if row_idx >= len(table.rows):
            break
        for ci, val in enumerate(row_data):
            if ci < len(table.columns):
                set_cell_text(table.cell(row_idx, ci), val)

    _fix_table_style(shape, FONT_LATIN, FONT_EA,
                     SLIDE3_TABLE_WIDTH_CM, SLIDE3_TABLE_HEIGHT_CM)


def fill_slide2(prs, data):
    """填充第二页汇总表"""
    slide2 = prs.slides[1]
    rows = build_slide2_table_data(data)

    # 1) 第一个 xxxx 占位符 -> 与首页一致的项目名 + 需要进行GT_DOE验证可行性
    project_name = data.get("物料信息", {}).get("project_name")
    for shape in slide2.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if "XXXXX" in run.text or "xxxxx" in run.text:
                    run.text = run.text.replace("XXXXX", "").replace("xxxxx", "")
                    if not run.text.strip():
                        run.text = f"{project_name}需要进行GT_DOE验证可行性"

    # 2-4) 表格：项目/精度SPEC/精度结果/Die shear SPEC/Die shear结果/结论
    table_shapes = [s for s in slide2.shapes if s.has_table]
    if table_shapes:
        table = table_shapes[0].table
        headers = ["NO.", "项目", "精度SPEC/mm", "精度结果/mm",
                   "Die shear SPEC/g", "Die shear结果/g","空洞率SPEC/%","空洞率结果/%", "结论", "备注"]
        while len(table.columns) < len(headers):
            _add_table_col(table)
        for ci, h in enumerate(headers):
            if ci < len(table.columns):
                set_cell_text(table.cell(0, ci), h)

        while len(table.rows) < len(rows) + 1:
            _add_table_row(table)
        # 实际数据行少于模板预置行时，删掉多余空行（否则末尾残留一行只有 NO. 的空白行）
        if len(table.rows) > len(rows) + 1:
            _remove_table_rows(table, len(rows) + 1)

        for ri, row_data in enumerate(rows):
            row_idx = ri + 1
            if row_idx >= len(table.rows):
                break
            set_cell_text(table.cell(row_idx, 0), str(ri + 1))
            for ci, val in enumerate(row_data):
                if ci + 1 < len(table.columns):
                    set_cell_text(table.cell(row_idx, ci + 1), val)

        _fix_slide2_table_style(table_shapes[0], SLIDE2_TABLE_FONT,
                                SLIDE2_TABLE_WIDTH_CM, SLIDE2_TABLE_HEIGHT_CM)

        # 结论列配色：PASS 绿色加粗 / FAIL 红色加粗
        concl_ci = headers.index("结论")
        for ri in range(1, len(table.rows)):
            cell = table.cell(ri, concl_ci)
            concl = cell.text.strip().upper()
            if concl not in ("PASS", "FAIL"):
                continue
            color = RGBColor(0x00, 0x80, 0x00) if concl == "PASS" else RGBColor(0xC0, 0x00, 0x00)
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.bold = True
                    run.font.color.rgb = color

        # 全部 PASS 时在表格下方加一行加粗总结论
        shp = table_shapes[0]
        # 增删行后同步图形框高度为实际行高之和，避免与表格错位
        total_h = sum(r.height for r in table.rows)
        if total_h:
            shp.height = total_h
        concl_cells = [table.cell(ri, concl_ci).text.strip().upper()
                       for ri in range(1, len(table.rows))]
        if concl_cells and all(c == "PASS" for c in concl_cells):
            tb = slide2.shapes.add_textbox(
                shp.left, shp.top + shp.height + Cm(0.2),
                shp.width, Cm(1.0))
            tf = tb.text_frame
            tf.word_wrap = True
            run = tf.paragraphs[0].add_run()
            run.text = "结论：所有项目均满足SPEC，DOE PASS"
            _set_run_font(run, FONT_LATIN, FONT_EA, 16)
            run.font.bold = True


def build_position_summary(data):
    """构建位置精度结论"""
    position_data = data.get("位置精度", {})
    conclusions = []
    for group_name, group_data in position_data.items():
        if group_name == "统计":
            continue
        for dev_name, dev_data in group_data.items():
            if dev_name == "统计":
                continue
            stats = dev_data.get("统计", {})
            if stats and "mean" in stats:
                mean_dict = stats["mean"]
                parts = [f"{k}={v:.4f}" for k, v in mean_dict.items() if isinstance(v, (int, float))]
                mean_str = ", ".join(parts)
                conclusions.append(f"{group_name}-{dev_name}: {mean_str}")
    return "; ".join(conclusions) if conclusions else "待分析"


def build_reliability_summary(data):
    """构建剪切力结论"""
    rel_keys = [k for k in ["可靠性-T0", "可靠性-T168", "可靠性-T500"] if k in data]
    return f"共{len(rel_keys)}个温度条件数据待分析"


def insert_image(slide, image_path, left=0.5, top=0.5, width=8.0):
    """向slide插入图片"""
    if os.path.exists(image_path):
        slide.shapes.add_picture(image_path, Inches(left), Inches(top),
                                  width=Inches(width))
        print(f"  [IMG] {os.path.basename(image_path)}")
    else:
        print(f"  [IMG] NOT FOUND: {os.path.basename(image_path)}")


def find_text_shapes(slide, keyword):
    """查找包含关键词的文本框"""
    results = []
    for shape in slide.shapes:
        if shape.has_text_frame and keyword in shape.text_frame.text:
            results.append(shape)
    return results


def fill_pptx():
    """主函数：填充PPTX模板。

    输入/输出路径在 parameter.py 顶部配置。
    """
    data_file = P.DATA_FILE
    template_path = P.TEMPLATE_PATH
    cp_base_dir = P.CP_BASE_DIR

    missing = [p for p in (data_file, template_path) if not os.path.isfile(p)]
    if missing:
        raise FileNotFoundError(
            "以下路径不存在，请检查 parameter.py 中的配置：\n  " + "\n  ".join(missing))

    P.ensure_image_dirs()

    print("1/4 加载数据...")
    print(f"  数据: {data_file}")
    print(f"  模板: {template_path}")
    data = load_data(data_file)
    specs = P.build_specs(data.get("物料信息", {}).get("spec", {}))

    # 输出报告名：{project_name}_GT工艺验证报告.pptx（project_name 取自 Excel）
    import re as _re
    project_name = str(data.get("物料信息", {}).get("project_name") or "").strip()
    safe_name = _re.sub(r'[\\/:*?"<>|]', "_", project_name) or "DOE"
    output_path = P.OUTPUT_PATH_FMT.format(project_name=safe_name)
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except Exception:
            pass

    print("=" * 50)
    print("2/4 生成图表...")
    print(f"  图片: {P.IMAGES_DIR}")
    generate_charts(data, specs)

    print("=" * 50)
    print("3/4 打开模板并填充...")
    prs = Presentation(template_path)

    n_slides = len(prs.slides)
    if n_slides < EXPECTED_TEMPLATE_SLIDES:
        raise ValueError(
            f"模板页数不足：需要至少 {EXPECTED_TEMPLATE_SLIDES} 页，实际 {n_slides} 页。"
            f"各 slide 模块依赖固定页序（位置精度=第5页、可靠性=第6页、"
            f"切片=第7页），请检查 {os.path.basename(template_path)}"
        )

    doe_params = build_doe_params(data)
    pos_summary = build_position_summary(data)
    rel_summary = build_reliability_summary(data)

    # --- Slide 1: 封面 ---
    slide1 = prs.slides[0]
    project_name = data.get("物料信息", {}).get("project_name") 
    today = datetime.now().strftime("%Y-%m-%d")
    for shape in slide1.shapes:
        if shape.has_text_frame:
            tf = shape.text_frame
            full_text = tf.text
            if not full_text.strip():
                continue
            # 替换标题行中的 XXX 芯片 -> 项目名
            has_gt = "GT" in full_text
            if has_gt:
                for para in tf.paragraphs:
                    for run in para.runs:
                        if "XXX" in run.text:
                            run.text = run.text.replace("XXX", project_name)
                        if "芯片" in run.text:
                            run.text = run.text.replace("芯片", "")
            # 替换姓名/日期行中的 xxx/XXX 占位符
            if "姓名" in full_text and "日期" in full_text:
                # 同一文本框中有姓名和日期，按顺序替换
                name_replaced = False
                date_replaced = False
                for para in tf.paragraphs:
                    for run in para.runs:
                        if "XXX" in run.text or "xxx" in run.text:
                            if not name_replaced and "姓名" in full_text:
                                run.text = run.text.replace("XXX", "张凌峰").replace("xxx", "张凌峰")
                                name_replaced = True
                            elif not date_replaced and "日期" in full_text:
                                run.text = run.text.replace("XXX", today).replace("xxx", today)
                                date_replaced = True
                            else:
                                run.text = run.text.replace("XXX", project_name).replace("xxx", project_name)
            else:
                for para in tf.paragraphs:
                    for run in para.runs:
                        if "XXX" in run.text or "xxx" in run.text:
                            run.text = run.text.replace("XXX", project_name).replace("xxx", project_name)

    # --- Slide 2: 汇总表 ---
    fill_slide2(prs, data)

    # --- Slide 3: 验证信息 ---
    fill_slide3(prs, data)

    # --- Slide 4: DOE参数 ---
    from slide4_doe import fill_slide4
    fill_slide4(prs, data)

    # --- Slide 5+: 位置精度（每个profile一页） ---
    from slide_position import fill_position_pages
    n_pos = fill_position_pages(prs, data)
    off = max(n_pos - 1, 0)  # 相对原模板的索引偏移（插入了N-1页）

    # --- Slide 6+: 可靠性（每个器件一页） ---
    from slide_reliability import fill_reliability_pages
    n_rel = fill_reliability_pages(prs, data, template_index=5 + off)
    off = off + max(n_rel - 1, 0)

    # --- Slide 7+: CP切片（每个xxxCP文件夹一页；需在模板第七页被填充前复制） ---
    from slide_CP import fill_cp_pages
    n_cp = fill_cp_pages(prs, data, template_index=6 + off, base_dir=cp_base_dir)

    # 模板原切片页作废，删除（CP页已复制到其后）
    if n_cp:
        from pptx.oxml.ns import qn as _qn
        sldIdLst = prs.slides._sldIdLst
        _el = list(sldIdLst)[6 + off]
        prs.part.drop_rel(_el.get(_qn('r:id')))
        sldIdLst.remove(_el)

    off = off + n_cp - 1  # 新增n_cp页，删除原切片页1页

    # # --- Slide 8: 附件 ---
    # slide8 = prs.slides[7 + off]

    # --- 保存 ---
    print("=" * 50)
    print(f"4/4 保存报告 -> {output_path}")
    fallback = os.path.join(
        tempfile.gettempdir(), os.path.basename(output_path))
    candidates = [output_path, fallback]
    saved = None
    for path in candidates:
        try:
            prs.save(path)
            saved = path
            break
        except PermissionError:
            print(f"[WARN] {path} 被占用，尝试下一个路径")
    if saved is None:
        raise PermissionError(
            "所有输出路径均被占用：请关闭 PowerPoint 中打开的报告后重试。"
            f"已尝试 {', '.join(candidates)}")
    print("=" * 50)
    print(f"报告生成完成 -> {saved}")
    # 自动打开报告
    try:
        os.startfile(saved)
    except Exception as e:
        print(f"[WARN] 自动打开失败: {e}")
    return saved


if __name__ == "__main__":
    fill_pptx()
