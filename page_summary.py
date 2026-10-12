# -*- coding: utf-8 -*-
"""第二页：汇总表。

三项判定（精度 / Die shear / 空洞率）全部满足 SPEC 才 PASS。
原先 build_slide2_table_data / fill_slide2 / _fix_slide2_table_style
都塞在 report_generator.py 里，与 slide4 之后的模块化风格不一致，现独立成页。
"""
import numpy as np
from pptx.dml.color import RGBColor
from pptx.util import Cm

import parameter as P
import pptx_utils as U

# 表格尺寸强制值：填数字(cm)则覆盖模板原尺寸；None 则沿用模板原尺寸
TABLE_WIDTH_CM = None
TABLE_HEIGHT_CM = None
TABLE_FONT = "Times New Roman"

REL_KEYS = ["可靠性-T0", "可靠性-T168", "可靠性-T500"]

HEADERS = ["NO.", "项目", "精度SPEC/mm", "精度结果/mm",
           "Die shear SPEC/g", "Die shear结果/g",
           "空洞率SPEC/%", "空洞率结果/%", "结论", "备注"]

COLOR_PASS = RGBColor(0x00, 0x80, 0x00)
COLOR_FAIL = RGBColor(0xC0, 0x00, 0x00)

ALL_PASS_TEXT = "结论：所有项目均满足SPEC，DOE PASS"


# ============================================================
# 三项结果的计算
# ============================================================
def _precision_results(data, dev_names):
    """精度结果 = 各 profile 下"各样本减均值后标准差"的最大值（排除 angle）。"""
    position_data = data.get("位置精度", {})
    out = {}
    for dev_name in dev_names:
        profile_stds = []
        for group_name, group_data in position_data.items():
            if group_name == "统计" or dev_name not in group_data:
                continue
            dev_data = group_data[dev_name]
            for dim, raw in dev_data.items():
                if dim == "统计" or str(dim).lower() == "angle":
                    continue
                if not isinstance(raw, list):
                    continue
                vals = [float(v) for v in raw if v is not None]
                if not vals:
                    continue
                m = np.mean(vals)
                profile_stds.append(float(np.std([v - m for v in vals])))
        out[dev_name] = max(profile_stds) if profile_stds else 0.0
    return out


def _shear_results(data, dev_names):
    """Die shear 结果 = 所有可靠性条件 / 所有 profile / 所有样本原始值的最小值。

    （旧逻辑取"各组剪切力均值的最小值"，已弃用，见下方注释保留。）
    """
    out = {}
    for dev_name in dev_names:
        all_vals = []
        for key in REL_KEYS:
            if key not in data:
                continue
            for _block_name, block_content in data[key].items():
                for _prof_name, prof_content in block_content.items():
                    if dev_name not in prof_content:
                        continue
                    raw = prof_content[dev_name].get("剪切力", [])
                    if isinstance(raw, list):
                        all_vals.extend(
                            float(v) for v in raw
                            if isinstance(v, (int, float))
                            and not isinstance(v, bool))
                    elif isinstance(raw, (int, float)) and not isinstance(raw, bool):
                        all_vals.append(float(raw))
        out[dev_name] = min(all_vals) if all_vals else 0.0
    return out


def _void_results(data, dev_names):
    """空洞率结果 = 该器件所有 profile 实测值的最大值（最差）；无数据返回 None。"""
    void_data = data.get("空洞率", {}) or {}
    out = {}
    for dev_name in dev_names:
        vals = []
        for prof_vals in void_data.values():
            if not isinstance(prof_vals, dict):
                continue
            v = prof_vals.get(dev_name)
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                vals.append(float(v))
        out[dev_name] = max(vals) if vals else None
    return out


def _fmt_void(v):
    """空洞率格式化：单元格里已是百分数（0.6 = 0.6%），直接输出不换算。"""
    if v is None or v == "":
        return "N/A"
    try:
        return f"{float(v):g}"
    except (TypeError, ValueError):
        return str(v)


def build_rows(data):
    """构建汇总表数据行（不含 NO. 列）。"""
    spec = data.get("物料信息", {}).get("spec", {})
    dev_names = list(spec.keys())

    precision = _precision_results(data, dev_names)
    shear = _shear_results(data, dev_names)
    void = _void_results(data, dev_names)

    rows = []
    for dev_name, spec_val in spec.items():
        pos_spec = spec_val.get("位置", "N/A")
        shear_spec = spec_val.get("剪切力", "N/A")
        prec_result = precision.get(dev_name, 0.0)
        shear_result = shear.get(dev_name, 0.0)
        void_result = void.get(dev_name)

        # 判定
        pos_bound = P.parse_position_bound(pos_spec)
        pass_pos = pos_bound is not None and prec_result <= pos_bound
        pass_shear = (isinstance(shear_spec, (int, float))
                      and shear_result >= shear_spec)
        void_spec_raw = spec_val.get("空洞率")
        pass_void = None
        if (isinstance(void_spec_raw, (int, float))
                and not isinstance(void_spec_raw, bool)
                and void_result is not None):
            pass_void = void_result <= float(void_spec_raw)

        # 空洞率结果为空时不参与判定（只看精度 + Die shear）
        conclusion = ("FAIL" if (not pass_pos or not pass_shear
                                 or pass_void is False) else "PASS")

        shear_spec_s = (f"{shear_spec:.2f}"
                        if isinstance(shear_spec, (int, float)) else str(shear_spec))
        rows.append([
            dev_name,
            str(pos_spec),
            f"{prec_result:.4f}",
            shear_spec_s,
            f"{shear_result:.2f}",
            _fmt_void(spec_val.get("空洞率")),
            _fmt_void(void_result) if void_result is not None else "/",
            conclusion,
            "",
        ])
    return rows


# ============================================================
# 填充
# ============================================================
def _fill_project_placeholder(slide, project_name):
    """把页面上的 XXXXX 占位符替换为"<项目名>需要进行GT_DOE验证可行性"。"""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if "XXXXX" in run.text or "xxxxx" in run.text:
                    run.text = run.text.replace("XXXXX", "").replace("xxxxx", "")
                    if not run.text.strip():
                        run.text = f"{project_name}需要进行GT_DOE验证可行性"


def _fill_table(shape, rows):
    table = shape.table
    while len(table.columns) < len(HEADERS):
        U.add_table_col(table)
    for ci, h in enumerate(HEADERS):
        if ci < len(table.columns):
            U.set_cell_text(table.cell(0, ci), h)

    while len(table.rows) < len(rows) + 1:
        U.add_table_row(table)
    # 实际数据行少于模板预置行时删掉多余空行（否则末尾残留只有 NO. 的空白行）
    if len(table.rows) > len(rows) + 1:
        U.remove_table_rows(table, len(rows) + 1)

    for ri, row_data in enumerate(rows):
        row_idx = ri + 1
        if row_idx >= len(table.rows):
            break
        U.set_cell_text(table.cell(row_idx, 0), str(ri + 1))
        for ci, val in enumerate(row_data):
            if ci + 1 < len(table.columns):
                U.set_cell_text(table.cell(row_idx, ci + 1), val)

    U.apply_table_font(shape, TABLE_FONT, TABLE_FONT)
    U.fit_table_size(shape, TABLE_WIDTH_CM, TABLE_HEIGHT_CM)


def _color_conclusions(table):
    """结论列配色：PASS 绿色加粗 / FAIL 红色加粗。"""
    concl_ci = HEADERS.index("结论")
    for ri in range(1, len(table.rows)):
        cell = table.cell(ri, concl_ci)
        concl = cell.text.strip().upper()
        if concl not in ("PASS", "FAIL"):
            continue
        color = COLOR_PASS if concl == "PASS" else COLOR_FAIL
        for para in cell.text_frame.paragraphs:
            for run in para.runs:
                run.font.bold = True
                run.font.color.rgb = color


def _add_all_pass_summary(slide, shape):
    """全部 PASS 时在表格下方加一行加粗总结论。"""
    concl_ci = HEADERS.index("结论")
    concl_cells = [shape.table.cell(ri, concl_ci).text.strip().upper()
                   for ri in range(1, len(shape.table.rows))]
    if not concl_cells or not all(c == "PASS" for c in concl_cells):
        return None
    tb = slide.shapes.add_textbox(shape.left, shape.top + shape.height + Cm(0.2),
                                  shape.width, Cm(1.0))
    tf = tb.text_frame
    tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = ALL_PASS_TEXT
    U.set_run_font(run, P.FONT_LATIN_PPT, P.FONT_EA_PPT, 16)
    run.font.bold = True
    return tb


def fill_summary(prs, data):
    """填充第二页汇总表。返回写入的数据行数。"""
    slide = prs.slides[P.TEMPLATE_PAGE_INDEX["summary"]]
    rows = build_rows(data)

    project_name = data.get("物料信息", {}).get("project_name")
    _fill_project_placeholder(slide, project_name)

    shape = U.table_shape(slide)
    if shape is None:
        print("[WARN] 汇总页未找到表格")
        return 0

    _fill_table(shape, rows)
    # 增删行后同步图形框高度为实际行高之和，避免与表格错位
    U.sync_shape_height(shape, shape.table)
    _color_conclusions(shape.table)
    _add_all_pass_summary(slide, shape)
    return len(rows)
