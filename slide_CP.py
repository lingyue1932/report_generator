# -*- coding: utf-8 -*-
"""CP切片页：每个 xxxCP 文件夹生成一页
页面参考模板第七页（切片图 + 孔洞率表格）
"""
import copy
import os
import re

from PIL import Image
from pptx.oxml.ns import qn
from pptx.util import Cm

import parameter as P
from slide4_doe import _add_row, _remove_extra_rows
from slide_position import (_duplicate_slide, _reorder_slide, _set_run_font,
                            FONT_EA, FONT_LATIN)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_SLIDE_INDEX = 6  # 模板第七页（0-based）

FONT_SIZE = 14

# 切片图高度 / 结论文本框坐标（cm）：统一在 parameter.py 配置

# 表格列
COL_COND, COL_IMG, COL_HOLE, COL_SPEC, COL_CONCL = 0, 1, 2, 3, 4

IMG_EXTS = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp")
TITLE_PREFIX = "切片"


# ============================================================
# 器件 / 图片发现
# ============================================================
def get_devices(base_dir=BASE_DIR):
    """扫描 xxxCP 文件夹，返回 [(器件名, 文件夹路径)]"""
    out = []
    try:
        names = sorted(os.listdir(base_dir))
    except OSError:
        return out
    for name in names:
        path = os.path.join(base_dir, name)
        if not os.path.isdir(path):
            continue
        if name.startswith("__") or name.lower() == "thumbs.db":
            continue
        if name.upper().endswith("CP") and len(name) > 2:
            out.append((name[:-2], path))
    return out


def image_root(dev_folder):
    """优先取 xxxCP/CP 子目录，否则用 xxxCP 本身"""
    sub = os.path.join(dev_folder, "CP")
    return sub if os.path.isdir(sub) else dev_folder


def profile_no(name):
    """1-1 / 1 -> 1；2-x -> 2 ... 非数字开头返回 None"""
    m = re.match(r"\s*(\d+)", str(name))
    return int(m.group(1)) if m else None


def _list_entries(root):
    dirs, files = [], []
    try:
        names = sorted(os.listdir(root))
    except OSError:
        return dirs, files
    for name in names:
        if name.lower() == "thumbs.db":
            continue
        path = os.path.join(root, name)
        if os.path.isdir(path):
            dirs.append((name, path))
        else:
            files.append((name, path))
    return dirs, files


def _find_a_image(path):
    """文件夹中命名为 A 的图（排除 A- / AM）"""
    if not os.path.isdir(path):
        return None
    cands = []
    for name in os.listdir(path):
        stem, ext = os.path.splitext(name)
        if stem == "A" and ext.lower() in IMG_EXTS:
            cands.append(os.path.join(path, name))
    return sorted(cands)[0] if cands else None


def collect_profiles(root):
    """{profile号: (容器路径, A图路径)}，按 profile 号升序"""
    dirs, files = _list_entries(root)
    result = {}
    if dirs:
        for name, path in dirs:
            no = profile_no(name)
            if no is None or no in result:
                continue
            result[no] = (path, _find_a_image(path))
    else:
        # 平铺结构：文件名 1-x / 1 视为 profile1，A 图取根目录 A.*
        a_img = _find_a_image(root)
        nos = set()
        for name, _ in files:
            no = profile_no(name)
            if no is not None:
                nos.add(no)
        for no in nos:
            result[no] = (root, a_img)
    return {k: result[k] for k in sorted(result)}


# ============================================================
# SPEC / 实测值
# ============================================================
def hole_spec(data, device_name):
    """空洞率 SPEC（来自物料信息 SPEC 表的“空洞率”列）"""
    v = data.get("物料信息", {}).get("spec", {}).get(device_name, {}).get("空洞率")
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def hole_measured(data, device_name, profile_no_):
    """孔洞率实测值，暂无固定数据源时返回 None"""
    for key in ("空洞率", "孔洞率"):
        blk = data.get(key)
        if isinstance(blk, dict):
            sub = blk.get(profile_no_) or blk.get(f"Profile{profile_no_}")
            if isinstance(sub, dict) and device_name in sub:
                return _num(sub[device_name])
            if device_name in blk and not isinstance(blk[device_name], dict):
                return _num(blk[device_name])
    spec = data.get("物料信息", {}).get("spec", {}).get(device_name, {})
    for k in ("空洞率实测", "孔洞率实测"):
        if k in spec:
            return _num(spec[k])
    return None


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _pct(v):
    """0.1 -> 10%；10 -> 10%"""
    if v is None:
        return "/"
    return f"{v * 100:g}%" if abs(v) <= 1 else f"{v:g}%"


# ============================================================
# 表格
# ============================================================
def _set_cell_text(cell, text, ref_rPr=None):
    """写单元格文字；无 run 时复制参照单元格的 rPr 保持格式"""
    tf = cell.text_frame
    p = tf.paragraphs[0]
    if p.runs:
        p.runs[0].text = str(text)
        for r in p.runs[1:]:
            r.text = ""
    else:
        run = p.add_run()
        run.text = str(text)
        if ref_rPr is not None:
            old = run._r.find(qn('a:rPr'))
            new = copy.deepcopy(ref_rPr)
            if old is not None:
                run._r.replace(old, new)
            else:
                run._r.insert(0, new)
    for extra in tf.paragraphs[1:]:
        for r in extra.runs:
            r.text = ""


def _table_shape(slide):
    for shp in slide.shapes:
        if shp.has_table:
            return shp
    return None


def _ref_rPr(table):
    """取一个有文字单元格的 rPr 作为格式参照"""
    for row in table.rows:
        for cell in row.cells:
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    rPr = run._r.find(qn('a:rPr'))
                    if rPr is not None:
                        return copy.deepcopy(rPr)
    return None


def _fill_table(shape, data, device_name, profiles):
    """填充表格：一行一个 profile，返回 [(profile号, A图路径, 是否满足SPEC)]"""
    table = shape.table
    ref = _ref_rPr(table)

    need = 1 + len(profiles)
    while len(table.rows) < need:
        _add_row(table)
    if len(table.rows) > need:
        _remove_extra_rows(table, need)

    spec = hole_spec(data, device_name)
    results = []
    for ri, (no, (_, a_img)) in enumerate(profiles.items()):
        row = ri + 1
        meas = hole_measured(data, device_name, no)
        if spec is None or meas is None:
            row_concl = "待定"
            ok = None
        else:
            ok = meas <= spec
            row_concl = "符合" if ok else "不符合"
        vals = [f"Profile{no}", "", _pct(meas), _pct(spec), row_concl]
        for ci, v in enumerate(vals):
            if ci < len(table.columns):
                _set_cell_text(table.cell(row, ci), v, ref)
        results.append((no, a_img, ok))

    # 防御性清空富余行
    for row in range(need, len(table.rows)):
        for ci in range(len(table.columns)):
            _set_cell_text(table.cell(row, ci), "", ref)

    # 行数变化后同步表格外框高度
    shape.height = sum(r.height for r in table.rows)
    return results


def _insert_cell_picture(slide, shape, row, col, img_path,
                          height_cm=P.CP_IMG_HEIGHT_CM):
    """按表格单元格位置插入图片：高度固定，宽度等比例，单元格内居中"""
    if not img_path or not os.path.exists(img_path):
        print(f"  [CP] 切片图不存在: {img_path}")
        return False
    try:
        with Image.open(img_path) as im:
            w, h = im.size
    except Exception as e:
        print(f"  [CP] 图片无法读取 {img_path}: {e}")
        return False
    if not w or not h:
        return False

    img_h = Cm(height_cm)
    img_w = int(img_h * w / h)

    table = shape.table
    x = shape.left + sum(c.width for i, c in enumerate(table.columns) if i < col)
    y = shape.top + sum(r.height for i, r in enumerate(table.rows) if i < row)
    cw = table.columns[col].width
    rh = table.rows[row].height

    left = x + max(0, (cw - img_w) // 2)
    top = y + max(0, (rh - img_h) // 2)
    slide.shapes.add_picture(img_path, left, top, width=img_w, height=img_h)
    print(f"  [CP] 图片 -> {os.path.basename(img_path)} "
          f"({img_w / 360000:.2f}x{height_cm}cm)")
    return True


# ============================================================
# 单页填充
# ============================================================
def _set_title(slide, device_name):
    for shp in slide.shapes:
        if not shp.has_text_frame:
            continue
        if not shp.text_frame.text.strip().startswith(TITLE_PREFIX):
            continue
        tf = shp.text_frame
        title = f"{TITLE_PREFIX} - {device_name}"
        p = tf.paragraphs[0]
        if p.runs:
            p.runs[0].text = title
            for r in p.runs[1:]:
                r.text = ""
        else:
            run = p.add_run()
            run.text = title
        for extra in tf.paragraphs[1:]:
            for r in extra.runs:
                r.text = ""
        return True
    print(f"  [CP] 未找到标题文本框: {device_name}")
    return False


def _page_conclusion(device_name, results, spec):
    if spec is None:
        return f"结论：{device_name}空洞率SPEC未定义，无法判定"
    oks = [ok for _, _, ok in results]
    if all(ok is None for ok in oks):
        return f"结论：{device_name}空洞率暂无实测数据，无法判定"
    if any(ok is False for ok in oks):
        return f"结论：{device_name}空洞率存在不符合SPEC的数据"
    if all(ok for ok in oks):
        return f"结论：{device_name}空洞率符合SPEC"
    return f"结论：{device_name}空洞率暂无实测数据，无法判定"


def _set_conclusion(slide, text, shape_bottom_emu=None):
    target = None
    for shp in slide.shapes:
        if shp.has_text_frame and shp.text_frame.text.strip().startswith("结论"):
            target = shp
            break
    if target is None:
        top = Cm(P.CP_CONCL_DEFAULT_TOP_CM) if shape_bottom_emu is None else \
            shape_bottom_emu + Cm(P.CP_CONCL_GAP_CM)
        target = slide.shapes.add_textbox(Cm(P.CP_CONCL_LEFT_CM), top,
                                          Cm(P.CP_CONCL_W_CM), Cm(P.CP_CONCL_H_CM))
    else:
        target.left = Cm(P.CP_CONCL_LEFT_CM)
        target.width = Cm(P.CP_CONCL_W_CM)
        target.height = Cm(P.CP_CONCL_H_CM)
        if shape_bottom_emu is not None:
            target.top = shape_bottom_emu + Cm(P.CP_CONCL_GAP_CM)
    tf = target.text_frame
    tf.word_wrap = True
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    _set_run_font(run, FONT_LATIN, FONT_EA, FONT_SIZE)
    return text


def display_path(path, base_dir):
    """日志展示用路径：base_dir 内显示相对路径，跨盘符或外部目录显示绝对路径"""
    try:
        return os.path.relpath(path, base_dir)
    except ValueError:
        return path


def _fill_device_page(prs, slide, device_name, dev_folder, data, base_dir):
    _set_title(slide, device_name)
    root = image_root(dev_folder)
    profiles = collect_profiles(root)
    print(f"  [CP] {device_name}: profile {list(profiles.keys())} "
          f"(图片目录 {display_path(root, base_dir)})")

    shape = _table_shape(slide)
    if shape is None:
        print(f"  [CP] {device_name} 页面无表格")
        results = []
    else:
        results = _fill_table(shape, data, device_name, profiles)
        for ri, (_, a_img, _) in enumerate(results):
            _insert_cell_picture(slide, shape, ri + 1, COL_IMG, a_img)

    spec = hole_spec(data, device_name)
    text = _page_conclusion(device_name, results, spec)
    bottom = (shape.top + shape.height) if shape is not None else None
    print(f"  [CP] {device_name}: {_set_conclusion(slide, text, bottom)}")


def fill_cp_pages(prs, data, template_index=TEMPLATE_SLIDE_INDEX,
                  base_dir=None):
    """生成CP切片页：每个 xxxCP 文件夹一页
    模板第七页保持不动（原切片页），新页全部插在其后
    切片图片目录由 parameter.CP_BASE_DIR 指定
    返回生成的页数
    """
    base_dir = base_dir or P.CP_BASE_DIR
    devices = get_devices(base_dir)
    if not devices:
        print("[CP] 未发现 xxxCP 文件夹")
        return 0

    sldIdLst = prs.slides._sldIdLst
    for i in range(len(devices)):
        _duplicate_slide(prs, template_index)
        _reorder_slide(prs, len(sldIdLst) - 1, template_index + 1 + i)

    for i, (dev_name, folder) in enumerate(devices):
        slide = prs.slides[template_index + 1 + i]
        _fill_device_page(prs, slide, dev_name, folder, data, base_dir)

    print(f"[CP] 共生成 {len(devices)} 页CP切片: "
          f"{[d for d, _ in devices]}")
    return len(devices)
