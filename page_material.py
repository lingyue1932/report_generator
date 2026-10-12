# -*- coding: utf-8 -*-
"""第三页：验证信息（底部物料信息表）。

行结构：[0]标题行 [1]表头 [2..]数据行。
表格按"首行首格是否含物料信息"识别，不依赖它在页面中的先后顺序。
"""
import parameter as P
import pptx_utils as U

HEADERS = ["物料名称", "料号", "供应商", "批次"]

# 尺寸强制值：None 沿用模板原表尺寸
TABLE_WIDTH_CM = None
TABLE_HEIGHT_CM = None

TABLE_TITLE_KEY = "物料信息"


def build_material_rows(data):
    """构建物料表数据行：[名称, 料号, 供应商, 批次]。"""
    material_info = data.get("物料信息", {}).get("material_info", {})
    names = material_info.get("器件行") or []
    part_nos = material_info.get("料号") or []
    suppliers = material_info.get("供应商信息") or []
    batches = material_info.get("批次信息") or []

    def _as_list(v):
        return v if isinstance(v, list) else [v]

    names, part_nos = _as_list(names), _as_list(part_nos)
    suppliers, batches = _as_list(suppliers), _as_list(batches)

    rows = []
    n = max(len(names), len(part_nos), len(suppliers), len(batches))
    for i in range(n):
        def pick(lst):
            v = lst[i] if i < len(lst) else None
            return "" if v is None else str(v)
        rows.append([pick(names), pick(part_nos), pick(suppliers), pick(batches)])
    return rows


def _find_material_table(slide):
    """按首行首格内容识别物料表，避免依赖表在页面中的顺序。"""
    for shp in slide.shapes:
        if not shp.has_table:
            continue
        t = shp.table
        if len(t.rows) and len(t.columns) and \
                t.cell(0, 0).text.strip().startswith(TABLE_TITLE_KEY):
            return shp
    return None


def fill_material(prs, data):
    """填充第三页物料信息表。返回写入的数据行数。"""
    slide = prs.slides[P.TEMPLATE_PAGE_INDEX["material"]]
    if not any(s.has_table for s in slide.shapes):
        return 0

    shape = _find_material_table(slide)
    if shape is None:
        print("[WARN] 验证信息页未找到物料信息表")
        return 0

    table = shape.table
    rows = build_material_rows(data)

    for ci, h in enumerate(HEADERS):
        if ci < len(table.columns):
            U.set_cell_text(table.cell(1, ci), h)

    # 数据行从 row2 开始，不够则加行
    need = len(rows)
    have = len(table.rows) - 2
    while have < need:
        U.add_table_row(table)
        have += 1
    # 模板预置数据行多于实际数据时，删掉多余空行
    if have > need:
        U.remove_table_rows(table, need + 2)

    for ri, row_data in enumerate(rows):
        row_idx = ri + 2
        if row_idx >= len(table.rows):
            break
        for ci, val in enumerate(row_data):
            if ci < len(table.columns):
                U.set_cell_text(table.cell(row_idx, ci), val)

    U.apply_table_font(shape, P.FONT_LATIN_PPT, P.FONT_EA_PPT)
    U.fit_table_size(shape, TABLE_WIDTH_CM, TABLE_HEIGHT_CM)
    return len(rows)
