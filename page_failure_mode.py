# -*- coding: utf-8 -*-
"""失效模式页（模板第 7 页，单页不复制）。

表格：行 = profile，列 = 可靠性条件（T0 / 168TCT / 168UDH / 500TCT / 500UDH）。
单元格填该 profile 在该条件下的全部失效模式图（上大下小、居中、等比缩放）。
图片由 data_reader.extract_failure_images 从 Excel 内嵌图提取落盘后写入 data。
"""
import os
import re

from PIL import Image

import parameter as P
import pptx_utils as U

TITLE_PREFIX = "失效模式"
COL_COND = 0
MARGIN = 45720   # 图片距单元格边 (EMU, 0.05")
GAP = 45720      # 图片间距 (EMU)


def _cond_key(s):
    """表头/映射名归一化：'168 TCT' / '168tct' -> '168TCT'"""
    return re.sub(r"\s+", "", str(s)).upper()


def _cond_lookup():
    """归一化条件名 -> (sheet_key, block_name)，与 P.RELIABILITY_MAP 对齐。"""
    return {_cond_key(name): (sheet, block)
            for name, (sheet, block) in P.RELIABILITY_MAP.items()}


def _ordered_profiles(data, conds):
    """按条件列顺序收集 profile 名（去重、保序）。"""
    out = []
    for sheet, block in conds:
        if not sheet:
            continue
        blk = data.get(sheet, {}).get(block, {})
        if not isinstance(blk, dict):
            continue
        for pname, entry in blk.items():
            if isinstance(entry, dict) and pname not in out:
                out.append(pname)
    return out


def _collect_images(data, sheet, block, profile):
    if not sheet:
        return []
    entry = data.get(sheet, {}).get(block, {}).get(profile)
    if not isinstance(entry, dict):
        return []
    return [p for p in (entry.get("失效模式图") or []) if p]


def _insert_cell_images(slide, shape, row, col, img_paths):
    """单元格内填多张图。

    首行 1 张（Excel 里是上排大图），次行其余（小图）；
    每行等比缩放填满可用宽，整体按可用高再收缩，水平/垂直居中。
    返回实际插入的图片数。
    """
    table = shape.table
    items = []
    for p in img_paths:
        if not p or not os.path.exists(p):
            print(f"  [FM] 失效模式图不存在: {p}")
            continue
        try:
            with Image.open(p) as im:
                w, h = im.size
        except Exception as e:
            print(f"  [FM] 图片无法读取 {p}: {e}")
            continue
        if w and h:
            items.append((p, w, h))
    if not items:
        return 0

    x, y, cw, rh = U.cell_rect(shape, table, row, col)
    avail_w = cw - 2 * MARGIN
    avail_h = rh - 2 * MARGIN

    rows = [items[:1], items[1:]] if len(items) > 1 else [items]
    rows = [g for g in rows if g]

    # 每行按可用宽求缩放
    scales = []
    for grp in rows:
        sum_w = sum(w for _, w, _ in grp)
        scales.append((avail_w - GAP * (len(grp) - 1)) / sum_w if sum_w > 0 else 0.0)

    # 总高超可用高时整体等比收缩（保持行内宽度比例）
    natural = sum(max(h for _, _, h in grp) * scales[i]
                  for i, grp in enumerate(rows))
    gap_sum = GAP * (len(rows) - 1)
    factor = 1.0
    if natural > 0 and natural + gap_sum > avail_h:
        factor = max((avail_h - gap_sum) / natural, 0.05)

    total_h = int(natural * factor + gap_sum)
    cy = y + MARGIN + max(0, (avail_h - total_h) // 2)
    n = 0
    for grp, s in zip(rows, scales):
        sf = s * factor
        row_w = int(sum(w for _, w, _ in grp) * sf + GAP * (len(grp) - 1))
        row_h = int(max(h for _, _, h in grp) * sf)
        cx = x + MARGIN + max(0, (avail_w - row_w) // 2)
        for p, w, h in grp:
            iw, ih = int(w * sf), int(h * sf)
            if iw < 1 or ih < 1:
                continue
            slide.shapes.add_picture(p, cx, cy, width=iw, height=ih)
            print(f"  [FM] 图片 -> {os.path.basename(p)} "
                  f"({iw / 914400:.2f}x{ih / 914400:.2f}in)")
            cx += iw + GAP
            n += 1
        cy += row_h + GAP
    return n


def fill_failure_mode_page(prs, data, template_index=None):
    """填充失效模式页（单页，不复制）。返回填入图片的单元格数。"""
    if template_index is None:
        template_index = P.TEMPLATE_PAGE_INDEX["failure_mode"]
    if template_index >= len(prs.slides):
        print(f"[FM] 模板无失效模式页（index {template_index}）")
        return 0

    slide = prs.slides[template_index]
    shape = U.table_shape(slide)
    if shape is None:
        print("[FM] 失效模式页无表格")
        return 0
    table = shape.table
    ref = U.get_ref_rPr(table)

    # 列 -> 条件
    lookup = _cond_lookup()
    conds = []
    for ci in range(1, len(table.columns)):
        key = _cond_key(table.cell(0, ci).text)
        cond = lookup.get(key)
        if cond is None:
            print(f"  [FM] 表头条件未识别: {table.cell(0, ci).text!r}")
        conds.append(cond)

    # 行 -> profile
    profiles = _ordered_profiles(data, conds)
    if not profiles:
        print("[FM] 无 profile 数据")
        return 0

    need = 1 + len(profiles)
    while len(table.rows) < need:
        U.add_table_row(table)
    if len(table.rows) > need:
        U.remove_table_rows(table, need)
    U.sync_shape_height(shape, table)

    filled = 0
    for ri, pname in enumerate(profiles):
        U.set_cell_text(table.cell(ri + 1, COL_COND), pname, ref)
        for ci, cond in enumerate(conds, start=1):
            cell = table.cell(ri + 1, ci)
            paths = _collect_images(data, *(cond or ("", "")), pname) if cond else []
            U.set_cell_text(cell, "", ref)
            n = _insert_cell_images(slide, shape, ri + 1, ci, paths) if paths else 0
            if n:
                filled += 1
            else:
                U.set_cell_text(cell, "/", ref)

    print(f"[FM] 失效模式页: {len(profiles)} profile x "
          f"{sum(1 for c in conds if c)} 条件，{filled} 个单元格填图")
    return filled
