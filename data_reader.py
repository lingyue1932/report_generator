# -*- coding: utf-8 -*-
"""
通用读取器：按表头自适应解析，不依赖固定列偏移
数据行判断：该 profile 第一个器件第一个维度列必须是数值
可靠性：每个 profile 起始列默认为 SN
"""
import os
import pandas as pd
from openpyxl import load_workbook

import parameter as P

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ============================================================
# 工具
# ============================================================
def clean(v):
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    if isinstance(v, str) and v.strip() == "":
        return None
    return v


def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not pd.isna(v)


def load_sheet_raw(sheet_name, wb=None, file=None):
    """读取单个 sheet 为二维列表。

    file: Excel 路径，默认 P.DATA_FILE。
    wb: 传入可复用同一个 workbook，避免重复全量加载。
    """
    own = wb is None
    if own:
        wb = load_workbook(file or P.DATA_FILE, data_only=True)
    try:
        ws = wb[sheet_name]
        return [[clean(c) for c in row] for row in ws.iter_rows(values_only=True)]
    finally:
        if own:
            wb.close()


def _norm_text(s):
    if s is None:
        return ""
    return str(s).replace("\n", "").replace(" ", "").lower()


# DOE 参数区工艺列（顺序与模板一致，用于按表头名定位列）
DOE_PARAM_KEYS = [
    "Pre-heating Temp",
    "Ramp rate",
    "Solder Temp",
    "Press Time",
    "Solder Time",
    "Cooling Rate",
    "Cooling Temp",
    "Solder Force",
]


def build_doe_header_map(header_row):
    """按表头文本定位 DOE 参数区各列，返回 {列名: 列索引}。

    表头通常带单位后缀（如 "Pre-heating Temp/℃"、"Soler Temp\n℃"），故用前缀匹配；
    同时容忍模板中的 "Soler" 误拼。未找到的列不出现在返回结果中。
    """
    texts = [_norm_text(c) for c in header_row]
    col_of = {}
    taken = set()

    for ci, t in enumerate(texts):
        if t == "no":
            col_of["NO."] = ci
            taken.add(ci)
            break

    def locate(want):
        for ci, t in enumerate(texts):
            if ci not in taken and t.startswith(want):
                taken.add(ci)
                return ci
        return None

    for key in DOE_PARAM_KEYS:
        want = _norm_text(key)
        ci = locate(want)
        if ci is None and key.startswith("Solder"):
            ci = locate(want.replace("solder", "soler"))
        if ci is not None:
            col_of[key] = ci

    return col_of


def row_texts(row):
    return [c.strip() if isinstance(c, str) else None for c in row]


def is_profile_name(s):
    # profile 名不一定以 "profile" 开头（如 "Vendor profile"），按包含匹配
    return isinstance(s, str) and "profile" in s.strip().lower()


# ============================================================
# 1) 物料信息
# ============================================================
def parse_wuliao(raw):
    result = {"doe_params": [], "doe_groups": [], "material_info": {}}

    # --- 验证名称（项目名）：A1="验证名称"，A2=名称 ---
    for ri in range(min(4, len(raw))):
        row = raw[ri]
        for ci, cell in enumerate(row):
            if isinstance(cell, str) and cell.strip() == "验证名称":
                for cj in range(ci + 1, len(row)):
                    if clean(row[cj]) is not None:
                        result["project_name"] = str(row[cj]).strip()
                        break
                if "project_name" not in result and ri + 1 < len(raw):
                    nxt = clean(raw[ri + 1][ci]) if ci < len(raw[ri + 1]) else None
                    if nxt is not None:
                        result["project_name"] = str(nxt).strip()
                break
        if "project_name" in result:
            break

    # --- DOE参数区：表头含 Sample size / NO ---
    header_idx = None
    for ri, row in enumerate(raw):
        texts = [str(c).replace("\n", "").strip() for c in row if isinstance(c, str)]
        if any("Samplesize" in t.replace(" ", "") or t == "NO" for t in texts):
            header_idx = ri
            break
        if any(isinstance(c, str) and "DOE参数" in c for c in row):
            # 下一行通常是表头
            if ri + 1 < len(raw):
                header_idx = ri + 1
            break

    if header_idx is not None:
        # 按表头名定位工艺参数列，不再依赖固定列偏移
        hcol = build_doe_header_map(raw[header_idx])
        no_col = hcol.get("NO.")

        def cell_at(row, ci):
            if ci is None or ci < 0 or ci >= len(row):
                return None
            return clean(row[ci])

        missing = [k for k in DOE_PARAM_KEYS if k not in hcol]
        if missing:
            print(f"[WARN] DOE参数区: 表头未找到列 {missing}")

        groups = []
        current = None
        for r in range(header_idx + 1, len(raw)):
            row = raw[r]
            if all(clean(c) is None for c in row):
                if current is not None:
                    break
                continue
            first = cell_at(row, 0)
            sample = cell_at(row, 1)
            no = cell_at(row, no_col if no_col is not None else 2)
            # 新profile组
            if isinstance(first, str) and first.strip():
                current = {
                    "profile": first.strip(),
                    "sample_size": sample,
                    "devices": [],
                }
                groups.append(current)
            if current is None:
                continue
            if sample is not None and current.get("sample_size") is None:
                current["sample_size"] = sample
            if no is None and first is None:
                break
            device = {"NO.": no if no is not None else ""}
            for key in DOE_PARAM_KEYS:
                device[key] = cell_at(row, hcol.get(key))
            if device["NO."] != "" or any(
                v is not None for k, v in device.items() if k != "NO."
            ):
                current["devices"].append(device)
        result["doe_groups"] = groups
        # 兼容旧扁平结构
        flat = []
        for g in groups:
            for d in g["devices"]:
                item = dict(d)
                item["Profile"] = g["profile"]
                item["Sample size"] = g.get("sample_size")
                flat.append(item)
        result["doe_params"] = flat

    # 器件行定位锚点：「料号」标签的 (行, 列)，器件行 = 其上一行、列+1 起
    liao_pos = None
    for r in range(6, len(raw)):
        row = raw[r]
        for ci, cell in enumerate(row):
            if not isinstance(cell, str):
                continue
            txt = cell.strip()
            if txt in ("工单号：", "发料记录（是否欠料）：", "备注"):
                val = None
                for cj in range(ci + 1, len(row)):
                    if clean(row[cj]) is not None:
                        val = clean(row[cj])
                        break
                result["material_info"][txt] = val
            # 物料表：料号/供应商/批次 —— 取整行所有非空值
            if txt in ("料号", "供应商信息", "批次信息"):
                vals = [clean(row[cj]) for cj in range(ci + 1, len(row))
                        if clean(row[cj]) is not None]
                result["material_info"][txt] = vals
                if txt == "料号" and liao_pos is None:
                    liao_pos = (r, ci)

    # 器件行：「料号」行的上一行，从「料号」列+1 起向右取非空单元格
    if liao_pos is not None:
        r0, c0 = liao_pos
        if r0 - 1 >= 0:
            dev_row = raw[r0 - 1]
            devs = [clean(dev_row[cj]) for cj in range(c0 + 1, len(dev_row))
                    if clean(dev_row[cj]) is not None]
            if devs:
                result["material_info"]["器件行"] = devs
    if "器件行" not in result["material_info"]:
        print("[WARN] 物料信息: 未找到器件行（料号标签缺失？）")

    # --- 解析 SPEC 表格区域（行式结构）---
    #   SPEC 行      : [None, 'SPEC', ...]
    #   表头行       : ['器件', '位置精度', '尺寸-长', '尺寸-宽', '面积', '剪切力(g)']
    #   数据行(每器件): ['激光器', '±0.02', 0.72, 0.2, 0.144, 178.56]
    spec_header_idx = None
    for ri, row in enumerate(raw):
        if any(isinstance(c, str) and c.strip() == "SPEC" for c in row):
            spec_header_idx = ri
            break

    if spec_header_idx is not None:
        header_idx = spec_header_idx + 1
        if header_idx < len(raw):
            header = row_texts(raw[header_idx])

            # 按表头名定位列
            pos_col = shear_col = hole_col = None
            for ci, h in enumerate(header):
                if h is None:
                    continue
                if "位置精度" in h:
                    pos_col = ci
                if "剪切力" in h:
                    shear_col = ci
                if "空洞率" in h or "孔洞率" in h:
                    hole_col = ci

            spec_devices = {}
            for r in range(header_idx + 1, len(raw)):
                row = raw[r]
                if all(clean(c) is None for c in row):
                    break
                dev_name = row[0] if len(row) > 0 else None
                if not isinstance(dev_name, str) or not dev_name.strip():
                    break
                dev_name = dev_name.strip()
                # 下一个区块标题（如 '激光器' 物料行已在上方处理，SPEC 区以空行/非器件行结束）
                if dev_name in ("备注",):
                    break

                entry = {"位置": None, "剪切力": None, "空洞率": None}

                if pos_col is not None and pos_col < len(row):
                    v = clean(row[pos_col])
                    if v is not None:
                        entry["位置"] = v

                if shear_col is not None and shear_col < len(row):
                    v = clean(row[shear_col])
                    if v is not None:
                        try:
                            entry["剪切力"] = float(v)
                        except (ValueError, TypeError):
                            entry["剪切力"] = v

                if hole_col is not None and hole_col < len(row):
                    v = clean(row[hole_col])
                    if v is not None:
                        try:
                            entry["空洞率"] = float(v)
                        except (ValueError, TypeError):
                            entry["空洞率"] = v

                spec_devices[dev_name] = entry

            if spec_devices:
                result["spec"] = spec_devices

    return result


# ============================================================
# 1.5) 空洞率（CP 实测结果 sheet）
# ============================================================
def parse_kongdong(raw):
    """解析“空洞率”sheet：
      器件行（合并单元格，器件名在段首列）
      profile 行（每个器件分段下的 profile 名）
      空洞率 值行

    返回 (values, profiles)：
      values  = {profile序号: {器件: float}}   # 序号按器件内列顺序 1..N，
                                              # 与 CP 图片文件夹编号 1-x/2-x 对应
      profiles = {器件: {profile序号: profile名}}
    """
    dev_row = prof_row = val_row = None
    for ri, row in enumerate(raw):
        t = row_texts(row)
        first = t[0] if t else None
        if first == "器件":
            dev_row = ri
        elif first == "profile" and dev_row is not None:
            prof_row = ri
        elif first in ("空洞率", "孔洞率") and prof_row is not None:
            val_row = ri
            break
    if dev_row is None or prof_row is None or val_row is None:
        return {}, {}

    def _cells(ri):
        return row_texts(raw[ri]) if 0 <= ri < len(raw) else []

    def _num_cells(ri):
        # 值行：row_texts 会把非字符串清成 None，这里保留数值
        row = raw[ri] if 0 <= ri < len(raw) else []
        return [c.strip() if isinstance(c, str) else c for c in row]

    dev_cells = _cells(dev_row)
    prof_cells = _cells(prof_row)
    val_cells = _num_cells(val_row)

    # 器件分段：器件名在段首列（合并单元格只在首列有值），延伸到下一个器件名
    # 第 0 列是行标签“器件”，不参与分段
    starts = [ci for ci, c in enumerate(dev_cells) if c and ci > 0]
    spans = []
    for i, sc in enumerate(starts):
        ec = starts[i + 1] if i + 1 < len(starts) else max(len(prof_cells),
                                                           len(dev_cells))
        spans.append((dev_cells[sc], sc, ec))

    values, profiles = {}, {}
    for dev, sc, ec in spans:
        no = 0
        for ci in range(sc, ec):
            pname = prof_cells[ci] if ci < len(prof_cells) else None
            if not pname:
                continue
            no += 1
            profiles.setdefault(dev, {})[no] = pname
            v = val_cells[ci] if ci < len(val_cells) else None
            try:
                v = float(v)
            except (TypeError, ValueError):
                v = None
            if v is not None:
                values.setdefault(no, {})[dev] = v
    return values, profiles


# ============================================================
# 2) 位置精度（表头自适应）
# ============================================================
def parse_weizhi_jingdu(raw):
    """
    - profile 行：出现 "profile..." 的行
    - 器件行  ：profile 行下一行，非空单元格 → 器件名
    - 维度行  ：器件行下一行，从器件起始列到下一个器件起始列
    - 数据行  ：维度行下一行开始，该 profile 第一个器件第一个维度列必须是数值
    结构：data["位置精度"][profile][器件][维度] = [values]
                          [profile][器件]["统计"][mean/max/min/3σ][维度] = 值
    """
    # --- 找 profile 行 ---
    profile_row_idx = None
    for ri, row in enumerate(raw):
        if any(is_profile_name(c) for c in row):
            profile_row_idx = ri
            break
    if profile_row_idx is None:
        return {}

    profile_row = row_texts(raw[profile_row_idx])
    profile_cols = {}
    for ci, cell in enumerate(profile_row):
        if is_profile_name(cell):
            profile_cols[cell.strip()] = ci

    # --- 器件行 ---
    dev_row_idx = profile_row_idx + 1
    if dev_row_idx >= len(raw):
        return {}
    dev_row = row_texts(raw[dev_row_idx])

    # --- 维度行 ---
    dim_row_idx = dev_row_idx + 1
    dim_row = row_texts(raw[dim_row_idx]) if dim_row_idx < len(raw) else []

    profile_starts = sorted(profile_cols.values()) + [len(dev_row)]

    # --- 器件 + 维度识别 ---
    dims_per_profile = {}   # {profile: {device: [(dim, col), ...]}}
    for pname, pstart in profile_cols.items():
        next_starts = [s for s in profile_starts if s > pstart]
        pend = next_starts[0] if next_starts else len(dev_row)

        devs = []
        for ci in range(pstart, pend):
            cell = dev_row[ci] if ci < len(dev_row) else None
            if cell and "profile" not in cell.strip().lower():
                devs.append((cell.strip(), ci))

        dims_per_profile[pname] = {}
        for di, (dev_name, dstart) in enumerate(devs):
            if di + 1 < len(devs):
                dend = devs[di + 1][1]
            else:
                next_starts2 = [s for s in profile_starts if s > dstart]
                dend = next_starts2[0] if next_starts2 else len(dim_row)
            dims = []
            for ci in range(dstart, dend):
                cell = dim_row[ci] if ci < len(dim_row) else None
                if cell:
                    dims.append((cell.strip(), ci))
            dims_per_profile[pname][dev_name] = dims

    # --- 初始化 ---
    result = {}
    for pname, devs in dims_per_profile.items():
        result[pname] = {}
        for dev_name, dims in devs.items():
            entry = {d[0]: [] for d in dims}
            entry["统计"] = {}
            result[pname][dev_name] = entry

    # --- 数据行 ---
    data_start = dim_row_idx + 1
    stat_labels = {"mean", "max", "min", "3σ", "3sigma"}
    stat_rows = {pname: {} for pname in dims_per_profile}

    for r in range(data_start, len(raw)):
        row = raw[r]
        if all(clean(c) is None for c in row):
            continue

        first = next((clean(c) for c in row if clean(c) is not None), None)
        label = first.strip().lower() if isinstance(first, str) else None

        if label in stat_labels:
            for pname in dims_per_profile:
                stat_rows[pname][label] = row
            continue

        # 数据行判断：每个 profile 用自己的第一个器件第一个维度列判断
        for pname, devs in dims_per_profile.items():
            if not devs:
                continue
            first_dev = next(iter(devs))
            first_dim_list = devs[first_dev]
            if not first_dim_list:
                continue
            _, first_col = first_dim_list[0]
            val = clean(row[first_col]) if first_col < len(row) else None
            if not is_number(val):
                # 该 profile 的这行不是数据行，跳过
                continue
            # 该 profile 有效，填数据
            for dev_name, dims in devs.items():
                for dim_name, col in dims:
                    v = clean(row[col]) if col < len(row) else None
                    if v is not None:
                        result[pname][dev_name][dim_name].append(v)

    # --- 统计值 ---
    for pname, stats in stat_rows.items():
        for label, srow in stats.items():
            for dev_name, dims in dims_per_profile[pname].items():
                for dim_name, col in dims:
                    v = clean(srow[col]) if col < len(srow) else None
                    result[pname][dev_name]["统计"].setdefault(label, {})[dim_name] = v

    return result


# ============================================================
# 3) 可靠性（表头自适应 + SN 保留）
# ============================================================
def parse_reliability(raw, block_titles, default_block_name=None):
    """
    结构：
      data["可靠性-T168"]["TCT168"]["profile1"]["SN"]            = [1,2,3,...]
      data["可靠性-T168"]["TCT168"]["profile1"]["芯片"]["剪切力"]
      data["可靠性-T168"]["TCT168"]["profile1"]["芯片"]["失效模式"]
      data["可靠性-T168"]["TCT168"]["profile1"]["热敏"]["剪切力"]
      data["可靠性-T168"]["TCT168"]["profile1"]["热敏"]["失效模式"]
      data["可靠性-T168"]["TCT168"]["profile1"]["芯片"]["统计"]["mean"]["剪切力"]
    """
    # --- 找区块 ---
    block_pos = {}
    for ri, row in enumerate(raw):
        for cell in row:
            if isinstance(cell, str):
                t = cell.strip()
                for bt in block_titles:
                    if t == bt:
                        block_pos[bt] = ri

    if not block_pos and default_block_name:
        block_pos[default_block_name] = 0

    order = sorted(block_pos.items(), key=lambda x: x[1])
    result = {}

    for bi, (bname, start) in enumerate(order):
        end = order[bi + 1][1] if bi + 1 < len(order) else len(raw)

        # --- 找 profile 行 ---
        profile_row_idx = None
        for r in range(start, min(start + 4, end)):
            if any(is_profile_name(c) for c in raw[r]):
                profile_row_idx = r
                break
        if profile_row_idx is None:
            continue

        profile_row = row_texts(raw[profile_row_idx])

        # --- 器件行 ---
        dev_row_idx = profile_row_idx + 1
        if dev_row_idx >= end:
            continue
        dev_row = row_texts(raw[dev_row_idx])

        # --- 分组：SN 列作为分组边界，只识别写了 profile 名的分组 ---
        # profile 标签故意留空的分组不识别、不解析；若不以 SN 列切分边界，
        # 这些无名分组会被并入上一个分组，其 SN 列会被误当成器件名。
        labels = {ci: c.strip() for ci, c in enumerate(profile_row)
                  if is_profile_name(c)}
        sn_start_cols = [ci for ci, c in enumerate(dev_row)
                         if isinstance(c, str) and c.strip().lower() == "sn"]
        starts = sorted(set(labels) | set(sn_start_cols))
        if not starts:
            continue
        profile_cols = {}
        for col in starts:
            if col in labels:
                profile_cols[labels[col]] = col
            # else: profile 名为空 -> 故意不识别，跳过该分组

        # --- 维度行 ---
        dim_row_idx = dev_row_idx + 1
        dim_row = row_texts(raw[dim_row_idx]) if dim_row_idx < end else []

        profile_starts = sorted(profile_cols.values()) + [len(dev_row)]

        # --- 器件 + 维度 ---
        dims_per_profile = {}
        sn_cols = {}   # 每个 profile 的 SN 列

        for pname, pstart in profile_cols.items():
            next_starts = [s for s in profile_starts if s > pstart]
            pend = next_starts[0] if next_starts else len(dev_row)

            # ★ SN 列 = profile 起始列本身
            sn_cols[pname] = pstart

            devs = []
            for ci in range(pstart + 1, pend):   # ★ 从 pstart+1 开始，跳过 SN 列
                cell = dev_row[ci] if ci < len(dev_row) else None
                if cell and "profile" not in cell.strip().lower() \
                        and cell.strip().lower() != "sn":
                    devs.append((cell.strip(), ci))

            dims_per_profile[pname] = {}
            for di, (dev_name, dstart) in enumerate(devs):
                if di + 1 < len(devs):
                    dend = devs[di + 1][1]
                else:
                    next_starts2 = [s for s in profile_starts if s > dstart]
                    dend = next_starts2[0] if next_starts2 else len(dim_row)
                dims = []
                for ci in range(dstart, dend):
                    cell = dim_row[ci] if ci < len(dim_row) else None
                    if cell:
                        dims.append((cell.strip(), ci))
                dims_per_profile[pname][dev_name] = dims

        # --- 初始化 ---
        block_result = {}
        for pname, devs in dims_per_profile.items():
            entry = {"SN": []}
            for dev_name, dims in devs.items():
                e = {d[0]: [] for d in dims}
                e["统计"] = {}
                entry[dev_name] = e
            block_result[pname] = entry

        stat_labels = {"mean", "max", "min"}
        stat_rows = {pname: {} for pname in dims_per_profile}
        data_start = dim_row_idx + 1

        for r in range(data_start, end):
            row = raw[r]
            if all(clean(c) is None for c in row):
                continue

            first = next((clean(c) for c in row if clean(c) is not None), None)
            label = first.strip().lower() if isinstance(first, str) else None

            if label in stat_labels:
                for pname in dims_per_profile:
                    stat_rows[pname][label] = row
                continue

            if isinstance(first, str) and "失效图示" in first:
                continue

            # 数据行判断：每个 profile 用 SN 列判断
            for pname, devs in dims_per_profile.items():
                sn_col = sn_cols[pname]
                sn = clean(row[sn_col]) if sn_col < len(row) else None
                if not is_number(sn):
                    continue
                block_result[pname]["SN"].append(sn)
                for dev_name, dims in devs.items():
                    for dim_name, col in dims:
                        v = clean(row[col]) if col < len(row) else None
                        block_result[pname][dev_name][dim_name].append(v)

        # --- 统计值 ---
        for pname, stats in stat_rows.items():
            for label, srow in stats.items():
                for dev_name, dims in dims_per_profile[pname].items():
                    for dim_name, col in dims:
                        v = clean(srow[col]) if col < len(srow) else None
                        block_result[pname][dev_name]["统计"]\
                            .setdefault(label, {})[dim_name] = v

        result[bname] = block_result

    return result


# ============================================================
# 主流程
# ============================================================
def main(data_file=None):
    """解析 DOE_data.xlsx 全部所需 sheet。

    data_file: Excel 路径，默认 P.DATA_FILE。
    只 load_workbook 一次，各 sheet 复用同一个 workbook。
    """
    wb = load_workbook(data_file or P.DATA_FILE, data_only=True)
    try:
        all_sheets = wb.sheetnames
        DATA = {}

        if "物料信息" in all_sheets:
            DATA["物料信息"] = parse_wuliao(load_sheet_raw("物料信息", wb))
        if "位置精度" in all_sheets:
            DATA["位置精度"] = parse_weizhi_jingdu(load_sheet_raw("位置精度", wb))
        if "可靠性-T0" in all_sheets:
            DATA["可靠性-T0"] = parse_reliability(
                load_sheet_raw("可靠性-T0", wb),
                ["T0"], default_block_name="T0")
        if "可靠性-T168" in all_sheets:
            DATA["可靠性-T168"] = parse_reliability(
                load_sheet_raw("可靠性-T168", wb),
                ["TCT168", "UDH168"])
        if "可靠性-T500" in all_sheets:
            DATA["可靠性-T500"] = parse_reliability(
                load_sheet_raw("可靠性-T500", wb),
                ["TCT500", "UDH500"])
        if "空洞率" in all_sheets:
            vals, profs = parse_kongdong(load_sheet_raw("空洞率", wb))
            DATA["空洞率"] = vals
            DATA["空洞率_profiles"] = profs

        return DATA
    finally:
        wb.close()


if __name__ == "__main__":
    data = main()
    print(f"数据文件: {P.DATA_FILE}")
    for sheet, content in data.items():
        print(f"【{sheet}】ok")