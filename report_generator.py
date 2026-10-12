# -*- coding: utf-8 -*-
"""DOE 工艺验证报告自动生成器 —— 主流程编排。

职责边界：本模块只负责"按什么顺序、把哪一页填成什么"。
  - PPT 操作基元      -> pptx_utils.py
  - 各页面内容逻辑    -> page_*.py（一页一模块）
  - 图表绘制          -> position.py / reliability.py
  - Excel 解析        -> data_reader.py
  - 全部可调参数      -> parameter.py

对外入口是 build_report()，既供 cli.py 调用，也供后续 GUI 调用。
"""
import os
import re
import sys
import tempfile

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from pptx import Presentation

import data_reader
import page_cp
import page_cover
import page_doe
import page_failure_mode
import page_material
import page_position
import page_reliability
import page_summary
import parameter as P
import pptx_utils as U

REL_SUMMARY_DIM_FALLBACK = "剪切力"


# ============================================================
# 工具
# ============================================================
def _first_dim(view, dev_name):
    """从可靠性视图里取该器件的第一个有效维度名（排除 SN / 统计）。"""
    for cond_data in view.values():
        for prof_data in cond_data.values():
            dev_data = prof_data.get(dev_name, {})
            for k in dev_data:
                if k not in ("SN", "统计"):
                    return k
    return None


def _resolve_output_path(data, explicit=None):
    """确定输出文件路径。

    explicit 与 OUTPUT_PATH_FMT 语法一致：都是"模板串"，含 {project_name}
    时用 Excel 里的项目名填充；不含占位符就当普通路径原样使用。

    显式路径也支持占位符，是为了让 GUI / 拖拽模式能把输出固定到某个目录下
    （例如 <项目目录>/{project_name}_GT工艺验证报告.pptx），
    而不必自己先解析一遍 Excel。
    """
    project_name = str(data.get("物料信息", {}).get("project_name") or "").strip()
    safe_name = re.sub(r'[\\/:*?"<>|]', "_", project_name) or "DOE"
    template = explicit or P.OUTPUT_PATH_FMT
    return template.format(project_name=safe_name) if "{" in template else template


def load_data(data_file=None):
    """加载并解析所有 sheet。"""
    return data_reader.main(data_file or P.DATA_FILE)


# ============================================================
# 图表生成
# ============================================================
def generate_charts(data, specs):
    """生成所有图表。data / specs 由调用方提供，避免重复解析 Excel。"""
    P.ensure_image_dirs()

    print("=" * 50)
    print("生成位置精度图...")
    import position

    groups_data, group_names, devices_per_group = (
        position.build_precision_inputs(data["位置精度"])
    )
    for group in group_names:
        position.plot_single_profile(
            profile_name=group,
            profile_data=groups_data[group],
            devices=devices_per_group.get(group, []),
            out_path=os.path.join(P.POSITION_IMG_DIR, f"{group}_position.png"),
            specs=specs,
        )

    print("=" * 50)
    print("生成可靠性图...")
    import reliability

    view = data_reader.build_reliability_view(data)
    for dev_name in data_reader.list_reliability_devices(view):
        dim_name = _first_dim(view, dev_name) or REL_SUMMARY_DIM_FALLBACK
        reliability.plot_device(
            dev_name, dim_name, view,
            os.path.join(P.RELIABILITY_IMG_DIR, f"{dev_name}_reliability.png"),
            specs,
        )


# ============================================================
# 页面计划
# ============================================================
def plan_page_indices(prs, n_pos, n_rel, n_cp):
    """把模板页展开成最终页数，并一次性算出各页面组的起始索引。

    模板页序（0-based）：cover summary material doe | position reliability
    failure_mode cp。竖线右侧按实际数量展开：

      count >= 1  保留模板页，复制 count-1 份 -> 占用 count 页
      count == 0  删除该模板页（否则会残留一张未填充的模板页）

    ★ 展开只在这里做一次。各 page_* 的 fill_*_pages 只负责往
      [template_index, template_index+count) 这段里填内容，不再自行复制页，
      否则会把页面重复展开（页数翻倍、索引错位）。

    原先的做法是在 fill_pptx 里用 `off` 逐个模块累加偏移量
    （template_index=5+off / 6+off / 7+off），下游模块一旦增删页就要重算；
    这里改为按实际数量一次算清，索引可审计。
    """
    tpl = P.TEMPLATE_PAGE_INDEX
    idx = {"cover": tpl["cover"], "summary": tpl["summary"],
           "material": tpl["material"], "doe": tpl["doe"]}

    cursor = tpl["position"]
    for key, count in (("position", n_pos), ("reliability", n_rel)):
        if count <= 0:
            U.drop_slide(prs, cursor)
            idx[key] = cursor
            continue
        U.expand_pages(prs, cursor, count)
        idx[key] = cursor
        cursor += count

    # 失效模式：单页，不展开
    idx["failure_mode"] = cursor
    cursor += 1

    # 切片
    if n_cp <= 0:
        U.drop_slide(prs, cursor)
        idx["cp"] = cursor
    else:
        U.expand_pages(prs, cursor, n_cp)
        idx["cp"] = cursor
        cursor += n_cp

    return idx


def _expected_total(n_pos, n_rel, n_cp):
    """按计划最终应有的页数（模板 cp 页会被删除，故减 1）。

    仅用于日志核对；实际校验用的是"规划后页数 == 填充后页数"这一不变式。
    """
    tpl = P.EXPECTED_TEMPLATE_SLIDES
    delta = 0
    delta += (n_pos - 1) if n_pos > 0 else -1
    delta += (n_rel - 1) if n_rel > 0 else -1
    delta += n_cp - 1
    return tpl + delta


# ============================================================
# 保存
# ============================================================
def _save_pptx(prs, output_path):
    """保存报告；目标被占用时回退到系统临时目录的同名文件。

    目标被占用几乎总是因为这份报告还在 PowerPoint 里开着——Windows 上
    被独占打开的文件无法覆盖。退到 TEMP 至少让本次生成不白跑，
    但必须把话说清楚：**原文件没有更新**，不然后面会拿旧稿反复对版式。
    """
    fallback = os.path.join(tempfile.gettempdir(), os.path.basename(output_path))
    try:
        prs.save(output_path)
        return output_path
    except PermissionError:
        print("[WARN] 目标文件写不进去，通常是因为它还在 PowerPoint 里开着：")
        print(f"       {output_path}")
        print("       本次改存到临时目录，原文件不会被更新。")
        print("       想覆盖原文件，请先关掉打开的报告再生成一次。")

    try:
        prs.save(fallback)
        print(f"       -> {fallback}")
        return fallback
    except PermissionError:
        raise PermissionError(
            "目标文件和临时文件都被占用，无法保存：\n"
            f"  {output_path}\n  {fallback}\n"
            "请关闭已打开的报告后重试。")


# ============================================================
# 主流程
# ============================================================
def build_report(data_file=None, template_path=None, output_path=None,
                 images_dir=None, cp_base_dir=None, author=None,
                 report_date=None, open_after=True):
    """生成 DOE 工艺验证报告。

    所有路径参数缺省时回落到 parameter.py 中的配置；显式传入则覆盖，
    下游模块无需改动即可换目录。返回生成的文件路径。

    参数
    ----
    data_file      输入 Excel（DOE_data.xlsx）
    template_path  输入 PPT 模板（GT工艺验证模板.pptx）
    output_path    输出 pptx 路径；None 则按 Excel 项目名自动命名
    images_dir     图表输出目录
    cp_base_dir    切片图片根目录（下面放 <器件>CP 文件夹）
    author         封面作者名；None 取 parameter.AUTHOR_NAME
    report_date    封面日期字符串；None 取今天
    open_after     生成后是否自动打开（仅 Windows 有效）
    """
    cfg = P.configure(data_file=data_file, template_path=template_path,
                      output_path=output_path, images_dir=images_dir,
                      cp_base_dir=cp_base_dir)
    data_file = cfg["data_file"]
    template_path = cfg["template_path"]
    cp_base_dir = cfg["cp_base_dir"]

    missing = [p for p in (data_file, template_path) if not os.path.isfile(p)]
    if missing:
        raise FileNotFoundError(
            "以下路径不存在，请检查 parameter.py 中的配置或用命令行参数指定：\n  "
            + "\n  ".join(missing))

    P.ensure_image_dirs()

    print("1/4 加载数据...")
    print(f"  数据: {data_file}")
    print(f"  模板: {template_path}")
    data = load_data(data_file)
    specs = P.build_specs(data.get("物料信息", {}).get("spec", {}))

    output_path = _resolve_output_path(data, output_path)
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
    if n_slides < P.EXPECTED_TEMPLATE_SLIDES:
        raise ValueError(
            f"模板页数不足：需要至少 {P.EXPECTED_TEMPLATE_SLIDES} 页，"
            f"实际 {n_slides} 页。各页面模块依赖固定页序"
            f"（{P.TEMPLATE_PAGE_INDEX}），请检查 {os.path.basename(template_path)}")

    # --- 先算数量，再一次性规划页序 ---
    n_pos = page_position.count_pages(data)
    n_rel = page_reliability.count_pages(data)
    n_cp = page_cp.count_pages(base_dir=cp_base_dir)
    idx = plan_page_indices(prs, n_pos, n_rel, n_cp)
    n_planned = len(prs.slides)   # 规划后即最终页数
    print(f"  页面计划: 位置精度 {n_pos} 页 @{idx['position']}，"
          f"可靠性 {n_rel} 页 @{idx['reliability']}，"
          f"失效模式 @{idx['failure_mode']}，切片 {n_cp} 页 @{idx['cp']}"
          f"（预计共 {_expected_total(n_pos, n_rel, n_cp)} 页，"
          f"规划后 {n_planned} 页）")

    # --- 逐页填充 ---
    page_cover.fill_cover(prs, data, author=author, today=report_date)
    page_summary.fill_summary(prs, data)
    page_material.fill_material(prs, data)
    page_doe.fill_doe(prs, data)
    page_position.fill_position_pages(prs, data, template_index=idx["position"])
    page_reliability.fill_reliability_pages(prs, data,
                                            template_index=idx["reliability"])
    page_failure_mode.fill_failure_mode_page(prs, data,
                                             template_index=idx["failure_mode"])
    page_cp.fill_cp_pages(prs, data, template_index=idx["cp"],
                          base_dir=cp_base_dir)

    # --- 不变式：页数由"规划"决定，填充阶段不允许再增删页 ---
    if len(prs.slides) != n_planned:
        print(f"[WARN] 填充阶段页数发生变化：规划 {n_planned} 页，"
              f"实际 {len(prs.slides)} 页。"
              f"页面复制/删除只应发生在 plan_page_indices 里。")

    # --- 保存 ---
    print("=" * 50)
    print(f"4/4 保存报告 -> {output_path}")
    saved = _save_pptx(prs, output_path)
    print("=" * 50)
    print(f"报告生成完成 -> {saved}（共 {len(prs.slides)} 页）")

    if open_after:
        try:
            os.startfile(saved)
        except Exception as e:
            print(f"[WARN] 自动打开失败: {e}")
    return saved


if __name__ == "__main__":
    build_report()
