import os
import sys
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from scipy.stats import gaussian_kde

# ============================================================
# 路径
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

# ============================================================
# 导入模块
# ============================================================
try:
    import importlib
    reader = importlib.import_module("data_reader")
    import parameter as P
except Exception as e:
    print(f"[ERROR] 导入失败：{e}")
    sys.exit(1)

# ============================================================
# 字体（从 parameter.py 读取）
# ============================================================
matplotlib.rcParams["font.sans-serif"] = [
    P.FONT_CHINESE, "SimHei", "Arial Unicode MS"
]
matplotlib.rcParams["font.serif"] = [
    P.FONT_ENGLISH, "Times New Roman", "SimSun"
]
matplotlib.rcParams["axes.unicode_minus"] = False
matplotlib.rcParams["axes.linewidth"] = 1.0

# 固定随机种子，保证多次生成图片一致
RNG = np.random.default_rng(P.RNG_SEED)


# ============================================================
# 工具函数
# ============================================================
def half_violin(ax, data, x_pos, color, width=0.35):
    if len(data) < 2:
        return
    try:
        kde = gaussian_kde(data, bw_method=0.25)
    except Exception:
        return
    y_range = np.linspace(min(data) - 1, max(data) + 1, 100)
    density = kde(y_range)
    density = density / density.max() * width
    ax.fill_betweenx(y_range, x_pos, x_pos + density,
                     facecolor=color, alpha=0.35, edgecolor=color,
                     linewidth=1.0, linestyle="--")


def jitter_scatter(ax, data, x_pos, color, jitter_width=0.18, size=10):
    if not data:
        return
    jitter = RNG.uniform(-jitter_width, jitter_width, len(data))
    ax.scatter(x_pos + jitter, data, s=size, c=color,
               alpha=0.6, edgecolors="white", linewidths=0.3, zorder=3)


# 未配置颜色的 profile 使用的备选色（按出现顺序循环取）
FALLBACK_PROFILE_COLORS = [
    "#4CAF50", "#9C27B0", "#FF5722", "#009688", "#795548", "#607D8B",
]


# ============================================================
# 单个器件绘图
# ============================================================
def plot_device(device_name, dim_name, all_data, out_path, specs):
    """
    all_data: {condition: {profile: {device: {dim: [values]}}}}
    specs: parameter.SpecConfig
    """
    # profile 列表按数据实际出现顺序（Excel 列顺序）动态识别，
    # 不再依赖 RELIABILITY_PROFILE_COLORS 的 key，否则 Vendor profile 等
    # 未配色的 profile 会被整段跳过
    profiles = []
    for cond_data in all_data.values():
        for prof in cond_data:
            if prof not in profiles:
                profiles.append(prof)
    colors = {p: P.RELIABILITY_PROFILE_COLORS.get(p)
              or FALLBACK_PROFILE_COLORS[i % len(FALLBACK_PROFILE_COLORS)]
              for i, p in enumerate(profiles)}
    cond_groups = P.RELIABILITY_COND_GROUPS

    # ---------- 收集 slots ----------
    slots = []
    x = 1.0
    PROF_GAP = 0.75
    COND_GAP = 1.3
    GROUP_GAP = 2.5

    for gi, group in enumerate(cond_groups):
        for ci, cond in enumerate(group):
            cond_data = all_data.get(cond, {})
            for pi, prof in enumerate(profiles):
                vals = cond_data.get(prof, {}).get(device_name, {}).get(dim_name, [])
                vals = [float(v) for v in vals if v is not None]
                if not vals:
                    continue
                slots.append({"cond": cond, "prof": prof, "vals": vals, "x": x})
                x += PROF_GAP
            x += COND_GAP - PROF_GAP
        x += GROUP_GAP - COND_GAP

    if not slots:
        print(f"[WARN] {device_name} 无数据")
        return

    all_vals = [v for si in slots for v in si["vals"]]
    data_min, data_max = min(all_vals), max(all_vals)

    # ---------- 断轴区间计算 ----------
    # 下段：0 ~ break_low；上段：break_high ~ data_max+margin
    # 让下段只占很小高度，用于显示 0 及低值区
    span = data_max - data_min if data_max != data_min else 50
    margin = span * 0.12

    # 断点：数据最小值下方留一点，作为下段顶部
    break_low = max(0, data_min - margin)
    break_high = break_low + span * 0.02   # 上段起点略高于断点，制造"跳过"感

    # 如果数据本身就在 0 附近（没有明显断层），退化为普通单轴
    use_broken = data_min > span * 0.25   # 数据整体远离 0 时才断

    # ---------- 画布：上下两个子图 ----------
    n_slots = len(slots)
    fig_w = max(10, n_slots * 0.85 + 4)

    if use_broken:
        fig = plt.figure(figsize=(fig_w, 6.5))
        gs = fig.add_gridspec(
            2, 1, height_ratios=[6, 0.9], hspace=0.04,
            left=0.08, right=0.95, top=0.92, bottom=0.22,
        )
        ax_top = fig.add_subplot(gs[0])
        ax_bot = fig.add_subplot(gs[1], sharex=ax_top)
        axes = [ax_top, ax_bot]
    else:
        fig, ax_top = plt.subplots(figsize=(fig_w, 6))
        ax_bot = None
        axes = [ax_top]
        plt.subplots_adjust(left=0.08, right=0.95, top=0.92, bottom=0.22)

    # ---------- 在指定轴上绘制 raincloud ----------
    def draw_on(ax, ylim_lo, ylim_hi):
        for si in slots:
            xp, vals = si["x"], si["vals"]
            color = colors.get(si["prof"], "#888")

            bp = ax.boxplot(
                [vals], positions=[xp], widths=P.BOX_WIDTH,
                patch_artist=True, showfliers=False,
                whiskerprops=dict(color="#444", linewidth=1.0),
                capprops=dict(color="#444", linewidth=1.0),
                medianprops=dict(color="#222", linewidth=1.8),
            )
            bp["boxes"][0].set_facecolor(color)
            bp["boxes"][0].set_alpha(0.85)
            bp["boxes"][0].set_edgecolor("#333")

            ax.scatter(xp, np.mean(vals), marker="x", s=P.MEAN_SIZE,
                       c="#222", linewidths=1.5, zorder=5)

            # jitter_scatter(ax, vals, xp, color,
            #                jitter_width=P.JITTER_WIDTH, size=P.DOT_SIZE)
            half_violin(ax, vals, xp, color, width=P.VIOLIN_WIDTH)

        ax.set_ylim(ylim_lo, ylim_hi)
        ax.grid(axis="y", linestyle=":", linewidth=0.8, alpha=0.4, color="#AAA")
        ax.set_axisbelow(True)
        ax.spines["right"].set_visible(True)
        ax.spines["top"].set_visible(True)

    # ---------- 绘制 ----------
    if use_broken:
        draw_on(ax_top, break_high, data_max + margin)
        draw_on(ax_bot, 0, break_low)

        # 隐藏中间相邻边框
        ax_top.spines["bottom"].set_visible(False)
        ax_bot.spines["top"].set_visible(False)
        ax_top.tick_params(axis="x", which="both", bottom=False, labelbottom=False)

        # 断轴斜杠（在两轴交界处，左右各一组）
        d = 0.005
        kw = dict(transform=ax_top.transAxes, color="#333",
                  clip_on=False, linewidth=1.2)
        # 左侧
        ax_top.plot((-d, +d), (-d, +d), **kw)
        ax_top.plot((-d, +d), (-3*d, -d), **kw)
        # 右侧
        ax_top.plot((1-d, 1+d), (-d, +d), **kw)
        ax_top.plot((1-d, 1+d), (-3*d, -d), **kw)

        ylabel_ax = ax_top
    else:
        draw_on(ax_top, max(0, data_min - margin), data_max + margin)
        ylabel_ax = ax_top

    # ---------- SPEC 线（画在对应区间所在的轴上）----------
    spec_info = specs.get_shear(device_name)
    if spec_info is not None:
        spec_val, spec_label = spec_info
    else:
        spec_val = None
    if spec_val is not None:
        for ax in axes:
            lo, hi = ax.get_ylim()
            if lo <= spec_val <= hi:
                ax.axhline(spec_val, color="red", linestyle="-.",
                           linewidth=1.5, alpha=0.9, zorder=1)
                ax.text(slots[0]["x"], spec_val + (hi - lo) * 0.015,
                        spec_label, fontsize=P.FONT_SIZE_SPEC,
                        color="red", va="bottom", ha="left")

    # ---------- 底部条件标签（画在最下面的轴上）----------
    label_ax = ax_bot if use_broken else ax_top
    cond_centers = {}
    for si in slots:
        cond_centers.setdefault(si["cond"], []).append(si["x"])

    for cond, xs in cond_centers.items():
        cx = np.mean(xs)
        label_ax.text(cx, -0.06, cond, ha="center", va="top",
                      transform=label_ax.get_xaxis_transform(),
                      fontsize=P.FONT_SIZE_LABEL,
                      fontweight="bold", color="#222")

    # ---------- 组分隔线 ----------
    for gi in range(len(cond_groups)):
        group_slots = [si for si in slots
                       for cond in cond_groups[gi] if si["cond"] == cond]
        if group_slots and gi > 0:
            left = min(si["x"] for si in group_slots)
            for ax in axes:
                ax.axvline(left - 0.6, color="#D0D0D0",
                           linestyle="--", linewidth=1.0, zorder=0)

    # ---------- 标题 / 轴标签 ----------
    ax_top.set_title(f"{device_name} - Die Shear 可靠性",
                     fontsize=P.FONT_SIZE_TITLE, fontweight="bold", pad=12)
    if use_broken:
        # 用 fig.text 让 Y 轴标签居中于整个断轴区域
        fig.text(0.02, 0.57, "Die shear (g)", rotation=90,
                 va="center", ha="center", fontsize=P.FONT_SIZE_LABEL)
        ax_top.tick_params(labelbottom=False)
        ax_bot.set_xticks([])
    else:
        ylabel_ax.set_ylabel("Die shear (g)", fontsize=P.FONT_SIZE_LABEL)
        ylabel_ax.set_xticks([])

    # ---------- 图例 ----------
    legend_handles = [
        Patch(facecolor=colors[p], edgecolor="#333", label=p)
        for p in profiles
    ]
    ax_top.legend(handles=legend_handles, loc="upper left",
                  fontsize=P.FONT_SIZE_TICK, frameon=True,
                  framealpha=0.95, ncol=2)

    # ---------- 保存 ----------
    plt.savefig(out_path, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"[OK] {out_path}")

# ============================================================
# Main
# ============================================================
if __name__ == "__main__":
    data = reader.main()

    # 从物料信息表构建 SPEC
    specs = P.build_specs(data.get("物料信息", {}).get("spec", {}))
    P.ensure_image_dirs()

    # 整理数据：{condition: {profile: {device: {dim: [values]}}}}
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

    # 从数据中动态识别所有器件
    all_devices = set()
    for cond_data in all_data.values():
        for prof_data in cond_data.values():
            for dev_name in prof_data:
                if dev_name != "SN":
                    all_devices.add(dev_name)

    # 每个器件单独出图
    for dev_name in sorted(all_devices):
        # 自动检测维度名：从数据中取第一个有效维度
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
            dim_name = "剪切力"  # fallback

        out_path = os.path.join(P.RELIABILITY_IMG_DIR,
                                f"{dev_name}_reliability.png")
        plot_device(dev_name, dim_name, all_data, out_path, specs)
        os.startfile(out_path)
