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

def normalize_dim(dim):
    d = str(dim).lower()
    if d.startswith("x"):
        return "X"
    if d.startswith("y"):
        return "Y"
    return dim


def get_spec(device, specs):
    """获取某个设备的 SPEC (lo, hi, label)"""
    return specs.get_position(device)


def get_profile_color(profile):
    return P.PROFILE_COLORS.get(profile, P.PROFILE_COLORS_DEFAULT)


def get_dim_color(dim_key):
    return P.DIM_COLORS.get(dim_key, "#888888")


def half_violin(ax, data, x_pos, color, width=0.35):
    if len(data) < 2:
        return
    try:
        kde = gaussian_kde(data, bw_method=0.25)
    except Exception:
        return
    y_range = np.linspace(min(data) - 0.002, max(data) + 0.002, 100)
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


# ============================================================
# 单个 Profile 绘图
# ============================================================
def plot_single_profile(profile_name, profile_data, devices, out_path, specs):
    """
    profile_data: groups_data[profile_name]
    devices: [(dev_name, [dim1, dim2, ...]), ...]
    specs: parameter.SpecConfig
    """
    skip_dims = {"angle", "Angle", "ANGLE"}

    # 收集 slots
    slots = []
    for dev_name, dims in devices:
        for d in dims:
            if d in skip_dims:
                continue
            raw = profile_data.get(dev_name, {}).get(d, [])
            raw = [float(v) for v in raw if v is not None]
            if not raw:
                continue
            m = float(np.mean(raw))
            vals = [v - m for v in raw]
            slots.append({"dev": dev_name, "dim": d, "vals": vals, "mean": m})

    if not slots:
        print(f"[WARN] {profile_name} 无数据")
        return

    # 读取尺寸
    w, h = P.FIGURE_SIZE.get(profile_name, P.FIGURE_SIZE_DEFAULT)
    fig, ax = plt.subplots(figsize=(w, h))

    # 布局位置
    BOX_W = P.BOX_WIDTH
    XY_GAP = 0.9
    DEV_GAP = 1.5

    x = 1.0
    positions = []
    for si in slots:
        positions.append(x)
        si["x"] = x
        x += XY_GAP
        # device 内最后一个 dim 后加 DEV_GAP
        next_dev = None
        idx = slots.index(si)
        if idx + 1 < len(slots) and slots[idx + 1]["dev"] != si["dev"]:
            x += DEV_GAP - XY_GAP

    # 画 raincloud
    prof_color = get_profile_color(profile_name)
    for si in slots:
        xp = si["x"]
        vals = si["vals"]
        dim_key = normalize_dim(si["dim"])
        color = get_dim_color(dim_key)

        # Box
        bp = ax.boxplot([vals], positions=[xp], widths=BOX_W,
                        patch_artist=True, showfliers=False,
                        whiskerprops=dict(color="#444", linewidth=1.0),
                        capprops=dict(color="#444", linewidth=1.0),
                        medianprops=dict(color="#222", linewidth=1.8))
        bp["boxes"][0].set_facecolor(color)
        bp["boxes"][0].set_alpha(0.85)
        bp["boxes"][0].set_edgecolor("#333")

        # 均值菱形
        mean_val = np.mean(vals)
        ax.scatter(xp, mean_val, marker="D", s=P.MEAN_SIZE, c="white",
                   edgecolors="#222", linewidths=1.2, zorder=5)

        # 散点（注释掉不显示）
        # jitter_scatter(ax, vals, xp, color,
        #                jitter_width=P.JITTER_WIDTH, size=P.DOT_SIZE)

        # 半边 violin
        half_violin(ax, vals, xp, color, width=P.VIOLIN_WIDTH)

    # SPEC 线（每个 device 一条）
    seen_devs = {}
    for si in slots:
        dev = si["dev"]
        if dev not in seen_devs:
            seen_devs[dev] = si["x"]
    for dev, x_center in seen_devs.items():
        x_center+=0.5
        spec_lo, spec_hi, spec_label = get_spec(dev, specs)
        if spec_hi is not None:
            ax.axhline(spec_hi, color="#1f3a93", linestyle="--",
                       linewidth=1.0, alpha=0.8, zorder=1)
            ax.text(x_center, spec_hi - 0.001, f" {spec_label}",
                    fontsize=P.FONT_SIZE_SPEC, color="#1f3a93", va="top", ha="center")
        if spec_lo is not None:
            ax.axhline(spec_lo, color="#1f3a93", linestyle="--",
                       linewidth=1.0, alpha=0.8, zorder=1)
            ax.text(x_center, spec_lo + 0.001, f" {spec_label}",
                    fontsize=P.FONT_SIZE_SPEC, color="#1f3a93", va="bottom", ha="center")

    # 零线
    ax.axhline(0, color="#999", linestyle="-", linewidth=0.8, alpha=0.5, zorder=0)

    # 底部标签：只写器件名称
    dev_centers = {}
    for si in slots:
        key = si["dev"]
        if key not in dev_centers:
            dev_centers[key] = []
        dev_centers[key].append(si["x"])

    for dev, xs in dev_centers.items():
        cx = np.mean(xs)
        ax.text(cx, -0.04, dev, ha="center", va="top",
                transform=ax.get_xaxis_transform(),
                fontsize=P.FONT_SIZE_LABEL, fontweight="bold", color="#222")

    # 标题
    ax.set_title(f"{profile_name} - 位置精度",
                 fontsize=P.FONT_SIZE_TITLE, fontweight="bold", pad=12)
    ax.set_ylabel("标准差", fontsize=P.FONT_SIZE_LABEL)
    ax.set_xticks([])
    ax.grid(axis="y", linestyle=":", linewidth=0.8, alpha=0.4, color="#AAA")
    ax.set_axisbelow(True)

    # 封闭边框
    ax.spines["right"].set_visible(True)
    ax.spines["top"].set_visible(True)

    # Y 轴范围
    all_vals = [v for si in slots for v in si["vals"]]
    if all_vals:
        ymin, ymax = min(all_vals), max(all_vals)
        data_range = ymax - ymin if ymax != ymin else 1
        margin = data_range * 0.15
        lo = ymin - margin
        hi = ymax + margin
        # 考虑 SPEC
        for dev in seen_devs:
            spec_lo, spec_hi, _ = get_spec(dev, specs)
            if spec_lo is not None:
                lo = min(lo, spec_lo - margin * 0.3)
            if spec_hi is not None:
                hi = max(hi, spec_hi + margin * 0.3)
        ax.set_ylim(lo, hi)

    # 图例
    legend_handles = [
        Patch(facecolor=get_dim_color("X"), edgecolor="#333", label="X"),
        Patch(facecolor=get_dim_color("Y"), edgecolor="#333", label="Y"),
    ]
    ax.legend(handles=legend_handles, loc="upper left",
              fontsize=P.FONT_SIZE_TICK, frameon=True, framealpha=0.95, ncol=2)

    plt.subplots_adjust(left=0.08, right=0.95, top=0.90, bottom=0.25)

    # 保存
    plt.savefig(out_path, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"[OK] {profile_name} -> {out_path}")


# ============================================================
# 数据转换
# ============================================================
def build_precision_inputs(precision_data):
    group_names = list(precision_data.keys())
    devices_per_group = {}
    groups_data = {}

    for group in group_names:
        groups_data[group] = {}
        device_list = []
        for device_name, device_content in precision_data[group].items():
            if device_name == "统计":
                continue
            if not isinstance(device_content, dict):
                continue
            dims = [k for k in device_content.keys() if k != "统计"]
            device_list.append((device_name, dims))
            groups_data[group][device_name] = {
                d: device_content[d] for d in dims
            }
        devices_per_group[group] = device_list

    return groups_data, group_names, devices_per_group


# ============================================================
# Main
# ============================================================
if __name__ == "__main__":
    data = reader.main()

    # 从物料信息表构建 SPEC
    specs = P.build_specs(data.get("物料信息", {}).get("spec", {}))

    groups_data, group_names, devices_per_group = (
        build_precision_inputs(data["位置精度"])
    )

    print("识别到分组：", group_names)

    P.ensure_image_dirs()

    for group in group_names:
        devs = devices_per_group.get(group, [])
        print(f"  {group}: {[(d, dims) for d, dims in devs]}")

        out_path = os.path.join(P.POSITION_IMG_DIR, f"{group}_position.png")
        plot_single_profile(
            profile_name=group,
            profile_data=groups_data[group],
            devices=devs,
            out_path=out_path,
            specs=specs,
        )

        # 自动打开
        os.startfile(out_path)
