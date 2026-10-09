# ============================================================
# 图尺寸 (width, height) 单位英寸
# ============================================================
import os
import re

PROJECT_DIR  = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = r"D:\工作\进行中\macom 70mW DOE\macom70&285-0814"



DATA_FILE = os.path.join(BASE_DIR, "DOE_data.xlsx")        # 输入：数据 Excel
TEMPLATE_PATH = os.path.join(PROJECT_DIR , "GT工艺验证模板.pptx")  # 输入：PPT 模板
OUTPUT_PATH_FMT = os.path.join(BASE_DIR, "{project_name}_GT工艺验证报告.pptx")  # 输出：报告名（{project_name} 取自 Excel 项目名）
IMAGES_DIR = os.path.join(BASE_DIR, "images")               # 输出：图表（position/ reliability 两个子目录）
CP_BASE_DIR = BASE_DIR                                       # 输入：切片图片目录（下面放 <器件>CP 文件夹）

POSITION_IMG_DIR = os.path.join(IMAGES_DIR, "position")
RELIABILITY_IMG_DIR = os.path.join(IMAGES_DIR, "reliability")


def ensure_image_dirs():
    os.makedirs(POSITION_IMG_DIR, exist_ok=True)
    os.makedirs(RELIABILITY_IMG_DIR, exist_ok=True)


FIGURE_SIZE = {
    "profile1": (8, 6),
    "profile2": (8, 6),
    "profile3": (8, 6),
}
# 默认尺寸（未配置的 profile 使用）
FIGURE_SIZE_DEFAULT = (8, 6)

# ============================================================
# SPEC 默认值
# ============================================================
SPEC_DEFAULT = (-0.02, 0.02, "SPEC")
RELIABILITY_SPEC_DEFAULT = None   # 设为 None 则未配置器件不画 SPEC 线


class SpecConfig:
    """SPEC 配置容器，显式创建/传递。

    取代原来的模块级可变字典（SPEC / RELIABILITY_DEVICE_SPEC），
    避免依赖 load_specs_from_data 的调用顺序。
      - position_spec: {device: (lower, upper, label)}
      - shear_spec:    {device: (value, label)}
    """

    def __init__(self, position_spec=None, shear_spec=None):
        self.position_spec = dict(position_spec or {})
        self.shear_spec = dict(shear_spec or {})

    def get_position(self, device):
        return self.position_spec.get(device, SPEC_DEFAULT)

    def get_shear(self, device):
        return self.shear_spec.get(device, RELIABILITY_SPEC_DEFAULT)


def build_specs(spec_data):
    """从 data_reader 解析的物料信息 SPEC 表构建 SpecConfig。

    spec_data: {device_name: {"位置": str, "剪切力": float, "空洞率": float}}
    """
    cfg = SpecConfig()
    for dev_name, vals in (spec_data or {}).items():
        # 位置精度 SPEC：解析 "(±0.02，0.02)" 格式
        pos_str = vals.get("位置")
        if pos_str and isinstance(pos_str, str):
            m = re.search(r'[\u00b1+-]?([\d.]+)', pos_str)
            if m:
                bound = float(m.group(1))
                cfg.position_spec[dev_name] = (-bound, bound, f"{dev_name} SPEC")

        # 剪切力 SPEC
        shear_val = vals.get("剪切力")
        if shear_val is not None:
            try:
                cfg.shear_spec[dev_name] = (
                    float(shear_val),
                    f"{dev_name} Die shear SPEC"
                )
            except (ValueError, TypeError):
                pass
    return cfg

# ============================================================
# 颜色配置
# ============================================================

# X / Y 维度颜色（所有 profile 共用）
DIM_COLORS = {
    "X": "#4C78A8",
    "Y": "#F58518",
}

# 每个 profile 的主色调（用于半透明 violin 填充）
PROFILE_COLORS = {
    "profile1": "#E4A72E",
    "profile2": "#C76DA5",
    "profile3": "#5DA5DA",
}
PROFILE_COLORS_DEFAULT = "#888888"

# ============================================================
# 图表通用样式
# ============================================================
BOX_WIDTH = 0.5
JITTER_WIDTH = 0.18
VIOLIN_WIDTH = 0.35
DOT_SIZE = 10
MEAN_SIZE = 30

# ============================================================
# 字体配置
# ============================================================
FONT_CHINESE = "Microsoft YaHei"
FONT_ENGLISH = "Times New Roman"
FONT_SIZE_TITLE = 15
FONT_SIZE_LABEL = 12
FONT_SIZE_SPEC = 9
FONT_SIZE_TICK = 10

# ============================================================
# 可靠性配置
# ============================================================

# 可靠性图尺寸
RELIABILITY_FIG_SIZE = (14, 6)

# 可靠性条件顺序
RELIABILITY_CONDITIONS = ["T0", "168 TCT", "168 UDH", "500 TCT", "500 UDH"]

# 可靠性数据键值映射：显示名 -> (sheet_key, block_name)
RELIABILITY_MAP = {
    "T0":     ("可靠性-T0", "T0"),
    "168 TCT": ("可靠性-T168", "TCT168"),
    "168 UDH": ("可靠性-T168", "UDH168"),
    "500 TCT": ("可靠性-T500", "TCT500"),
    "500 UDH": ("可靠性-T500", "UDH500"),
}

# 每个 profile 的颜色（可靠性图用）
RELIABILITY_PROFILE_COLORS = {
    "profile1": "#E4A72E",
    "profile2": "#C76DA5",
    "profile3": "#5DA5DA",
}

# 可靠性条件分组：[组1, 组2, ...]，组内条件并列展示
RELIABILITY_COND_GROUPS = [
    ["T0"],
    ["168 TCT", "168 UDH"],
    ["500 TCT", "500 UDH"],
]

# ============================================================
# 随机数种子（jitter 等随机元素，保证多次生成图片一致）
# ============================================================
RNG_SEED = 20260928

# ============================================================
# PPT 版面尺寸（cm）
# ============================================================

# 位置精度页
POSITION_IMG_W_CM = 25
POSITION_IMG_H_CM = 15
POSITION_IMG_LEFT_CM = 4.43
POSITION_IMG_TOP_CM = 2.1
POSITION_CONCL_NEW_BOX = (1.0, 17.3, 31.87, 1.4)   # 新建结论文本框 (left, top, w, h)
POSITION_CONCL_FIT_BOX = (1.79, 30.0)               # 复用模板文本框时 (left, width)

# 可靠性页
RELIABILITY_IMG_W_CM = 33
RELIABILITY_IMG_H_CM = 11
RELIABILITY_IMG_TOP_CM = 2.4
RELIABILITY_CONCL_TOP_CM = 13.9
RELIABILITY_CONCL_LEFT_CM = 1.0
RELIABILITY_CONCL_W_CM = 31.87
RELIABILITY_CONCL_H_CM = 1.6

# CP 切片页：切片图在表格中的固定高度（cm），宽度等比例
CP_IMG_HEIGHT_CM = 3.5
CP_CONCL_DEFAULT_TOP_CM = 16.9   # 无表格参照时结论文本框的默认 top
CP_CONCL_LEFT_CM = 1.0
CP_CONCL_W_CM = 31.87
CP_CONCL_H_CM = 1.2
CP_CONCL_GAP_CM = 0.3
