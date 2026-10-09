# DOE 工艺验证报告自动生成器

基于 `DOE_data.xlsx` 中的验证数据，自动绘制位置精度图 / 可靠性图，并套用到 `GT工艺验证模板.pptx`，
最终输出一份可直接交付的 `DOE_工艺验证报告.pptx`。

## 快速开始

```bash
pip install -r requirements.txt

# 生成完整报告（推荐）
python report_generator.py

# 仅重新生成图表
python run.py
```

## 路径配置

输入数据在哪、生成结果放哪，只需要改 `parameter.py` 顶部这一处，其他代码都不用动：

```python
# ★ 路径配置：输入在哪 / 输出到哪，只需要改这里 ★
DATA_FILE     = os.path.join(BASE_DIR, "DOE_data.xlsx")         # 输入：数据 Excel
TEMPLATE_PATH = os.path.join(BASE_DIR, "GT工艺验证模板.pptx")    # 输入：PPT 模板
OUTPUT_PATH   = os.path.join(BASE_DIR, "DOE_工艺验证报告.pptx")  # 输出：生成的报告
IMAGES_DIR    = os.path.join(BASE_DIR, "images")                # 输出：图表
CP_BASE_DIR   = BASE_DIR                                        # 输入：切片图片目录
```

| 项 | 说明 |
|---|---|
| `DATA_FILE` | 输入数据 Excel |
| `TEMPLATE_PATH` | 输入 PPT 模板 |
| `OUTPUT_PATH` | 报告输出路径（所在目录不存在会自动创建） |
| `IMAGES_DIR` | 图表输出目录，下分 `position\` 和 `reliability\` 两个子目录 |
| `CP_BASE_DIR` | 切片图片目录，下面放 `<器件>CP` 文件夹 |

写绝对路径即可换目录，例如 `DATA_FILE = r"D:\项目\DOE_data.xlsx"`。
数据文件或模板不存在时会直接报错并指出是哪个路径。

## 目录结构

```
├── DOE_data.xlsx               # 输入：验证数据源（唯一数据入口）
├── GT工艺验证模板.pptx          # 输入：报告模板
├── DOE_工艺验证报告.pptx        # 输出：生成的报告
├── report_generator.py          # 主入口：读数据 → 画图 → 填充PPT → 保存
├── run.py                       # 只跑图表（位置精度 + 可靠性）
├── requirements.txt             # Python 依赖（已测试版本见文件内注释）
│
├── data_reader.py               # Excel 解析：按表头自适应，不依赖固定列偏移
├── parameter.py                 # 全局配置：路径/图尺寸/颜色/字体/SPEC/可靠性条件/PPT版面坐标
│
├── position.py                  # 位置精度图：violin + boxplot + jitter 散点
├── reliability.py               # 可靠性图：按器件横向对比 T0/T168/T500
│
├── slide4_doe.py                # 第4页：DOE 参数表
├── slide_position.py            # 第5页+：位置精度（每个 profile 一页）
├── slide_reliability.py         # 第6页+：可靠性（每个器件一页）
├── slide_CP.py                  # 第7页+：CP 切片（每个 xxxCP 文件夹一页）
│
├── images/                      # 图表输出
│   ├── position/profileN_position.png
│   └── reliability/<器件>_reliability.png
│
├── 激光器CP/  长电容CP/ ...      # 输入：CP 切片图片，命名 <器件>CP
│   └── CP/<profile>-<序号>/A.*   # 取每张 profile 中名为 A 的图
│
└── 需求.txt                     # 开发需求记录
```

## 数据流

```
DOE_data.xlsx
      │  data_reader.main()
      ▼
  { 物料信息, 位置精度, 可靠性-T0, 可靠性-T168, 可靠性-T500 }
      │
       ├─► parameter.build_specs()   # 从物料信息 SPEC 表构建 SpecConfig（显式传参）
      │
      ├─► position.py  ──► images/position/*.png
      ├─► reliability.py ─► images/reliability/*.png
      │
      └─► report_generator.fill_pptx()       # 填充模板各页
              ├─ Slide 1  封面（项目名 / 姓名 / 日期）
              ├─ Slide 2  汇总表
              ├─ Slide 3  验证信息（物料）
              ├─ Slide 4  DOE 参数
              ├─ Slide 5+ 位置精度（每 profile 一页）
              ├─ Slide 6+ 可靠性（每器件一页）
              ├─ Slide 7+ CP 切片（每器件一页，表格 + 切片图）
              └─ 附件页
```

`data_reader` 按表头文本自适应定位（如 `NO`、`料号`、`SPEC`、`位置精度`、`剪切力`、`空洞率`），
新增/删除列不影响解析。

## 输出页面生成规则

| 页 | 模块 | 分页依据 | 结论文案 |
|---|---|---|---|
| 5+ | `slide_position.py` | 每个 profile 一页，图 25×15cm 居中 | 全部满足 → "结论：xxx 下，所有器件位置精度均满足SPEC" |
| 6+ | `slide_reliability.py` | 每个器件一页，图 33×11cm 居中 | 全部满足 → "结论：xxx在可靠性测试中Die shear均满足SPEC" |
| 7+ | `slide_CP.py` | 每个 `xxxCP` 文件夹一页；表格行数 = 实际 profile 数 | 全部满足 → "结论：xxx器件空洞率符合SPEC" |

CP 切片图路径约定：`<器件>CP/CP/<profile号>-<序号>/A.*`，插入表格时等比缩放至高度 3.5cm。

## 可靠性条件

定义在 `parameter.py`：

```python
RELIABILITY_CONDITIONS = ["T0", "168 TCT", "168 UDH", "500 TCT", "500 UDH"]
RELIABILITY_MAP = {
    "T0":      ("可靠性-T0",  "T0"),
    "168 TCT": ("可靠性-T168", "TCT168"),
    "168 UDH": ("可靠性-T168", "UDH168"),
    "500 TCT": ("可靠性-T500", "TCT500"),
    "500 UDH": ("可靠性-T500", "UDH500"),
}
```

## 常用配置（parameter.py）

| 项 | 说明 |
|---|---|
| `FIGURE_SIZE` / `RELIABILITY_FIG_SIZE` | 图尺寸（英寸） |
| `SPEC_DEFAULT` | 位置精度 SPEC 默认值 `(-0.02, 0.02, "SPEC")` |
| `RELIABILITY_SPEC_DEFAULT` | 剪切力 SPEC 默认值（`None` 表示不画 SPEC 线） |
| `SpecConfig` / `build_specs()` | SPEC 配置容器与构造函数；由调用方显式传入绘图函数 |
| `POSITION_*` / `RELIABILITY_*` / `CP_*` | PPT 版面尺寸与结论文本框坐标（cm） |
| `DIM_COLORS` / `PROFILE_COLORS` | X/Y 维度色、profile 主色 |
| `FONT_CHINESE` / `FONT_ENGLISH` | 中文字体 `Microsoft YaHei`，西文 `Times New Roman` |
| `RELIABILITY_COND_GROUPS` | 可靠性条件分组展示 |
| `RNG_SEED` | 随机数种子（jitter 等随机元素），保证多次生成图片一致 |

SPEC 通过 `build_specs()` 从 `DOE_data.xlsx` 的物料信息 SPEC 表构建为 `SpecConfig` 实例，
再由 `plot_single_profile(...)` / `plot_device(...)` 显式接收，缺失时回落到默认值。

## 注意事项

- `report_generator.py` 会先删除 `OUTPUT_PATH` 再重写，请确保报告未被打开占用；
  被占用时会回退保存到系统临时目录（`%TEMP%`）下同名文件。
- 生成成功后会自动打开报告文件（`os.startfile`，仅 Windows）。
- 模板页数在 `fill_pptx()` 入口校验：`GT工艺验证模板.pptx` 必须至少 8 页
  （`EXPECTED_TEMPLATE_SLIDES`），否则直接报错。改动模板页序需同步
  `slide_position.py` / `slide_reliability.py` / `slide_CP.py` 中的 `TEMPLATE_SLIDE_INDEX`。
- `data_reader.main()` 只 `load_workbook` 一次，各 sheet 复用同一个 workbook。
- `~$*.pptx` 为 Office 临时锁文件，可忽略/删除。
