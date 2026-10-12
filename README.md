# DOE 工艺验证报告自动生成器

基于 `DOE_data.xlsx` 中的验证数据，自动绘制位置精度图 / 可靠性图 / 失效模式图，
套用到 `GT工艺验证模板.pptx`，最终输出一份可直接交付的 `<项目名>_GT工艺验证报告.pptx`。

## 快速开始

```bash
pip install -r requirements.txt
```

**给不想碰命令行的人：直接拖。**

把 `DOE_data.xlsx`（或整个项目文件夹）拖到 `run.bat` 上即可；
双击 `run.bat` 会弹选择框挑文件。路径全自动推断，见「拖拽与批量」一节。

**命令行**（三种粒度，按需选）：

```bash
# 1. 零配置推断：给个入口就行，其余自动推断（拖拽的后端）
python launcher.py "D:\项目\DOE_data.xlsx"
python launcher.py "D:\项目A" "D:\项目B" --no-open    # 批量
python launcher.py "D:\项目" --dry-run               # 只看会怎么跑

# 2. 显式参数：路径一条条写清楚（GUI / 脚本 / 自动化调用）
python cli.py --data "D:\项目\DOE_data.xlsx" \
              --output "D:\项目\out\报告.pptx" \
              --cp-dir "D:\项目" \
              --no-open

# 3. 老办法：读 parameter.py 里的配置
python cli.py
# 或
python report_generator.py

# 只看当前生效的配置
python cli.py --print-config

# 仅重新生成图表
python run.py
```

> `--output` 支持 `{project_name}` 占位符，例如
> `--output "D:\项目\{project_name}_GT工艺验证报告.pptx"`，
> 文件名会由 Excel 里的项目名填充。

## 拖拽与批量

`run.bat` 是给人工操作用的薄壳，真正的推断逻辑在 `launcher.py`。

| 你拖进去的东西 | 会怎么跑 |
|---|---|
| `DOE_data.xlsx` | 用它作数据源，所在目录同时当切片图根目录和输出目录 |
| 项目文件夹 | 在里面找 `DOE_data.xlsx`（找不到就取唯一的 xlsx） |
| 多个数据 / 文件夹 | 批量逐个生成，**默认不自动打开**，免得弹一屏窗口 |
| 名字含"模板"的 `.pptx` | 当作 PPT 模板 |

其余全部自动推断：模板按「拖入的 → 数据同目录的 `GT工艺验证模板.pptx`
→ 同目录唯一的 `*模板*.pptx` → 项目目录自带」顺序找；报告输出到
`<数据所在目录>/{项目名}_GT工艺验证报告.pptx`；图表输出到 `<数据所在目录>/images`。

常用开关：`--dry-run`（只看计划不生成）、`--no-open`、`--template`、`--out-dir`。

### ⚠ 拖拽的一个硬限制：路径含 `&`

Windows 把文件拖到 `.bat` 上时，**只会对含空格的路径自动加引号**。
路径里含 `&` 又不含空格时（例如历史目录 `macom70&285-0814`）不会加引号，
cmd 会在 `&` 处把命令行截断，Python 拿不到完整路径。

三条出路，任选：

1. **双击 `run.bat` 用弹框选文件** —— 路径由 Python 直接取，完全不经过命令行解析，最省事；
2. 命令行加引号：`python launcher.py "D:\工作\进行中\macom70&285-0814"`；
3. 用 `cli.py` 显式参数。

含空格路径不用担心：即使被 cmd 拆成好几段，`launcher.regroup()`
也会按"拼回去能不能命中真实路径"重新拼合。

## 路径配置

两条路，任选其一。

**1. 改 `parameter.py` 顶部**（一次配置，长期复用）：

```python
BASE_DIR      = r"D:\工作\进行中\<项目目录>"
DATA_FILE     = os.path.join(BASE_DIR, "DOE_data.xlsx")        # 输入：数据 Excel
TEMPLATE_PATH = os.path.join(PROJECT_DIR, "GT工艺验证模板.pptx")  # 输入：PPT 模板
OUTPUT_PATH_FMT = os.path.join(BASE_DIR, "{project_name}_GT工艺验证报告.pptx")
IMAGES_DIR    = os.path.join(BASE_DIR, "images")               # 输出：图表
CP_BASE_DIR   = BASE_DIR                                       # 输入：切片图片目录
```

**2. 用命令行参数覆盖**（见上），由 `parameter.configure()` 在运行时改写，
下游模块无需改动。这是 GUI / 批量模式的统一入口。

> ⚠ `BASE_DIR` 目前指向的是某个历史项目路径，**而且那个目录已经不存在了**。
> 现在更省事：不用动这一行，直接用 `run.bat` / `launcher.py` / `cli.py --data` 指定路径；
> 只有当你想把某个项目固化成"默认"时才需要改它。数据或模板不存在时会明确报出是哪个路径。

## 目录结构

```
├── DOE_data.xlsx               # 输入：验证数据源（唯一数据入口）
├── GT工艺验证模板.pptx          # 输入：报告模板（必须 ≥ 8 页）
├── <项目名>_GT工艺验证报告.pptx  # 输出：生成的报告
│
├── parameter.py                 # 全局配置：路径 / SPEC / 颜色 / 字体 / PPT 版面坐标 / 模板页序
├── pptx_utils.py                # PPT 操作基元：字体、单元格、表格增删行列、幻灯片复制重排、图片定位
├── data_reader.py               # Excel 解析：按表头自适应 + 可靠性视图归一化
│
├── position.py                  # 位置精度图：violin + boxplot（raincloud）
├── reliability.py               # 可靠性图：按器件横向对比 T0/T168/T500，自动断轴
│
├── page_cover.py                # 第 1 页  封面
├── page_summary.py              # 第 2 页  汇总表（三项判定 + PASS/FAIL 配色）
├── page_material.py             # 第 3 页  验证信息（物料表）
├── page_doe.py                  # 第 4 页  DOE 参数表
├── page_position.py             # 第 5 页+ 位置精度（每个 profile 一页）
├── page_reliability.py          # 第 6 页+ 可靠性（每个器件一页）
├── page_failure_mode.py         # 第 7 页  失效模式（单页）
├── page_cp.py                   # 第 8 页+ CP 切片（每个 xxxCP 文件夹一页）
│
├── report_generator.py          # 主流程编排：build_report()
├── cli.py                       # 命令行入口（显式参数；给脚本 / GUI 用）
├── launcher.py                  # 拖拽 / 批量入口（零配置推断；run.bat 的后端）
├── run.bat                      # 双击 / 拖拽启动器（薄壳，只负责找 Python）
├── run.py                       # 只跑两张图（调试用）
│
├── images/                      # 图表输出
│   ├── position/profileN_position.png
│   └── reliability/<器件>_reliability.png
│
├── 激光器CP/ 长电容CP/ ...        # 输入：CP 切片图片，命名 <器件>CP
│   └── CP/<profile号>-<序号>/A.*  # 取每张 profile 中名为 A 的图
│
└── 需求.txt                     # 开发需求记录
```

## 分层约定

新增或修改页面的正确做法：

| 你要做的事 | 改哪里 |
|---|---|
| 改文字、配色、图片尺寸、文本框坐标 | `parameter.py`（不要散落在页面模块里） |
| 新增一种 PPT 操作（如合并单元格） | `pptx_utils.py` |
| 改某一页填什么内容 | 对应的 `page_*.py` |
| 改页序、加一种新页面 | `report_generator.plan_page_indices()` |
| 改 Excel 解析规则 | `data_reader.py` |
| 加一个命令行参数 | `cli.py`（显式接口）或 `launcher.py`（推断接口） |
| 改拖拽时的路径推断规则 | `launcher.build_jobs()` / `find_data_file()` / `find_template()` |

**关键约定：页面复制/删除只能发生在 `plan_page_indices()` 里。**
各 `page_*.fill_*_pages()` 只往 `[template_index, template_index+count)`
这段区间里填内容，绝不自行复制页——否则页面会被重复展开导致页数翻倍、索引错位。
主流程末尾有不变式校验：规划后页数必须等于填充后页数。

## 数据流

```
DOE_data.xlsx
      │  data_reader.main()
      ▼
  { 物料信息, 位置精度, 可靠性-T0/T168/T500, 空洞率 }
      │
       ├─► parameter.build_specs()      # 从物料信息 SPEC 表构建 SpecConfig
       ├─► data_reader.build_reliability_view()   # 归一化为 {条件:{profile:{器件:{维度:[值]}}}}
       │
       ├─► position.py    ──► images/position/*.png
       ├─► reliability.py ──► images/reliability/*.png
       │
       └─► report_generator.build_report()
              ├─ plan_page_indices()    # ① 按实际数量一次性规划页序，展开模板页
              └─ 逐页填充               # ② 只填内容，不再增删页
                   第1页 封面（项目名 / 姓名 / 日期）
                   第2页 汇总表
                   第3页 验证信息（物料）
                   第4页 DOE 参数
                   第5页+ 位置精度（每 profile 一页）
                   第6页+ 可靠性（每器件一页）
                   第7页  失效模式（单页）
                   第8页+ CP 切片（每器件一页）
```

`data_reader` 按表头文本自适应定位（如 `NO`、`料号`、`SPEC`、`位置精度`、
`剪切力`、`空洞率`），新增/删除列不影响解析。

## 输出页面生成规则

| 页 | 模块 | 分页依据 | 结论文案 |
|---|---|---|---|
| 5+ | `page_position` | 每个 profile 一页，图 25×15cm | 全部满足 → "结论：xxx 下，所有器件位置精度均满足SPEC" |
| 6+ | `page_reliability` | 每个器件一页，图 33×11cm 居中 | 全部满足 → "结论：xxx在可靠性测试中Die shear均满足SPEC" |
| 7 | `page_failure_mode` | 单页；行=profile，列=可靠性条件 | 无（单元格内填失效模式图） |
| 8+ | `page_cp` | 每个 `xxxCP` 文件夹一页；表格行数 = 实际 profile 数 | 全部满足 → "结论：xxx器件空洞率符合SPEC" |

CP 切片图路径约定：`<器件>CP/CP/<profile号>-<序号>/A.*`，插入表格时等比缩放至高度 3.5cm。

## 可靠性条件

定义在 `parameter.py`：

```python
RELIABILITY_CONDITIONS = ["T0", "168 TCT", "168 UDH", "500 TCT", "500 UDH"]
RELIABILITY_MAP = {
    "T0":      ("可靠性-T0",   "T0"),
    "168 TCT": ("可靠性-T168", "TCT168"),
    "168 UDH": ("可靠性-T168", "UDH168"),
    "500 TCT": ("可靠性-T500", "TCT500"),
    "500 UDH": ("可靠性-T500", "UDH500"),
}
```

## 常用配置（parameter.py）

| 项 | 说明 |
|---|---|
| `TEMPLATE_PAGE_INDEX` / `EXPECTED_TEMPLATE_SLIDES` | 模板页序（0-based）与最少页数 |
| `AUTHOR_NAME` | 封面作者名（原先硬编码在 `fill_pptx` 里） |
| `FONT_EA_PPT` / `FONT_LATIN_PPT` / `FONT_SIZE_PPT_BODY` | PPT 中西文字体与正文字号 |
| `FONT_EA_PPT_DOE_TABLE` | 第 4 页 DOE 表字体（见下方"已知不一致"） |
| `FIGURE_SIZE` / `RELIABILITY_FIG_SIZE` | 图尺寸（英寸） |
| `SPEC_DEFAULT` / `RELIABILITY_SPEC_DEFAULT` | SPEC 默认值 |
| `SpecConfig` / `build_specs()` / `parse_position_bound()` | SPEC 容器、构造函数、`±0.02` 解析 |
| `POSITION_*` / `RELIABILITY_*` / `CP_*` | PPT 版面尺寸与结论文本框坐标（cm） |
| `DIM_COLORS` / `PROFILE_COLORS` | X/Y 维度色、profile 主色 |
| `FONT_CHINESE` / `FONT_ENGLISH` | 图表中文字体 `Microsoft YaHei`，西文 `Times New Roman` |
| `RELIABILITY_COND_GROUPS` | 可靠性条件分组展示 |
| `RNG_SEED` | 随机数种子（jitter 等），保证多次生成图片一致 |

SPEC 通过 `build_specs()` 从 `DOE_data.xlsx` 的物料信息 SPEC 表构建为
`SpecConfig` 实例，再由 `plot_single_profile(...)` / `plot_device(...)`
显式接收，缺失时回落到默认值。

## 已知不一致 / 待办

- **第 4 页 DOE 参数表用"宋体"，其余各页用"微软雅黑"。** 这是历史遗留，
  当前重构**刻意保持原样**以免改动视觉输出。如需统一，把
  `parameter.FONT_EA_PPT_DOE_TABLE` 改成 `FONT_EA_PPT` 即可。
- 第 5/6 页模板上只有标题文本框，结论文本框是运行时新建的
  （坐标见 `POSITION_CONCL_NEW_BOX` / `RELIABILITY_CONCL_*`）。
- 若某个数据源整体缺失（位置精度无 profile / 无 `xxxCP` 文件夹），
  对应的模板页会被**删除**而不是留一张空白页。

## 注意事项

- **生成前会先删掉旧报告再重写**，所以报告还在 PowerPoint 里开着就写不进去。
  这种情况下会自动回退保存到 `%TEMP%` 下的同名文件，**但原文件不会更新**——
  脚本会明确警告，别把它当成"生成失败"或拿旧稿去对版式。
- 生成成功后会自动打开报告（`os.startfile`，仅 Windows）；
  批量运行时建议加 `--no-open`。
- `run.bat` 会依次试探 `py -3` / `python` / `python3` / 常见安装路径，
  **以能否 `import pptx,matplotlib,pandas,openpyxl` 为准**挑选解释器，
  而不是看谁在 PATH 上排前面（PATH 第一名可能是个没装依赖的空环境）。
- `GT工艺验证模板.pptx` 必须至少 8 页，否则直接报错。改动模板页序需同步
  `parameter.TEMPLATE_PAGE_INDEX`。
- `data_reader.main()` 只 `load_workbook` 一次，各 sheet 复用同一个 workbook。
- `~$*.pptx` 为 Office 临时锁文件，可忽略/删除。
