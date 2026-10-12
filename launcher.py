# -*- coding: utf-8 -*-
"""拖拽 / 批量启动器（run.bat 的后端）。

与 cli.py 的分工
----------------
  cli.py       "显式参数"接口 —— 路径一条条写清楚，给 GUI / 脚本 / 自动化用
  launcher.py  "零配置推断"接口 —— 只给一个入口，路径全自动推断，给人拖拽用

两者最终都调 report_generator.build_report()，不存在第二套生成逻辑。

推断规则
--------
  拖入对象               推断结果
  ---------------------  ----------------------------------------------------
  *.xlsx / *.xlsm        数据文件；其所在目录同时作为切片图根目录与输出目录
  文件夹                 在里面找 DOE_data.xlsx（找不到就取唯一的 xlsx）
  *.pptx 且名字含"模板"   作为 PPT 模板
  多个对象               批量，逐个生成（默认不自动打开，免得弹一屏窗口）

  模板查找顺序：拖入的 > 数据同目录的 GT工艺验证模板.pptx
                > 数据同目录里唯一的 "*模板*.pptx" > 项目目录自带的模板
  输出路径：<输出目录>/{project_name}_GT工艺验证报告.pptx

已知限制（重要）
----------------
Windows 资源管理器把文件拖到 .bat 上时，含空格的路径**可能不带引号**，
会被 cmd 按空格拆成好几段。本文件用 regroup() 把被拆散的片段重新拼回真实路径。

但路径里含 `&`（例如历史目录 macom70&285-0814）且不带引号时，
cmd 会在 .bat 那一层就把命令行截断，Python 根本收不到完整路径——这个救不了。
遇到这种情况改用带引号的命令行方式：

    python launcher.py "D:\\工作\\进行中\\macom70&285-0814"
"""
from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

# 拖拽场景只要出图文件，不要交互窗口；显式切到无窗口后端，
# 避免在受限环境里因为没有显示设备而报错。
try:
    import matplotlib
    matplotlib.use("Agg")
except ImportError:
    pass

import parameter as P
import report_generator

XLSX_EXT = (".xlsx", ".xlsm", ".xls")
TEMPLATE_HINTS = ("模板", "template")
BUNDLED_TEMPLATE_NAME = "GT工艺验证模板.pptx"


# ============================================================
# 拖拽参数修复
# ============================================================
def regroup(raw):
    """把被空格拆散的拖拽路径重新拼回完整路径。

    Explorer 拖拽时若未给含空格的路径加引号，sys.argv 会是
      ['D:\\\\工作\\\\进行中\\\\macom 70mW', 'DOE\\\\macom70&285-0814']
    而不是一个整体。逐段贪心：从当前位置起向右不断拼接，
    直到拼出一个真实存在的路径为止；拼不出来就原样保留（后面会报"不存在"）。
    """
    out = []
    i, n = 0, len(raw)
    while i < n:
        if os.path.exists(raw[i]):
            out.append(raw[i])
            i += 1
            continue
        merged, hit, j = raw[i], None, i + 1
        while j < n:
            merged = merged + " " + raw[j]
            if os.path.exists(merged):
                hit = merged
                break
            j += 1
        if hit is None:
            out.append(raw[i])
            i += 1
        else:
            out.append(hit)
            i = j + 1
    return out


# ============================================================
# 路径推断
# ============================================================
def find_data_file(folder):
    """在文件夹里找数据 Excel：优先 DOE_data.xlsx，否则取唯一的 xlsx。"""
    preferred = os.path.join(folder, "DOE_data.xlsx")
    if os.path.isfile(preferred):
        return preferred

    candidates = []
    try:
        names = sorted(os.listdir(folder))
    except OSError:
        return None
    for name in names:
        # ~$xxx.xlsx 是 Excel 打开文件时生成的锁文件
        if name.startswith("~$") or not name.lower().endswith(XLSX_EXT):
            continue
        candidates.append(os.path.join(folder, name))

    if not candidates:
        return None
    if len(candidates) > 1:
        print(f"[提示] {folder} 下有 {len(candidates)} 个 Excel，"
              f"取第一个：{os.path.basename(candidates[0])}")
        print("       （想指定就用命令行：python launcher.py \"具体的xlsx路径\"）")
    return candidates[0]


def find_template(base_dir, override=None):
    """按 拖入的 > 同目录标准名 > 同目录唯一含"模板" > 项目自带 的顺序找模板。"""
    if override:
        return override

    local = os.path.join(base_dir, BUNDLED_TEMPLATE_NAME)
    if os.path.isfile(local):
        return local

    try:
        hits = [
            os.path.join(base_dir, name)
            for name in sorted(os.listdir(base_dir))
            if name.lower().endswith(".pptx")
            and any(h in name.lower() for h in TEMPLATE_HINTS)
        ]
    except OSError:
        hits = []
    if len(hits) == 1:
        return hits[0]

    bundled = os.path.join(BASE_DIR, BUNDLED_TEMPLATE_NAME)
    return bundled if os.path.isfile(bundled) else None


@dataclass
class Job:
    """一次生成任务：一个数据文件 + 全套路径。"""
    source: str          # 拖入的原始对象，仅用于日志
    data_file: str
    template_path: str
    output_path: str     # 含 {project_name} 占位符，由 build_report 填充
    images_dir: str
    cp_dir: str


def build_jobs(paths, template_override=None, images_dir=None, out_dir=None):
    """把拖入的一堆路径翻译成生成任务清单。"""
    xlsx_files, folders, templates = [], [], []

    for raw in paths:
        path = os.path.abspath(raw)
        if os.path.isdir(path):
            folders.append(path)
        elif os.path.isfile(path):
            lower = path.lower()
            if lower.endswith(XLSX_EXT):
                xlsx_files.append(path)
            elif lower.endswith(".pptx"):
                if any(h in os.path.basename(path).lower() for h in TEMPLATE_HINTS):
                    templates.append(path)
                else:
                    print(f"[忽略] 这是报告不是模板，已跳过：{path}")
            else:
                print(f"[忽略] 不支持的文件类型：{path}")
        else:
            print(f"[错误] 路径不存在：{raw}")
            if "&" in raw:
                print("       路径含 & 时，拖拽可能被 cmd 截断，"
                      "请改用带引号的命令行方式")

    tpl_override = template_override or (templates[0] if templates else None)

    jobs, seen = [], set()

    def add(data_file, cp_dir, origin):
        key = os.path.normcase(os.path.abspath(data_file))
        if key in seen:
            return
        seen.add(key)

        base = os.path.dirname(data_file)
        template = find_template(base, tpl_override)
        if not template:
            print(f"[错误] 找不到 PPT 模板。请把它命名为 {BUNDLED_TEMPLATE_NAME} "
                  f"放在 {base} 下，或用 --template 指定")
            return

        jobs.append(Job(
            source=origin,
            data_file=data_file,
            template_path=template,
            output_path=os.path.join(out_dir or base,
                                     "{project_name}_GT工艺验证报告.pptx"),
            images_dir=images_dir or os.path.join(base, "images"),
            cp_dir=cp_dir,
        ))

    # 直接拖入的 xlsx 各自成任务（cp-dir 取数据同目录）
    for data_file in xlsx_files:
        add(data_file, os.path.dirname(data_file), data_file)

    # 拖入的文件夹：找到数据文件再成任务（cp-dir 取该文件夹）
    for folder in folders:
        data_file = find_data_file(folder)
        if data_file:
            add(data_file, folder, folder)
        else:
            print(f"[错误] {folder} 下没找到 Excel 数据文件")

    return jobs


# ============================================================
# 执行
# ============================================================
def run_job(job, open_after, author=None, report_date=None):
    print("=" * 66)
    print(f"数据:   {job.data_file}")
    print(f"模板:   {job.template_path}")
    print(f"切片图: {job.cp_dir}")
    print(f"图表:   {job.images_dir}")
    print(f"输出:   {_pretty(job.output_path)}")
    print("=" * 66)
    return report_generator.build_report(
        data_file=job.data_file,
        template_path=job.template_path,
        output_path=job.output_path,
        images_dir=job.images_dir,
        cp_base_dir=job.cp_dir,
        author=author,
        report_date=report_date,
        open_after=open_after,
    )


def _pretty(path):
    """把 {project_name} 占位符显示成人话，用于 --dry-run。"""
    return path.replace("{project_name}", "<Excel里的项目名>")


def pick_data_file_interactively():
    """双击（没拖东西）时弹选择框挑数据文件。

    这条路的价值不只是方便：路径是 Python 自己取到的，**完全不经过
    cmd 的命令行解析**，所以对 `&`、`!` 之类的字符天然免疫——
    正好补上"拖拽 .bat 遇到特殊字符就会被截断"那个洞。

    返回所选路径；用户取消或环境里没有 GUI 支持时返回 None。
    """
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        return None

    try:
        root = tk.Tk()
    except Exception:
        return None

    root.withdraw()
    root.update()
    path = filedialog.askopenfilename(
        title="选择 DOE 数据文件（DOE_data.xlsx）",
        filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
    )
    root.destroy()
    return path or None


def print_help():
    print("=" * 66)
    print(" DOE 工艺验证报告生成器 —— 拖拽启动器")
    print("=" * 66)
    print(" 用法：用鼠标把下面任意一种拖到 run.bat 的图标上")
    print()
    print("   1) DOE_data.xlsx      推荐。自动在它所在目录找切片图与模板")
    print("   2) 项目文件夹         自动找里面的 DOE_data.xlsx")
    print("   3) 多个数据 / 文件夹   批量生成（默认不逐个弹窗打开）")
    print()
    print(" 命令行等价用法：")
    print('   python launcher.py "D:\\项目\\DOE_data.xlsx"')
    print('   python launcher.py "D:\\项目A" "D:\\项目B" --no-open')
    print('   python launcher.py "D:\\项目" --dry-run     只看会怎么跑')
    print()
    print(" 切片图目录：数据所在目录下应有 <器件>CP 子文件夹")
    print("=" * 66)


def build_parser():
    parser = argparse.ArgumentParser(
        prog="launcher.py",
        description="拖拽 / 批量启动器：自动推断路径后生成报告",
        add_help=True,
    )
    parser.add_argument("paths", nargs="*",
                        help="拖入的文件或文件夹（可多个）")
    parser.add_argument("--template", "-t", metavar="PPTX",
                        help="显式指定 PPT 模板")
    parser.add_argument("--images-dir", metavar="DIR",
                        help="图表输出目录（缺省：数据同目录下 images）")
    parser.add_argument("--out-dir", metavar="DIR",
                        help="报告输出目录（缺省：数据同目录）")
    parser.add_argument("--author", metavar="NAME", help="封面作者名")
    parser.add_argument("--date", metavar="YYYY-MM-DD", help="封面日期")
    parser.add_argument("--no-open", dest="open_after", action="store_false",
                        default=None, help="生成后不打开报告")
    parser.add_argument("--open", dest="open_after", action="store_true",
                        help="强制生成后打开（批量时默认不打开）")
    parser.add_argument("--dry-run", action="store_true",
                        help="只打印将要执行的任务，不真正生成")
    parser.add_argument("--pick", action="store_true",
                        help="弹选择框挑数据文件（双击运行时默认就会弹）")
    parser.add_argument("--debug", action="store_true",
                        help="出错时打印完整堆栈")
    return parser


def main(argv=None):
    # 控制台编码受限时（比如项目名里有 GBK 表示不了的字符），
    # 宁可显示成 ? 也不要在绘图跑完之后因为 print 失败而整个任务中断。
    try:
        sys.stdout.reconfigure(errors="replace")
        sys.stderr.reconfigure(errors="replace")
    except Exception:
        pass

    argv = regroup(list(sys.argv[1:] if argv is None else argv))
    parser = build_parser()
    # 拖拽时 Explorer 可能塞进无关参数，宽容处理：忽略不认识的
    args, unknown = parser.parse_known_args(argv)
    if unknown:
        print(f"[忽略] 无法识别的参数：{' '.join(unknown)}")

    if not args.paths:
        default_data = P.DATA_FILE
        if args.pick or not os.path.isfile(default_data):
            # 双击（或 --pick）：弹框选文件。这条路绕开了命令行解析，
            # 路径里有什么特殊字符都不怕。
            picked = pick_data_file_interactively()
            if not picked:
                print_help()
                if not os.path.isfile(default_data):
                    print(f"\n[提示] parameter.py 里的默认数据也不存在：{default_data}")
                    return 2
                return 0
            args.paths = [picked]
        else:
            print("未拖入任何文件，按 parameter.py 的默认配置运行。")
            print("（想换项目：把 DOE_data.xlsx 拖到 run.bat 上，或双击后选文件）")

    jobs = build_jobs(args.paths,
                      template_override=args.template,
                      images_dir=args.images_dir,
                      out_dir=args.out_dir)

    if not jobs:
        print("\n[错误] 没有可执行的任务。")
        print_help()
        return 2

    # 单个任务默认自动打开；批量默认不打开（免得弹一屏）
    open_after = args.open_after
    if open_after is None:
        open_after = (len(jobs) == 1)

    if args.dry_run:
        print(f"\n将要执行 {len(jobs)} 个任务（--dry-run，不实际生成）：\n")
        for i, job in enumerate(jobs, 1):
            print(f"[{i}/{len(jobs)}] {job.source}")
            print(f"      数据:   {job.data_file}")
            print(f"      模板:   {job.template_path}")
            print(f"      切片图: {job.cp_dir}")
            print(f"      图表:   {job.images_dir}")
            print(f"      输出:   {_pretty(job.output_path)}")
        print(f"\n生成后自动打开: {'是' if open_after else '否'}")
        return 0

    succeeded, failed = [], []
    print(f"\n共 {len(jobs)} 个任务，开始生成。"
          f"（生成后自动打开：{'是' if open_after else '否'}）")

    for i, job in enumerate(jobs, 1):
        print(f"\n\n########## [{i}/{len(jobs)}] {job.source} ##########")
        try:
            saved = run_job(job, open_after=open_after,
                            author=args.author, report_date=args.date)
            succeeded.append(saved)
        except KeyboardInterrupt:
            print("\n[中断] 用户取消")
            return 130
        except Exception as exc:            # noqa: BLE001 —— 批量不能因单个失败而全停
            failed.append((job.source, exc))
            print(f"\n[失败] {job.source}\n       {type(exc).__name__}: {exc}")
            if args.debug:
                import traceback
                traceback.print_exc()

    # ---------------- 汇总 ----------------
    print("\n" + "=" * 66)
    print(f"完成：成功 {len(succeeded)} 个，失败 {len(failed)} 个")
    for path in succeeded:
        print(f"  [OK]   {path}")
    for source, exc in failed:
        print(f"  [FAIL] {source}  ->  {exc}")
    print("=" * 66)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
