# -*- coding: utf-8 -*-
"""命令行入口。

用法示例：

    # 用 parameter.py 里的配置（和以前一样）
    python cli.py

    # 指定某个项目的路径，不改源码
    python cli.py --data "D:\\项目\\DOE_data.xlsx" \\
                  --template "GT工艺验证模板.pptx" \\
                  --output "D:\\项目\\out\\报告.pptx" \\
                  --cp-dir "D:\\项目" \\
                  --no-open

    # 只看当前生效配置
    python cli.py --print-config

这是后续 GUI / 批量模式的统一后端入口：GUI 只需拼出同样的参数，
再调用 report_generator.build_report()。
"""
import argparse
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

import parameter as P
import report_generator


def build_parser():
    parser = argparse.ArgumentParser(
        prog="cli.py",
        description="DOE 工艺验证报告自动生成器：读 DOE_data.xlsx -> 出图 -> 填模板 -> 出报告",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="所有路径参数缺省时回落到 parameter.py 中的配置。",
    )
    parser.add_argument("--data", "-d", metavar="XLSX",
                        help="输入数据 Excel（缺省：parameter.DATA_FILE）")
    parser.add_argument("--template", "-t", metavar="PPTX",
                        help="PPT 模板（缺省：parameter.TEMPLATE_PATH）")
    parser.add_argument("--output", "-o", metavar="PPTX",
                        help="输出报告路径（缺省：按 Excel 项目名自动命名）")
    parser.add_argument("--images-dir", metavar="DIR",
                        help="图表输出目录（缺省：parameter.IMAGES_DIR）")
    parser.add_argument("--cp-dir", metavar="DIR",
                        help="切片图片根目录，下面放 <器件>CP 文件夹")
    parser.add_argument("--author", metavar="NAME",
                        help=f"封面作者名（缺省：{P.AUTHOR_NAME}）")
    parser.add_argument("--date", metavar="YYYY-MM-DD",
                        help="封面日期（缺省：今天）")
    parser.add_argument("--no-open", action="store_true",
                        help="生成后不自动打开报告（批量时建议加上）")
    parser.add_argument("--print-config", action="store_true",
                        help="只打印当前生效配置后退出，不生成报告")
    return parser


def _effective_config(args):
    return P.configure(
        data_file=args.data,
        template_path=args.template,
        output_path=args.output,
        images_dir=args.images_dir,
        cp_base_dir=args.cp_dir,
    )


def main(argv=None):
    args = build_parser().parse_args(argv)

    if args.print_config:
        for key, value in _effective_config(args).items():
            print(f"{key:18}{value}")
        return 0

    try:
        saved = report_generator.build_report(
            data_file=args.data,
            template_path=args.template,
            output_path=args.output,
            images_dir=args.images_dir,
            cp_base_dir=args.cp_dir,
            author=args.author,
            report_date=args.date,
            open_after=not args.no_open,
        )
    except KeyboardInterrupt:
        print("\n[中断] 用户取消", file=sys.stderr)
        return 130
    except (FileNotFoundError, ValueError, PermissionError) as e:
        print(f"\n[ERROR] {e}", file=sys.stderr)
        return 1

    print(f"\n输出: {saved}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
