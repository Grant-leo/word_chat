"""CLI parser and dispatch helpers for the one-click pipeline runner."""
from __future__ import annotations

import argparse
import os

from .io import choose_file, choose_mode, exit_from_result, normalize_mode, scan_inputs


def build_arg_parser(default_golden_dir=os.path.join("TestData", "GoldenBaselines")):
    parser = argparse.ArgumentParser(description="Word 论文排版流水线")
    parser.add_argument("--template", "-t", help="模板文件名，位于 Templates/，支持 .docx/.pdf")
    parser.add_argument("--content", "-c", help="内容文件名，位于 Inputs/，支持 .docx/.md")
    parser.add_argument("--md", help="单个 MD 文件：格式+内容合一的纯 MD 模式")
    parser.add_argument(
        "--mode",
        choices=["auto", "user", "developer"],
        default="auto",
        help="工作模式：user 只改 build_generated.py；developer 只改核心引擎；auto 交互时询问，参数模式默认 user",
    )
    parser.add_argument(
        "--qa-level",
        choices=["basic", "strict", "visual"],
        default="strict",
        help="QA 级别：basic=结构检查；strict=结构+DOCX合规；visual=再加 PDF/全页PNG/黄金基线",
    )
    parser.add_argument(
        "--golden-dir",
        default=default_golden_dir,
        help="visual QA 的黄金基线 JSON 目录，默认 TestData/GoldenBaselines",
    )
    parser.add_argument(
        "--update-golden",
        action="store_true",
        help="visual QA 时创建或更新当前模板+内容的黄金基线",
    )
    parser.add_argument(
        "--require-wps",
        action="store_true",
        help="visual QA 时如果 WPS 导出不可用则记为 error",
    )
    parser.add_argument("--no-qa", action="store_true", help="跳过生成后的 QA 检测")
    return parser


def print_cli_banner():
    print("=" * 50)
    print("  Word 论文排版流水线")
    print("=" * 50)


def dispatch_cli(args, *, run_pipeline, template_dir, inputs_dir):
    print_cli_banner()

    interactive = not args.md and not (args.template and args.content)
    mode = choose_mode() if args.mode == "auto" and interactive else normalize_mode("user" if args.mode == "auto" else args.mode)

    if args.md:
        exit_from_result(
            run_pipeline(
                None,
                None,
                md_file=args.md,
                mode=mode,
                run_qa=not args.no_qa,
                qa_level=args.qa_level,
                golden_dir=args.golden_dir,
                update_golden=args.update_golden,
                require_wps=args.require_wps,
            )
        )

    if args.template and args.content:
        template_file = args.template
        content_file = args.content
    else:
        templates = scan_inputs(template_dir, exts=(".docx", ".pdf"))
        contents = scan_inputs(inputs_dir, exts=(".docx", ".md"))

        if not templates:
            print("\n[INFO] Templates/ 下没有 .docx 或 .pdf 模板文件。")
            print("  纯 MD 模式请用: python run_pipeline.py --md <文件名>")
            md_files = scan_inputs(inputs_dir, exts=(".md",))
            if md_files:
                print("\n  Inputs/ 下找到 .md 文件，可直接纯 MD 模式:")
                for filename in md_files:
                    print(f"    python run_pipeline.py --md {filename}")
            raise SystemExit(1)
        if not contents:
            print("\n[ERROR] Inputs/ 下没有 .docx 或 .md 内容文件，请放入内容文件后重试。")
            raise SystemExit(1)

        template_file = choose_file(templates, "选择模板")
        content_file = choose_file(contents, "选择内容")

    exit_from_result(
        run_pipeline(
            template_file,
            content_file,
            mode=mode,
            run_qa=not args.no_qa,
            qa_level=args.qa_level,
            golden_dir=args.golden_dir,
            update_golden=args.update_golden,
            require_wps=args.require_wps,
        )
    )


def main_cli(*, run_pipeline, template_dir, inputs_dir, argv=None):
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    dispatch_cli(args, run_pipeline=run_pipeline, template_dir=template_dir, inputs_dir=inputs_dir)
