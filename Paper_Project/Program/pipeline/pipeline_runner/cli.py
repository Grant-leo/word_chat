"""CLI parser and dispatch helpers for the one-click pipeline runner."""
from __future__ import annotations

import argparse
import os

from .io import choose_file, choose_mode, exit_from_result, normalize_mode, scan_inputs
from .summary import write_agent_preflight_report


DEFAULT_GOLDEN_DIR = os.path.join("TestData", "GoldenBaselines")


def build_arg_parser(default_golden_dir=None):
    parser = argparse.ArgumentParser(description="Word 论文排版流水线")
    parser.add_argument("--template", "-t", help="模板文件名，位于 Templates/，支持 .docx/.pdf")
    parser.add_argument("--content", "-c", help="内容文件名，位于 Inputs/，支持 .docx/.md")
    parser.add_argument("--md", help="单个 MD 文件：格式+内容合一的纯 MD 模式")
    parser.add_argument(
        "--agent-auto",
        action="store_true",
        help="Agent 自动入口：非交互扫描文件，唯一候选直接运行，普通用户模式会开启自动修复并写出 agent_summary.md/json。",
    )
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
        help="visual QA 的黄金基线 JSON 目录；不传则不启用 golden 对比，--update-golden 未指定目录时使用 TestData/GoldenBaselines",
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
    parser.add_argument(
        "--auto-repair",
        action="store_true",
        help="普通用户自动修复闭环：只允许修改 Outputs/<本轮>/build_generated.py，并在每轮后重建与重跑 QA。",
    )
    parser.add_argument("--repair-max-rounds", type=int, default=5, help="自动修复最多运行轮数。")
    parser.add_argument(
        "--repair-stop-no-improve",
        type=int,
        default=2,
        help="连续多少轮没有减少 QA error 后停止自动修复。",
    )
    return parser


def _visual_golden_dir(args):
    if args.golden_dir:
        return args.golden_dir
    if args.update_golden:
        return DEFAULT_GOLDEN_DIR
    return None


def _agent_outputs_dir(template_dir):
    return os.path.join(os.path.dirname(os.path.abspath(template_dir)), "Outputs")


def _rel_for_terminal(path):
    try:
        return os.path.relpath(path, os.getcwd()).replace(os.sep, "/")
    except ValueError:
        return str(path).replace(os.sep, "/")


def _write_agent_preflight(preflight_dir, *, status, message, next_steps, candidates=None):
    if not preflight_dir:
        return
    _json_path, md_path = write_agent_preflight_report(
        preflight_dir,
        status=status,
        message=message,
        next_steps=next_steps,
        candidates=candidates,
    )
    print(f"  [NEXT] 已写入预检报告: {_rel_for_terminal(md_path)}")


def _agent_selection_role(label):
    if label == "模板":
        return "模板"
    if label == "内容":
        return "内容"
    if label == "纯 Markdown 文件":
        return "纯 Markdown 输入"
    return label


def _agent_selection_next_steps(label, folder_label, files=None):
    role = _agent_selection_role(label)
    if label == "模板":
        steps = [
            "请告诉 Agent 要使用哪一个模板文件名。",
            f"然后让 Agent 重跑自动入口，并明确说“使用 {folder_label}/文件名 作为模板”。",
        ]
    elif label == "内容":
        steps = [
            "请告诉 Agent 要使用哪一个内容文件名。",
            f"然后让 Agent 重跑自动入口，并明确说“使用 {folder_label}/文件名 作为内容”。",
        ]
    elif label == "纯 Markdown 文件":
        steps = [
            "请告诉 Agent 要使用哪一个 Markdown 文件名。",
            f"然后让 Agent 重跑自动入口，并明确说“使用 {folder_label}/文件名 作为纯 Markdown 输入”。",
        ]
    else:
        steps = [
            f"请告诉 Agent 要使用哪一个 {label} 文件名。",
            "然后让 Agent 重跑自动入口。",
        ]
    for filename in files or []:
        steps.append(f"可以直接回复：使用 {folder_label}/{filename} 作为{role}。")
    return steps


def _print_agent_candidates(label, folder_label, files):
    print(f"\n[AGENT-AUTO] {label}存在多个候选，不能替用户盲选。请让用户指定一个文件名后重跑：")
    for filename in files:
        print(f"  - {folder_label}/{filename}")


def _agent_select_single(files, *, label, folder_label, preflight_dir=None):
    if len(files) == 1:
        selected = files[0]
        print(f"[AGENT-AUTO] {label}: {folder_label}/{selected}")
        return selected
    if len(files) > 1:
        _print_agent_candidates(label, folder_label, files)
        _write_agent_preflight(
            preflight_dir,
            status="blocked_ambiguous_input",
            message=f"{label}存在多个候选，Agent 不能替用户盲选。",
            next_steps=_agent_selection_next_steps(label, folder_label, files),
            candidates={folder_label: files},
        )
        raise SystemExit(2)
    return None


def _agent_missing_inputs_message(preflight_dir=None, *, reason="没有找到可运行的模板和内容。"):
    print(f"\n[AGENT-AUTO] {reason}")
    print("  请把模板 DOCX/PDF 放入 Templates/，把内容 DOCX/Markdown 放入 Inputs/ 后重跑。")
    print("  若是纯 Markdown 模式，请在 Inputs/ 放入一个同时包含格式说明和正文的 .md 文件。")
    _write_agent_preflight(
        preflight_dir,
        status="blocked_missing_input",
        message=reason,
        next_steps=[
            "把模板 DOCX/PDF 放入 Templates/，把内容 DOCX 或 Markdown 放入 Inputs/。",
            "如果是纯 Markdown 模式，只放一个同时包含格式说明和正文的 .md 文件到 Inputs/。",
            "放好后告诉 Agent 重新运行自动入口。",
        ],
    )


def _dispatch_agent_auto(args, *, run_pipeline, template_dir, inputs_dir, mode, run_qa, golden_dir, auto_repair, preflight_dir):
    if args.md:
        exit_from_result(
            run_pipeline(
                None,
                None,
                md_file=args.md,
                mode=mode,
                run_qa=run_qa,
                qa_level=args.qa_level,
                golden_dir=golden_dir,
                update_golden=args.update_golden,
                require_wps=args.require_wps,
                auto_repair=auto_repair,
                agent_auto=True,
                repair_max_rounds=args.repair_max_rounds,
                repair_stop_no_improve=args.repair_stop_no_improve,
            )
        )

    templates = scan_inputs(template_dir, exts=(".docx", ".pdf"))
    contents = scan_inputs(inputs_dir, exts=(".docx", ".md"))
    md_files = scan_inputs(inputs_dir, exts=(".md",))

    template_file = args.template
    content_file = args.content
    if not template_file and content_file and str(content_file).lower().endswith(".md") and not templates:
        exit_from_result(
            run_pipeline(
                None,
                None,
                md_file=content_file,
                mode=mode,
                run_qa=run_qa,
                qa_level=args.qa_level,
                golden_dir=golden_dir,
                update_golden=args.update_golden,
                require_wps=args.require_wps,
                auto_repair=auto_repair,
                agent_auto=True,
                repair_max_rounds=args.repair_max_rounds,
                repair_stop_no_improve=args.repair_stop_no_improve,
            )
        )

    if not template_file and not templates:
        md_file = _agent_select_single(md_files, label="纯 Markdown 文件", folder_label="Inputs", preflight_dir=preflight_dir)
        if md_file:
            exit_from_result(
                run_pipeline(
                    None,
                    None,
                    md_file=md_file,
                    mode=mode,
                    run_qa=run_qa,
                    qa_level=args.qa_level,
                    golden_dir=golden_dir,
                    update_golden=args.update_golden,
                    require_wps=args.require_wps,
                    auto_repair=auto_repair,
                    agent_auto=True,
                    repair_max_rounds=args.repair_max_rounds,
                    repair_stop_no_improve=args.repair_stop_no_improve,
                )
            )
        _agent_missing_inputs_message(preflight_dir)
        raise SystemExit(1)

    if not content_file and not contents:
        _agent_missing_inputs_message(preflight_dir, reason="没有找到内容文件。")
        raise SystemExit(1)

    if not template_file:
        template_file = _agent_select_single(templates, label="模板", folder_label="Templates", preflight_dir=preflight_dir)
    if not content_file:
        content_file = _agent_select_single(contents, label="内容", folder_label="Inputs", preflight_dir=preflight_dir)

    exit_from_result(
        run_pipeline(
            template_file,
            content_file,
            mode=mode,
            run_qa=run_qa,
            qa_level=args.qa_level,
            golden_dir=golden_dir,
            update_golden=args.update_golden,
            require_wps=args.require_wps,
            auto_repair=auto_repair,
            agent_auto=True,
            repair_max_rounds=args.repair_max_rounds,
            repair_stop_no_improve=args.repair_stop_no_improve,
        )
    )


def print_cli_banner():
    print("=" * 50)
    print("  Word 论文排版流水线")
    print("=" * 50)


def dispatch_cli(args, *, run_pipeline, template_dir, inputs_dir):
    print_cli_banner()

    interactive = not args.agent_auto and not args.md and not (args.template and args.content)
    mode = choose_mode() if args.mode == "auto" and interactive else normalize_mode("user" if args.mode == "auto" else args.mode)
    auto_repair = bool(args.auto_repair or (args.agent_auto and mode == "user"))
    run_qa = True if args.agent_auto or auto_repair else not args.no_qa
    golden_dir = _visual_golden_dir(args)
    preflight_dir = _agent_outputs_dir(template_dir)

    if args.agent_auto:
        _dispatch_agent_auto(
            args,
            run_pipeline=run_pipeline,
            template_dir=template_dir,
            inputs_dir=inputs_dir,
            mode=mode,
            run_qa=run_qa,
            golden_dir=golden_dir,
            auto_repair=auto_repair,
            preflight_dir=preflight_dir,
        )

    if args.md:
        exit_from_result(
            run_pipeline(
                None,
                None,
                md_file=args.md,
                mode=mode,
                run_qa=run_qa,
                qa_level=args.qa_level,
                golden_dir=golden_dir,
                update_golden=args.update_golden,
                require_wps=args.require_wps,
                auto_repair=auto_repair,
                agent_auto=False,
                repair_max_rounds=args.repair_max_rounds,
                repair_stop_no_improve=args.repair_stop_no_improve,
            )
        )

    if args.template and args.content:
        template_file = args.template
        content_file = args.content
    else:
        template_file = args.template
        content_file = args.content
        templates = scan_inputs(template_dir, exts=(".docx", ".pdf"))
        contents = scan_inputs(inputs_dir, exts=(".docx", ".md"))

        if not template_file and not templates:
            print("\n[INFO] Templates/ 下没有 .docx 或 .pdf 模板文件。")
            print("  纯 MD 模式请用: python run_pipeline.py --md <文件名>")
            print("  下一步：把模板放入 Templates/，或把格式+正文合一的 Markdown 放入 Inputs/ 后运行 --md。")
            md_files = scan_inputs(inputs_dir, exts=(".md",))
            if md_files:
                print("\n  Inputs/ 下找到 .md 文件，可直接纯 MD 模式:")
                for filename in md_files:
                    print(f"    python run_pipeline.py --md {filename}")
            raise SystemExit(1)
        if not content_file and not contents:
            print("\n[ERROR] Inputs/ 下没有 .docx 或 .md 内容文件，请放入内容文件后重试。")
            print("  下一步：把正文 DOCX/Markdown 放入 Inputs/，然后运行 python run_pipeline.py --agent-auto。")
            raise SystemExit(1)

        if not template_file:
            template_file = choose_file(templates, "选择模板")
        if not content_file:
            content_file = choose_file(contents, "选择内容")

    exit_from_result(
        run_pipeline(
            template_file,
            content_file,
            mode=mode,
            run_qa=run_qa,
            qa_level=args.qa_level,
            golden_dir=golden_dir,
            update_golden=args.update_golden,
            require_wps=args.require_wps,
            auto_repair=auto_repair,
            agent_auto=False,
            repair_max_rounds=args.repair_max_rounds,
            repair_stop_no_improve=args.repair_stop_no_improve,
        )
    )


def main_cli(*, run_pipeline, template_dir, inputs_dir, argv=None):
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    dispatch_cli(args, run_pipeline=run_pipeline, template_dir=template_dir, inputs_dir=inputs_dir)
