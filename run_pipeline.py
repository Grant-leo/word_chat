"""
run_pipeline.py —— 一键工作流入口
===================================

两种用法:
  Agent 自动入口:  python run_pipeline.py --agent-auto
             → 给 VSCode Agent 使用；自动扫描、单候选运行、普通用户自动修复、写出 agent_summary

  交互模式:  python run_pipeline.py
             → 自动扫描文件，编号选择

  参数模式:  python run_pipeline.py --template 模版.docx --content 论文.docx
             → 直接运行，无交互（Skill / 脚本调用）

  结果自动输出到 Outputs/{日期}_{内容名}/
  build_generated.py 是用户模式的微调入口；开发者模式的可复用修复请修改核心引擎脚本后重跑。
"""

import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.join(BASE, 'Paper_Project')
PIPELINE = os.path.join(PROJ, 'Program', 'pipeline')
TEMPLATE_DIR = os.path.join(BASE, 'Templates')
INPUTS_DIR = os.path.join(BASE, 'Inputs')
OUTPUTS_DIR = os.path.join(BASE, 'Outputs')

for d in [TEMPLATE_DIR, INPUTS_DIR, OUTPUTS_DIR]:
    os.makedirs(d, exist_ok=True)

# Prefer scripts placed beside this runner.  Fall back to the old project
# layout only when someone keeps the historical directory structure.
sys.path.insert(0, BASE)
if os.path.isdir(PIPELINE):
    sys.path.insert(1, PIPELINE)

from format_extractor import extract as extract_format
from content_parser import extract as extract_content
from script_generator import generate as generate_script
from pipeline_runner.artifacts import write_content_artifacts, write_extraction_failure_report, write_format_artifacts
from pipeline_runner.build_phase import generate_and_build_docx_phase
from pipeline_runner.cli import main_cli
from pipeline_runner.contracts import (
    validate_content_data,
    validate_format_data,
)
from pipeline_runner.context import (
    create_unique_output_dir,
    normalize_qa_level,
    resolve_inputs,
    write_workflow_mode,
)
from pipeline_runner.dependencies import load_optional_dependencies
from pipeline_runner.execution import run_generated_script
from pipeline_runner.io import normalize_mode
from pipeline_runner.qa import QADependencies, run_qa_phases
from pipeline_runner.repair_loop import run_repair_loop
from pipeline_runner.reports import print_contract_issues, step
from pipeline_runner.summary import print_completion_summary, write_agent_preflight_report, write_agent_summary
from pipeline_runner.template_phase import write_template_profile_phase, write_template_requirements_phase
from pipeline_runner.verification import VerificationError, double_verify

OPTIONAL_DEPS = load_optional_dependencies()


def _write_agent_preflight(status, message, next_steps):
    json_path, md_path = write_agent_preflight_report(
        OUTPUTS_DIR,
        status=status,
        message=message,
        next_steps=next_steps,
    )
    try:
        display = os.path.relpath(md_path, BASE).replace(os.sep, "/") if os.path.isabs(md_path) else md_path
    except ValueError:
        display = str(md_path).replace(os.sep, "/")
    print(f"  [NEXT] 已写入预检报告: {display}")
    return json_path, md_path


def run(
    template_file,
    content_file,
    md_file=None,
    mode='user',
    run_qa=True,
    qa_level='strict',
    golden_dir=None,
    update_golden=False,
    require_wps=False,
    auto_repair=False,
    agent_auto=False,
    repair_max_rounds=5,
    repair_stop_no_improve=2,
):
    """Core pipeline: takes filenames, runs all phases, returns output directory.

    If md_file is provided, uses MD file for BOTH format and content (single-file mode).
    Otherwise routes by file extension: .docx/.pdf use format/content extractors, .md uses md_parser.
    """
    mode = normalize_mode(mode)
    qa_level = normalize_qa_level(qa_level)
    auto_repair = bool(auto_repair)
    if auto_repair and mode != 'user':
        print('[ERROR] --auto-repair only supports --mode user because it may edit only Outputs/<run>/build_generated.py.')
        if agent_auto:
            _write_agent_preflight(
                "blocked_invalid_mode",
                "--auto-repair 只能在 user 模式下运行。",
                [
                    "请让 Agent 改用 user 模式重跑自动入口。",
                    "如果你确实要做开发者级修复，请关闭 --auto-repair，并明确要求开发者模式。",
                ],
            )
        return None
    if auto_repair:
        run_qa = True

    resolution = resolve_inputs(template_file, content_file, md_file, TEMPLATE_DIR, INPUTS_DIR)
    if not resolution.ok:
        print(resolution.error)
        if agent_auto:
            _write_agent_preflight(
                "blocked_input_resolution",
                resolution.error,
                [
                    "请确认文件名拼写和扩展名与 Templates/ 或 Inputs/ 中的实际文件一致。",
                    "如果不确定文件名，请让 Agent 使用 --agent-auto 自动扫描候选文件。",
                ],
            )
        return None
    inputs = resolution.inputs
    template_path = inputs.template_path
    content_path = inputs.content_path
    content_name = inputs.content_name
    use_md_format = inputs.use_md_format
    use_md_content = inputs.use_md_content
    if md_file and (OPTIONAL_DEPS.extract_md_format is None or OPTIONAL_DEPS.extract_md_content is None):
        message = f'[ERROR] md_parser.py 不可用，不能处理 Markdown 单文件模式。{OPTIONAL_DEPS.optional_import_detail("md_parser")}'
        print(message)
        if agent_auto:
            _write_agent_preflight(
                "blocked_missing_dependency",
                message,
                [
                    "请修复 md_parser.py 或相关依赖后重跑。",
                    "如果不是纯 Markdown 模式，请提供一个模板 DOCX/PDF 和一个内容 DOCX/Markdown 文件。",
                ],
            )
        return None
    if (use_md_format or use_md_content) and (OPTIONAL_DEPS.extract_md_format is None or OPTIONAL_DEPS.extract_md_content is None):
        message = f'[ERROR] md_parser.py 不可用，不能处理 Markdown 输入。{OPTIONAL_DEPS.optional_import_detail("md_parser")}'
        print(message)
        if agent_auto:
            _write_agent_preflight(
                "blocked_missing_dependency",
                message,
                [
                    "请修复 md_parser.py 或相关依赖后重跑。",
                    "如果急需继续，请把 Markdown 内容转成 DOCX 后放入 Inputs/。",
                ],
            )
        return None

    out_dir, folder_name = create_unique_output_dir(OUTPUTS_DIR, content_name)

    print(f'  输出目录: Outputs/{folder_name}/')
    print(f'  模版: {os.path.basename(template_path)}')
    print(f'  内容: {os.path.basename(content_path)}')
    print(f'  工作模式: {"普通用户" if mode == "user" else "开发者"}')

    write_workflow_mode(
        out_dir,
        mode=mode,
        template_path=template_path,
        content_path=content_path,
        run_qa=run_qa,
        qa_level=qa_level,
        golden_dir=golden_dir,
        update_golden=update_golden,
        require_wps=require_wps,
        auto_repair=auto_repair,
        agent_auto=agent_auto,
        repair_max_rounds=repair_max_rounds,
        repair_stop_no_improve=repair_stop_no_improve,
    )

    # ── Phase 1: Format ──
    step('Phase 1/6: 提取模版格式')
    fmt_extractor = OPTIONAL_DEPS.extract_md_format if use_md_format else extract_format
    try:
        fmt, md_text = double_verify(fmt_extractor, template_path, 'Format')
    except VerificationError as exc:
        write_extraction_failure_report(
            out_dir,
            mode=mode,
            label='Format',
            error=str(exc),
            target='format_extractor.py / md_parser.py',
        )
        write_agent_summary(out_dir, folder_name, "最终论文.docx", mode, pipeline_status="failed", note="模板格式提取多次验证无法收敛。")
        print('  [ERROR] 模板格式提取多次验证无法收敛，已生成 qa_report.md / qa_repair_plan.md。')
        return None
    print_contract_issues('format.json', validate_format_data(fmt))

    fmt_json_path, fmt_md_path = write_format_artifacts(fmt, md_text, out_dir)
    print(f'  段落:{len(fmt["paragraphs"])}  表格:{len(fmt["tables"])}  节:{len(fmt["sections"])}')

    # ── Phase 2: Template profile ──
    step('Phase 2/6: 生成模板画像')
    write_template_profile_phase(fmt, out_dir, project_root=BASE, write_template_profile=OPTIONAL_DEPS.write_template_profile)

    # ── Phase 2: Content ──
    step('Phase 3/6: 提取文本内容')
    cnt_extractor = OPTIONAL_DEPS.extract_md_content if use_md_content else extract_content
    try:
        content = double_verify(cnt_extractor, content_path, 'Content', output_dir=out_dir)
    except VerificationError as exc:
        write_extraction_failure_report(
            out_dir,
            mode=mode,
            label='Content',
            error=str(exc),
            target='content_parser.py / md_parser.py',
        )
        write_agent_summary(out_dir, folder_name, "最终论文.docx", mode, pipeline_status="failed", note="文本内容提取多次验证无法收敛。")
        print('  [ERROR] 文本内容提取多次验证无法收敛，已生成 qa_report.md / qa_repair_plan.md。')
        return None
    print_contract_issues('content.json', validate_content_data(content))

    cnt_json_path, cnt_md_path = write_content_artifacts(content, out_dir, content_path)

    print(f'  章节:{len(content["sections"])}  参考文献:{len(content["references"])}  图片:{content["_meta"]["images_extracted"]}')
    write_template_requirements_phase(
        fmt,
        content,
        out_dir,
        write_template_requirements=OPTIONAL_DEPS.write_template_requirements,
        optional_import_detail=OPTIONAL_DEPS.optional_import_detail,
    )

    output_docx = '最终论文.docx'
    if not generate_and_build_docx_phase(
        fmt_json_path,
        cnt_json_path,
        out_dir,
        folder_name,
        output_docx_name=output_docx,
        generate_script=generate_script,
        run_generated_script=run_generated_script,
        python_executable=sys.executable,
        step=step,
    ):
        write_agent_summary(out_dir, folder_name, output_docx, mode, pipeline_status="failed", note="最终 DOCX 构建失败，请查看 build_generated.py 的执行错误。")
        return None

    qa_deps = QADependencies(
        qa_check_and_write=OPTIONAL_DEPS.qa_check_and_write,
        conformance_check_and_write=OPTIONAL_DEPS.conformance_check_and_write,
        visual_check_and_write=OPTIONAL_DEPS.visual_check_and_write,
        optional_import_detail=OPTIONAL_DEPS.optional_import_detail,
    )

    # -- Phase 5: QA --
    if run_qa:
        step('Phase 6/6: QA 检测（发现 error 会阻断流水线）')
        qa_ok = run_qa_phases(
            out_dir,
            mode=mode,
            output_docx_name=output_docx,
            qa_level=qa_level,
            project_root=BASE,
            golden_dir=golden_dir,
            update_golden=update_golden,
            require_wps=require_wps,
            deps=qa_deps,
        )
        if not qa_ok:
            if not auto_repair:
                write_agent_summary(out_dir, folder_name, output_docx, mode, pipeline_status="failed", note="QA 未通过，请查看 qa_report.md / qa_repair_plan.md。")
                return None
        if auto_repair:
            step('Phase 7/7: Auto repair loop')
            repair_result = run_repair_loop(
                out_dir,
                mode=mode,
                output_docx_name=output_docx,
                qa_level=qa_level,
                project_root=BASE,
                max_rounds=repair_max_rounds,
                stop_no_improve=repair_stop_no_improve,
                deps=qa_deps,
                run_generated_script=run_generated_script,
                python_executable=sys.executable,
                golden_dir=golden_dir,
                update_golden=update_golden,
                require_wps=require_wps,
            )
            print(f'  [AUTO-REPAIR] {repair_result.status}: final_errors={repair_result.final_errors}')
            repair_report_path = repair_result.report_path
            if os.path.isabs(repair_report_path):
                repair_report_path = os.path.relpath(repair_report_path, BASE)
            print(f'  [AUTO-REPAIR] report -> {repair_report_path.replace(os.sep, "/")}')
            if not repair_result.ok:
                write_agent_summary(out_dir, folder_name, output_docx, mode, pipeline_status="needs_user_action", note="自动修复闭环停止，仍需要用户文件、依赖或人工确认。")
                return None

    # -- Done --
    step('完成')
    write_agent_summary(out_dir, folder_name, output_docx, mode, pipeline_status="completed")
    print_completion_summary(folder_name, output_docx, mode)
    return out_dir


def main():
    main_cli(run_pipeline=run, template_dir=TEMPLATE_DIR, inputs_dir=INPUTS_DIR)


if __name__ == '__main__':
    main()
