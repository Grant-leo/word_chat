"""Run-context helpers for the one-click pipeline runner."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
import json
import os


@dataclass(frozen=True)
class ResolvedInputs:
    template_path: str
    content_path: str
    content_name: str
    use_md_format: bool
    use_md_content: bool


@dataclass(frozen=True)
class InputResolution:
    inputs: ResolvedInputs | None = None
    error: str | None = None

    @property
    def ok(self) -> bool:
        return self.inputs is not None and not self.error


def normalize_qa_level(qa_level: str | None) -> str:
    level = (qa_level or "strict").strip().lower()
    return level if level in ("basic", "strict", "visual") else "strict"


def _display_input_path(folder_label: str, requested: str) -> str:
    requested = str(requested or "").replace(os.sep, "/")
    if os.path.isabs(requested):
        return os.path.basename(requested) or "<absolute path>"
    return f"{folder_label}/{requested}".replace("//", "/")


def _workflow_project_root(out_dir: str) -> str:
    normalized = os.path.normpath(os.path.abspath(os.fspath(out_dir or ".")))
    cursor = normalized
    while True:
        if os.path.basename(cursor).lower() == "outputs":
            parent = os.path.dirname(cursor)
            return parent or cursor
        parent = os.path.dirname(cursor)
        if parent == cursor:
            return normalized
        cursor = parent


def _safe_relative_path(path: str, folder_path: str) -> str:
    normalized_path = os.path.normpath(os.path.abspath(path))
    normalized_folder = os.path.normpath(os.path.abspath(folder_path))
    try:
        common = os.path.commonpath([os.path.normcase(normalized_path), os.path.normcase(normalized_folder)])
    except ValueError:
        return ""
    if common != os.path.normcase(normalized_folder):
        return ""
    rel = os.path.relpath(normalized_path, normalized_folder)
    if rel in ("", ".") or rel.startswith(".."):
        return ""
    return rel.replace("\\", "/")


def _workflow_file_arg(path: str, folder_name: str, project_root: str | None = None) -> str:
    """Return a non-absolute argument that can be reused from Templates/Inputs."""
    path_text = os.fspath(path or "")
    if not path_text:
        return ""
    if not os.path.isabs(path_text):
        return path_text.replace("\\", "/").lstrip("./")
    if project_root:
        return _safe_relative_path(path_text, os.path.join(project_root, folder_name))
    return ""


def _workflow_location_warning(path: str, folder_name: str, role_label: str, project_root: str | None = None) -> str:
    path_text = os.fspath(path or "")
    if not path_text or not os.path.isabs(path_text):
        return ""
    if _workflow_file_arg(path_text, folder_name, project_root):
        return ""
    basename = os.path.basename(path_text) or "<unknown>"
    return (
        f"{role_label} `{basename}` 不在 `{folder_name}/` 下，报告不会生成可能失效的重跑命令；"
        f"请先把该文件放入 `{folder_name}/`，再用文件名重新运行流水线。"
    )


def resolve_inputs(template_file, content_file, md_file, template_dir, inputs_dir) -> InputResolution:
    """Resolve CLI filenames into concrete input paths."""
    if md_file:
        md_path = os.path.abspath(md_file) if not os.path.isabs(md_file) else md_file
        if not os.path.exists(md_path):
            md_path = os.path.join(inputs_dir, md_file)
        if not os.path.exists(md_path):
            return InputResolution(error=f"[ERROR] MD 文件不存在: {_display_input_path('Inputs', md_file)}")
        return InputResolution(
            inputs=ResolvedInputs(
                template_path=md_path,
                content_path=md_path,
                content_name=os.path.splitext(os.path.basename(md_path))[0],
                use_md_format=True,
                use_md_content=True,
            )
        )

    template_path = template_file if os.path.isabs(template_file) else os.path.join(template_dir, template_file)
    content_path = content_file if os.path.isabs(content_file) else os.path.join(inputs_dir, content_file)
    if not os.path.exists(template_path):
        return InputResolution(error=f"[ERROR] 模版文件不存在: {_display_input_path('Templates', template_file)}")
    if not os.path.exists(content_path):
        return InputResolution(error=f"[ERROR] 内容文件不存在: {_display_input_path('Inputs', content_file)}")
    return InputResolution(
        inputs=ResolvedInputs(
            template_path=template_path,
            content_path=content_path,
            content_name=os.path.splitext(os.path.basename(content_file))[0],
            use_md_format=str(template_file).lower().endswith(".md"),
            use_md_content=str(content_file).lower().endswith(".md"),
        )
    )


def create_unique_output_dir(outputs_dir, content_name, today=None):
    """Create and return a unique output directory for a content name."""
    day = today or date.today().isoformat()
    base_folder_name = f"{day}_{content_name}"
    folder_name = base_folder_name
    out_dir = os.path.join(outputs_dir, folder_name)
    suffix = 2
    while os.path.exists(out_dir):
        folder_name = f"{base_folder_name}_{suffix}"
        out_dir = os.path.join(outputs_dir, folder_name)
        suffix += 1
    os.makedirs(out_dir, exist_ok=True)
    return out_dir, folder_name


def write_workflow_mode(
    out_dir,
    *,
    mode,
    template_path,
    content_path,
    run_qa,
    qa_level,
    golden_dir,
    update_golden,
    require_wps,
    auto_repair=False,
    agent_auto=False,
    repair_max_rounds=5,
    repair_stop_no_improve=2,
):
    workflow_path = os.path.join(out_dir, "workflow_mode.json")
    project_root = _workflow_project_root(out_dir)
    md_file = ""
    same_markdown_input = os.path.abspath(template_path) == os.path.abspath(content_path) and str(content_path).lower().endswith(".md")
    if same_markdown_input:
        md_file = _workflow_file_arg(content_path, "Inputs", project_root)
    template_file = _workflow_file_arg(template_path, "Templates", project_root)
    content_file = _workflow_file_arg(content_path, "Inputs", project_root)
    input_location_warnings = []
    if same_markdown_input and not md_file:
        warning = _workflow_location_warning(content_path, "Inputs", "Markdown 输入", project_root)
        if warning:
            input_location_warnings.append(warning)
    elif not same_markdown_input:
        template_warning = _workflow_location_warning(template_path, "Templates", "模板文件", project_root)
        content_warning = _workflow_location_warning(content_path, "Inputs", "内容文件", project_root)
        if template_warning:
            input_location_warnings.append(template_warning)
        if content_warning:
            input_location_warnings.append(content_warning)
    with open(workflow_path, "w", encoding="utf-8") as f:
        json.dump(
            {
                "mode": mode,
                "template": template_file,
                "content": content_file,
                "md": md_file,
                "user_fix_target": "build_generated.py",
                "developer_fix_target": "Paper_Project/Program/pipeline/",
                "qa_enabled": bool(run_qa),
                "qa_level": qa_level,
                "golden_dir": golden_dir,
                "update_golden": bool(update_golden),
                "require_wps": bool(require_wps),
                "auto_repair": bool(auto_repair),
                "agent_auto": bool(agent_auto),
                "repair_max_rounds": int(repair_max_rounds or 0),
                "repair_stop_no_improve": int(repair_stop_no_improve or 0),
                "input_location_warnings": input_location_warnings,
            },
            f,
            ensure_ascii=False,
            indent=2,
        )
    return workflow_path
