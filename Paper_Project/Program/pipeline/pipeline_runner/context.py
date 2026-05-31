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


def _workflow_file_arg(path: str, folder_name: str) -> str:
    """Return a non-absolute argument that can be reused from Templates/Inputs."""
    normalized = os.path.normpath(os.path.abspath(path))
    parts = normalized.split(os.sep)
    folder_lower = folder_name.lower()
    for idx in range(len(parts) - 1, -1, -1):
        if parts[idx].lower() == folder_lower and idx < len(parts) - 1:
            return "/".join(parts[idx + 1 :])
    return os.path.basename(path)


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
    md_file = ""
    if os.path.abspath(template_path) == os.path.abspath(content_path) and str(content_path).lower().endswith(".md"):
        md_file = _workflow_file_arg(content_path, "Inputs")
    template_file = _workflow_file_arg(template_path, "Templates")
    content_file = _workflow_file_arg(content_path, "Inputs")
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
            },
            f,
            ensure_ascii=False,
            indent=2,
        )
    return workflow_path
