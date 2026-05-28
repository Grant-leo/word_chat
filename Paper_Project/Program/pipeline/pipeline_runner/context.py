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


def resolve_inputs(template_file, content_file, md_file, template_dir, inputs_dir) -> InputResolution:
    """Resolve CLI filenames into concrete input paths."""
    if md_file:
        md_path = os.path.abspath(md_file) if not os.path.isabs(md_file) else md_file
        if not os.path.exists(md_path):
            md_path = os.path.join(inputs_dir, md_file)
        if not os.path.exists(md_path):
            return InputResolution(error=f"[ERROR] MD 文件不存在: {md_file}")
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
        return InputResolution(error=f"[ERROR] 模版文件不存在: {template_path}")
    if not os.path.exists(content_path):
        return InputResolution(error=f"[ERROR] 内容文件不存在: {content_path}")
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
    repair_max_rounds=5,
    repair_stop_no_improve=2,
):
    workflow_path = os.path.join(out_dir, "workflow_mode.json")
    with open(workflow_path, "w", encoding="utf-8") as f:
        json.dump(
            {
                "mode": mode,
                "template": os.path.basename(template_path),
                "content": os.path.basename(content_path),
                "user_fix_target": "build_generated.py",
                "developer_fix_target": "Paper_Project/Program/pipeline/",
                "qa_enabled": bool(run_qa),
                "qa_level": qa_level,
                "golden_dir": golden_dir,
                "update_golden": bool(update_golden),
                "require_wps": bool(require_wps),
                "auto_repair": bool(auto_repair),
                "repair_max_rounds": int(repair_max_rounds or 0),
                "repair_stop_no_improve": int(repair_stop_no_improve or 0),
            },
            f,
            ensure_ascii=False,
            indent=2,
        )
    return workflow_path
