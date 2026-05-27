"""Template-profile and requirements phase helpers for run_pipeline."""
from __future__ import annotations


def write_template_profile_phase(fmt, out_dir, *, project_root, write_template_profile):
    if write_template_profile is None:
        print("  [WARN] template_profiler.py 不可用，已跳过模板画像")
        return None
    profile = write_template_profile(fmt, out_dir, project_root=project_root)
    caps = profile.get("capabilities") or {}
    risks = profile.get("risk_flags") or {}
    active_risks = [key for key, value in risks.items() if value]
    print("  [OK] template_profile.json / template_profile.md")
    print(
        f'  能力: cover={caps.get("has_cover")} '
        f'headings={caps.get("has_heading_styles")} '
        f'captions={caps.get("has_caption_styles")}'
    )
    if active_risks:
        print(f'  风险标记: {", ".join(active_risks[:6])}')
    return profile


def write_template_requirements_phase(
    fmt,
    content,
    out_dir,
    *,
    write_template_requirements,
    optional_import_detail,
):
    if write_template_requirements is None:
        detail = optional_import_detail("qa_conformance") if optional_import_detail else ""
        print(f"  [WARN] qa_conformance.py 不可用，已跳过 template_requirements。{detail}")
        return False
    write_template_requirements(fmt, content, out_dir)
    print("  [OK] template_requirements.json / template_requirements.md")
    return True
