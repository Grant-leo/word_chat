"""Optional dependency loading for the one-click pipeline runner."""
from __future__ import annotations

from dataclasses import dataclass, field
import importlib


@dataclass(frozen=True)
class OptionalPipelineDependencies:
    qa_check_and_write: object = None
    visual_check_and_write: object = None
    conformance_check_and_write: object = None
    write_template_requirements: object = None
    write_template_profile: object = None
    extract_md_format: object = None
    extract_md_content: object = None
    import_errors: dict = field(default_factory=dict)

    def optional_import_detail(self, name):
        exc = self.import_errors.get(name)
        return f" ({exc})" if exc else ""


def _import_attrs(module_name, attr_names, errors, import_module):
    try:
        module = import_module(module_name)
        return tuple(getattr(module, attr_name) for attr_name in attr_names)
    except (ImportError, AttributeError) as exc:
        errors[module_name] = exc
        return tuple(None for _ in attr_names)


def load_optional_dependencies(import_module=None):
    import_module = import_module or importlib.import_module
    errors = {}

    (qa_check_and_write,) = _import_attrs("qa_checker", ("check_and_write",), errors, import_module)
    (visual_check_and_write,) = _import_attrs("qa_visual", ("check_and_write",), errors, import_module)
    conformance_check_and_write, write_template_requirements = _import_attrs(
        "qa_conformance",
        ("check_and_write", "write_requirements"),
        errors,
        import_module,
    )
    (write_template_profile,) = _import_attrs("template_profiler", ("write_profile",), errors, import_module)
    extract_md_format, extract_md_content = _import_attrs(
        "md_parser",
        ("extract_format", "extract_content"),
        errors,
        import_module,
    )

    return OptionalPipelineDependencies(
        qa_check_and_write=qa_check_and_write,
        visual_check_and_write=visual_check_and_write,
        conformance_check_and_write=conformance_check_and_write,
        write_template_requirements=write_template_requirements,
        write_template_profile=write_template_profile,
        extract_md_format=extract_md_format,
        extract_md_content=extract_md_content,
        import_errors=errors,
    )
