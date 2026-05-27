"""Build-script generation and DOCX build phase helpers for run_pipeline."""
from __future__ import annotations

import os


def generate_and_build_docx_phase(
    fmt_json_path,
    cnt_json_path,
    out_dir,
    folder_name,
    *,
    output_docx_name,
    generate_script,
    run_generated_script,
    python_executable,
    step,
):
    step("Phase 4/6: 生成构建脚本")

    gen_size = generate_script(fmt_json_path, cnt_json_path, out_dir, output_docx_name)
    gen_py_path = os.path.join(out_dir, "build_generated.py")
    print(f"  生成脚本: build_generated.py ({gen_size} chars)")

    step("Phase 5/6: 构建最终 docx（生成静态目录；可用 Word COM 时写入页码）")

    build_result = run_generated_script(gen_py_path, out_dir, python_executable=python_executable)
    if build_result.returncode != 0:
        print(f"  [ERROR] {build_result.stderr[:500]}")
        return False

    print(build_result.stdout.strip())
    print(f"  [OK] 最终 docx -> Outputs/{folder_name}/{output_docx_name}")
    print("  [OK] 已生成静态目录；若当前环境可调用 Word COM，会自动写入页码")
    return True
