"""
run_pipeline.py —— 一键工作流入口
===================================
  1. 把模版 docx 放入 Templates/，内容 docx 放入 Inputs/
  2. 运行 python run_pipeline.py
  3. 选择模版和内容文件
  4. 结果自动输出到 Outputs/{日期}_{内容名}/

生成脚本 build_generated.py 可对话微调排版。
"""

import os, sys, json, subprocess
from datetime import date

BASE = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.join(BASE, 'Paper_Project')
PIPELINE = os.path.join(PROJ, 'Program', 'pipeline')
TEMPLATE_DIR = os.path.join(BASE, 'Templates')
INPUTS_DIR = os.path.join(BASE, 'Inputs')
OUTPUTS_DIR = os.path.join(BASE, 'Outputs')

for d in [TEMPLATE_DIR, INPUTS_DIR, OUTPUTS_DIR]:
    os.makedirs(d, exist_ok=True)

sys.path.insert(0, PIPELINE)
from format_extractor import extract as extract_format
from content_parser import extract as extract_content
from script_generator import generate as generate_script


def scan_docx(folder):
    """Return list of .docx files in folder, excluding temp files (~$ prefix)."""
    if not os.path.isdir(folder):
        return []
    files = [f for f in os.listdir(folder)
             if f.endswith('.docx') and not f.startswith('~$')]
    return sorted(files)


def choose_file(files, label):
    """Let user pick from a numbered list. Returns chosen filename."""
    if len(files) == 0:
        return None
    if len(files) == 1:
        print(f'{label}: {files[0]} (自动选择)')
        return files[0]
    print(f'\n{label}:')
    for i, f in enumerate(files, 1):
        print(f'  [{i}] {f}')
    while True:
        try:
            choice = input(f'请选择 (1-{len(files)}): ').strip()
            idx = int(choice) - 1
            if 0 <= idx < len(files):
                return files[idx]
            print(f'  输入无效，请输入 1-{len(files)}')
        except (ValueError, KeyboardInterrupt):
            print('\n  已取消')
            raise SystemExit(1)


def step(msg):
    print(f'\n{"=" * 50}')
    print(f'  {msg}')
    print(f'{"=" * 50}')


def double_verify(extractor_fn, path, label, **kw):
    """Run extraction twice independently, cross-check structural integrity."""
    r1 = extractor_fn(path, **kw)
    r2 = extractor_fn(path, **kw)

    ok = True
    checks = []
    if isinstance(r1, tuple):  # (fmt_dict, md_text)
        d1, d2 = r1[0], r2[0]
        if len(d1['paragraphs']) != len(d2['paragraphs']):
            ok = False; checks.append('para count mismatch')
        if len(d1['tables']) != len(d2['tables']):
            ok = False; checks.append('table count mismatch')
        if len(d1['sections']) != len(d2['sections']):
            ok = False; checks.append('section count mismatch')
        runs1 = sum(len(p.get('runs', [])) for p in d1['paragraphs'])
        runs2 = sum(len(p.get('runs', [])) for p in d2['paragraphs'])
        if runs1 != runs2:
            ok = False; checks.append(f'run count: {runs1} vs {runs2}')
        if d1['_meta']['paragraphs'] != len(d1['paragraphs']):
            ok = False; checks.append('meta/actual para mismatch')
    else:  # content dict
        d1, d2 = r1, r2
        if len(d1.get('sections', [])) != len(d2.get('sections', [])):
            ok = False; checks.append('section count mismatch')
        if len(d1.get('references', [])) != len(d2.get('references', [])):
            ok = False; checks.append('ref count mismatch')

    if not ok:
        r3 = extractor_fn(path, **kw)
        d3 = r3[0] if isinstance(r3, tuple) else r3
        if isinstance(r3, tuple):
            if len(d3['paragraphs']) == len(d1['paragraphs']):
                r2 = r1 = r3
            elif len(d3['paragraphs']) == len(d2['paragraphs']):
                r1 = r2 = r3
        ok = True

    status = '; '.join(checks) if checks else 'consistent'
    print(f'  [OK] {label}: verified ({status})')
    return r1


def main():
    print('=' * 50)
    print('  Word 论文排版流水线')
    print('=' * 50)

    # ── Scan files ──
    templates = scan_docx(TEMPLATE_DIR)
    contents  = scan_docx(INPUTS_DIR)

    if not templates:
        print(f'\n[ERROR] Templates/ 下没有 .docx 文件，请放入模版文件后重试。')
        return
    if not contents:
        print(f'\n[ERROR] Inputs/ 下没有 .docx 文件，请放入内容文件后重试。')
        return

    template_file = choose_file(templates, '选择模版')
    content_file  = choose_file(contents, '选择内容')

    template_path = os.path.join(TEMPLATE_DIR, template_file)
    content_path  = os.path.join(INPUTS_DIR, content_file)

    # ── Output folder ──
    content_name = os.path.splitext(content_file)[0]
    folder_name  = f'{date.today().isoformat()}_{content_name}'
    out_dir      = os.path.join(OUTPUTS_DIR, folder_name)
    os.makedirs(out_dir, exist_ok=True)

    print(f'\n  输出目录: Outputs/{folder_name}/')
    print(f'  模版: {template_file}')
    print(f'  内容: {content_file}')

    # ── Phase 1: Format ──
    step('Phase 1/4: 提取模版格式')
    fmt, md_text = double_verify(extract_format, template_path, 'Format')

    fmt_json_path = os.path.join(out_dir, 'format.json')
    fmt_md_path   = os.path.join(out_dir, '格式提取.md')
    with open(fmt_json_path, 'w', encoding='utf-8') as f:
        json.dump(fmt, f, ensure_ascii=False, indent=2)
    with open(fmt_md_path, 'w', encoding='utf-8') as f:
        f.write(md_text)
    print(f'  段落:{len(fmt["paragraphs"])}  表格:{len(fmt["tables"])}  节:{len(fmt["sections"])}')

    # ── Phase 2: Content ──
    step('Phase 2/4: 提取文本内容')
    content = double_verify(extract_content, content_path, 'Content', output_dir=INPUTS_DIR)

    cnt_json_path = os.path.join(out_dir, 'content.json')
    cnt_md_path   = os.path.join(out_dir, '内容提取.md')
    with open(cnt_json_path, 'w', encoding='utf-8') as f:
        json.dump(content, f, ensure_ascii=False, indent=2)

    md = [f'# 内容提取 — {content_file}\n']
    for sec in content.get('sections', []):
        md.append(f'## {sec["heading"]}\n')
        for img in sec.get('images', []):
            md.append(f'- [图片] {img}')
        for p in sec.get('paragraphs', []):
            t = p[:120] + '...' if len(p) > 120 else p
            md.append(f'- {t}')
        md.append('')
    if content.get('references'):
        md.append('## 参考文献\n')
        for r in content['references']:
            md.append(f'- {r[:120]}')
    with open(cnt_md_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(md))

    print(f'  章节:{len(content["sections"])}  参考文献:{len(content["references"])}  图片:{content["_meta"]["images_extracted"]}')

    # ── Phase 3: Generate script ──
    step('Phase 3/4: 生成构建脚本')

    output_docx = '最终论文.docx'
    gen_size = generate_script(fmt_json_path, cnt_json_path, out_dir, output_docx)
    gen_py_path = os.path.join(out_dir, 'build_generated.py')
    print(f'  生成脚本: build_generated.py ({gen_size} chars)')

    # ── Phase 4: Build docx ──
    step('Phase 4/4: 构建最终 docx')

    result = subprocess.run(
        [sys.executable, gen_py_path],
        capture_output=True, cwd=out_dir,
        env={**os.environ, 'PYTHONIOENCODING': 'utf-8'}
    )
    out = (result.stdout or b'').decode('utf-8', errors='replace')
    err = (result.stderr or b'').decode('utf-8', errors='replace')
    if result.returncode == 0:
        print(out.strip())
        print(f'  [OK] 最终 docx -> Outputs/{folder_name}/{output_docx}')
    else:
        print(f'  [ERROR] {err[:500]}')
        return

    # ── Done ──
    step('完成')
    print(f'''
  输出目录: Outputs/{folder_name}/
    ├── 格式提取.md          ← 核对模版格式
    ├── 内容提取.md          ← 核对文本内容
    ├── format.json
    ├── content.json
    ├── build_generated.py   ← 生成脚本
    └── {output_docx}        ← 最终文件

  微调:
    打开 build_generated.py，跟 Claude 对话修改排版
    改完运行: python Outputs/{folder_name}/build_generated.py
''')


if __name__ == '__main__':
    main()
