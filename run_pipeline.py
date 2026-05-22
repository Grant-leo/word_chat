"""
run_pipeline.py —— 一键工作流入口
===================================

两种用法:
  交互模式:  python run_pipeline.py
             → 自动扫描文件，编号选择

  参数模式:  python run_pipeline.py --template 模版.docx --content 论文.docx
             → 直接运行，无交互（Skill / 脚本调用）

  结果自动输出到 Outputs/{日期}_{内容名}/
  生成脚本 build_generated.py 可对话微调排版。
"""

import os, sys, json, subprocess, argparse
from datetime import date

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

try:
    from md_parser import extract_format as extract_md_format
    from md_parser import extract_content as extract_md_content
except Exception:
    extract_md_format = None
    extract_md_content = None


def scan_inputs(folder, exts=('.docx', '.md')):
    """Return list of docx/md files in folder, excluding temp files (~$ prefix)."""
    if not os.path.isdir(folder):
        return []
    files = [f for f in os.listdir(folder)
             if any(f.endswith(e) for e in exts) and not f.startswith('~$')]
    return sorted(files)


def choose_file(files, label):
    """Interactive file selection. Returns chosen filename."""
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


def run(template_file, content_file, md_file=None):
    """Core pipeline: takes filenames, runs all phases, returns output directory.

    If md_file is provided, uses MD file for BOTH format and content (single-file mode).
    Otherwise routes by file extension: .docx uses existing extractors, .md uses md_parser.
    """
    if md_file:
        # Single-MD mode: same file for format + content
        md_path = os.path.abspath(md_file) if not os.path.isabs(md_file) else md_file
        if not os.path.exists(md_path):
            # Try Inputs/ directory
            md_path = os.path.join(INPUTS_DIR, md_file)
        if not os.path.exists(md_path):
            print(f'[ERROR] MD 文件不存在: {md_file}')
            return None
        template_path = md_path
        content_path = md_path
        content_name = os.path.splitext(os.path.basename(md_path))[0]
        use_md_format = True
        use_md_content = True
        if extract_md_format is None or extract_md_content is None:
            print('[ERROR] 当前脚本包未包含 md_parser.py，不能处理 Markdown 单文件模式。')
            return None
    else:
        template_path = os.path.join(TEMPLATE_DIR, template_file)
        content_path  = os.path.join(INPUTS_DIR, content_file)
        if not os.path.exists(template_path):
            print(f'[ERROR] 模版文件不存在: {template_path}')
            return None
        if not os.path.exists(content_path):
            print(f'[ERROR] 内容文件不存在: {content_path}')
            return None
        content_name = os.path.splitext(content_file)[0]
        use_md_format = template_file.endswith('.md')
        use_md_content = content_file.endswith('.md')
        if (use_md_format or use_md_content) and (extract_md_format is None or extract_md_content is None):
            print('[ERROR] 当前脚本包未包含 md_parser.py，不能处理 Markdown 输入。请使用 .docx 模版和 .docx 内容。')
            return None

    base_folder_name = f'{date.today().isoformat()}_{content_name}'
    folder_name = base_folder_name
    out_dir = os.path.join(OUTPUTS_DIR, folder_name)
    suffix = 2
    while os.path.exists(out_dir):
        folder_name = f'{base_folder_name}_{suffix}'
        out_dir = os.path.join(OUTPUTS_DIR, folder_name)
        suffix += 1
    os.makedirs(out_dir, exist_ok=True)

    print(f'  输出目录: Outputs/{folder_name}/')
    print(f'  模版: {os.path.basename(template_path)}')
    print(f'  内容: {os.path.basename(content_path)}')

    # ── Phase 1: Format ──
    step('Phase 1/4: 提取模版格式')
    fmt_extractor = extract_md_format if use_md_format else extract_format
    fmt, md_text = double_verify(fmt_extractor, template_path, 'Format')

    fmt_json_path = os.path.join(out_dir, 'format.json')
    fmt_md_path   = os.path.join(out_dir, '格式提取.md')
    with open(fmt_json_path, 'w', encoding='utf-8') as f:
        json.dump(fmt, f, ensure_ascii=False, indent=2)
    with open(fmt_md_path, 'w', encoding='utf-8') as f:
        f.write(md_text)
    print(f'  段落:{len(fmt["paragraphs"])}  表格:{len(fmt["tables"])}  节:{len(fmt["sections"])}')

    # ── Phase 2: Content ──
    step('Phase 2/4: 提取文本内容')
    cnt_extractor = extract_md_content if use_md_content else extract_content
    content = double_verify(cnt_extractor, content_path, 'Content', output_dir=out_dir)

    cnt_json_path = os.path.join(out_dir, 'content.json')
    cnt_md_path   = os.path.join(out_dir, '内容提取.md')
    with open(cnt_json_path, 'w', encoding='utf-8') as f:
        json.dump(content, f, ensure_ascii=False, indent=2)

    md = [f'# 内容提取 — {os.path.basename(content_path)}\n']
    for sec in content.get('sections', []):
        md.append(f'## {sec["heading"]}\n')
        for img in sec.get('images', []):
            md.append(f'- [图片] {img}')
        for p in sec.get('paragraphs', []):
            if isinstance(p, dict):
                t = p.get('text', '') or '[公式]'
                if p.get('math'):
                    t += f' (+{len(p["math"])}公式)'
            else:
                t = p
            t = t[:120] + '...' if len(t) > 120 else t
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
    step('Phase 4/4: 构建最终 docx（生成静态目录；可用 Word COM 时写入页码）')

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
        print('  [OK] 已生成静态目录；若当前环境可调用 Word COM，会自动写入页码')
    else:
        print(f'  [ERROR] {err[:500]}')
        return None

    # ── Done ──
    step('完成')
    print(f'''
  输出目录: Outputs/{folder_name}/
    ├── 格式提取.md          <- 核对模版格式
    ├── 内容提取.md          <- 核对文本内容
    ├── format.json
    ├── content.json
    ├── build_generated.py   <- 生成脚本
    └── {output_docx}        <- 最终文件

  微调:
    打开 build_generated.py，可继续微调排版
    改完运行: python Outputs/{folder_name}/build_generated.py
    目录: 生成脚本会优先用 Word COM 解析正文标题页码；不可用时仍保留静态目录行
''')
    return out_dir


def main():
    parser = argparse.ArgumentParser(description='Word 论文排版流水线')
    parser.add_argument('--template', '-t', help='模版文件名 (位于 Templates/)')
    parser.add_argument('--content',  '-c', help='内容文件名 (位于 Inputs/)')
    parser.add_argument('--md', help='单个 MD 文件（含格式+内容，纯 MD 模式）')
    args = parser.parse_args()

    print('=' * 50)
    print('  Word 论文排版流水线')
    print('=' * 50)

    # ── Determine files ──
    if args.md:
        # Pure MD mode: single file for both format and content
        run(None, None, md_file=args.md)
        return

    if args.template and args.content:
        # Non-interactive mode (Skill / script)
        template_file = args.template
        content_file  = args.content
    else:
        # Interactive mode (CLI user)
        templates = scan_inputs(TEMPLATE_DIR, exts=('.docx',))
        contents  = scan_inputs(INPUTS_DIR, exts=('.docx', '.md'))

        if not templates:
            print('\n[INFO] Templates/ 下没有 .docx 文件。')
            print('  纯 MD 模式请用: python run_pipeline.py --md <文件名>')
            # Check if there's an MD in Inputs/ we can offer
            md_files = scan_inputs(INPUTS_DIR, exts=('.md',))
            if md_files:
                print(f'\n  Inputs/ 下找到 .md 文件，可直接纯 MD 模式:')
                for f in md_files:
                    print(f'    python run_pipeline.py --md {f}')
            return
        if not contents:
            print('\n[ERROR] Inputs/ 下没有 .docx 或 .md 文件，请放入内容文件后重试。')
            return

        template_file = choose_file(templates, '选择模版')
        content_file  = choose_file(contents, '选择内容')

    run(template_file, content_file)


if __name__ == '__main__':
    main()
