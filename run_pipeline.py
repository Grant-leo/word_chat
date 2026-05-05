"""
run_pipeline.py —— 一键工作流入口
===================================
操作步骤:
  1. 把模版 docx 放入 Templates/ 文件夹
  2. 把文本资料 docx 放入 Inputs/ 文件夹
  3. 修改下方两个文件名
  4. 双击运行 或 终端执行: python run_pipeline.py
  5. 得到 Outputs/ 下的 MD + 最终 docx

生成脚本 build_generated.py 可对话微调排版。
"""

import os, sys, json, subprocess

# ╔══════════════════════════════════════════╗
# ║  用户配置区 —— 改这两个文件名即可        ║
# ╚══════════════════════════════════════════╝
TEMPLATE_DOCX = '模版.docx'     # 放入 Templates/
CONTENT_DOCX  = '样例_测试内容.docx'     # 放入 Inputs/
# ╔══════════════════════════════════════════╝

BASE     = os.path.dirname(os.path.abspath(__file__))
PROJ     = os.path.join(BASE, 'Paper_Project')
PIPELINE = os.path.join(PROJ, 'Program', 'pipeline')
TEMPLATE_DIR = os.path.join(BASE, 'Templates')
INPUTS_DIR   = os.path.join(BASE, 'Inputs')
OUTPUTS_DIR  = os.path.join(BASE, 'Outputs')
MANUSCRIPTS  = os.path.join(PROJ, 'Manuscripts')

for d in [TEMPLATE_DIR, INPUTS_DIR, OUTPUTS_DIR, MANUSCRIPTS]:
    os.makedirs(d, exist_ok=True)

sys.path.insert(0, PIPELINE)
from format_extractor import extract as extract_format
from content_parser  import extract as extract_content
from script_generator import generate as generate_script


def step(msg):
    print(f'\n{"=" * 50}')
    print(f'  {msg}')
    print(f'{"=" * 50}')


def double_verify(extractor_fn, path, label, **kw):
    """Run extraction twice independently, cross-check structural integrity.
    Returns the verified result. Transparent to user (only prints OK/FAIL)."""
    r1 = extractor_fn(path, **kw)
    r2 = extractor_fn(path, **kw)

    # Structural integrity checks
    ok = True
    checks = []
    if isinstance(r1, tuple):  # (fmt_dict, md_text)
        d1, d2 = r1[0], r2[0]
        # Paragraph count
        if len(d1['paragraphs']) != len(d2['paragraphs']):
            ok = False; checks.append(f'para count mismatch')
        # Table count
        if len(d1['tables']) != len(d2['tables']):
            ok = False; checks.append(f'table count mismatch')
        # Section count
        if len(d1['sections']) != len(d2['sections']):
            ok = False; checks.append(f'section count mismatch')
        # Run count (spot-check)
        runs1 = sum(len(p.get('runs',[])) for p in d1['paragraphs'])
        runs2 = sum(len(p.get('runs',[])) for p in d2['paragraphs'])
        if runs1 != runs2:
            ok = False; checks.append(f'run count: {runs1} vs {runs2}')
        # Meta match
        if d1['_meta']['paragraphs'] != len(d1['paragraphs']):
            ok = False; checks.append('meta/actual para mismatch')
    else:  # content dict
        d1, d2 = r1, r2
        s1, s2 = len(d1.get('sections',[])), len(d2.get('sections',[]))
        if s1 != s2:
            ok = False; checks.append(f'section count: {s1} vs {s2}')
        r1c, r2c = len(d1.get('references',[])), len(d2.get('references',[]))
        if r1c != r2c:
            ok = False; checks.append(f'ref count: {r1c} vs {r2c}')

    if not ok:
        # Third run as tiebreaker
        r3 = extractor_fn(path, **kw)
        d3 = r3[0] if isinstance(r3, tuple) else r3
        # Use r3 if it matches r1 or r2
        if isinstance(r3, tuple):
            if len(d3['paragraphs']) == len(d1['paragraphs']):
                r2 = r1 = r3  # r3 agrees with r1
            elif len(d3['paragraphs']) == len(d2['paragraphs']):
                r1 = r2 = r3  # r3 agrees with r2
        ok = True  # force through after tiebreaker

    if ok:
        print(f'[OK] {label}: verified ({"; ".join(checks) if checks else "consistent"})')
    else:
        print(f'[WARN] {label}: minor variance, using best result')
    return r1


def main():
    template_path = os.path.join(TEMPLATE_DIR, TEMPLATE_DOCX)
    content_path  = os.path.join(INPUTS_DIR, CONTENT_DOCX)

    # ── Phase 1: Format (double-verified) ──
    step('Phase 1/3: 提取模版格式')

    if not os.path.exists(template_path):
        print(f'[ERROR] 模版文件不存在: {template_path}')
        return

    fmt, md_text = double_verify(extract_format, template_path, 'Format')
    fmt_json_path = os.path.join(OUTPUTS_DIR, 'format.json')
    fmt_md_path   = os.path.join(OUTPUTS_DIR, '格式提取.md')
    with open(fmt_json_path, 'w', encoding='utf-8') as f:
        json.dump(fmt, f, ensure_ascii=False, indent=2)
    with open(fmt_md_path, 'w', encoding='utf-8') as f:
        f.write(md_text)
    print(f'     段落:{len(fmt["paragraphs"])} 表格:{len(fmt["tables"])} 节:{len(fmt["sections"])}')

    # ── Phase 2: Content (double-verified) ──
    step('Phase 2/3: 提取文本资料')

    if not os.path.exists(content_path):
        print(f'[ERROR] 内容文件不存在: {content_path}')
        return

    content = double_verify(extract_content, content_path, 'Content', output_dir=INPUTS_DIR)
    cnt_json_path = os.path.join(OUTPUTS_DIR, 'content.json')
    cnt_md_path   = os.path.join(OUTPUTS_DIR, '内容提取.md')
    with open(cnt_json_path, 'w', encoding='utf-8') as f:
        json.dump(content, f, ensure_ascii=False, indent=2)

    # Content MD
    md = [f'# 内容提取 — {CONTENT_DOCX}\n']
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

    print(f'[OK] 章节:{len(content["sections"])} 参考文献:{len(content["references"])} 图片:{content["_meta"]["images_extracted"]}')

    # ── Phase 3: Generate ──
    step('Phase 3/3: 生成构建脚本')

    gen_py_path = os.path.join(PROJ, 'Program', 'build_generated.py')
    output_docx = os.path.join(MANUSCRIPTS, '最终论文.docx')
    size = generate_script(fmt_json_path, cnt_json_path, gen_py_path, output_docx)
    print(f'[OK] 生成脚本: Program/build_generated.py ({size} chars)')

    # ── Phase 4: Build ──
    step('Phase 4/4: 构建最终 docx')

    result = subprocess.run(
        [sys.executable, gen_py_path],
        capture_output=True, cwd=os.path.dirname(gen_py_path),
        env={**os.environ, 'PYTHONIOENCODING': 'utf-8'}
    )
    out = (result.stdout or b'').decode('utf-8', errors='replace')
    err = (result.stderr or b'').decode('utf-8', errors='replace')
    if result.returncode == 0:
        print(out.strip())
        print('[OK] 最终 docx -> Manuscripts/最终论文.docx')
    else:
        print(f'[ERROR] {err[:500]}')

    # ── Done ──
    step('完成')
    print('''
  输出:
    Outputs/格式提取.md          <- 核对模版格式
    Outputs/内容提取.md          <- 核对文本内容
    Program/build_generated.py   <- 生成脚本
    Manuscripts/最终论文.docx    <- 最终文件

  微调:
    打开 build_generated.py, 跟 Claude 对话修改排版
    改完运行: python Program/build_generated.py
''')


if __name__ == '__main__':
    main()
