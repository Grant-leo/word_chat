"""
master.py — 论文项目主控脚本
按顺序串联：数据清洗 → 分析 → 生成图表

执行方式：
    python master.py

步骤：
    1. Program/Clean/   — 数据清洗
    2. Program/Analysis/ — 统计分析
    3. Results/         — 输出图表
"""
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.abspath(__file__))

STEPS = [
    # ("步骤名", "脚本相对路径"),
    # ("01_clean_data", "Program/Clean/clean_data.py"),
    # ("02_analysis", "Program/Analysis/run_analysis.py"),
    # ("03_make_tables", "Program/Analysis/make_tables.py"),
    # ("04_make_figures", "Program/Analysis/make_figures.py"),
]

def run_step(name, script):
    path = os.path.join(BASE, script)
    print(f"[{name}] running {path} ...")
    result = subprocess.run([sys.executable, path], capture_output=True, text=True)
    if result.returncode != 0:
        print(f"  ERROR: {result.stderr}")
        sys.exit(1)
    print(f"  OK")

def main():
    print("=== 论文项目主控脚本 ===")
    for name, script in STEPS:
        run_step(name, script)
    print("=== 全部完成 ===")

if __name__ == "__main__":
    main()
