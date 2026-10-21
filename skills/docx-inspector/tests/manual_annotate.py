# -*- coding: utf-8 -*-
"""手动验证渲染标注链路（--annotate）

为什么不在 CI 里跑：
    标注需要把 docx 渲染成 PDF，进而依赖 Word COM（Windows）或 LibreOffice。
    GitHub runner 两者都没有——装 LibreOffice 也要求与本地一致的渲染结果，
    成本远高于收益。

什么时候跑：
    改动了 annotate.py、或升级了 python-docx / PyMuPDF 之后。

用法：
    python tests/manual_annotate.py                # 用内置 fixture
    python tests/manual_annotate.py <某个.docx>    # 用指定文档

人工检查点：
    1. 每页都生成了 PNG，且页数与 Word 打开时一致
    2. 顶部状态条颜色/文字与报告一致（绿=通过、橙=只有 warning、红=有 error）
    3. 有溢出的页面上，红框从版心右边界起、覆盖真正越界的区域，并标注了溢出厘米数
    4. 底部说明行显示的本页溢出次数合理
"""
import os
import sys
import tempfile

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
sys.path.insert(0, ROOT)

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:                                   # noqa: BLE001
        pass

from docx_inspector.cli import main                    # noqa: E402


def run(docx_path):
    out_dir = os.path.join(tempfile.gettempdir(), "docx-inspector-annotate-check")
    os.makedirs(out_dir, exist_ok=True)
    print("输入:", docx_path)
    print("输出:", out_dir)
    print("-" * 60)
    code = main([docx_path, "--annotate", out_dir, "--dpi", "110"])
    print("-" * 60)
    print("退出码:", code, "(0 通过 / 1 有 error)")

    pages = sorted(f for f in os.listdir(out_dir) if f.endswith(".png"))
    print("生成页面:", len(pages))
    for p in pages:
        print("  %-14s %8d bytes" % (p, os.path.getsize(os.path.join(out_dir, p))))
    if not pages:
        print("!! 未生成任何图片，渲染链路有问题")
        return 1
    print("\n请人工打开这些 PNG，按文件头的检查点核对。")
    return 0


if __name__ == "__main__":
    if len(sys.argv) > 1:
        target = sys.argv[1]
    else:
        target = os.path.join(ROOT, "case_overflow.docx")
    if not os.path.exists(target):
        print("文件不存在:", target)
        sys.exit(2)
    sys.exit(run(target))
