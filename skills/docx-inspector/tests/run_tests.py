# -*- coding: utf-8 -*-
"""回归测试：断言退出码与规则命中，不靠肉眼看输出

运行：  python tests/run_tests.py
退出码：0 全通过 / 1 有断言失败
"""
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
sys.path.insert(0, ROOT)

# Windows 控制台 / CI runner 的 stdout 未必是 UTF-8，中文输出会抛
# UnicodeEncodeError（本地 cp1252 下的实测）。这里强制改成 UTF-8 兜底。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:                                   # noqa: BLE001
        pass

from docx_inspector.core import Ctx, Report, run_rules          # noqa: E402
from docx_inspector.cli import rules_for_preset                 # noqa: E402
from docx_inspector.rules import installed_fonts                # noqa: E402

# C10 依赖“本机是否能枚举字体”。Linux CI runner 上枚举不到，此时优雅跳过而不是报错。
HAS_FONT_ENV = bool(installed_fonts())


def inspect(path, preset="plain"):
    rep = Report(os.path.abspath(path))
    ctx = Ctx(os.path.abspath(path))
    try:
        run_rules(ctx, rep, rules_for_preset(preset), set())
    finally:
        ctx.close()
    return rep


CASES = [
    # (说明, 文件名, preset, 期望退出码, 必须命中, 必须不命中)
    ("fixed 布局超宽表 → error",
     "case_overflow.docx", "plain", 1, {"B7"}, set()),

    ("auto 布局声明超宽 → 仅 warning，不算失败",
     "case_auto_overflow.docx", "plain", 0, {"B7"}, set()),

    ("合规文档 → 全部通过",
     "case_clean.docx", "plain", 0, set(), {"B7", "C10"}),

    ("缺字体 仿宋_GB2312 → C10 命中",
     "case_font.docx", "plain", 0, {"C10"}, set()),

    ("正文字体不一致 → C14 仅在 thesis 下启用",
     "case_font.docx", "thesis", 0, {"C10", "C14"}, set()),
]


def main():
    fails = 0
    for name, fname, preset, want_exit, want_rules, forbid_rules in CASES:
        path = os.path.join(ROOT, fname)
        print("-" * 68)
        print("#", name)
        if not os.path.exists(path):
            print("  [SKIP] 文件不存在:", path)
            fails += 1
            continue
        rep = inspect(path, preset)
        got_ids = {f.rule for f in rep.findings if f.severity in ("error", "warning")}
        got_exit = rep.exit_code()

        ok = True
        if got_exit != want_exit:
            print("  [FAIL] 退出码 期望 %s 实际 %s" % (want_exit, got_exit))
            ok = False
        for rid in want_rules:
            if rid == "C10" and not HAS_FONT_ENV:
                print("  [SKIP] 本机无法枚举字体（非 Windows 或字体目录缺失），跳过 C10 断言")
                continue
            if rid not in got_ids:
                print("  [FAIL] 应命中 %s 但未命中（命中: %s）" % (rid, sorted(got_ids)))
                ok = False
        for rid in forbid_rules:
            if rid in got_ids:
                print("  [FAIL] 不应命中 %s 却命中了" % rid)
                ok = False
        if ok:
            print("  [PASS] exit=%s 命中=%s" % (got_exit, sorted(got_ids) or "无"))
        else:
            fails += 1

    print("=" * 68)
    print("%d 个用例，%d 失败" % (len(CASES), fails))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
