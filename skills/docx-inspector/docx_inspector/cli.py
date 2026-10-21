# -*- coding: utf-8 -*-
"""docx-inspector 命令行入口

退出码：
  0  全部通过（可能仍有 warning/info）
  1  存在 error
  2  无法解析该文件

用法：
  python -m docx_inspector file.docx
  python -m docx_inspector file.docx --json
  python -m docx_inspector file.docx --preset thesis
  python -m docx_inspector file.docx --annotate out/ --dpi 150
"""
from __future__ import annotations

import argparse
import json
import os
import sys

from .core import RULES, Ctx, Report, run_rules, SEVERITY_ORDER

VERSION = "0.1.2"

PRESETS = {
    # 默认：8 条最致命，零配置
    "plain": None,
    # 论文/长报告：额外开启 theme 覆盖、正文字体不一致、表头重复
    "thesis": {"add": {"C12", "C14", "D17"}},
    # 只看"打不开/很丑"这类硬伤
    "blockers": {"only": {"A1", "A3", "A4", "B7", "B8"}},
}


def rules_for_preset(name: str) -> set[str]:
    """返回本次要跑的规则 id 集合"""
    if name not in PRESETS:
        raise SystemExit("未知 preset: %s（可选 %s）" % (name, ", ".join(PRESETS)))
    meta = PRESETS[name]
    if meta is None:
        return {rid for rid, m in RULES.items() if m.default_on}
    if "only" in meta:
        return set(meta["only"])
    base = {rid for rid, m in RULES.items() if m.default_on}
    return base | set(meta.get("add", set()))


def _parse_ids(s):
    if not s:
        return set()
    return {x.strip().upper() for x in s.split(",") if x.strip()}


def render_text(rep: Report, ctx: Ctx | None):
    out = []
    out.append("docx-inspector %s" % VERSION)
    out.append("文件: %s  %.1f KB" % (os.path.basename(rep.path), os.path.getsize(rep.path) / 1024))
    if ctx is not None:
        out.append("版心: %.2f × %.2f cm   页边距: 左%.2f 右%.2f 上%.2f 下%.2f cm"
                   % (ctx.usable_w.cm, ctx.usable_h.cm,
                      ctx.margin_l.cm, ctx.margin_r.cm,
                      ctx.margin_t.cm, ctx.margin_b.cm))
    out.append("")

    if rep.error:
        out.append("无法解析: %s" % rep.error)
        return "\n".join(out)

    items = sorted(rep.findings, key=lambda f: (SEVERITY_ORDER.get(f.severity, 9), f.rule))
    noise = [f for f in items if f.severity == "info" and f.rule == "B5"]
    items = [f for f in items if f not in noise]      # B5 的 info 只是报平安，不刷屏

    if not items:
        out.append("未发现问题。")
    for f in items:
        out.append(f.to_line())

    c = rep.counts()
    out.append("")
    out.append("%d 项检查 · %d error / %d warning / %d info"
               % (len(rep.rules_run), c["error"], c["warning"], c["info"]))
    if rep.rules_skipped:
        out.append("未启用: %s" % ", ".join(rep.rules_skipped))
    code = rep.exit_code()
    verdict = {0: "通过", 1: "不通过", 2: "无法解析"}[code]
    out.append("结果: %s (exit %d)" % (verdict, code))
    return "\n".join(out)


def main(argv=None):
    # Windows 控制台/CI 的 stdout 未必是 UTF-8，中文输出会抛 UnicodeEncodeError
    for _stream in (sys.stdout, sys.stderr):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:                           # noqa: BLE001
            pass

    ap = argparse.ArgumentParser(
        prog="docx-inspector",
        description="对任意 .docx 做排版体检：包结构 / 页面几何 / 字体 / 布局 / 编号")
    ap.add_argument("file")
    ap.add_argument("--json", action="store_true", help="输出机器可读 JSON")
    ap.add_argument("--preset", default="plain", choices=list(PRESETS))
    ap.add_argument("--only", default="", help="只跑这些规则，逗号分隔，如 A1,B7")
    ap.add_argument("--skip", default="", help="跳过这些规则")
    ap.add_argument("--annotate", metavar="DIR", default=None,
                    help="渲染并在图上标注问题，输出到该目录")
    ap.add_argument("--dpi", type=int, default=120)
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args(argv)

    if not os.path.exists(a.file):
        print("文件不存在: %s" % a.file, file=sys.stderr)
        return 2

    rep = Report(os.path.abspath(a.file))

    # 预检：不是 ZIP 就没必要继续
    with open(a.file, "rb") as f:
        head = f.read(4)
    if head[:4] != b"PK\x03\x04":
        from .core import Finding
        rep.error = ("不是有效的 DOCX（文件头 %s）"
                     % (head.hex(" ").upper() if head else "空文件"))
        rep.findings.append(Finding("A1", "error", rep.error, location="文件头"))
        _emit(rep, None, a)
        return 2

    try:
        ctx = Ctx(a.file)
    except Exception as e:                              # noqa: BLE001
        rep.error = "%s: %s" % (type(e).__name__, e)
        _emit(rep, None, a)
        return 2

    try:
        include = _parse_ids(a.only) or rules_for_preset(a.preset)
        exclude = _parse_ids(a.skip)
        run_rules(ctx, rep, include, exclude)

        if a.annotate and not a.json:
            try:
                from .annotate import annotate as do_annotate
                files = do_annotate(ctx, rep, a.annotate, a.dpi)
                if not a.quiet:
                    print("标注图 %d 张 -> %s" % (len(files), a.annotate), file=sys.stderr)
            except Exception as e:                      # noqa: BLE001
                print("标注失败（不影响检测结果）: %s: %s"
                      % (type(e).__name__, e), file=sys.stderr)
    finally:
        ctx.close()

    _emit(rep, ctx, a)
    return rep.exit_code()


def _emit(rep: Report, ctx, a):
    if a.json:
        print(json.dumps(rep.to_dict(), ensure_ascii=False, indent=2))
    elif not a.quiet:
        print(render_text(rep, ctx))


if __name__ == "__main__":
    sys.exit(main())
