# -*- coding: utf-8 -*-
"""docx_inspector.rules — v1.0 的 12 条规则

每条规则只回答一个确定性问题，不做审美判断。
规则内不得读取外部配置：凡需基准的判定，一律进 L1 preset 或 L2 --against。
"""
from __future__ import annotations

import functools
import os
import shutil
import struct
import subprocess
import zipfile

from docx.oxml.ns import qn
from docx.shared import Emu

from .core import Ctx, Finding, rule

# ------------------------------------------------------------------ 字体环境
# Windows 字体目录常见文件名 → family 名（本地化名优先，因为中文 Word 写中文名）
_FONT_FILES = {
    "simkai.ttf": "楷体",
    "simsun.ttc": "宋体",
    "simhei.ttf": "黑体",
    "simfang.ttf": "仿宋",
    "msyh.ttc": "微软雅黑",
    "msyhbd.ttc": "微软雅黑",
    "Deng.ttf": "等线",
    "times.ttf": "Times New Roman",
    "arial.ttf": "Arial",
    "calibri.ttf": "Calibri",
    "cambria.ttc": "Cambria",
    "consola.ttf": "Consolas",
    "symbol.ttf": "Symbol",
}

# 这些"字体名"不是真字体，不该报缺失
_PSEUDO_FONTS = {"", "default", "inherit"}


def installed_fonts() -> set[str]:
    """本机已安装字体族名集合。

    探测顺序：Windows 注册表 → Windows 字体目录 → fontconfig → macOS。
    底层结果缓存（frozenset），避免同进程内重复枚举。
    """
    return set(_installed_fonts_cached())


@functools.lru_cache(maxsize=1)
def _installed_fonts_cached() -> frozenset:
    names: set[str] = set()

    # ---- 1) Windows：注册表 ----
    try:
        import winreg
        key = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(root, key) as k:
                    i = 0
                    while True:
                        try:
                            vname, _data, _t = winreg.EnumValue(k, i)
                        except OSError:
                            break
                        # 值名形如 "楷体 (TrueType)" / "SimSun & NSimSun (TrueType)"
                        base = vname.split("(")[0].strip()
                        if base:
                            names.add(base)
                            for part in base.replace("&", ",").split(","):
                                p = part.strip()
                                if p:
                                    names.add(p)
                        i += 1
            except OSError:
                pass
    except Exception:                              # noqa: BLE001
        pass

    # ---- 2) Windows：字体目录文件名兜底（中文 Word 常写中文名） ----
    fdir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    if os.path.isdir(fdir):
        try:
            existing = {f.lower() for f in os.listdir(fdir)}
            for fn, family in _FONT_FILES.items():
                if fn.lower() in existing:
                    names.add(family)
        except OSError:
            pass

    # ---- 3) Linux / 其他 fontconfig 系统 ----
    if not names:
        names |= _fonts_from_fontconfig()

    # ---- 4) macOS 兜底 ----
    if not names:
        names |= _fonts_from_macos()

    return frozenset(names)


def _fonts_from_fontconfig() -> set[str]:
    """用 fc-list 枚举字体族（Linux 及安装了 fontconfig 的系统）。"""
    out: set[str] = set()
    fc = shutil.which("fc-list")
    if not fc:
        return out
    try:
        r = subprocess.run([fc, ":", "family"], capture_output=True, text=True,
                           timeout=20, errors="replace")
        for line in (r.stdout or "").splitlines():
            for part in line.split(","):
                p = part.strip()
                if p:
                    out.add(p)
    except Exception:                              # noqa: BLE001
        pass
    return out


def _fonts_from_macos() -> set[str]:
    """macOS 兜底：system_profiler 抓 Family / Full Name（较慢，仅首次调用）。"""
    out: set[str] = set()
    sp = shutil.which("system_profiler")
    if not sp:
        return out
    try:
        r = subprocess.run([sp, "SPFontsDataType"], capture_output=True,
                           text=True, timeout=60, errors="replace")
        for line in (r.stdout or "").splitlines():
            t = line.strip()
            for prefix in ("Family:", "Full Name:"):
                if t.startswith(prefix):
                    v = t[len(prefix):].strip()
                    if v:
                        out.add(v)
    except Exception:                              # noqa: BLE001
        pass
    return out


# ================================================================== A 包结构
@rule("A1", "package", "ZIP 头非法（含 DLP 加密或非标准 zip）", "error", True, False)
def r_a1_zip(ctx: Ctx):
    with open(ctx.path, "rb") as f:
        head = f.read(8)
    out = []
    if head[:4] == b"PK\x03\x04":
        return out
    if head[:4] == b"\xE0\xA8\x91\xE7":
        out.append(Finding("A1", "error",
                           "文件被透明加密（头 E0 A8 91 E7），普通程序读到的是密文",
                           location="文件头", detail={"head": head[:8].hex(" ")}))
    elif head[:2] == b"PK":
        out.append(Finding("A1", "error",
                           "ZIP 头非本地文件头（可能是空包或分卷）",
                           location="文件头", detail={"head": head[:8].hex(" ")}))
    else:
        out.append(Finding("A1", "error",
                           "不是有效的 DOCX/ZIP 容器",
                           location="文件头", detail={"head": head[:8].hex(" ")}))
    return out


@rule("A3", "package", "关系目标缺失（图片/嵌入对象找不到）", "error", True, False)
def r_a3_rel(ctx: Ctx):
    out = []
    rels_map = {}
    for n in ctx.names:
        if not n.endswith(".rels"):
            continue
        base_dir = os.path.dirname(os.path.dirname(n))       # _rels 的上一级
        try:
            xml = ctx.zf.read(n).decode("utf-8", "replace")
        except Exception:                                   # noqa: BLE001
            continue
        import re
        for m in re.finditer(r'<Relationship\b[^>]*>', xml):
            tag = m.group(0)
            if 'TargetMode="External"' in tag:
                continue
            tm = re.search(r'Target="([^"]+)"', tag)
            if not tm:
                continue
            target = tm.group(1)
            if target.startswith("/"):
                resolved = target.lstrip("/")
            else:
                resolved = os.path.normpath(os.path.join(base_dir, target)).replace("\\", "/")
            rels_map.setdefault(n, []).append((target, resolved))
            if resolved not in ctx.name_set:
                out.append(Finding("A3", "error",
                                   "关系目标不存在：%s" % target,
                                   location=n.split("/")[-1],
                                   detail={"target": target, "resolved": resolved}))
    return out


@rule("A4", "package", "媒体文件损坏或格式与扩展名不符", "error", True, False)
def r_a4_media(ctx: Ctx):
    out = []
    media = [n for n in ctx.names if n.startswith("word/media/") and not n.endswith("/")]
    if not media:
        return out
    bad = set(ctx.bad_crc_entries())
    for n in media:
        if n in bad:
            out.append(Finding("A4", "error", "媒体文件解压失败（数据损坏）",
                               location=n.split("/")[-1]))
            continue
        try:
            with ctx.zf.open(n) as f:
                head = f.read(16)
        except Exception:                                   # noqa: BLE001
            out.append(Finding("A4", "error", "媒体文件无法读取",
                               location=n.split("/")[-1]))
            continue
        ext = n.rsplit(".", 1)[-1].lower()
        ok = _magic_matches(ext, head)
        if ok is False:
            out.append(Finding("A4", "warning",
                               "扩展名 .%s 与文件实际格式不符" % ext,
                               location=n.split("/")[-1],
                               detail={"head": head.hex(" ")}))
    return out


def _magic_matches(ext: str, head: bytes):
    sigs = {
        "png": [b"\x89PNG"],
        "jpg": [b"\xff\xd8\xff"],
        "jpeg": [b"\xff\xd8\xff"],
        "gif": [b"GIF8"],
        "bmp": [b"BM"],
        "emf": [b"\x01\x00\x00\x00"],
        "wmf": [b"\xd7\xcd\xc6\x9a", b"\x01\x00\x09\x00"],
        "tif": [b"II*\x00", b"MM\x00*"],
        "tiff": [b"II*\x00", b"MM\x00*"],
        "svg": [b"<?xml", b"<svg"],
    }
    if ext not in sigs:
        return None                                        # 未知类型不判断
    return any(head.startswith(s) for s in sigs[ext])


# ================================================================== B 页面几何
@rule("B5", "geometry", "页边距或版心异常", "warning", True, False)
def r_b5_margin(ctx: Ctx):
    out = []
    P = ctx
    if P.usable_w.cm < 10:
        out.append(Finding("B5", "error",
                           "可用版心宽度仅 %.2fcm，正文几乎无法排版" % P.usable_w.cm,
                           location="页面设置",
                           detail={"usable_cm": round(P.usable_w.cm, 2)}))
    for name, val in (("左", P.margin_l), ("右", P.margin_r),
                      ("上", P.margin_t), ("下", P.margin_b)):
        if val.cm < 0:
            out.append(Finding("B5", "error",
                               "%s边距为负值 %.2fcm" % (name, val.cm),
                               location="页面设置"))
        elif val.cm > P.page_w.cm / 2:
            out.append(Finding("B5", "warning",
                               "%s边距 %.2fcm 超过页宽一半" % (name, val.cm),
                               location="页面设置"))
    if not out:
        out.append(Finding("B5", "info",
                           "版心 %.2fcm × %.2fcm，页边距正常"
                           % (P.usable_w.cm, P.usable_h.cm), location="页面设置"))
    return out


@rule("B7", "geometry", "表格总宽超出可用版心", "warning", True, True)
def r_b7_table_width(ctx: Ctx):
    out = []
    usable = int(ctx.usable_w)
    for idx, t in enumerate(ctx.doc.tables):
        tbl = t._tbl
        tblPr = tbl.tblPr

        # 布局模式决定"是否真的会溢出"：
        #   fixed → Word 严格按声明宽度画，必然溢出（确定）
        #   auto  → Word 会压缩回版心，但在 WPS / 在线预览 / 打印驱动下可能不同
        is_fixed = False
        if tblPr is not None:
            lay = tblPr.find(qn("w:tblLayout"))
            is_fixed = lay is not None and lay.get(qn("w:type")) == "fixed"

        # 1) tblW 声明值
        declared = None
        if tblPr is not None:
            tw = tblPr.find(qn("w:tblW"))
            if tw is not None and tw.get(qn("w:type")) == "dxa":
                try:
                    declared = int(tw.get(qn("w:w"))) * 635      # twips → EMU
                except (TypeError, ValueError):
                    declared = None
        # 2) tblGrid 求和
        grid_total = 0
        grid = tbl.find(qn("w:tblGrid"))
        if grid is not None:
            for gc in grid.findall(qn("w:gridCol")):
                try:
                    grid_total += int(gc.get(qn("w:w"))) * 635
                except (TypeError, ValueError):
                    pass
        # 3) column.width 兜底
        col_total = sum((c.width for c in t.columns if c.width), 0)

        width = max(declared or 0, grid_total, col_total)
        if width <= 0:
            continue
        cols = len(t.columns)
        if cols == 1:
            continue                                        # 单列多为引用块，不判
        if width <= usable + 2000:                          # 容差 ~0.06cm
            continue

        over_cm = Emu(width - usable).cm
        if is_fixed:
            msg = ("表格总宽 %.2fcm，超出可用版心 %.2fcm（溢出 %.2fcm）。"
                   "fixed 布局，渲染时必然溢出" % (Emu(width).cm, ctx.usable_w.cm, over_cm))
            sev = "error"
        else:
            msg = ("表格声明总宽 %.2fcm，超出可用版心 %.2fcm（超大 %.2fcm）。"
                   "自动布局下 Word 会压缩回版心，但 WPS / 在线预览 / 其他软件可能真溢出"
                   % (Emu(width).cm, ctx.usable_w.cm, over_cm))
            sev = "warning"

        out.append(Finding(
            "B7", sev, msg,
            location="第 %d 个表格（%d 列）" % (idx + 1, cols),
            detail={"table_index": idx, "width_cm": round(Emu(width).cm, 2),
                    "usable_cm": round(ctx.usable_w.cm, 2),
                    "overflow_cm": round(over_cm, 2),
                    "layout": "fixed" if is_fixed else "auto"}))
    return out


@rule("B8", "geometry", "图片显示宽度超出可用版心", "error", True, True)
def r_b8_image_width(ctx: Ctx):
    out = []
    usable = int(ctx.usable_w)
    for i, shp in enumerate(ctx.doc.inline_shapes):
        try:
            w = shp.width
        except Exception:                                   # noqa: BLE001
            continue
        if w and w > usable + 2000:
            out.append(Finding(
                "B8", "error",
                "第 %d 张图片宽 %.2fcm，超出可用版心 %.2fcm"
                % (i + 1, Emu(w).cm, ctx.usable_w.cm),
                location="内嵌图片 %d" % (i + 1),
                detail={"image_index": i, "width_cm": round(Emu(w).cm, 2),
                        "usable_cm": round(ctx.usable_w.cm, 2)}))
    return out


# ================================================================== C 字体系统
@rule("C10", "font", "文档引用的字体本机缺失", "warning", True, False)
def r_c10_font_missing(ctx: Ctx):
    out = []
    have = installed_fonts()
    if not have:
        return [Finding("C10", "info", "无法枚举本机字体，跳过缺失检查")]
    used = ctx.fonts_used
    missing = []
    for name, wheres in used.items():
        if name.strip().lower() in _PSEUDO_FONTS:
            continue
        if name in have:
            continue
        # 双名兜底：英文名/中文名互相认
        if name.lower() in {h.lower() for h in have}:
            continue
        missing.append((name, wheres))
    for name, wheres in missing:
        where = sorted(wheres)[:3]
        out.append(Finding(
            "C10", "warning",
            "字体「%s」本机未安装，Word 会静默替换为其他字体" % name,
            location=", ".join(where),
            detail={"font": name, "used_in": sorted(wheres)}))
    return out


@rule("C12", "font", "theme 属性覆盖显式字体名", "warning", False, False)
def r_c12_theme_font(ctx: Ctx):
    out = []
    targets = []
    used = ctx.used_style_names
    for style in ctx.doc.styles:
        if style.name not in used:
            continue                    # 未引用的内置样式（如 macro）不检查，避免误报
        rPr = style.element.find(qn("w:rPr"))
        if rPr is None:
            continue
        rf = rPr.find(qn("w:rFonts"))
        if rf is None:
            continue
        targets.append(("样式 " + style.name, rf))
    for p in ctx.doc.paragraphs[:200]:                        # 抽样，避免巨量输出
        for r in p.runs:
            rPr = r._element.rPr
            if rPr is None:
                continue
            rf = rPr.find(qn("w:rFonts"))
            if rf is not None:
                targets.append(("正文 run", rf))

    for where, rf in targets:
        theme_attrs = [a for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme")
                       if rf.get(qn(a))]
        explicit = [a for a in ("w:ascii", "w:hAnsi", "w:eastAsia")
                    if rf.get(qn(a))]
        if theme_attrs and explicit:
            out.append(Finding(
                "C12", "warning",
                "%s 同时声明了 theme 字体和显式字体名，theme 优先，显式字体不生效" % where,
                location=where,
                detail={"theme_attrs": theme_attrs, "explicit": explicit}))
            if len(out) >= 10:
                return out
    return out


@rule("C14", "font", "同类元素字体不一致", "warning", False, False)
def r_c14_font_consistency(ctx: Ctx):
    out = []
    body_fonts: dict[str, int] = {}
    style = ctx.doc.styles["Normal"]
    if style.font.size:
        body_pt = style.font.size.pt
    else:
        body_pt = 10.5
    for p in ctx.doc.paragraphs:
        if not p.text.strip() or p.style.name != "Normal":
            continue
        pPr = p._p.pPr
        if pPr is not None and pPr.find(qn("w:numPr")) is not None:
            continue
        is_title = any(r.font.size and r.font.size.pt > body_pt * 1.5 for r in p.runs)
        if is_title:
            continue
        for r in p.runs:
            rPr = r._element.rPr
            if rPr is None:
                continue
            rf = rPr.find(qn("w:rFonts"))
            if rf is None:
                continue
            ea = rf.get(qn("w:eastAsia"))
            if ea:
                body_fonts[ea] = body_fonts.get(ea, 0) + 1
    if len(body_fonts) > 1:
        ranked = sorted(body_fonts.items(), key=lambda kv: -kv[1])
        out.append(Finding(
            "C14", "warning",
            "正文使用了 %d 种中文字体：%s"
            % (len(ranked), "、".join("%s(%d)" % (k, v) for k, v in ranked)),
            location="正文",
            detail={"fonts": dict(ranked)}))
    return out


# ================================================================== D 流式布局
@rule("D15", "layout", "列宽总和与 tblGrid 声明不符", "warning", True, False)
def r_d15_grid_mismatch(ctx: Ctx):
    out = []
    bad = []
    for idx, t in enumerate(ctx.doc.tables):
        tbl = t._tbl
        grid = tbl.find(qn("w:tblGrid"))
        if grid is None:
            continue
        grid_vals = []
        for gc in grid.findall(qn("w:gridCol")):
            try:
                grid_vals.append(int(gc.get(qn("w:w"))))
            except (TypeError, ValueError):
                grid_vals.append(0)
        grid_sum = sum(grid_vals)
        if grid_sum <= 0:
            continue
        # 首行各单元格 tcW 求和
        tcw_vals = []
        first_row = t.rows[0]
        for cell in first_row.cells:
            tcPr = cell._tc.tcPr
            w = 0
            if tcPr is not None:
                e = tcPr.find(qn("w:tcW"))
                if e is not None and e.get(qn("w:type")) == "dxa":
                    try:
                        w = int(e.get(qn("w:w")))
                    except (TypeError, ValueError):
                        w = 0
            tcw_vals.append(w)
        if any(v == 0 for v in tcw_vals):
            continue                                        # 未显式声明，不判
        tcw_sum = sum(tcw_vals)
        if abs(tcw_sum - grid_sum) > max(20, grid_sum * 0.02):
            bad.append((idx + 1, grid_sum, tcw_sum))

    if not bad:
        return out
    # 同一文档里多个表格的同一问题合并为一条，避免重复刷屏
    idxs = [b[0] for b in bad]
    g0, t0 = bad[0][1], bad[0][2]
    if len(bad) == 1:
        msg = ("tblGrid 声明总宽 %.2fcm，但单元格实际声明 %.2fcm（差 %.2fcm）"
               % (g0 * 635 / 360000, t0 * 635 / 360000, (g0 - t0) * 635 / 360000))
        loc = "第 %d 个表格" % idxs[0]
    else:
        msg = ("%d 个表格的 tblGrid 与单元格列宽声明不一致"
               "（第 %s 个；例如声明 %.2fcm / 实际 %.2fcm）"
               % (len(bad), "、".join(str(i) for i in idxs),
                  g0 * 635 / 360000, t0 * 635 / 360000))
        loc = "、".join("第 %d 个表格" % i for i in idxs)
    out.append(Finding(
        "D15", "warning", msg, location=loc,
        detail={"tables": [{"index": b[0] - 1, "grid_twips": b[1], "tcw_twips": b[2]}
                           for b in bad]}))
    return out


@rule("D17", "layout", "多行表格缺少跨页表头重复标记", "info", False, True)
def r_d17_header_repeat(ctx: Ctx):
    out = []
    for idx, t in enumerate(ctx.doc.tables):
        if len(t.columns) <= 1 or len(t.rows) < 5:
            continue                                        # 短表不易跨页，不判
        first = t.rows[0]._tr
        trPr = first.find(qn("w:trPr"))
        has = trPr is not None and trPr.find(qn("w:tblHeader")) is not None
        if not has:
            out.append(Finding(
                "D17", "info",
                "%d 行表格未标记重复表头，跨页后第二页无表头" % len(t.rows),
                location="第 %d 个表格" % (idx + 1),
                detail={"table_index": idx, "rows": len(t.rows)}))
    return out


# ================================================================== E 引用编号
@rule("E20", "refs", "编号 numId 跨批复用（导致编号串号）", "error", True, True)
def r_e20_numbering_reuse(ctx: Ctx):
    seq: list[tuple[int, str]] = []
    for i, p in enumerate(ctx.doc.paragraphs):
        pPr = p._p.pPr
        if pPr is None:
            continue
        npr = pPr.find(qn("w:numPr"))
        if npr is None:
            continue
        nid_el = npr.find(qn("w:numId"))
        if nid_el is None:
            continue
        nid = nid_el.get(qn("w:val"))
        if nid:
            seq.append((i, nid))

    # 连续相同 numId 归为一批
    batches: list[list] = []
    for idx, nid in seq:
        if batches and batches[-1][0] == nid and batches[-1][2] == idx - 1:
            batches[-1][1].append(idx)
            batches[-1][2] = idx
        else:
            batches.append([nid, [idx], idx])

    by_nid: dict[str, int] = {}
    for nid, _idxes, _last in batches:
        by_nid[nid] = by_nid.get(nid, 0) + 1

    out = []
    for nid, count in by_nid.items():
        if count > 1:
            out.append(Finding(
                "E20", "error",
                "编号 numId=%s 被 %d 批列表共用，第 2 批起会接着上一批继续编号（串号）"
                % (nid, count),
                location="列表 numId=%s" % nid,
                detail={"num_id": nid, "batch_count": count}))
    return out
