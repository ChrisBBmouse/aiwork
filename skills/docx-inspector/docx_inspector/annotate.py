# -*- coding: utf-8 -*-
"""docx_inspector.annotate — 把检测结果画到渲染图上

为什么用渲染结果而不是 OOXML 声明做坐标：
    声明说"表格宽 14.66cm"，渲染才是真相。两者不一致时，看图的用户
    相信图。所以标注框一律从 PDF 里量出来（与 gongwen 的坐标量测同思路）。

链路：Word COM → PDF → PyMuPDF 量坐标 → Pillow 画框 → PNG
后端不可用时抛异常，由调用方降级处理。
"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile

PT_PER_CM = 28.3464567


# ------------------------------------------------------------------ 渲染
def _to_pdf(docx_path: str, pdf_path: str) -> str:
    """返回使用的后端名：com / soffice"""
    # 1) Windows Word COM
    try:
        import win32com.client as win32
        app = win32.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        try:
            doc = app.Documents.Open(os.path.abspath(docx_path), ReadOnly=True,
                                     AddToRecentFiles=False)
            try:
                doc.ExportAsFixedFormat(os.path.abspath(pdf_path), 17)
            finally:
                doc.Close(0)
        finally:
            app.Quit()
        return "com"
    except Exception:                                   # noqa: BLE001
        pass

    # 2) LibreOffice
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        outdir = os.path.dirname(os.path.abspath(pdf_path))
        subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                        "--outdir", outdir, os.path.abspath(docx_path)],
                       capture_output=True, timeout=180)
        auto = os.path.join(
            outdir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
        if os.path.exists(auto) and os.path.abspath(auto) != os.path.abspath(pdf_path):
            os.replace(auto, pdf_path)
        if os.path.exists(pdf_path):
            return "soffice"

    raise RuntimeError("没有可用的渲染后端（需要 Word COM 或 LibreOffice）")


# ------------------------------------------------------------------ 标注
def _norm(s):
    return os.path.abspath(s)


def _cluster(boxes, gap_pt=6.0):
    """把纵向相邻的越界元素合并成簇，避免画出几十个小框"""
    groups = []
    for b in sorted(boxes, key=lambda x: x[1]):
        for g in groups:
            if not (b[1] > g[3] + gap_pt or b[3] < g[1] - gap_pt):
                g[0] = min(g[0], b[0])
                g[1] = min(g[1], b[1])
                g[2] = max(g[2], b[2])
                g[3] = max(g[3], b[3])
                break
        else:
            groups.append(list(b))
    return groups


def annotate(ctx, report, out_dir: str, dpi: int = 120) -> list[str]:
    import pymupdf
    from PIL import Image, ImageDraw, ImageFont

    os.makedirs(out_dir, exist_ok=True)
    tmp_pdf = os.path.join(tempfile.gettempdir(),
                           "_docx_inspector_%d.pdf" % os.getpid())
    backend = _to_pdf(ctx.path, tmp_pdf)

    # 版心（PDF points）
    page_w_pt = ctx.page_w.cm * PT_PER_CM
    left_pt = ctx.margin_l.cm * PT_PER_CM
    right_pt = page_w_pt - ctx.margin_r.cm * PT_PER_CM

    # 待标注的几何问题
    geo = [f for f in report.findings if f.rule in ("B7", "B8") and f.severity == "error"]
    n_err = report.counts()["error"]
    n_warn = report.counts()["warning"]

    font = None
    for cand in (r"C:\Windows\Fonts\simhei.ttf",
                 r"C:\Windows\Fonts\msyh.ttc",
                 r"C:\Windows\Fonts\arial.ttf"):
        if os.path.exists(cand):
            try:
                font = ImageFont.truetype(cand, 22)
                small = ImageFont.truetype(cand, 16)
                break
            except Exception:                           # noqa: BLE001
                continue

    made = []
    doc = pymupdf.open(tmp_pdf)
    try:
        for pno, page in enumerate(doc):
            pix = page.get_pixmap(dpi=dpi)
            png_path = os.path.join(out_dir, "page-%d.png" % (pno + 1))
            pix.save(png_path)

            img = Image.open(png_path).convert("RGB")
            drw = ImageDraw.Draw(img, "RGBA")
            sx = dpi / 72.0
            bar_h = 34

            # 版心右边界参考线
            rx = right_pt * sx
            for y in range(0, img.height, 12):
                drw.line([(rx, y), (rx, min(y + 6, img.height))],
                         fill=(120, 120, 120, 140), width=1)

            # 量出真正越界的元素（只看横向溢出）
            boxes = []
            for b in page.get_text("dict")["blocks"]:
                for line in b.get("lines", []):
                    for span in line.get("spans", []):
                        if not span.get("text", "").strip():
                            continue
                        x0, y0, x1, y1 = span["bbox"]
                        if x1 > right_pt + 1.0:
                            boxes.append((x0, y0, x1, y1))
            for dr in page.get_drawings():
                r = dr["rect"]
                if r.x1 > right_pt + 1.0 and r.width < page_w_pt:
                    boxes.append((r.x0, r.y0, r.x1, r.y1))

            for grp in _cluster(boxes):
                bx0, by0, bx1, by1 = grp
                left_x = min(bx0, right_pt) * sx - 2      # 从版心边界起画，直观看出伸出去多少
                drw.rectangle([left_x, by0 * sx - 4, bx1 * sx + 6, by1 * sx + 4],
                              outline=(220, 38, 38, 240), width=4)
                over_cm = (bx1 - right_pt) / PT_PER_CM
                if over_cm > 0 and font:
                    ty = max(by0 * sx - 26, bar_h + 2)
                    drw.rectangle([left_x, ty - 2, left_x + 150, ty + 20],
                                  fill=(220, 38, 38, 220))
                    drw.text((left_x + 6, ty), "溢出 %.2f cm" % over_cm,
                             fill=(255, 255, 255), font=small)

            # 顶部状态条
            if n_err:
                drw.rectangle([0, 0, img.width, bar_h], fill=(220, 38, 38, 235))
                text = "ERROR %d · WARN %d · %s" % (n_err, n_warn, os.path.basename(ctx.path))
            elif n_warn:
                drw.rectangle([0, 0, img.width, bar_h], fill=(217, 119, 6, 235))
                text = "WARN %d · %s" % (n_warn, os.path.basename(ctx.path))
            else:
                drw.rectangle([0, 0, img.width, bar_h], fill=(22, 163, 74, 235))
                text = "PASS · %s" % os.path.basename(ctx.path)

            if font:
                drw.text((12, 6), text, fill=(255, 255, 255), font=small)

            # 底部说明
            if geo or boxes:
                n_box = len(_cluster(boxes)) if boxes else 0
                label = ("红框 = 渲染后实际越过版心右边界的区域（本页 %d 处）"
                         % n_box if n_box else "本页无横向溢出")
                drw.rectangle([0, img.height - 28, img.width, img.height],
                              fill=(255, 255, 255, 220))
                if font:
                    drw.text((12, img.height - 24), label,
                             fill=(80, 80, 80), font=small)

            img.save(png_path)
            made.append(png_path)
    finally:
        doc.close()
        try:
            os.remove(tmp_pdf)
        except OSError:
            pass

    return made
