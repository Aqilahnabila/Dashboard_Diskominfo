# -*- coding: utf-8 -*-
"""
report.py
---------
Generator Laporan Analisis Data CCTV Lalu Lintas (format Word .docx).

Struktur laporan mengikuti laporan pedoman (Laporan Analisis Data Pengunjung PSC 2025):
    Sampul -> Kata Pengantar -> Daftar Isi
    BAB I   Pendahuluan          (1.1 Latar Belakang, 1.2 Tujuan)
    BAB II  Metodologi           (2.1 Data, 2.2 Metode Analisis)
    BAB III Hasil dan Pembahasan (3.1 Statistik Deskriptif, 3.2 Tren, 3.3 Perbandingan
                                  Lokasi, 3.4 Komposisi, 3.5 Kualitas Data)
    BAB IV  Insight dan Rekomendasi
    BAB V   Kesimpulan
    Lampiran

Seluruh angka, grafik, dan tabel DIHITUNG LANGSUNG dari data yang sedang aktif di dashboard
(app.py): KPI, tren total, tren per lokasi, peringkat lokasi tertinggi/terendah, komposisi
kategori (tabel, distribusi, per lokasi), dan status kualitas data. Narasi, insight, dan
rekomendasi dibangkitkan dari hasil perhitungan tersebut (bukan teks tetap), sehingga
berubah mengikuti data & filter.

Antarmuka tidak berubah (kompatibel dengan app.py):
    generate_docx_report(df, dq, meta) -> io.BytesIO
"""

import io
import re
import base64
import textwrap
import calendar
from datetime import datetime

import numpy as np
import pandas as pd

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from matplotlib.patches import Rectangle, Wedge, Circle

from docx import Document
from docx.shared import Pt, Cm, Emu, RGBColor
from docx.enum.text import (
    WD_ALIGN_PARAGRAPH,
    WD_LINE_SPACING,
    WD_TAB_ALIGNMENT,
    WD_TAB_LEADER,
)
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml

from parser import CATEGORIES, BULAN_ID, format_tanggal_id

# ============================================================================
# KONFIGURASI
# ============================================================================
INSTANSI = "Dinas Komunikasi dan Informatika Kota Madiun"
KOTA = "Madiun"
JABATAN_PENANDATANGAN = "Plt. Kepala Dinas Komunikasi dan Informatika\nKota Madiun"
NAMA_PENANDATANGAN = "Noor Aflah, S.Kom."
NIP_PENANDATANGAN = "197609072003121007"

# Palet warna (mengikuti laporan pedoman: teal, oranye, hijau, navy)
TEAL = "12B3AA"
TEAL_DARK = "0B7F79"
TEAL_LIGHT = "E6F6F5"
ORANGE = "F59E32"
ORANGE_LIGHT = "FEF3E2"
NAVY = "0B3D5C"
GREEN = "6BBF59"
RED = "C0392B"
RED_LIGHT = "FDECEC"
GREY = "5A5A5A"
GREY_LIGHT = "F2F5F7"
LINE = "D5DCE0"

FONT_BODY = "Times New Roman"
FONT_HEAD = "Arial"

# Warna grafik: disamakan dengan tampilan dashboard (Streamlit / Altair)
BAR_COLOR = "#4C78A8"  # batang default Altair
LINE_SINGLE = "#FF4B4B"  # garis tunggal st.line_chart
LINE_MULTI = [
    "#0068C9",
    "#83C9FF",
    "#FF2B2B",
    "#FFABAB",
    "#29B09D",
    "#7DEFA1",
    "#FF8700",
    "#FFD16A",
    "#6D3FC0",
    "#D5DAE5",
]
CAT_COLORS = {
    "Bus": "#4C78A8",
    "Mobil": "#F58518",
    "Motor": "#E45756",
    "Orang": "#72B7B2",
    "Truk": "#54A24B",
}

ST_INACTIVE = "Tidak ada data / kamera tidak aktif"
ST_PARTIAL = "Sebagian data kosong"
ST_MISMATCH = "Total tidak cocok"
ST_ZERO = "Perlu verifikasi nilai nol"
ST_OK = "Data lengkap"

STATUS_FILL = {
    ST_OK: "E4F5E9",
    ST_PARTIAL: "FFF3D9",
    ST_MISMATCH: "FDE5D0",
    ST_ZERO: "FDE5D0",
    ST_INACTIVE: "FBE7E4",
}
STATUS_FONT = {
    ST_OK: "1E7B3E",
    ST_PARTIAL: "8A6100",
    ST_MISMATCH: "A8480B",
    ST_ZERO: "A8480B",
    ST_INACTIVE: "B5391F",
}

BULAN_SHORT = [
    "",
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "Mei",
    "Jun",
    "Jul",
    "Agu",
    "Sep",
    "Okt",
    "Nov",
    "Des",
]

TEXT_WIDTH_CM = 15.5


def _rgb(hex_):
    return RGBColor.from_string(hex_.upper())


# ============================================================================
# Format angka gaya Indonesia
# ============================================================================
def _fmt_id(value, decimals=0):
    """Ribuan '.', desimal ','  (mis. 1.234.567 / 42.667,3)."""
    if (
        value is None
        or (isinstance(value, float) and np.isnan(value))
        or pd.isna(value)
    ):
        return "-"
    s = f"{float(value):,.{decimals}f}"
    return s.replace(",", "§").replace(".", ",").replace("§", ".")


def _p(x, d=1):
    """Persen gaya Indonesia: 44,3%."""
    return _fmt_id(x, d) + "%"


def _n(x, d=1):
    """Angka desimal gaya Indonesia tanpa ribuan: 2,4."""
    return f"{x:.{d}f}".replace(".", ",")


def _compact(v):
    """Format ringkas untuk label grafik: 11,6 jt / 42,7 rb / 9.492."""
    v = float(v)
    if abs(v) >= 1e6:
        s = (
            f"{v / 1e6:.1f}".rstrip("0").rstrip(".")
            if v / 1e6 < 100
            else f"{v / 1e6:.0f}"
        )
        return s.replace(".", ",") + " jt"
    if abs(v) >= 1e4:
        s = f"{v / 1e3:.1f}".rstrip("0").rstrip(".")
        return s.replace(".", ",") + " rb"
    return _fmt_id(v)


def _join_id(items, conj="dan"):
    items = [str(i) for i in items]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} {conj} {items[1]}"
    return ", ".join(items[:-1]) + f", {conj} {items[-1]}"


def _list_ex(items, n=3):
    """Daftar ringkas: 'A, B, dan C' atau 'A, B, C, dan lainnya' bila lebih dari n."""
    items = [str(i) for i in items]
    if len(items) <= n:
        return _join_id(items)
    return ", ".join(items[:n]) + ", dan lainnya"


def _roman(n):
    return ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII"][n]


# ============================================================================
# Helper python-docx (level rendah)
# ============================================================================
def _set_run_font(
    run, name=None, size=None, bold=None, italic=None, color=None, spacing=None
):
    if name:
        run.font.name = name
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        for att in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rfonts.set(qn(att), name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if color:
        run.font.color.rgb = _rgb(color)
    if spacing is not None:  # letter-spacing (twentieths of a point)
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), str(int(spacing)))
        run._element.get_or_add_rPr().append(sp)


def _style_set_fonts(style, name):
    style.font.name = name
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for att in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(att), name)
    for att in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(att)) is not None:
            del rfonts.attrib[qn(att)]


def _para_border(paragraph, side="bottom", color=TEAL, sz=12, space=4):
    ppr = paragraph._p.get_or_add_pPr()
    pbdr = ppr.find(qn("w:pBdr"))
    if pbdr is None:
        pbdr = OxmlElement("w:pBdr")
        ppr.append(pbdr)
    el = OxmlElement(f"w:{side}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), str(space))
    el.set(qn("w:color"), color)
    pbdr.append(el)


def _style_para_border(style, side="bottom", color=TEAL, sz=12, space=4):
    ppr = style.element.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    el = OxmlElement(f"w:{side}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), str(space))
    el.set(qn("w:color"), color)
    pbdr.append(el)
    ppr.append(pbdr)


def _cell_bg(cell, hex_):
    tcpr = cell._tc.get_or_add_tcPr()
    for old in tcpr.findall(qn("w:shd")):
        tcpr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_)
    tcpr.append(shd)


def _cell_borders(cell, **kw):
    """kw: top/bottom/left/right = (sz, color) atau None -> nil."""
    tcpr = cell._tc.get_or_add_tcPr()
    tcb = tcpr.find(qn("w:tcBorders"))
    if tcb is None:
        tcb = OxmlElement("w:tcBorders")
        tcpr.append(tcb)
    for side in ("top", "left", "bottom", "right"):
        if side in kw:
            el = OxmlElement(f"w:{side}")
            v = kw[side]
            if v is None:
                el.set(qn("w:val"), "nil")
            else:
                el.set(qn("w:val"), "single")
                el.set(qn("w:sz"), str(v[0]))
                el.set(qn("w:space"), "0")
                el.set(qn("w:color"), v[1])
            tcb.append(el)


def _cell_margins(cell, top=60, bottom=60, left=90, right=90):
    tcpr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for side, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(v))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcpr.append(mar)


def _table_borders(table, color=LINE, sz=4, inside=True):
    tblpr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    sides = ["top", "left", "bottom", "right"] + (
        ["insideH", "insideV"] if inside else []
    )
    for side in sides:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblpr.append(borders)


def _table_no_borders(table):
    tblpr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblpr.append(borders)


def _table_fixed(table, widths_cm):
    table.autofit = False
    for row in table.rows:
        for i, w in enumerate(widths_cm):
            row.cells[i].width = Cm(w)
    # tblGrid
    grid = table._tbl.tblGrid
    for i, gc in enumerate(grid.findall(qn("w:gridCol"))):
        if i < len(widths_cm):
            gc.set(qn("w:w"), str(int(widths_cm[i] * 567)))


def _row_no_split(row, header=False):
    trpr = row._tr.get_or_add_trPr()
    cs = OxmlElement("w:cantSplit")
    trpr.append(cs)
    if header:
        th = OxmlElement("w:tblHeader")
        trpr.append(th)


def _keep_next(paragraph, on=True):
    paragraph.paragraph_format.keep_with_next = on


def _add_field(paragraph, instr, font=None, size=None, color=None, bold=None):
    """Sisipkan field Word (PAGE, dll)."""

    def _run():
        r = paragraph.add_run()
        _set_run_font(r, font, size, bold, None, color)
        return r

    r1 = _run()
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), "begin")
    r1._r.append(fc)
    r2 = _run()
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    r2._r.append(it)
    r3 = _run()
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), "separate")
    r3._r.append(fc)
    r4 = _run()
    r4.text = "1"
    r5 = _run()
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), "end")
    r5._r.append(fc)


_BOLD_RE = re.compile(r"(\*\*.+?\*\*)")


def _rich(
    paragraph, text, font=None, size=None, color=None, italic=None, bold_all=None
):
    """Tambahkan teks; bagian **tebal** dirender sebagai bold."""
    for part in _BOLD_RE.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            r = paragraph.add_run(part[2:-2])
            _set_run_font(r, font, size, True, italic, color)
        else:
            r = paragraph.add_run(part)
            _set_run_font(r, font, size, bold_all, italic, color)
    return paragraph


# ============================================================================
# Helper tingkat dokumen
# ============================================================================
class _Counter:
    def __init__(self):
        self.tabel = 0
        self.gambar = 0


def _setup_document():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin, sec.right_margin = Cm(3.0), Cm(2.5)
    sec.top_margin, sec.bottom_margin = Cm(2.5), Cm(2.3)
    sec.header_distance, sec.footer_distance = Cm(1.1), Cm(1.0)
    sec.different_first_page_header_footer = True  # sampul tanpa header/footer

    st = doc.styles
    n = st["Normal"]
    _style_set_fonts(n, FONT_BODY)
    n.font.size = Pt(12)
    n.paragraph_format.space_after = Pt(6)
    n.paragraph_format.line_spacing = 1.3

    h1 = st["Heading 1"]
    _style_set_fonts(h1, FONT_HEAD)
    h1.font.size = Pt(20)
    h1.font.bold = True
    h1.font.color.rgb = _rgb(TEAL_DARK)
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.page_break_before = True
    h1.paragraph_format.space_before = Pt(0)
    h1.paragraph_format.space_after = Pt(16)
    h1.paragraph_format.keep_with_next = True
    h1.paragraph_format.line_spacing = 1.1
    _style_para_border(h1, "bottom", ORANGE, 18, 8)

    h2 = st["Heading 2"]
    _style_set_fonts(h2, FONT_HEAD)
    h2.font.size = Pt(13)
    h2.font.bold = True
    h2.font.color.rgb = _rgb(TEAL_DARK)
    h2.paragraph_format.space_before = Pt(16)
    h2.paragraph_format.space_after = Pt(6)
    h2.paragraph_format.keep_with_next = True
    h2.paragraph_format.line_spacing = 1.1

    # Gaya TOC (dipakai Word saat Daftar Isi diperbarui)
    for name, left, bold in (("TOC 1", 0, True), ("TOC 2", 0.8, False)):
        try:
            s = st.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        except Exception:
            s = st[name]
        s.base_style = st["Normal"]
        _style_set_fonts(s, FONT_HEAD)
        s.font.size = Pt(10.5)
        s.font.bold = bold
        s.paragraph_format.left_indent = Cm(left)
        s.paragraph_format.space_after = Pt(3 if not bold else 4)
        s.paragraph_format.space_before = Pt(6 if bold else 0)
        s.paragraph_format.line_spacing = 1.1
        s.paragraph_format.tab_stops.add_tab_stop(
            Cm(TEXT_WIDTH_CM), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
        )
    return doc


def _force_update_fields_on_open(doc):
    """Word akan memperbarui field (Daftar Isi) saat dokumen dibuka."""
    settings = doc.settings.element
    uf = OxmlElement("w:updateFields")
    uf.set(qn("w:val"), "true")
    anchors = [
        "hdrShapeDefaults",
        "footnotePr",
        "endnotePr",
        "compat",
        "docVars",
        "rsids",
        "mathPr",
        "themeFontLang",
        "clrSchemeMapping",
        "shapeDefaults",
        "decimalSymbol",
        "listSeparator",
    ]
    for a in anchors:
        el = settings.find(qn(f"w:{a}"))
        if el is not None:
            el.addprevious(uf)
            return
    settings.append(uf)


def _body(
    doc,
    text,
    indent=True,
    size=None,
    after=6,
    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    color=None,
    italic=None,
    keep_next=False,
):
    p = doc.add_paragraph()
    p.alignment = align
    if indent:
        p.paragraph_format.first_line_indent = Cm(1.0)
    p.paragraph_format.space_after = Pt(after)
    _rich(p, text, size=size, color=color, italic=italic)
    if keep_next:
        _keep_next(p)
    return p


def _h1(doc, num, title, label=None):
    p = doc.add_paragraph(style="Heading 1")
    r1 = p.add_run(label or f"BAB {_roman(num)}")
    _set_run_font(r1, FONT_HEAD, 12, True, None, ORANGE, spacing=60)
    r1.add_break()
    r2 = p.add_run(title.upper())
    _set_run_font(r2, FONT_HEAD, 20, True, None, TEAL_DARK)
    return p


def _h2(doc, text):
    return doc.add_paragraph(text, style="Heading 2")


def _h3(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    _keep_next(p)
    r = p.add_run(text)
    _set_run_font(r, FONT_HEAD, 11, True, None, NAVY)
    return p


def _bullet(doc, text, level_indent=1.0, size=None, after=3, marker="•"):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(level_indent + 0.6)
    p.paragraph_format.first_line_indent = Cm(-0.6)
    p.paragraph_format.space_after = Pt(after)
    r = p.add_run(f"{marker}\t")
    _set_run_font(r, None, size, True, None, TEAL)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(level_indent + 0.6))
    _rich(p, text, size=size)
    return p


def _caption_table(doc, cnt, text):
    cnt.tabel += 1
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    _keep_next(p)
    r = p.add_run(f"Tabel {cnt.tabel}. ")
    _set_run_font(r, FONT_HEAD, 9.5, True, None, TEAL_DARK)
    r = p.add_run(text)
    _set_run_font(r, FONT_HEAD, 9.5, False, None, GREY)
    return p


def _caption_figure(doc, cnt, text):
    cnt.gambar += 1
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(f"Gambar {cnt.gambar}. ")
    _set_run_font(r, FONT_HEAD, 9.5, True, None, TEAL_DARK)
    r = p.add_run(text)
    _set_run_font(r, FONT_HEAD, 9.5, False, None, GREY)
    return p


def _figure(doc, stream, width_cm=TEXT_WIDTH_CM):
    doc.add_picture(stream, width=Cm(width_cm))
    p = doc.paragraphs[-1]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0
    _keep_next(p)
    return p


def _table(
    doc,
    headers,
    rows,
    widths_cm,
    aligns=None,
    font_size=8.5,
    header_bg=TEAL,
    zebra=True,
    cell_styler=None,
    bold_last_row=False,
):
    """Tabel berwarna: header teal, baris zebra. cell_styler(r, c, value)->dict(bg, color, bold)."""
    aligns = aligns or ["left"] * len(headers)
    amap = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _table_borders(t)
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = ""
        _cell_bg(c, header_bg)
        _cell_margins(c, 70, 70, 80, 80)
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run(str(h))
        _set_run_font(r, FONT_HEAD, font_size, True, None, "FFFFFF")
    _row_no_split(t.rows[0], header=True)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        _row_no_split(t.rows[-1])
        last = bold_last_row and ri == len(rows) - 1
        for ci, val in enumerate(row):
            c = cells[ci]
            c.text = ""
            _cell_margins(c, 45, 45, 80, 80)
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            style = {}
            if cell_styler:
                style = cell_styler(ri, ci, val) or {}
            bg = style.get("bg") or (TEAL_LIGHT if (zebra and ri % 2 == 1) else None)
            if last:
                bg = style.get("bg") or ORANGE_LIGHT
            if bg:
                _cell_bg(c, bg)
            p = c.paragraphs[0]
            p.alignment = amap[aligns[ci]]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            r = p.add_run(str(val))
            _set_run_font(
                r,
                FONT_HEAD,
                font_size,
                style.get("bold", last),
                None,
                style.get("color"),
            )
    _table_fixed(t, widths_cm)
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(4)
    sp.paragraph_format.line_spacing = 0.6
    return t


def _kpi_cards(doc, kpis):
    """kpis: list (label, value, sub|None). Kartu KPI seperti dashboard."""
    n = len(kpis)
    widths = [TEXT_WIDTH_CM / n] * n
    if n == 5:
        widths = [3.0, 2.9, 2.5, 2.2, 4.9]
    t = doc.add_table(rows=1, cols=n)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _table_no_borders(t)
    for i, (label, value, sub) in enumerate(kpis):
        c = t.rows[0].cells[i]
        c.text = ""
        _cell_bg(c, TEAL_LIGHT)
        _cell_borders(
            c,
            top=(24, TEAL if i % 2 == 0 else ORANGE),
            left=(12, "FFFFFF"),
            right=(12, "FFFFFF"),
        )
        _cell_margins(c, 90, 90, 70, 70)
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run(label)
        _set_run_font(r, FONT_HEAD, 8, True, None, GREY)
        p2 = c.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after = Pt(0)
        p2.paragraph_format.line_spacing = 1.0
        small = len(str(value)) > 12
        r = p2.add_run(str(value))
        _set_run_font(r, FONT_HEAD, 10.5 if small else 14, True, None, TEAL_DARK)
        if sub:
            p3 = c.add_paragraph()
            p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p3.paragraph_format.space_after = Pt(0)
            p3.paragraph_format.line_spacing = 1.0
            r = p3.add_run(sub)
            _set_run_font(r, FONT_HEAD, 8.5, False, None, ORANGE)
    _table_fixed(t, widths)
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(4)
    return t


def _callout(
    doc,
    title,
    body,
    accent=ORANGE,
    bg=ORANGE_LIGHT,
    title_color=None,
    tag=None,
    tag_color=None,
    body_size=10.5,
):
    """Kotak bergaris-aksen kiri (untuk catatan / rekomendasi)."""
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _table_no_borders(t)
    c = t.rows[0].cells[0]
    c.text = ""
    _cell_bg(c, bg)
    _cell_borders(c, left=(36, accent))
    _cell_margins(c, 90, 90, 160, 140)
    _row_no_split(t.rows[0])
    p = c.paragraphs[0]
    if tag:
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.keep_with_next = True
        r = p.add_run(tag)
        _set_run_font(r, FONT_HEAD, 7.5, True, None, tag_color or accent, spacing=30)
        p = c.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.1
    p.paragraph_format.keep_with_next = True
    r = p.add_run(title)
    _set_run_font(r, FONT_HEAD, 10.5, True, None, title_color or NAVY)
    chunks = body if isinstance(body, list) else [body]
    for ch in chunks:
        p2 = c.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p2.paragraph_format.space_after = Pt(3)
        p2.paragraph_format.line_spacing = 1.15
        _rich(p2, ch, font=FONT_BODY, size=body_size)
    _table_fixed(t, [TEXT_WIDTH_CM])
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(2)
    sp.paragraph_format.line_spacing = 0.5
    return t


def _add_anchored_picture(paragraph, stream, width_cm, height_cm, x_cm=0, y_cm=0):
    """Gambar latar (di belakang teks), diposisikan relatif terhadap halaman."""
    run = paragraph.add_run()
    shape = run.add_picture(stream, width=Cm(width_cm), height=Cm(height_cm))
    inline = shape._inline
    graphic = inline.find(qn("a:graphic"))
    docpr = inline.find(qn("wp:docPr"))
    ns = (
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    )
    anchor = parse_xml(
        f'<wp:anchor {ns} distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" '
        f'behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{int(Cm(x_cm))}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{int(Cm(y_cm))}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{int(Cm(width_cm))}" cy="{int(Cm(height_cm))}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{docpr.get("id")}" name="Latar Sampul"/>'
        f"<wp:cNvGraphicFramePr/></wp:anchor>"
    )
    anchor.append(graphic)
    drawing = inline.getparent()
    drawing.remove(inline)
    drawing.append(anchor)


def _fig_to_stream(fig, dpi=170):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


# ============================================================================
# Agregasi periodik (mengikuti pilihan periode di dashboard)
# ============================================================================
def _period_key(d: pd.DataFrame, period: str) -> pd.Series:
    if period == "Harian":
        return d["Tanggal"].dt.strftime("%Y-%m-%d")
    if period == "Mingguan":
        iso = d["Tanggal"].dt.isocalendar()
        return iso["year"].astype(str) + "-W" + iso["week"].astype(str).str.zfill(2)
    if period == "Bulanan":
        return d["Tanggal"].dt.strftime("%Y-%m")
    return pd.Series("ALL", index=d.index)


def _periodic_series(df: pd.DataFrame, period: str) -> pd.DataFrame:
    """Total seluruh lokasi per periode, urut kronologis. Kolom: _key, Tanggal, Label, kategori, Total."""
    d = df.copy()
    d["Tanggal"] = pd.to_datetime(d["Tanggal"])
    d["_key"] = _period_key(d, period)
    agg = {c: "sum" for c in CATEGORIES + ["Total"]}
    agg["Tanggal"] = "min"
    g = (
        d.groupby("_key", as_index=False)
        .agg(agg)
        .sort_values("Tanggal")
        .reset_index(drop=True)
    )
    if period == "Harian":
        g["Label"] = g["Tanggal"].apply(format_tanggal_id)
    elif period == "Mingguan":
        g["Label"] = g["_key"].apply(
            lambda s: f"Minggu ke-{int(s.split('-W')[1])} {s.split('-W')[0]}"
        )
    elif period == "Bulanan":
        g["Label"] = g["Tanggal"].apply(lambda x: f"{BULAN_ID[x.month]} {x.year}")
    else:
        lo, hi = d["Tanggal"].min(), d["Tanggal"].max()
        g["Label"] = f"{format_tanggal_id(lo)} – {format_tanggal_id(hi)}"
    return g


# ============================================================================
# Pasca-proses: urutkan elemen XML sesuai skema OOXML (agar Word tidak menolak file)
# ============================================================================
_ORDERS = {
    "pPr": [
        "pStyle",
        "keepNext",
        "keepLines",
        "pageBreakBefore",
        "framePr",
        "widowControl",
        "numPr",
        "suppressLineNumbers",
        "pBdr",
        "shd",
        "tabs",
        "suppressAutoHyphens",
        "kinsoku",
        "wordWrap",
        "overflowPunct",
        "topLinePunct",
        "autoSpaceDE",
        "autoSpaceDN",
        "bidi",
        "adjustRightInd",
        "snapToGrid",
        "spacing",
        "ind",
        "contextualSpacing",
        "mirrorIndents",
        "suppressOverlap",
        "jc",
        "textDirection",
        "textAlignment",
        "textboxTightWrap",
        "outlineLvl",
        "divId",
        "cnfStyle",
        "rPr",
        "sectPr",
        "pPrChange",
    ],
    "tcPr": [
        "cnfStyle",
        "tcW",
        "gridSpan",
        "hMerge",
        "vMerge",
        "tcBorders",
        "shd",
        "noWrap",
        "tcMar",
        "textDirection",
        "tcFitText",
        "vAlign",
        "hideMark",
    ],
    "tblPr": [
        "tblStyle",
        "tblpPr",
        "tblOverlap",
        "bidiVisual",
        "tblStyleRowBandSize",
        "tblStyleColBandSize",
        "tblW",
        "jc",
        "tblCellSpacing",
        "tblInd",
        "tblBorders",
        "shd",
        "tblLayout",
        "tblCellMar",
        "tblLook",
        "tblCaption",
        "tblDescription",
    ],
    "rPr": [
        "rStyle",
        "rFonts",
        "b",
        "bCs",
        "i",
        "iCs",
        "caps",
        "smallCaps",
        "strike",
        "dstrike",
        "outline",
        "shadow",
        "emboss",
        "imprint",
        "noProof",
        "snapToGrid",
        "vanish",
        "webHidden",
        "color",
        "spacing",
        "w",
        "kern",
        "position",
        "sz",
        "szCs",
        "highlight",
        "u",
        "effect",
        "bdr",
        "shd",
        "fitText",
        "vertAlign",
        "rtl",
        "cs",
        "em",
        "lang",
        "eastAsianLayout",
        "specVanish",
        "oMath",
    ],
}
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _reorder_children(root):
    for tag, order in _ORDERS.items():
        idx = {f"{{{_W_NS}}}{n}": i for i, n in enumerate(order)}
        for el in root.iter(f"{{{_W_NS}}}{tag}"):
            kids = list(el)
            if len(kids) < 2:
                continue
            srt = sorted(kids, key=lambda k: idx.get(k.tag, 999))
            if srt != kids:
                for k in kids:
                    el.remove(k)
                for k in srt:
                    el.append(k)


def _fix_package_order(doc):
    for part in doc.part.package.iter_parts():
        el = getattr(part, "element", None)
        if el is not None:
            try:
                _reorder_children(el)
            except Exception:
                pass
    # atribut zoom wajib 'percent'
    for z in doc.settings.element.iter(f"{{{_W_NS}}}zoom"):
        if z.get(qn("w:percent")) is None:
            z.set(qn("w:percent"), "100")


# ============================================================================
# MESIN ANALISIS: semua angka laporan dihitung di sini dari data dashboard
# ============================================================================
def _third_means(vals):
    k = max(1, len(vals) // 3)
    return float(np.mean(vals[:k])), float(np.mean(vals[-k:]))


def _trend_info(first, last, thr=10.0):
    """Kembalikan (pct, arah) arah in {'naik','turun','stabil','na'}."""
    if first <= 0:
        return None, "na"
    pct = (last - first) / first * 100
    if pct > thr:
        return pct, "naik"
    if pct < -thr:
        return pct, "turun"
    return pct, "stabil"


def _analyze(df, dq, meta):
    S = {}
    period = meta.get("periode_label", "Harian")
    d = df.copy()
    d["Tanggal"] = pd.to_datetime(d["Tanggal"])
    S["df"], S["dq"], S["period"] = d, dq, period

    S["tgl_min"], S["tgl_max"] = d["Tanggal"].min(), d["Tanggal"].max()
    S["n_hari"] = int(d["Tanggal"].nunique())
    S["span_hari"] = int((S["tgl_max"] - S["tgl_min"]).days) + 1
    S["n_lokasi"] = int(d["Lokasi"].nunique())
    S["multi"] = S["n_lokasi"] > 1
    S["total_all"] = float(d["Total"].sum())

    daily = d.groupby("Tanggal")["Total"].sum(min_count=1)
    S["daily"] = daily
    daily_check = d.groupby("Tanggal").agg(
        total=("Total", lambda x: x.sum(min_count=1)),
        n_lokasi=("Lokasi", "nunique"),
        n_total_valid=(
            "Total",
            lambda x: pd.to_numeric(x, errors="coerce").notna().sum(),
        ),
        n_positive=("Total", lambda x: (pd.to_numeric(x, errors="coerce") > 0).sum()),
    )
    S["all_zero_dates"] = daily_check.index[
        (daily_check["total"] == 0)
        & (daily_check["n_lokasi"] == S["n_lokasi"])
        & (daily_check["n_total_valid"] == S["n_lokasi"])
        & (daily_check["n_positive"] == 0)
    ].tolist()
    S["avg_daily"] = float(daily.mean())
    S["day_max"], S["day_min"] = daily.idxmax(), daily.idxmin()

    # ---- periode ----
    g = _periodic_series(d, period)
    S["g"], S["n_periods"] = g, len(g)
    vals = g["Total"].astype(float).values
    pos = vals[vals > 0]
    med = float(np.median(pos)) if len(pos) else 0.0
    thr = 0.10 * med
    low = (
        (vals < thr) if len(vals) >= 7 and med > 0 else np.zeros(len(vals), dtype=bool)
    )
    tail = 0
    for v in vals[::-1]:
        if len(vals) >= 7 and med > 0 and v < thr:
            tail += 1
        else:
            break
    S["low_mask"], S["tail_low"], S["median_pos"] = low, tail, med
    valid = ~low
    S["valid_mask"] = valid
    vv = vals[valid]
    S["mean_valid"] = float(vv.mean()) if len(vv) else 0.0
    S["cv"] = (
        round(float(vv.std(ddof=1) / vv.mean()), 2)
        if len(vv) >= 3 and vv.mean() > 0
        else None
    )
    if len(vv) >= 4:
        S["first_m"], S["last_m"] = _third_means(vv)
        S["trend_pct"], S["trend_dir"] = _trend_info(S["first_m"], S["last_m"])
    else:
        S["first_m"] = S["last_m"] = None
        S["trend_pct"], S["trend_dir"] = None, "na"
    S["peak_idx"] = int(np.argmax(vals)) if len(vals) else 0
    S["peak_ratio"] = (
        (vals.max() / S["mean_valid"]) if S["mean_valid"] > 0 and len(vals) else None
    )
    if valid.any():
        vi = np.where(valid)[0]
        S["trough_idx"] = int(vi[np.argmin(vals[vi])])
    else:
        S["trough_idx"] = 0

    # ---- komposisi ----
    comp = d[CATEGORIES].sum()
    comp_present = comp[comp > 0].sort_values(ascending=False)
    S["cats_present"] = [c for c in CATEGORIES if comp.get(c, 0) > 0]
    S["comp_present"] = comp_present
    S["comp_sum"] = float(comp_present.sum())
    S["comp_pct"] = (
        (comp_present / comp_present.sum() * 100)
        if comp_present.sum() > 0
        else comp_present
    )
    cp = S["comp_pct"]
    S["share_priv"] = (
        float(cp.reindex(["Motor", "Mobil"]).fillna(0).sum()) if len(cp) else 0.0
    )
    S["share"] = {c: float(cp.get(c, 0.0)) for c in CATEGORIES}

    # ---- lokasi ----
    by_lok = d.groupby("Lokasi")["Total"].sum().sort_values(ascending=False)
    S["by_lok"] = by_lok
    stat = dict(zip(dq["Lokasi"], dq["status"]))
    S["status"] = stat
    S["n_inactive"] = int((dq["status"] == ST_INACTIVE).sum())
    S["n_partial"] = int((dq["status"] == ST_PARTIAL).sum())
    S["n_mismatch"] = int((dq["status"] == ST_MISMATCH).sum())
    S["n_zero"] = int((dq["status"] == ST_ZERO).sum())
    S["n_ok"] = int((dq["status"] == ST_OK).sum())
    S["n_aktif"] = int((dq["status"] != ST_INACTIVE).sum())
    S["aktif_locs"] = dq.loc[dq["status"] != ST_INACTIVE, "Lokasi"].tolist()
    S["top_n"] = min(10, S["n_lokasi"])
    S["top_lok"] = by_lok.index[0] if len(by_lok) else "-"
    S["top_val"] = float(by_lok.iloc[0]) if len(by_lok) else 0.0
    tot_l = float(by_lok.sum())
    S["share_top1"] = S["top_val"] / tot_l * 100 if tot_l > 0 else 0.0
    S["share_top3"] = float(by_lok.head(3).sum()) / tot_l * 100 if tot_l > 0 else 0.0
    S["share_top5"] = float(by_lok.head(5).sum()) / tot_l * 100 if tot_l > 0 else 0.0
    S["rank_low"] = (
        by_lok[by_lok.index.isin(S["aktif_locs"])].nsmallest(S["top_n"]).sort_values()
    )
    sm = by_lok[by_lok.index.isin(S["aktif_locs"])]
    S["small_locs"] = sm[(sm / tot_l * 100 < 1.0)].index.tolist() if tot_l > 0 else []

    # ---- tren per lokasi (top 5, sama dengan default dashboard) ----
    top5 = list(by_lok.head(5).index)
    S["top5"] = top5
    if S["n_periods"] >= 2:
        key = _period_key(d, period)
        pv = (
            d.assign(_key=key)
            .groupby(["_key", "Lokasi"])["Total"]
            .sum()
            .unstack(fill_value=0)
            .reindex(g["_key"])
            .fillna(0)
        )
        S["pivot"] = pv
        chg = {}
        for loc in top5:
            if loc in pv.columns:
                arr = pv[loc].values[valid]
                if len(arr) >= 6:
                    f, l = _third_means(arr)
                    pct, dr = _trend_info(f, l)
                    chg[loc] = (pct, dr)
        S["loc_trend"] = chg
    else:
        S["pivot"], S["loc_trend"] = None, {}

    # ---- bulanan (hanya bulan lengkap) ----
    mon = (
        d.assign(_m=d["Tanggal"].dt.to_period("M"))
        .groupby("_m")
        .agg(total=("Total", "sum"), hari=("Tanggal", "nunique"))
    )
    rows = []
    for m, r in mon.iterrows():
        dim = calendar.monthrange(m.year, m.month)[1]
        rows.append((m, float(r["total"]), int(r["hari"]), dim, int(r["hari"]) == dim))
    S["monthly"] = rows
    comp_m = [x for x in rows if x[4]]
    # abaikan bulan lengkap yang periode akhirnya anomali (volume < 10% median)
    S["monthly_complete"] = comp_m
    return S


# ============================================================================
# INSIGHT & REKOMENDASI (bercabang mengikuti hasil data)
# ============================================================================
def _lbl(S, i):
    return S["g"].loc[i, "Label"]


def _build_insights(S):
    """Daftar dict: kelompok, judul, insight, rek_judul, rek, prioritas (Tinggi/Sedang/Rendah)."""
    out = []
    g, n = S["g"], S["n_periods"]
    period = S["period"].lower()

    def add(kel, judul, ins, rj, rek, prio):
        out.append(
            dict(
                kelompok=kel, judul=judul, insight=ins, rek_judul=rj, rek=rek, prio=prio
            )
        )

    # ---------- A. Pola waktu ----------
    if n >= 3 and S["peak_ratio"]:
        pi = S["peak_idx"]
        pl, pv = _lbl(S, pi), g.loc[pi, "Total"]
        ratio = S["peak_ratio"]
        if ratio >= 2:
            add(
                "Pola Waktu",
                "Terdapat lonjakan volume ekstrem pada periode puncak",
                f"Puncak volume terjadi pada **{pl}** sebesar **{_fmt_id(pv)}**, atau **{_n(ratio)} kali** "
                f"rata-rata per periode ({_fmt_id(S['mean_valid'])}). Selisih sebesar ini menandakan "
                f"periode dengan tekanan operasional jauh di atas kondisi normal.",
                "Verifikasi periode puncak sebelum menetapkan kebutuhan operasional",
                f"Periksa kondisi lalu lintas pada {pl} pada beberapa waktu yang mewakili periode tersebut. "
                f"Jika ditemukan antrean yang mengular, tempatkan petugas di pangkal antrean dan titik "
                f"pertemuan arus untuk mengarahkan kendaraan serta mencegah kendaraan berhenti di "
                f"persimpangan. Jika hambatan berasal dari parkir atau aktivitas antar-jemput, arahkan "
                f"kendaraan ke area parkir atau kantong berhenti yang tersedia dan tertibkan kendaraan "
                f"yang berhenti pada badan jalan. Jika penyeberang kesulitan melintas, tempatkan petugas "
                f"di titik penyeberangan pada periode ramai. Penyesuaian lampu lalu lintas atau perubahan "
                f"arah kendaraan perlu dibahas dengan instansi berwenang setelah pola antrean dan "
                f"penyebabnya dipastikan di lapangan.",
                "Sedang",
            )
        elif ratio >= 1.5:
            add(
                "Pola Waktu",
                "Periode puncak lebih ramai dibanding kondisi rata-rata",
                f"Puncak volume tercatat pada **{pl}** sebesar **{_fmt_id(pv)}**, sekitar "
                f"**{_n(ratio)} kali** rata-rata per periode ({_fmt_id(S['mean_valid'])}). Lonjakan ada, "
                f"namun masih dalam rentang yang dapat diantisipasi.",
                "Jadwalkan petugas berdasarkan pola periode ramai",
                f"Pada periode yang polanya menyerupai {pl}, penempatan petugas dapat difokuskan pada "
                f"pintu masuk-keluar kawasan, persimpangan, dan titik penyeberangan yang ramai. Petugas "
                f"dapat membantu mengatur giliran kendaraan ketika arus saling bertemu, mengarahkan "
                f"kendaraan ke tempat parkir, serta mencegah kendaraan berhenti di mulut persimpangan. "
                f"Jika tidak ditemukan antrean atau konflik pergerakan, pengawasan rutin dan penertiban "
                f"parkir secara berkala sudah dapat dipertahankan.",
                "Sedang",
            )
        else:
            add(
                "Pola Waktu",
                "Tidak ada lonjakan tajam; volume relatif merata antar periode",
                f"Puncak volume pada **{pl}** ({_fmt_id(pv)}) hanya sekitar **{_n(ratio)} kali** "
                f"rata-rata per periode ({_fmt_id(S['mean_valid'])}), sehingga tekanan operasional "
                f"relatif seragam.",
                "Pertahankan pola layanan standar dan fokus pada kualitas",
                "Pertahankan pengawasan rutin di seluruh lokasi, dengan pemeriksaan berkala pada marka, "
                "rambu, trotoar, penyeberangan, dan area parkir. Petugas dapat melakukan patroli bergilir "
                "untuk memastikan kendaraan tidak berhenti di persimpangan, akses darurat, atau jalur "
                "pejalan kaki. Jika ditemukan titik yang berulang kali menjadi tempat berhenti sembarangan, "
                "pasang petunjuk larangan berhenti dan siapkan area berhenti yang lebih sesuai.",
                "Rendah",
            )

    if S["trend_dir"] != "na":
        pct, f, l = S["trend_pct"], S["first_m"], S["last_m"]
        cvtxt = ""
        if S["cv"] is not None:
            lab = (
                "sangat fluktuatif"
                if S["cv"] >= 0.5
                else ("cukup fluktuatif" if S["cv"] >= 0.25 else "relatif stabil")
            )
            cvtxt = f" Variasi antar periode tergolong {lab} (koefisien variasi ≈ {_n(S['cv'], 2)})."
        fleks = (
            " Karena variasinya tinggi, tingkatkan frekuensi pemantauan pada periode yang menunjukkan "
            "perubahan besar dan evaluasi kembali setelah data tervalidasi."
            if (S["cv"] or 0) >= 0.5
            else ""
        )
        if S["trend_dir"] == "naik":
            add(
                "Pola Waktu",
                "Tren volume menunjukkan peningkatan",
                f"Rata-rata pada sepertiga akhir periode ({_fmt_id(l)}) sekitar **{_p(pct)}** lebih tinggi "
                f"daripada sepertiga awal ({_fmt_id(f)}).{cvtxt}",
                "Antisipasi pertumbuhan volume secara bertahap",
                "Periksa lokasi yang menunjukkan peningkatan aktivitas dan bandingkan kondisi pada beberapa "
                "hari. Jika kendaraan mengantre hingga menutup persimpangan atau pintu kawasan, tempatkan "
                "petugas di titik masuk-keluar untuk mengatur giliran kendaraan. Jika hambatan berasal "
                "dari parkir, tetapkan area parkir yang diperbolehkan dan larang berhenti pada mulut "
                "persimpangan, akses keluar-masuk, serta trotoar. Sediakan titik khusus untuk antar-jemput "
                "agar kendaraan tidak menunggu di lajur berjalan. Jika pejalan kaki harus berbagi ruang "
                "dengan kendaraan, kosongkan trotoar dan arahkan penyeberangan ke titik yang terlihat jelas. "
                "Tempatkan petugas pada periode yang terbukti ramai; penambahan atau perubahan fasilitas "
                "ditentukan setelah hambatan di lapangan dicatat." + fleks,
                "Tinggi" if pct >= 50 else "Sedang",
            )
        elif S["trend_dir"] == "turun":
            add(
                "Pola Waktu",
                "Tren volume menunjukkan penurunan",
                f"Rata-rata pada sepertiga akhir periode ({_fmt_id(l)}) sekitar **{_p(abs(pct))}** lebih "
                f"rendah daripada sepertiga awal ({_fmt_id(f)}).{cvtxt}",
                "Validasi dan pantau perubahan tren volume",
                "Pastikan terlebih dahulu kamera tetap aktif dan catatan pada periode akhir sama lengkapnya "
                "dengan periode awal. Jika penurunan terkonfirmasi, periksa apakah ada perubahan akses, "
                "pengalihan arus, penutupan jalan, atau berkurangnya kegiatan di sekitar lokasi. Pengaturan "
                "parkir, penyeberangan, dan akses kendaraan yang masih dibutuhkan dapat tetap dipertahankan. "
                "Jangan mengurangi petugas, fasilitas, atau layanan hanya berdasarkan penurunan angka; "
                "sesuaikan setelah kebutuhan pengguna jalan diperiksa." + fleks,
                "Sedang",
            )
        else:
            add(
                "Pola Waktu",
                "Tren volume relatif stabil",
                f"Rata-rata sepertiga akhir periode ({_fmt_id(l)}) berbeda hanya {_p(abs(pct))} dari "
                f"sepertiga awal ({_fmt_id(f)}), sehingga tidak terlihat kecenderungan naik atau turun "
                f"yang berarti.{cvtxt}",
                "Jadikan periode ini sebagai baseline pemantauan",
                "Gunakan rata-rata saat ini sebagai pembanding. Buat tanda peringatan jika jumlah objek berubah jauh dari kebiasaan "
                "(misalnya penyimpangan ±20% dari rata-rata) agar perubahan pola cepat terdeteksi."
                + fleks,
                "Rendah",
            )

    k = int(S["low_mask"].sum())
    if k > 0:
        lows = g.loc[S["low_mask"], "Label"].tolist()
        contoh = _list_ex(lows, 3)
        if S["tail_low"] == k:
            tail_txt = (
                " Seluruhnya berada di ujung akhir periode data, yang mengarah pada data yang belum "
                "lengkap atau gangguan pencatatan."
            )
        elif S["tail_low"]:
            tail_txt = (
                f" Sebanyak {S['tail_low']} di antaranya berada di ujung akhir periode data, yang "
                f"mengarah pada data yang belum lengkap atau gangguan pencatatan."
            )
        else:
            tail_txt = ""
        zero_dates = S.get("all_zero_dates", [])
        zero_note = ""
        if zero_dates:
            zero_labels = [format_tanggal_id(x) for x in zero_dates]
            zero_note = (
                f" Secara khusus, terdapat {len(zero_dates)} tanggal ketika total tercatat nol "
                f"pada seluruh {S['n_lokasi']} lokasi, antara lain {_list_ex(zero_labels, 4)}."
            )
        add(
            "Pola Waktu",
            "Ada periode dengan volume sangat rendah yang perlu diverifikasi",
            f"Terdapat **{k} periode** ({period}) dengan volume di bawah 10% dari median "
            f"({_fmt_id(S['median_pos'])}), antara lain {contoh}.{tail_txt}{zero_note} Nilai sekecil itu tidak "
            f"selalu mencerminkan aktivitas riil.",
            "Verifikasi periode bervolume sangat rendah sebelum dipakai untuk keputusan",
            "Periksa rekap sumber dan log perangkat CCTV pada tanggal tersebut di seluruh lokasi. "
            "Pastikan apakah angka nol berarti tidak ada objek terdeteksi atau data tidak terekam. "
            "Jika terjadi gangguan atau rekap belum lengkap, tandai sebagai data tidak tersedia, "
            "lengkapi setelah diverifikasi, dan jangan gunakan periode tersebut untuk menyimpulkan tren.",
            "Tinggi" if S["tail_low"] else "Sedang",
        )

    mc = S["monthly_complete"]
    if len(mc) >= 3:
        # bulan yang ujung datanya anomali tidak dipakai sebagai "terendah"
        cand = mc
        mx = max(cand, key=lambda x: x[1])
        mn = min(cand, key=lambda x: x[1])
        if mx[0] != mn[0] and mn[1] > 0:
            nm = lambda m: f"{BULAN_ID[m.month]} {m.year}"
            add(
                "Pola Waktu",
                "Pola bulanan dapat dijadikan dasar penjadwalan kegiatan",
                f"Dari {len(mc)} bulan dengan data lengkap, volume tertinggi terjadi pada **{nm(mx[0])}** "
                f"({_fmt_id(mx[1])}) dan terendah pada **{nm(mn[0])}** ({_fmt_id(mn[1])}); selisihnya "
                f"{_n(mx[1] / mn[1])} kali lipat.",
                "Gunakan pola bulanan sebagai acuan sementara",
                f"Gunakan perbedaan volume antarbulan sebagai bahan pemantauan awal, bukan sebagai "
                f"kesimpulan pola musiman. Bandingkan bulan yang sama pada tahun berikutnya dan "
                f"catat pengaruh hari libur, kegiatan, serta perubahan kondisi lokasi sebelum "
                f"menjadikannya dasar penjadwalan program.",
                "Rendah",
            )

    # ---------- B. Komposisi ----------
    cp = S["comp_pct"]
    if len(cp) >= 2:
        top_c, top_p = cp.index[0], float(cp.iloc[0])
        sec_c, sec_p = cp.index[1], float(cp.iloc[1])
        orang = S["share"]["Orang"]
        if S["share_priv"] >= 50:
            extra = ""
            if orang >= 20:
                extra = f" Pejalan kaki ({_p(orang)}) juga signifikan sehingga fasilitas pejalan kaki tetap perlu dijaga."
            add(
                "Komposisi",
                "Volume sangat bergantung pada kendaraan pribadi",
                f"Motor dan mobil bersama-sama menyumbang **{_p(S['share_priv'])}** dari seluruh volume; "
                f"kategori terbesar adalah **{top_c}** ({_p(top_p)}), disusul **{sec_c}** ({_p(sec_p)}).{extra}",
                "Evaluasi pengelolaan parkir dan akses kendaraan",
                "Pada lokasi dengan aktivitas kendaraan pribadi tinggi, tandai area parkir resmi dan larang "
                "parkir pada tikungan, mulut persimpangan, depan akses bangunan, halte, serta trotoar. "
                "Sediakan titik naik-turun penumpang yang tidak memotong lajur utama dan tetapkan batas "
                "waktu berhenti agar kendaraan tidak menunggu terlalu lama. Jika kendaraan keluar dari "
                "area parkir langsung memotong arus, arahkan titik keluar ke bagian yang memiliki jarak "
                "pandang lebih baik. Tempatkan petugas pada akses yang mengalami konflik pergerakan. "
                "Tindakan penanganan antrean dilakukan jika antrean memang terlihat saat pemeriksaan; "
                "jumlah kendaraan tercatat saja tidak membuktikan kemacetan.",
                "Sedang",
            )
            if 0 < orang < 15:
                add(
                    "Komposisi",
                    "Proporsi pejalan kaki masih rendah",
                    f"Pejalan kaki hanya **{_p(orang)}** dari total volume, jauh di bawah kendaraan "
                    f"bermotor ({_p(S['share_priv'])}).",
                    "Audit fasilitas keselamatan pejalan kaki",
                    "Periksa apakah trotoar tersedia pada kedua sisi jalan yang dibutuhkan, tersambung ke sekolah, "
                    "pasar, kantor, halte, atau fasilitas umum, dan bebas dari parkir maupun hambatan lain. "
                    "Periksa pula marka penyeberangan, penerangan, rambu peringatan, dan jarak pandang. "
                    "Jika fasilitas belum memadai, susun daftar titik yang perlu dibersihkan, diperbaiki, "
                    "atau diberi penyeberangan. Proporsi pejalan kaki dalam data CCTV saja belum cukup "
                    "untuk menyimpulkan kondisi fasilitas atau minat masyarakat berjalan kaki.",
                    "Sedang",
                )
        elif top_c == "Orang" and top_p >= 40:
            add(
                "Komposisi",
                "Pejalan kaki mendominasi volume",
                f"Kategori **Orang** menyumbang **{_p(top_p)}** dari total volume, diikuti **{sec_c}** "
                f"({_p(sec_p)}), yang menunjukkan kawasan aktif dimanfaatkan secara langsung oleh masyarakat.",
                "Utamakan fasilitas dan keselamatan pejalan kaki",
                "Pastikan trotoar memiliki jalur yang menerus dan tidak digunakan untuk parkir, tempat menunggu, "
                "atau menaruh barang. Pada bagian trotoar yang terputus atau rusak, catat titik dan tingkat "
                "kerusakannya untuk dimasukkan ke daftar perbaikan. Periksa marka penyeberangan, jarak ke "
                "tujuan umum, penerangan, dan jarak pandang pengendara. Jika penyeberang sulit mendapat "
                "kesempatan melintas, tempatkan petugas pada waktu ramai atau kaji kebutuhan fasilitas "
                "penyeberangan yang lebih aman bersama instansi terkait. Pastikan akses juga dapat digunakan "
                "oleh pengguna kursi roda, lansia, dan orang yang membawa anak.",
                "Sedang",
            )
        else:
            add(
                "Komposisi",
                "Komposisi relatif beragam tanpa satu kategori yang sangat dominan",
                f"**{top_c}** ({_p(top_p)}) dan **{sec_c}** ({_p(sec_p)}) merupakan dua kategori terbesar; "
                f"tidak ada kategori tunggal yang menguasai separuh volume.",
                "Terapkan pengelolaan terpadu lintas moda",
                "Bagi ruang jalan secara jelas: trotoar untuk pejalan kaki, lajur berjalan untuk kendaraan, "
                "dan area khusus untuk parkir maupun antar-jemput. Pindahkan aktivitas menunggu dari "
                "mulut persimpangan dan akses bangunan. Periksa apakah kendaraan yang keluar dari "
                "parkir memotong jalur pejalan kaki atau menutup pandangan pengendara; jika iya, "
                "atur ulang titik keluar atau arahkan petugas untuk membantu pada waktu ramai. "
                "Tertibkan pelanggaran secara bertahap melalui petugas dan rambu yang mudah dipahami.",
                "Rendah",
            )

    tr = S["share"]["Truk"]
    if tr >= 5:
        add(
            "Komposisi",
            "Kendaraan berat terpantau dalam komposisi lalu lintas",
            f"Truk mencapai **{_p(tr)}** dari total volume pada titik-titik CCTV yang dianalisis.",
            "Pantau distribusi kendaraan berat sebelum menetapkan pembatasan",
            "Identifikasi lokasi dan waktu kemunculan truk, lalu padukan dengan fungsi jalan, ketentuan "
            "kelas jalan, dan catatan keselamatan. Data agregat ini belum cukup untuk menetapkan "
            "pembatasan jam atau jalur khusus.",
            "Rendah",
        )
    elif tr >= 1:
        add(
            "Komposisi",
            "Kendaraan berat (truk) berproporsi kecil namun berdampak pada ruang jalan",
            f"Truk hanya **{_p(tr)}** dari total volume, tetapi setiap truk memerlukan ruang dan "
            f"membawa risiko yang lebih besar dibanding kendaraan lain.",
            "Pantau pergerakan truk di titik padat",
            "Pantau tren truk dari data CCTV dan siapkan pengaturan jam melintas jika proporsinya "
            "meningkat.",
            "Rendah",
        )

    # Jumlah bus dari titik CCTV saja tidak cukup untuk menyimpulkan kebutuhan
    # perluasan rute atau kebijakan angkutan umum; karena itu tidak dibuat
    # rekomendasi kebijakan bus otomatis dari proporsi kategori ini.

    # ---------- C. Lokasi ----------
    if S["multi"] and S["n_lokasi"] >= 3:
        names = by3 = S["by_lok"].head(3).index.tolist()
        t3 = S["share_top3"]
        ratio12 = ""
        if len(S["by_lok"]) >= 2 and S["by_lok"].iloc[1] > 0:
            ratio12 = (
                f" Lokasi teratas ({S['top_lok']}) {_n(S['by_lok'].iloc[0] / S['by_lok'].iloc[1])} kali "
                f"lebih ramai daripada peringkat kedua."
            )
        if t3 >= 60:
            add(
                "Lokasi",
                "Keramaian sangat terkonsentrasi pada beberapa lokasi",
                f"Tiga lokasi teratas ({_join_id(names)}) menyumbang **{_p(t3)}** dari total volume.{ratio12}",
                "Prioritaskan sumber daya pada lokasi inti",
                f"Lakukan pemeriksaan langsung di {_join_id(names)} dan catat titik masalah secara spesifik. "
                f"Jika parkir memakai badan jalan, tetapkan area parkir yang diperbolehkan dan kosongkan "
                f"mulut persimpangan serta akses bangunan. Jika kendaraan berhenti untuk menunggu "
                f"penumpang, sediakan titik antar-jemput dan arahkan petugas agar kendaraan tidak menumpuk. "
                f"Jika arus dari dua akses saling memotong, tempatkan petugas pada titik pertemuan dan "
                f"atur giliran kendaraan sementara sesuai kondisi. Jika penyeberangan tidak aman, jaga "
                f"area penyeberangan tetap bebas parkir dan tempatkan petugas saat dibutuhkan. Volume "
                f"tinggi menjadi alasan untuk memeriksa lokasi, bukan bukti bahwa lokasi tersebut macet.",
                "Tinggi",
            )
        elif t3 >= 40:
            add(
                "Lokasi",
                "Keramaian cukup terkonsentrasi pada beberapa lokasi",
                f"Tiga lokasi teratas ({_join_id(names)}) menyumbang **{_p(t3)}** dari total volume.{ratio12}",
                "Berikan prioritas proporsional pada lokasi teratas",
                f"Periksa {_join_id(names)} secara berkala. Pastikan parkir dan tempat berhenti tidak mengganggu "
                f"arus kendaraan maupun jalur pejalan kaki. Jika ditemukan antrean, konflik pergerakan, "
                f"atau penyeberangan yang tidak aman, lakukan penertiban dan tempatkan petugas sesuai "
                f"kebutuhan. Penambahan fasilitas atau perubahan arus dilakukan setelah masalah di "
                f"lapangan dipastikan.",
                "Sedang",
            )
        else:
            add(
                "Lokasi",
                "Sebaran volume antar lokasi relatif merata",
                f"Tiga lokasi teratas hanya menyumbang **{_p(t3)}** dari total volume, sehingga beban "
                f"tersebar ke banyak titik.{ratio12}",
                "Terapkan pengelolaan merata berbasis kebutuhan tiap lokasi",
                "Lakukan pemeriksaan singkat pada setiap lokasi dengan daftar yang sama: kondisi trotoar, "
                "marka dan rambu, titik parkir, akses keluar-masuk, serta tempat penyeberangan. Dahulukan "
                "perbaikan pada lokasi yang memiliki hambatan keselamatan, parkir yang menutup akses, "
                "atau konflik antara kendaraan dan pejalan kaki. Jika tidak ditemukan masalah, pertahankan "
                "pengawasan rutin dan jadwal patroli yang merata.",
                "Rendah",
            )

    sm = S["small_locs"]
    if S["multi"] and sm:
        kp = [x for x in sm if S["status"].get(x) != ST_OK]
        contoh = _list_ex(sm, 3)
        if kp:
            add(
                "Lokasi",
                "Sejumlah lokasi aktif berkontribusi sangat kecil dan memiliki catatan kualitas data",
                f"**{len(sm)} dari {S['n_aktif']} lokasi aktif** menyumbang kurang dari 1% volume (antara "
                f"lain {contoh}); **{len(kp)}** di antaranya memiliki catatan kualitas data.",
                "Pastikan data valid sebelum menyimpulkan lokasi sepi",
                "Cek perangkat dan kelengkapan data pada lokasi tersebut terlebih dahulu. Setelah data "
                "valid, baru evaluasi fungsi kamera atau pertimbangkan relokasi ke titik yang lebih strategis.",
                "Sedang",
            )
        else:
            add(
                "Lokasi",
                "Sejumlah lokasi memiliki volume sangat kecil",
                f"**{len(sm)} dari {S['n_aktif']} lokasi aktif** menyumbang kurang dari 1% volume (antara "
                f"lain {contoh}) dengan data yang lengkap, sehingga rendahnya volume kemungkinan nyata.",
                "Evaluasi fungsi dan penempatan kamera pada lokasi berkontribusi rendah",
                "Pertimbangkan relokasi kamera ke titik yang lebih strategis atau ubah fungsinya "
                "(misalnya keamanan), agar sumber daya pemantauan lebih efisien.",
                "Rendah",
            )

    # ---------- D. Kualitas data ----------
    if S["n_inactive"] > 0:
        nm = S["dq"].loc[S["dq"]["status"] == ST_INACTIVE, "Lokasi"].tolist()
        add(
            "Kualitas Data",
            "Ada lokasi tanpa data / kamera tidak aktif",
            f"**{S['n_inactive']} lokasi** tidak memiliki data valid atau seluruh nilainya nol "
            f"({_list_ex(nm, 4)}).",
            "Perbaiki perangkat pada lokasi tanpa data",
            "Minta petugas memeriksa kamera, sambungan jaringan, dan alat pencatat di lokasi tersebut. Tetapkan "
            "batas waktu perbaikan, dan catat tanggal mulai gangguan sebagai bahan evaluasi vendor.",
            "Tinggi",
        )
    if S["n_partial"] > 0:
        share_p = S["n_partial"] / S["n_lokasi"] * 100
        top_aff = [x for x in S["top5"] if S["status"].get(x) == ST_PARTIAL]
        aff = (
            f" Termasuk lokasi ber-volume tinggi: {_join_id(top_aff)}, sehingga total agregat berpotensi "
            f"terkoreksi."
            if top_aff
            else ""
        )
        if share_p >= 30:
            add(
                "Kualitas Data",
                "Masalah data kosong bersifat meluas",
                f"**{S['n_partial']} dari {S['n_lokasi']} lokasi ({_p(share_p, 0)})** memiliki data yang "
                f"sebagian kosong.{aff}",
                "Bangun mekanisme pemantauan kesehatan perangkat",
                "Buat jadwal pemeriksaan perangkat, periksa kelengkapan data setiap hari, dan siapkan pemberitahuan saat data "
                "kosong/nol terdeteksi; evaluasi sistem rekap karena masalah muncul pada banyak titik.",
                "Tinggi",
            )
        else:
            add(
                "Kualitas Data",
                "Sebagian lokasi memiliki data kosong",
                f"**{S['n_partial']} dari {S['n_lokasi']} lokasi ({_p(share_p, 0)})** memiliki data yang "
                f"sebagian kosong.{aff}",
                "Verifikasi lapangan pada lokasi bermasalah",
                "Periksa perangkat dan log rekap pada lokasi terkait, lalu lengkapi data bila memungkinkan.",
                "Sedang",
            )
    if S["n_mismatch"] > 0:
        add(
            "Kualitas Data",
            "Nilai Total tidak konsisten dengan penjumlahan kategori",
            f"**{S['n_mismatch']} lokasi** memiliki hari dengan Total yang berbeda dari jumlah kategorinya.",
            "Lakukan audit konsistensi rekap data",
            "Periksa cara penjumlahan Total, kategori yang belum terisi, dan kemungkinan salah memasukkan angka agar hasil "
            "agregat dapat dipertanggungjawabkan.",
            "Tinggi",
        )
    if S["n_zero"] > 0:
        add(
            "Kualitas Data",
            "Terdapat lokasi dengan banyak hari bernilai nol",
            f"**{S['n_zero']} lokasi** memiliki lebih dari 30% hari dengan Total nol.",
            "Verifikasi nilai nol dan terapkan pencatatan status perangkat",
            "Periksa lokasi dengan proporsi hari nol yang tinggi melalui log perangkat dan verifikasi "
            "lapangan. Terapkan pencatatan status harian yang membedakan nol hasil pengamatan, kamera "
            "tidak aktif, dan data yang belum masuk agar evaluasi kinerja lokasi tidak bias.",
            "Tinggi",
        )

    # ---------- E. Perencanaan ----------
    if S["span_hari"] < 60:
        add(
            "Perencanaan",
            "Rentang data masih pendek untuk membaca pola musiman",
            f"Data baru mencakup **{S['span_hari']} hari**, belum cukup untuk memastikan pola musiman "
            f"atau pengaruh hari libur.",
            "Perpanjang periode pengumpulan data",
            "Lanjutkan pengumpulan data minimal satu siklus tahunan sebelum menetapkan kebijakan "
            "jangka panjang berbasis musim.",
            "Rendah",
        )
    elif S["span_hari"] < 365:
        add(
            "Perencanaan",
            "Rentang data belum mencakup satu tahun penuh",
            f"Data mencakup **{S['span_hari']} hari**, sehingga belum mewakili satu siklus tahunan penuh.",
            "Lengkapi pemantauan hingga mencakup satu tahun penuh",
            "Lanjutkan pengumpulan data secara konsisten hingga minimal satu tahun penuh. "
            "Gunakan perbandingan bulan yang sama pada tahun berikutnya sebelum menetapkan "
            "kebijakan yang mengandalkan pola musiman.",
            "Sedang",
        )
    else:
        add(
            "Perencanaan",
            "Data mencakup satu tahun penuh atau lebih",
            f"Data mencakup **{S['span_hari']} hari**, sehingga perbandingan lintas tahun dapat mulai dilakukan.",
            "Lakukan perbandingan tahun ke tahun",
            "Bandingkan bulan yang sama antar tahun untuk membedakan pola musiman dari perubahan struktural.",
            "Rendah",
        )
    return out


# ============================================================================
# BANK REKOMENDASI KEBIJAKAN TAMBAHAN
# Aturan di bawah hanya menambahkan rekomendasi jika pemicunya ditemukan pada S.
# Rekomendasi yang memerlukan data di luar CCTV (kemacetan, kecelakaan, kapasitas
# jalan, jam puncak, asal-tujuan) dibatasi sebagai evaluasi/verifikasi, bukan
# langsung menyarankan intervensi fisik atau pembatasan.
# ============================================================================
def _build_policy_recommendations(S):
    """Bank rekomendasi rinci; setiap tindakan hanya muncul jika didukung temuan data."""
    recs = []

    def add(key, title, group, priority, evidence, action):
        recs.append(
            {
                "key": key,
                "rek_judul": title,
                "kelompok": group,
                "prio": priority,
                "evidence": evidence,
                "rek": action,
                "source": "Aturan rekomendasi berbasis temuan",
            }
        )

    shares = S.get("share", {})
    motor = float(shares.get("Motor", 0))
    mobil = float(shares.get("Mobil", 0))
    orang = float(shares.get("Orang", 0))
    bus = float(shares.get("Bus", 0))
    truk = float(shares.get("Truk", 0))
    cats = S.get("comp_present", pd.Series(dtype=float))
    by_lok = S.get("by_lok", pd.Series(dtype=float))
    dq = S.get("dq", pd.DataFrame())

    # Kualitas pencatatan: didahulukan karena memengaruhi keandalan keputusan lain.
    if S.get("n_inactive", 0) > 0:
        locs = []
        if not dq.empty and "status" in dq:
            locs = dq.loc[dq["status"] == ST_INACTIVE, "Lokasi"].astype(str).tolist()
        contoh = (
            _list_ex(locs, 4)
            if locs
            else "lokasi yang belum memiliki catatan yang dapat digunakan"
        )
        add(
            "repair_inactive",
            "Pulihkan pencatatan CCTV yang tidak berjalan",
            "Kualitas Data",
            "Tinggi",
            f"Terdapat {S['n_inactive']} lokasi tanpa data yang dapat digunakan, antara lain {contoh}.",
            f"Minta petugas memeriksa kamera, sumber listrik, sambungan jaringan, dan alat pencatat di {contoh}. Catat tanggal mulai gangguan, penyebabnya, petugas yang menangani, serta tanggal perangkat kembali berfungsi. Selama masalah belum selesai, beri tanda bahwa data lokasi tersebut belum tersedia dan jangan menganggap angka nol sebagai bukti bahwa tidak ada aktivitas.",
        )

    if S.get("n_mismatch", 0) > 0:
        add(
            "reconcile_totals",
            "Benahi perbedaan antara jumlah total dan rincian kategori",
            "Kualitas Data",
            "Tinggi",
            f"Terdapat {S['n_mismatch']} lokasi dengan catatan ketika angka Total berbeda dari penjumlahan kategori.",
            "Periksa contoh tanggal yang bermasalah bersama petugas pengelola data. Cocokkan angka Orang, Motor, Mobil, Truk, dan Bus dengan angka Total pada sumber aslinya, lalu perbaiki rumus atau proses penyalinan yang keliru. Simpan catatan perubahan agar angka yang sudah diperbaiki dapat ditelusuri dan gunakan data tersebut untuk laporan setelah pemeriksaan selesai.",
        )

    if S.get("n_partial", 0) > 0:
        locs = []
        if not dq.empty and "status" in dq:
            locs = dq.loc[dq["status"] == ST_PARTIAL, "Lokasi"].astype(str).tolist()
        contoh = (
            _list_ex(locs, 4)
            if locs
            else "lokasi dengan isian kategori yang belum lengkap"
        )
        add(
            "complete_records",
            "Lengkapi catatan kategori yang masih kosong",
            "Kualitas Data",
            (
                "Tinggi"
                if S.get("n_partial", 0) / max(S.get("n_lokasi", 1), 1) >= 0.3
                else "Sedang"
            ),
            f"Sebagian catatan pada {S['n_partial']} lokasi belum memuat seluruh kategori, antara lain {contoh}.",
            f"Periksa tanggal dan jenis catatan yang kosong di {contoh}. Pastikan apakah alat tidak merekam, data belum terkirim, atau kategori memang tidak terdeteksi. Minta pengelola melengkapi data dari sumber yang sah jika tersedia; jika tidak, tandai sebagai data kosong dan jangan mengisinya dengan angka perkiraan.",
        )

    if S.get("n_zero", 0) > 0:
        add(
            "verify_zero",
            "Pastikan angka nol benar-benar menunjukkan tidak ada objek yang tercatat",
            "Kualitas Data",
            "Tinggi",
            f"Terdapat {S['n_zero']} lokasi dengan banyak hari yang mencatat Total nol.",
            "Periksa tanggal-tanggal bernilai nol dengan melihat catatan kerja perangkat dan rekaman yang masih tersedia. Bedakan tiga keadaan: kamera berfungsi tetapi tidak mendeteksi objek, kamera atau jaringan bermasalah, dan data belum terkirim. Terapkan pilihan status yang jelas pada laporan harian agar angka nol tidak disalahartikan sebagai lokasi yang tidak memiliki aktivitas.",
        )

    coverage = (
        (S.get("n_hari", 0) / S.get("span_hari", 1)) if S.get("span_hari", 0) else 0
    )
    if S.get("span_hari", 0) >= 7 and coverage < 0.90:
        add(
            "date_coverage",
            "Tutup kekosongan pencatatan pada tanggal yang belum memiliki data",
            "Kualitas Data",
            "Tinggi",
            f"Data tersedia pada {S.get('n_hari', 0)} dari {S.get('span_hari', 0)} hari dalam rentang pengamatan.",
            "Buat daftar tanggal yang belum memiliki catatan untuk setiap lokasi, lalu cocokkan dengan jadwal operasional dan laporan gangguan kamera. Jika data masih tersimpan di perangkat, masukkan kembali setelah diperiksa. Jika tidak dapat dipulihkan, beri keterangan bahwa data tidak tersedia dan hindari membandingkan total antarperiode seolah-olah seluruh harinya tercatat lengkap.",
        )

    # Pejalan kaki.
    if orang >= 20:
        add(
            "pedestrian_facilities",
            "Periksa keamanan jalur pejalan kaki di lokasi dengan banyak pejalan kaki",
            "Pejalan Kaki",
            "Sedang",
            f"Pejalan kaki mencapai {_p(orang)} dari seluruh objek yang tercatat.",
            "Periksa lokasi yang banyak mencatat pejalan kaki. Pastikan trotoar tersambung dan tidak terhalang, penyeberangan mudah dijangkau, penerangan berfungsi, serta tersedia akses yang dapat digunakan oleh lansia dan penyandang disabilitas. Catat bagian yang rusak atau terputus dan susun perbaikan berdasarkan tingkat bahaya serta banyaknya orang yang harus melewati bagian tersebut.",
        )
    elif 0 < orang < 20:
        add(
            "pedestrian_low",
            "Periksa apakah fasilitas pejalan kaki mudah ditemukan dan digunakan",
            "Pejalan Kaki",
            "Rendah",
            f"Pejalan kaki tercatat sebesar {_p(orang)} dari seluruh objek yang terpantau.",
            "Periksa apakah trotoar tersedia dan tersambung ke sekolah, pasar, kantor, halte, atau fasilitas umum terdekat. Catat bagian yang terputus, rusak, terlalu sempit, tertutup parkir, atau digunakan untuk kegiatan lain. Periksa pula apakah tersedia penyeberangan pada jalur yang memang digunakan warga, apakah marka masih terlihat, dan apakah penerangan cukup pada malam hari. Dahulukan pembebasan trotoar dan perbaikan titik penyeberangan yang menimbulkan risiko keselamatan. Jumlah pejalan kaki yang tercatat rendah tidak dengan sendirinya berarti fasilitas pejalan kaki tidak dibutuhkan.",
        )

    elif orang == 0:
        add(
            "pedestrian_none",
            "Pastikan akses pejalan kaki tetap tersedia pada lokasi tanpa catatan pejalan kaki",
            "Pejalan Kaki",
            "Rendah",
            "Tidak terdapat objek pejalan kaki yang tercatat pada periode pengamatan.",
            "Periksa apakah sudut pandang kamera memang mencakup trotoar dan jalur penyeberangan. Saat pemeriksaan lapangan, pastikan jalur berjalan tersedia, tidak terhalang parkir atau aktivitas lain, dan tersambung ke fasilitas umum di sekitar lokasi. Periksa juga apakah warga harus berjalan di badan jalan karena trotoar terputus atau tidak tersedia tempat menyeberang. Tidak adanya catatan pejalan kaki tidak otomatis berarti tidak ada kebutuhan fasilitas pejalan kaki.",
        )

    # Sepeda motor dan mobil.
    if motor > 0:
        add(
            "motor_dominant",
            "Tata akses dan tempat berhenti sepeda motor",
            "Kendaraan",
            "Sedang" if motor >= 40 else "Rendah",
            f"Sepeda motor menyumbang {_p(motor)} dari seluruh objek yang tercatat.",
            "Pada lokasi sepeda motor terlihat, tentukan area parkir yang tidak mengambil trotoar, lajur berjalan, atau ruang pandang di persimpangan. Pisahkan tempat parkir dari titik antar-jemput agar kendaraan yang menunggu tidak menumpuk di pintu masuk. Beri marka atau papan petunjuk pada area yang diperbolehkan, dan tempatkan petugas untuk mengarahkan kendaraan saat area mulai penuh. Jika parkir meluber ke badan jalan, koordinasikan penataan kapasitas dan lokasi parkir dengan pengelola kawasan sebelum menerapkan pembatasan.",
        )
    if mobil > 0:
        add(
            "car_dominant",
            "Rapikan tempat parkir dan titik naik-turun penumpang mobil",
            "Kendaraan",
            "Sedang" if mobil >= 35 else "Rendah",
            f"Mobil menyumbang {_p(mobil)} dari seluruh objek yang tercatat.",
            "Tetapkan titik khusus untuk menurunkan dan menaikkan penumpang yang tidak berada di lajur utama, mulut persimpangan, halte, atau depan akses darurat. Atur agar kendaraan tidak menunggu setelah penumpang turun; sediakan area tunggu atau parkir terpisah. Beri marka batas parkir, pasang petunjuk arah masuk-keluar, dan tempatkan petugas pada pintu kawasan jika kendaraan saling memotong atau antrean mulai menutup akses. Periksa secara berkala apakah parkir meluber ke trotoar atau badan jalan, lalu sesuaikan pengaturan dengan kapasitas lokasi.",
        )
    if motor + mobil >= 65:
        add(
            "private_vehicle_mix",
            "Atur ruang jalan agar kendaraan dan pejalan kaki tidak saling menghalangi",
            "Kendaraan",
            "Sedang",
            f"Sepeda motor dan mobil bersama-sama mencapai {_p(motor + mobil)} dari seluruh objek yang tercatat.",
            "Pada lokasi dengan banyak kendaraan pribadi, pisahkan tiga kegiatan: kendaraan berjalan, kendaraan parkir, dan kendaraan antar-jemput. Tetapkan area parkir resmi, titik naik-turun penumpang, serta jalur masuk-keluar yang tidak saling berpotongan jika ruang memungkinkan. Kosongkan trotoar, mulut persimpangan, dan akses darurat dari kendaraan berhenti. Jika antrean kendaraan memasuki jalan utama, tempatkan petugas di pintu masuk untuk mengatur giliran masuk dan mengarahkan kendaraan ke area tunggu yang tersedia.",
        )

    # Angkutan umum dan kendaraan berat.
    if bus > 0:
        add(
            "bus_facilities",
            "Pastikan tempat berhenti bus aman dan mudah dijangkau penumpang",
            "Angkutan Umum",
            "Sedang" if bus >= 3 else "Rendah",
            f"Bus tercatat sebesar {_p(bus)} dari seluruh objek yang terpantau.",
            "Pada lokasi bus terlihat, periksa apakah bus berhenti di titik yang aman dan tidak menutup persimpangan, akses keluar-masuk, atau lajur berjalan. Jika belum tersedia titik berhenti yang layak, koordinasikan penetapan lokasi naik-turun dengan pengelola angkutan dan pemilik kawasan. Pastikan penumpang dapat menunggu di tempat yang tidak memakan trotoar dan memiliki jalur berjalan yang aman. Jika bus sering berhenti terlalu lama atau beberapa bus menunggu bersamaan, atur urutan kedatangan dan lama berhenti bersama operator. Jumlah bus dalam data ini sendiri belum menjadi dasar untuk menambah rute atau armada.",
        )
    if truk > 0:
        add(
            "truck_route_check",
            "Periksa kesesuaian jalur yang dilalui kendaraan berat",
            "Kendaraan Berat",
            "Sedang" if truk >= 3 else "Rendah",
            f"Truk tercatat sebesar {_p(truk)} dari seluruh objek yang terpantau.",
            "Catat ruas dan waktu truk terlihat, kemudian periksa kesesuaian dengan rambu, kelas jalan, "
            "batas muatan, dan ketentuan kendaraan berat yang berlaku. Periksa apakah truk berhenti "
            "untuk bongkar-muat di badan jalan, menutup akses usaha atau permukiman, atau membuat "
            "kendaraan lain sulit melintas. Jika kegiatan bongkar-muat mengganggu, koordinasikan "
            "penyediaan area bongkar-muat di luar lajur berjalan dan jadwal yang tidak mengganggu "
            "aktivitas sekitar. Penindakan atau pembatasan jalur dan waktu hanya dilakukan setelah "
            "pelanggaran serta kewenangan instansi terkait dipastikan.",
        )

    # Sebaran lokasi dan perubahan waktu.
    if S.get("multi") and S.get("n_lokasi", 0) >= 3 and S.get("share_top1", 0) >= 30:
        top_loc = (
            str(by_lok.index[0]) if len(by_lok) else "lokasi dengan volume tertinggi"
        )
        add(
            "top_location_focus",
            "Jadikan lokasi dengan catatan tertinggi sebagai sasaran pemeriksaan",
            "Lokasi",
            "Sedang",
            f"{top_loc} menyumbang {_p(S['share_top1'])} dari seluruh objek yang tercatat.",
            f"Lakukan pemeriksaan langsung di {top_loc} dengan memetakan titik parkir, tempat antar-jemput, mulut persimpangan, akses bangunan, dan penyeberangan. Jika parkir mengurangi lebar jalan, tetapkan area yang boleh digunakan dan arahkan kendaraan lain ke tempat parkir resmi. Jika kendaraan menunggu penumpang, sediakan titik berhenti singkat yang tidak menghalangi lajur. Jika terjadi antrean pada akses, tempatkan petugas untuk mengatur kendaraan masuk dan keluar. Catat lokasi dan waktu temuan agar penanganan dapat diarahkan pada masalah yang benar-benar terjadi; volume tinggi sendiri bukan bukti kemacetan.",
        )

    cv = S.get("cv")
    if cv is not None and cv >= 0.50:
        add(
            "high_variability",
            "Tingkatkan pengawasan pada periode ketika jumlah objek berubah besar",
            "Pola Waktu",
            "Sedang",
            "Jumlah objek yang tercatat berubah cukup besar antarperiode pengamatan.",
            "Pilih periode dengan perubahan terbesar dan periksa rekaman pada beberapa hari yang mewakili periode tersebut. Jika perubahan berulang pada hari atau periode yang sama, jadwalkan petugas pada titik masuk-keluar, persimpangan, dan penyeberangan yang membutuhkan pengaturan. Pada saat yang sama, pastikan parkir tidak meluber ke badan jalan dan kendaraan antar-jemput tidak menunggu di lajur berjalan. Karena ringkasan data harian atau bulanan tidak menunjukkan jam secara rinci, jadwal petugas perlu disesuaikan lagi setelah pengamatan langsung.",
        )

    if S.get("trend_dir") == "naik":
        add(
            "rising_volume",
            "Pantau lokasi yang menunjukkan kenaikan jumlah objek dari waktu ke waktu",
            "Pola Waktu",
            "Sedang",
            "Rata-rata jumlah objek pada bagian akhir periode lebih tinggi daripada bagian awal.",
            "Periksa lokasi yang mengalami kenaikan paling jelas dan pastikan kenaikan bukan akibat perbedaan kelengkapan data. Jika parkir mulai meluber, atur kembali area parkir dan larang kendaraan berhenti di mulut persimpangan atau akses bangunan. Jika antrean kendaraan menutup pintu kawasan, tempatkan petugas untuk mengatur giliran masuk dan keluar. Jika aktivitas antar-jemput meningkat, sediakan titik berhenti singkat dan area tunggu terpisah. Pastikan trotoar dan penyeberangan tetap terbuka. Penambahan fasilitas atau perubahan arus dipertimbangkan jika hambatan tersebut benar-benar ditemukan di lapangan.",
        )
    elif S.get("trend_dir") == "turun":
        add(
            "falling_volume",
            "Cari tahu penyebab penurunan sebelum mengubah layanan atau fasilitas",
            "Pola Waktu",
            "Rendah",
            "Rata-rata jumlah objek pada bagian akhir periode lebih rendah daripada bagian awal.",
            "Pastikan terlebih dahulu kamera tetap aktif dan catatan pada periode akhir sama lengkapnya dengan periode awal. Jika data dapat dipercaya, periksa apakah ada perubahan kegiatan, akses kawasan, atau penggunaan lokasi. Pertahankan layanan yang masih dibutuhkan dan jangan mengurangi fasilitas hanya berdasarkan penurunan angka dalam satu periode pengamatan.",
        )

    if S.get("multi") and len(by_lok[by_lok > 0]) >= 3:
        low_loc = str(by_lok[by_lok > 0].idxmin())
        low_value = float(by_lok[by_lok > 0].min())
        median_value = float(by_lok[by_lok > 0].median())
        if median_value > 0 and low_value < median_value * 0.25:
            add(
                "low_location_review",
                "Tinjau kembali fungsi pemantauan di lokasi dengan catatan rendah",
                "Lokasi",
                "Rendah",
                f"{low_loc} memiliki catatan jauh di bawah sebagian besar lokasi lain yang datanya bernilai positif.",
                f"Periksa lebih dahulu apakah kamera di {low_loc} berfungsi dan apakah seluruh hari pengamatan tercatat. Jika data lengkap, tinjau apakah titik kamera masih sesuai dengan tujuan pemantauan dan apakah lokasinya masih mewakili aktivitas yang ingin diamati. Relokasi kamera hanya dipertimbangkan setelah kebutuhan pengawasan di lokasi tersebut dan calon lokasi baru dibandingkan.",
            )

    # Kategori yang tidak muncul perlu dicek, bukan diasumsikan tidak ada.
    if S.get("n_lokasi", 0) > 0:
        missing_categories = [c for c in CATEGORIES if c not in cats.index]
        if missing_categories:
            add(
                "category_zero",
                "Pastikan semua kamera mencatat jenis objek dengan cara yang sama",
                "Kualitas Data",
                "Rendah",
                "Tidak ada catatan positif untuk kategori: "
                + ", ".join(missing_categories)
                + ".",
                "Periksa pengaturan kamera dan cara pengisian laporan untuk kategori tersebut. Pastikan kategori memang tidak terlihat selama pengamatan atau ternyata belum dapat dicatat oleh perangkat. Samakan cara pencatatan antar lokasi sebelum membandingkan jumlah orang dan jenis kendaraan.",
            )

    # Data yang singkat belum cukup untuk keputusan jangka panjang.
    if S.get("span_hari", 0) < 365:
        add(
            "extend_observation",
            "Lanjutkan pencatatan sebelum menetapkan kebijakan jangka panjang",
            "Perencanaan",
            "Rendah",
            f"Rentang data yang tersedia mencakup {S.get('span_hari', 0)} hari.",
            "Lanjutkan pencatatan dengan cara dan jadwal yang sama agar hasil antarbulan dapat dibandingkan. Setelah tersedia data satu tahun penuh, bandingkan bulan yang sama dan perhatikan hari libur atau kegiatan khusus. Gunakan hasil tersebut bersama pemeriksaan lapangan sebelum menetapkan perubahan layanan atau pembangunan fasilitas.",
        )
    else:
        add(
            "annual_comparison",
            "Bandingkan kondisi pada bulan yang sama setiap tahun",
            "Perencanaan",
            "Rendah",
            f"Rentang data mencakup {S.get('span_hari', 0)} hari sehingga perbandingan tahunan dapat dilakukan.",
            "Simpan hasil pencatatan setiap tahun dan bandingkan bulan yang sama, bukan hanya membandingkan total tahunan. Catat perubahan penting seperti kegiatan besar, perubahan akses, atau gangguan kamera agar pemerintah dapat membedakan perubahan aktivitas dari perubahan cara pencatatan.",
        )

    # Dasar prioritas dan nilai aktual diturunkan dari indikator yang benar-benar
    # digunakan oleh aturan di atas. Tidak ada skor prioritas baru yang dibuat.
    n_lokasi = max(int(S.get("n_lokasi", 0) or 0), 1)
    n_hari = int(S.get("n_hari", 0) or 0)
    span_hari = int(S.get("span_hari", 0) or 0)
    coverage_pct = (n_hari / span_hari * 100) if span_hari else 0.0
    basis = {
        "repair_inactive": ("Jumlah lokasi tanpa data yang dapat digunakan > 0", f"{_fmt_id(S.get('n_inactive', 0))} lokasi"),
        "reconcile_totals": ("Jumlah lokasi dengan Total tidak cocok > 0", f"{_fmt_id(S.get('n_mismatch', 0))} lokasi"),
        "complete_records": ("Proporsi lokasi dengan data sebagian kosong ≥ 30% = Tinggi; < 30% = Sedang", f"{_fmt_id(S.get('n_partial', 0))} dari {_fmt_id(S.get('n_lokasi', 0))} lokasi ({_p(S.get('n_partial', 0) / n_lokasi * 100)})"),
        "verify_zero": ("Jumlah lokasi dengan banyak hari bernilai Total nol > 0", f"{_fmt_id(S.get('n_zero', 0))} lokasi"),
        "date_coverage": ("Cakupan hari pencatatan < 90% pada rentang ≥ 7 hari", f"{_fmt_id(n_hari)} dari {_fmt_id(span_hari)} hari ({_p(coverage_pct)})"),
        "pedestrian_facilities": ("Proporsi pejalan kaki ≥ 20%", _p(S.get('share', {}).get('Orang', 0))),
        "pedestrian_low": ("Proporsi pejalan kaki > 0% dan < 20%", _p(S.get('share', {}).get('Orang', 0))),
        "pedestrian_none": ("Proporsi pejalan kaki = 0%", _p(S.get('share', {}).get('Orang', 0))),
        "motor_dominant": ("Proporsi sepeda motor ≥ 40% = Sedang; < 40% = Rendah", _p(S.get('share', {}).get('Motor', 0))),
        "car_dominant": ("Proporsi mobil ≥ 35% = Sedang; < 35% = Rendah", _p(S.get('share', {}).get('Mobil', 0))),
        "private_vehicle_mix": ("Gabungan proporsi motor dan mobil ≥ 65%", _p(float(S.get('share', {}).get('Motor', 0)) + float(S.get('share', {}).get('Mobil', 0)))),
        "bus_facilities": ("Proporsi bus ≥ 3% = Sedang; < 3% = Rendah", _p(S.get('share', {}).get('Bus', 0))),
        "truck_route_check": ("Proporsi truk ≥ 3% = Sedang; < 3% = Rendah", _p(S.get('share', {}).get('Truk', 0))),
        "top_location_focus": ("Lokasi aktif ≥ 3 dan kontribusi lokasi tertinggi ≥ 30%", _p(S.get('share_top1', 0))),
        "high_variability": ("Koefisien variasi (CV) ≥ 0,50", _n(float(S.get('cv', 0) or 0), 3)),
        "rising_volume": ("Rata-rata bagian akhir periode lebih tinggi daripada bagian awal", str(S.get('trend_dir', 'naik')).capitalize()),
        "falling_volume": ("Rata-rata bagian akhir periode lebih rendah daripada bagian awal", str(S.get('trend_dir', 'turun')).capitalize()),
        "low_location_review": ("Volume lokasi terendah < 25% median lokasi aktif", "Perbandingan terhadap median lokasi aktif memenuhi batas < 25%"),
        "category_zero": ("Kategori tidak memiliki catatan positif", "Kategori tanpa catatan: " + ", ".join([c for c in CATEGORIES if c not in cats.index])),
        "extend_observation": ("Rentang pengamatan < 365 hari", f"{_fmt_id(span_hari)} hari"),
        "annual_comparison": ("Rentang pengamatan ≥ 365 hari", f"{_fmt_id(span_hari)} hari"),
    }
    for item in recs:
        item["basis_prioritas"], item["nilai_analisis"] = basis.get(
            item.get("key"),
            ("Kondisi pemicu sesuai temuan analisis pada uraian rekomendasi", item.get("evidence") or "Lihat uraian temuan"),
        )
    return recs


# ============================================================================
# GRAFIK (meniru grafik pada dashboard: line chart, bar Altair, stacked bar)
# ============================================================================
def _style_ax(ax, grid_axis="y"):
    ax.set_facecolor("white")
    ax.grid(axis=grid_axis, color="#E6E9EC", linewidth=0.8)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_color("#BFC7CC")
    ax.tick_params(colors="#444444", labelsize=8)


def _yfmt():
    return FuncFormatter(lambda v, _: _compact(v) if abs(v) >= 1e4 else _fmt_id(v))


def _set_period_ticks(ax, g, period, max_ticks=8):
    n = len(g)
    idx = np.unique(np.linspace(0, n - 1, min(n, max_ticks)).round().astype(int))
    multi_year = g["Tanggal"].dt.year.nunique() > 1
    labels = []
    for i in idx:
        ts = g.loc[i, "Tanggal"]
        if period == "Bulanan":
            labels.append(f"{BULAN_SHORT[ts.month]} {ts.year}")
        else:
            labels.append(
                f"{ts.day} {BULAN_SHORT[ts.month]}"
                + (f" {ts.year}" if multi_year else "")
            )
    ax.set_xticks(idx)
    ax.set_xticklabels(
        labels, rotation=0 if n <= 8 else 25, ha="center" if n <= 8 else "right"
    )
    ax.set_xlim(-0.5, n - 0.5)


def _chart_trend_total(S):
    g, period = S["g"], S["period"]
    fig, ax = plt.subplots(figsize=(7.2, 3.1))
    x = np.arange(len(g))
    y = g["Total"].to_numpy()

    ax.plot(
        x,
        y,
        color=LINE_SINGLE,
        linewidth=1.6,
        marker="o" if len(g) <= 40 else None,
        markersize=3.5,
    )

    pi = int(S["peak_idx"])
    ax.scatter([pi], [y[pi]], color=LINE_SINGLE, s=28, zorder=5)

    # Tampilkan nilai setiap titik jika jumlah periode masih memungkinkan
    if len(g) <= 12:
        for i, nilai in enumerate(y):
            if i == pi:
                continue

            ax.annotate(
                _fmt_id(nilai),
                (i, nilai),
                xytext=(0, 5),
                textcoords="offset points",
                ha="center",
                va="bottom",
                fontsize=7,
                color="#444444",
            )

    # Nilai puncak selalu ditampilkan
    ax.annotate(
        f"Puncak: {_fmt_id(y[pi])}",
        (pi, y[pi]),
        xytext=(8, -2),
        textcoords="offset points",
        fontsize=8,
        color="#444444",
        va="center",
    )

    ax.set_title(
        f"Tren Volume Lalu Lintas ({period})",
        fontsize=11,
        fontweight="bold",
        loc="left",
        color="#222222",
    )
    ax.set_ylabel("Total", fontsize=8.5)
    ax.yaxis.set_major_formatter(_yfmt())
    _set_period_ticks(ax, g, period)
    ax.set_ylim(bottom=0)
    _style_ax(ax)
    fig.tight_layout()

    return _fig_to_stream(fig)


def _chart_trend_lokasi(S):
    g, period, pv = S["g"], S["period"], S["pivot"]
    fig, ax = plt.subplots(figsize=(7.2, 3.4))
    x = np.arange(len(g))

    for i, loc in enumerate(S["top5"]):
        if loc not in pv.columns:
            continue

        y = pv[loc].to_numpy()
        warna = LINE_MULTI[i % len(LINE_MULTI)]

        ax.plot(
            x,
            y,
            linewidth=1.5,
            color=warna,
            marker="o",
            markersize=3.5,
            label=loc,
        )

    ax.set_title(
        "Tren per Lokasi (5 lokasi tertinggi)",
        fontsize=11,
        fontweight="bold",
        loc="left",
        color="#222222",
    )
    ax.set_ylabel("Total", fontsize=8.5)
    ax.yaxis.set_major_formatter(_yfmt())
    _set_period_ticks(ax, g, period)
    ax.set_ylim(bottom=0)
    ax.legend(
        loc="upper center",
        bbox_to_anchor=(0.5, -0.2),
        ncol=3,
        fontsize=7.5,
        frameon=False,
    )
    _style_ax(ax)
    fig.tight_layout()

    return _fig_to_stream(fig)


def _table_tren_lokasi(doc, S):
    g = S["g"]
    pv = S["pivot"]
    top5 = [loc for loc in S["top5"] if loc in pv.columns]

    headers = ["Periode"] + top5
    rows = []

    for i in range(len(g)):
        label_periode = g.iloc[i]["Label"]

        nilai_lokasi = [_fmt_id(pv[loc].iloc[i]) for loc in top5]

        rows.append([label_periode] + nilai_lokasi)

    widths = [3.2] + [(TEXT_WIDTH_CM - 3.2) / len(top5)] * len(top5)
    aligns = ["left"] + ["right"] * len(top5)

    _table(
        doc,
        headers,
        rows,
        widths,
        aligns=aligns,
        font_size=8,
    )


def _chart_rank_bar(series, title, rotate=40):
    fig, ax = plt.subplots(figsize=(7.2, 4.3))
    x = np.arange(len(series))
    ax.bar(x, series.values, color=BAR_COLOR, width=0.68)
    for xi, v in zip(x, series.values):
        ax.text(
            xi, v, _compact(v), ha="center", va="bottom", fontsize=7.5, color="#333333"
        )
    labels = [textwrap.fill(str(s), 26) for s in series.index]
    ax.set_xticks(x)
    ax.set_xticklabels(
        labels, rotation=rotate, ha="right", rotation_mode="anchor", fontsize=7.8
    )
    ax.set_ylabel("Total Volume", fontsize=8.5)
    ax.yaxis.set_major_formatter(_yfmt())
    ax.set_title(title, fontsize=11, fontweight="bold", loc="left", color="#222222")
    ax.set_ylim(0, max(series.values.max() * 1.12, 1))
    _style_ax(ax)
    fig.tight_layout()
    return _fig_to_stream(fig)


def _chart_comp_bar(S):
    comp = S["comp_present"].sort_values(ascending=True)
    pct = S["comp_pct"]

    fig, ax = plt.subplots(figsize=(7.0, 2.9))

    y = np.arange(len(comp))
    ax.barh(
        y,
        comp.values,
        color=BAR_COLOR,
        height=0.6,
    )

    for yi, (k, v) in zip(y, comp.items()):
        ax.text(
            v,
            yi,
            f"  {_compact(v)}  ({_p(pct[k])})",
            va="center",
            fontsize=8,
            color="#333333",
        )

    ax.set_yticks(y)
    ax.set_yticklabels(comp.index, fontsize=9)
    ax.set_xlabel("Total Volume", fontsize=8.5)
    ax.xaxis.set_major_formatter(_yfmt())
    ax.set_xlim(0, comp.values.max() * 1.3)

    ax.set_title(
        "Distribusi per Kategori",
        fontsize=11,
        fontweight="bold",
        loc="left",
        color="#222222",
    )

    _style_ax(ax, "x")
    fig.tight_layout()

    return _fig_to_stream(fig)


def _chart_comp_pie(S):
    comp = S["comp_present"].sort_values(ascending=False)
    pct = S["comp_pct"]

    fig, ax = plt.subplots(figsize=(7.0, 3.3))

    colors = [CAT_COLORS.get(k, BAR_COLOR) for k in comp.index]

    wedges, _ = ax.pie(
        comp.values,
        startangle=90,
        counterclock=False,
        colors=colors,
        autopct=None,
        wedgeprops={
            "edgecolor": "white",
            "linewidth": 1.2,
        },
    )

    # Ambang untuk menentukan posisi label persentase.
    batas_persen = 5.0

    # Menampung label kecil yang perlu diletakkan di luar pie.
    label_luar = {"kiri": [], "kanan": []}

    for wedge, kategori in zip(wedges, comp.index):
        persen = float(pct.get(kategori, 0.0))

        # Sudut tengah irisan.
        sudut = np.deg2rad((wedge.theta1 + wedge.theta2) / 2)

        x = np.cos(sudut)
        y = np.sin(sudut)

        if persen >= batas_persen:
            # Persentase besar tetap berada di dalam irisan.
            ax.text(
                0.62 * x,
                0.62 * y,
                _p(persen),
                ha="center",
                va="center",
                fontsize=8,
                fontweight="bold",
                color="white",
            )

        else:
            # Persentase kecil dipindahkan ke luar irisan.
            sisi = "kanan" if x >= 0 else "kiri"

            label_luar[sisi].append(
                {
                    "x": x,
                    "y": y,
                    "persen": persen,
                    "y_label": 1.10 * y,
                }
            )

    # Menempatkan label luar agar tidak saling bertumpuk.
    for sisi, daftar in label_luar.items():
        if not daftar:
            continue

        # Urutkan dari posisi vertikal terendah ke tertinggi.
        daftar.sort(key=lambda item: item["y_label"])

        jarak_minimum = 0.15

        # Atur jarak vertikal antarteks.
        for i in range(1, len(daftar)):
            daftar[i]["y_label"] = max(
                daftar[i]["y_label"],
                daftar[i - 1]["y_label"] + jarak_minimum,
            )

        # Jika label terdorong terlalu tinggi, geser bersama-sama.
        batas_atas = 1.18
        batas_bawah = -1.18

        if daftar[-1]["y_label"] > batas_atas:
            geser = daftar[-1]["y_label"] - batas_atas

            for item in daftar:
                item["y_label"] -= geser

        if daftar[0]["y_label"] < batas_bawah:
            geser = batas_bawah - daftar[0]["y_label"]

            for item in daftar:
                item["y_label"] += geser

        for item in daftar:
            arah_x = 1 if sisi == "kanan" else -1

            ax.annotate(
                _p(item["persen"]),
                xy=(
                    0.92 * item["x"],
                    0.92 * item["y"],
                ),
                xytext=(
                    arah_x * 1.16,
                    item["y_label"],
                ),
                ha="left" if arah_x > 0 else "right",
                va="center",
                fontsize=7.5,
                fontweight="bold",
                color="#222222",
                arrowprops={
                    "arrowstyle": "-",
                    "color": "#777777",
                    "linewidth": 0.8,
                    "shrinkA": 2,
                    "shrinkB": 2,
                },
                annotation_clip=False,
            )

    # Legenda tetap memuat kategori.
    ax.legend(
        wedges,
        [str(k) for k in comp.index],
        title="Kategori",
        loc="center left",
        bbox_to_anchor=(1.0, 0.5),
        frameon=False,
        fontsize=8,
        title_fontsize=9,
    )

    # Sisakan ruang untuk label persentase di luar pie.
    ax.set_xlim(-1.38, 1.38)
    ax.set_ylim(-1.30, 1.30)
    ax.axis("equal")

    fig.subplots_adjust(
        left=0.03,
        right=0.67,
        top=0.88,
        bottom=0.05,
    )

    return _fig_to_stream(fig)


def _chart_comp_stacked(S):
    d = S["df"]
    locs = S["top5"]
    tab = d[d["Lokasi"].isin(locs)].groupby("Lokasi")[CATEGORIES].sum().reindex(locs)
    tab = tab.loc[:, tab.sum(axis=0) > 0]
    fig, ax = plt.subplots(figsize=(7.2, 4.0))
    x = np.arange(len(tab))
    bottom = np.zeros(len(tab))
    for cat in sorted(tab.columns):  # urutan legenda alfabetis seperti Altair
        ax.bar(
            x,
            tab[cat].values,
            bottom=bottom,
            color=CAT_COLORS[cat],
            label=cat,
            width=0.62,
        )
        bottom += tab[cat].values
    for xi, tot in zip(x, bottom):
        ax.text(
            xi,
            tot,
            _compact(tot),
            ha="center",
            va="bottom",
            fontsize=7.5,
            color="#333333",
        )
    ax.set_xticks(x)
    ax.set_xticklabels(
        [textwrap.fill(s, 18) for s in tab.index],
        rotation=30,
        ha="right",
        rotation_mode="anchor",
        fontsize=7.8,
    )
    ax.set_ylabel("Total Volume", fontsize=8.5)
    ax.yaxis.set_major_formatter(_yfmt())
    ax.set_ylim(0, max(bottom.max() * 1.12, 1))
    ax.legend(
        title="Jenis Kendaraan",
        fontsize=8,
        title_fontsize=8,
        frameon=False,
        loc="upper right",
    )
    ax.set_title(
        "Komposisi Kategori per Lokasi",
        fontsize=11,
        fontweight="bold",
        loc="left",
        color="#222222",
    )
    _style_ax(ax)
    fig.tight_layout()
    return _fig_to_stream(fig)


# ============================================================================
# SAMPUL: gambar latar (lingkaran dekoratif + mosaik KPI dari data) + logo
# ============================================================================
def _cover_image(S, tahun):
    W, H = 1240, 1754  # A4 @150dpi
    fig = plt.figure(figsize=(W / 150, H / 150), dpi=150)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, W)
    ax.set_ylim(H, 0)
    ax.axis("off")
    fig.patch.set_facecolor("white")

    # lingkaran dekoratif
    cx, cy = 40, 120
    for r in (235, 262):
        ax.add_patch(Circle((cx, cy), r, fill=False, ec="#" + TEAL, lw=3.2))
    ax.add_patch(
        Circle((cx + 262 * np.cos(0.35), cy + 262 * np.sin(0.35)), 11, color="#" + TEAL)
    )
    ox, oy = 1250, 1500
    for r in (225, 252):
        ax.add_patch(Circle((ox, oy), r, fill=False, ec="#" + ORANGE, lw=3.2))
    ax.add_patch(
        Circle(
            (ox - 252 * np.cos(0.45), oy - 252 * np.sin(0.45)), 11, color="#" + ORANGE
        )
    )

    # logo (opsional)
    try:
        from PIL import Image

        x0 = 840
        for key in ("kominfo", "madiun"):
            raw = _LOGOS.get(key)
            if not raw:
                continue
            im = Image.open(io.BytesIO(base64.b64decode(raw))).convert("RGBA")
            h = 150
            w = int(im.width * h / im.height)
            x0_ = x0
            ax.imshow(
                np.asarray(im),
                extent=(x0_, x0_ + w, 150 + h, 150),
                zorder=5,
                aspect="auto",
            )
            x0 += w + 28
        ax.set_xlim(0, W)
        ax.set_ylim(H, 0)
    except Exception:
        pass

    # mosaik 3x3 menggunakan foto milik pengguna
    from pathlib import Path
    from PIL import Image, ImageOps

    tile, gap = 270, 20
    gx0 = (W - (3 * tile + 2 * gap)) // 2
    gy0 = 330

    cover_dir = Path(__file__).resolve().parent / "assets" / "cover"

    foto_cover = [
        "gambar 1",
        "gambar 2",
        "gambar 3",
        "gambar 4",
        "gambar 5",
        "gambar 6",
        "gambar 7",
        "gambar 8",
        "gambar 9",
    ]

    ekstensi = [".jpg", ".jpeg", ".png"]

    for i, nama_foto in enumerate(foto_cover):
        baris = i // 3
        kolom = i % 3
        x = gx0 + kolom * (tile + gap)
        y = gy0 + baris * (tile + gap)

        path_foto = next(
            (
                cover_dir / f"{nama_foto}{ext}"
                for ext in ekstensi
                if (cover_dir / f"{nama_foto}{ext}").is_file()
            ),
            None,
        )

        if path_foto is None:
            raise FileNotFoundError(f"Foto {nama_foto} tidak ditemukan di {cover_dir}")

        with Image.open(path_foto) as img:
            img = ImageOps.fit(
                img.convert("RGB"), (tile, tile), method=Image.Resampling.LANCZOS
            )

            ax.imshow(img, extent=(x, x + tile, y + tile, y), zorder=4)

        ax.add_patch(
            Rectangle(
                (x, y), tile, tile, fill=False, edgecolor="white", linewidth=2, zorder=5
            )
        )

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)

    return buf, (gy0 + 3 * tile + 2 * gap)


# ============================================================================
# PERAKIT DOKUMEN
# ============================================================================
LET = "abcdefghijklmnop"


def _setup_header_footer(doc, tahun):
    sec = doc.sections[0]
    hp = sec.header.paragraphs[0]
    hp.text = ""
    r = hp.add_run(f"Laporan Analisis Data Lalu Lintas Tahun {tahun}")
    _set_run_font(r, FONT_HEAD, 8.5, False, None, GREY)
    _para_border(hp, "bottom", TEAL, 10, 4)
    fp = sec.footer.paragraphs[0]
    fp.text = ""
    fp.paragraph_format.tab_stops.add_tab_stop(
        Cm(TEXT_WIDTH_CM), WD_TAB_ALIGNMENT.RIGHT
    )
    _para_border(fp, "top", ORANGE, 8, 4)
    r = fp.add_run(INSTANSI)
    _set_run_font(r, FONT_HEAD, 8, False, None, GREY)
    r = fp.add_run("\t\tHalaman ")
    _set_run_font(r, FONT_HEAD, 8.5, False, None, GREY)
    _add_field(fp, "PAGE", FONT_HEAD, 9, TEAL_DARK, True)


def _build_cover(doc, S, tahun, meta):
    img, y_bottom_px = _cover_image(S, tahun)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(0)
    spacer.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    spacer.paragraph_format.line_spacing = Pt(y_bottom_px / 1754 * 842 - 71 + 18)
    _add_anchored_picture(spacer, img, 21.0, 29.7)

    t = doc.add_table(rows=1, cols=2)
    _table_no_borders(t)
    left, right = t.rows[0].cells
    left.text = right.text = ""
    yr = str(tahun)
    for i, part in enumerate((yr[:2], yr[2:])):
        p = left.paragraphs[0] if i == 0 else left.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p.paragraph_format.line_spacing = Pt(52)
        r = p.add_run(part)
        _set_run_font(r, FONT_HEAD, 54, True, None, TEAL)
    tgl_awal = S["tgl_min"]
    tgl_akhir = S["tgl_max"]

    periode_data = (
        f"Periode {tgl_awal.day} {BULAN_ID[tgl_awal.month]} "
        f"{tgl_awal.year} – "
        f"{tgl_akhir.day} {BULAN_ID[tgl_akhir.month]} "
        f"{tgl_akhir.year}"
    )

    lines = [
        ("LAPORAN ANALISIS DATA", 12.5, 40),
        ("LALU LINTAS ", 26, 0),
        (f"KOTA {meta.get('kota', KOTA).upper()}", 26, 0),
        (periode_data, 10.5, 0),
    ]

    for i, (txt, sz, sp_) in enumerate(lines):
        p = right.paragraphs[0] if i == 0 else right.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.05
        r = p.add_run(txt)
        _set_run_font(r, FONT_HEAD, sz, True, None, "1A1A1A", spacing=sp_)
    right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    _table_fixed(t, [3.4, 12.1])

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(meta.get("instansi", INSTANSI))
    _set_run_font(r, FONT_HEAD, 13, False, None, "333333")
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(f"Tahun {tahun}")
    _set_run_font(r, FONT_HEAD, 13, False, None, "333333")


def _build_kata_pengantar(doc, S, meta):
    p = doc.add_paragraph()
    p.paragraph_format.page_break_before = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(18)
    r = p.add_run("KATA PENGANTAR")
    _set_run_font(r, FONT_HEAD, 15, True, None, TEAL_DARK)
    _para_border(p, "bottom", ORANGE, 14, 6)

    a, b = format_tanggal_id(S["tgl_min"]), format_tanggal_id(S["tgl_max"])
    psc = any("psc" in str(x).lower() for x in S["by_lok"].index)
    psc_txt = ", termasuk kawasan Pahlawan Street Center (PSC)" if psc else ""
    lokasi = f"{S['n_lokasi']} titik" if S["multi"] else "satu titik"
    paras = [
        "Dengan mengucapkan puji syukur ke hadirat Tuhan Yang Maha Esa, Laporan Analisis Data Lalu Lintas "
        f"CCTV Kota {meta.get('kota', KOTA)} dapat disusun dan disajikan sebagai bagian dari upaya "
        "penyediaan data dan informasi yang akurat di lingkungan Pemerintah Kota "
        f"{meta.get('kota', KOTA)}.",
        f"Pemantauan melalui CCTV analytic pada {lokasi} di wilayah Kota {meta.get('kota', KOTA)} "
        "menghasilkan data pergerakan orang dan kendaraan secara berkelanjutan. Data tersebut menjadi "
        "sumber informasi strategis untuk memahami pola keramaian, karakter moda transportasi, serta "
        f"pemanfaatan ruang publik dan ruas jalan yang dipantau{psc_txt}.",
        f"Laporan ini memuat hasil analisis data periode {a} sampai {b}, meliputi statistik deskriptif, "
        "tren, perbandingan antar lokasi, komposisi, serta kualitas data, yang disertai insight dan "
        "rekomendasi. Diharapkan laporan ini menjadi bahan pertimbangan bagi perangkat daerah terkait "
        "dalam merumuskan kebijakan dan meningkatkan kualitas pelayanan kepada masyarakat, sekaligus "
        "menjadi bagian dari komitmen Dinas Komunikasi dan Informatika dalam mendukung pengambilan "
        "kebijakan berbasis data dan pemanfaatan teknologi informasi.",
        "Kami mengucapkan terima kasih kepada seluruh pihak yang telah berpartisipasi dan mendukung "
        "proses penyusunan laporan ini. Semoga laporan ini memberikan manfaat dan kontribusi positif bagi "
        f"pembangunan Kota {meta.get('kota', KOTA)} secara berkelanjutan.",
    ]
    for t in paras:
        _body(doc, t, indent=True, after=8)

    now = datetime.now()
    t = doc.add_table(rows=1, cols=2)
    _table_no_borders(t)
    c = t.rows[0].cells[1]
    c.text = ""
    items = [
        (
            f"{meta.get('kota', KOTA)}, {now.day} {BULAN_ID[now.month]} {now.year}",
            False,
        ),
        *[(x, False) for x in meta.get("jabatan", JABATAN_PENANDATANGAN).split("\n")],
    ]
    for i, (txt, bd) in enumerate(items):
        p = c.paragraphs[0] if i == 0 else c.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.15
        _set_run_font(p.add_run(txt), FONT_HEAD, 11)
    gap = c.add_paragraph()
    gap.paragraph_format.space_after = Pt(0)
    gap.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    gap.paragraph_format.line_spacing = Pt(62)
    p = c.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.1
    _set_run_font(p.add_run(meta.get("nama", NAMA_PENANDATANGAN)), FONT_HEAD, 11, True)
    p = c.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.1
    _set_run_font(
        p.add_run(f"NIP. {meta.get('nip', NIP_PENANDATANGAN)}"), FONT_HEAD, 11
    )
    _table_fixed(t, [8.5, 7.0])


def _build_toc(doc):
    p = doc.add_paragraph()
    p.paragraph_format.page_break_before = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(18)
    r = p.add_run("DAFTAR ISI")
    _set_run_font(r, FONT_HEAD, 15, True, None, TEAL_DARK)
    _para_border(p, "bottom", ORANGE, 14, 6)

    tp = doc.add_paragraph()

    def _fld(kind):
        r = tp.add_run()
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), kind)

        if kind == "begin":
        fc.set(qn("w:dirty"), "true")
        
        r._r.append(fc)
        return r

    _fld("begin")
    r = tp.add_run()
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = ' TOC \\o "1-2" \\h \\z \\u '
    r._r.append(it)
    _fld("separate")
    r = tp.add_run(
        "Daftar isi akan terisi otomatis saat dokumen dibuka di Microsoft Word "
        '(pilih "Yes/Ya" jika muncul pertanyaan pembaruan). Jika belum muncul: '
        "klik kanan di sini, pilih Update Field, lalu Update entire table."
    )
    _set_run_font(r, FONT_HEAD, 9.5, None, True, GREY)
    _fld("end")


def _static_rows_note(doc, n, shown):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(
        f"Catatan: tabel menampilkan {shown} dari {n} periode; data lengkap tersedia pada "
        f"tabel data detail di dashboard."
    )
    _set_run_font(r, FONT_HEAD, 8.5, None, True, GREY)


def _status_styler(col_idx):
    def f(ri, ci, val):
        if ci == col_idx:
            return {
                "bg": STATUS_FILL.get(val, None),
                "color": STATUS_FONT.get(val),
                "bold": True,
            }
        return None

    return f


def _span_hari_txt(S):
    return f"{S['n_hari']} hari"


def generate_docx_report(df: pd.DataFrame, dq: pd.DataFrame, meta: dict) -> io.BytesIO:
    """
    df   : data tidy yang SUDAH difilter dashboard (Lokasi, Tanggal, kategori..., Total)
    dq   : parser.data_quality_report(df)
    meta : {'periode_label', 'lokasi_terpilih', 'sumber_data', ...opsional: kota, instansi,
            kepala_nama, kepala_nip, jabatan}
    """

    meta = meta or {}
    S = _analyze(df, dq, meta)
    d, g, period = S["df"], S["g"], S["period"]
    n_p, multi = S["n_periods"], S["multi"]
    tahun = S["tgl_max"].year
    cnt = _Counter()
    a_tgl, b_tgl = format_tanggal_id(S["tgl_min"]), format_tanggal_id(S["tgl_max"])
    cats = S["cats_present"]
    kat_frasa = _join_id(cats) if cats else "beberapa kategori"
    lokasi_frasa = (
        f"{S['n_lokasi']} titik pemantauan"
        if multi
        else f"titik pemantauan {d['Lokasi'].iloc[0]}"
    )
    cp = S["comp_pct"]
    insights = _build_insights(S)

    doc = _setup_document()
    _setup_header_footer(doc, tahun)
    _build_cover(doc, S, tahun, meta)
    _build_kata_pengantar(doc, S, meta)
    _build_toc(doc)

    # ======================================================================
    # BAB I
    # ======================================================================
    _h1(doc, 1, "Pendahuluan")
    _h2(doc, "1.1 Latar Belakang")
    psc = any("psc" in str(x).lower() for x in S["by_lok"].index)
    _body(
        doc,
        "Pengelolaan ruang kota yang efektif memerlukan pemahaman yang baik tentang pergerakan "
        f"orang dan kendaraan. Pemantauan berbasis CCTV analytic pada {lokasi_frasa} menghasilkan data "
        f"pergerakan {kat_frasa} secara berkelanjutan, sehingga menjadi sumber informasi strategis "
        "untuk memahami pola keramaian, karakteristik moda yang digunakan, serta dinamika aktivitas "
        "pada kawasan yang dipantau.",
    )
    _body(
        doc,
        "Dinamika volume yang terus berubah menuntut pengelolaan kawasan yang tidak hanya efektif "
        "dan terencana, tetapi juga adaptif. Pengelolaan yang tidak didukung analisis data yang "
        "memadai berisiko mengabaikan periode kritis, ketimpangan beban antar lokasi, serta masalah "
        "pencatatan yang dapat menurunkan keandalan informasi.",
    )
    _body(
        doc,
        f"Oleh karena itu, analisis statistik terhadap data CCTV periode {a_tgl} hingga {b_tgl} "
        "dilakukan untuk mengidentifikasi tren, perbandingan antar lokasi, komposisi kategori, dan "
        "kualitas data. Hasilnya diharapkan memberikan insight yang objektif dan berbasis data "
        "sebagai landasan perumusan rekomendasi kebijakan dan strategi pengelolaan.",
    )

    _h2(doc, "1.2 Tujuan")
    _body(
        doc,
        "Tujuan dari laporan analisis ini adalah sebagai berikut:",
        indent=False,
        keep_next=True,
    )
    tujuan = [
        "Menggambarkan karakteristik dan kondisi umum volume lalu lintas (orang dan kendaraan) "
        "pada seluruh data yang dianalisis.",
        f"Mengidentifikasi tren dan pola volume dari waktu ke waktu secara {period.lower()}.",
    ]
    if multi:
        tujuan.append(
            "Membandingkan tingkat keramaian antar lokasi pemantauan untuk mendukung alokasi "
            "sumber daya."
        )
    tujuan.append(
        "Mengetahui komposisi volume berdasarkan kategori (orang dan jenis kendaraan), "
        "secara keseluruhan maupun per lokasi."
    )
    tujuan.append("Menilai kualitas data pencatatan CCTV pada setiap lokasi.")
    tujuan.append(
        "Memberikan insight dan rekomendasi berbasis data sebagai bahan pertimbangan "
        "pengelola dalam perencanaan dan pengambilan keputusan."
    )
    for i, t in enumerate(tujuan):
        _bullet(doc, t, marker=f"{LET[i]}.")

    # ======================================================================
    # BAB II
    # ======================================================================
    _h1(doc, 2, "Metodologi")
    _h2(doc, "2.1 Data")
    sumber = meta.get("sumber_data", "file data CCTV")
    _body(
        doc,
        f"Data yang digunakan merupakan data sekunder hasil pencatatan sistem CCTV analytic pada "
        f"{lokasi_frasa}, mencakup {S['n_hari']} hari data pada periode {a_tgl} sampai {b_tgl} "
        f"(sumber: {sumber}). Data memuat variabel waktu {period.lower()}, lokasi, serta jumlah menurut "
        f"kategori {kat_frasa}, dan total pada setiap pengamatan. Setelah proses pembersihan "
        f"(cleaning) diperoleh {_fmt_id(n_p)} baris data {period.lower()}.",
    )
    _body(
        doc,
        "Sistem menghitung orang dan kendaraan yang melintas pada titik pemasangan kamera. Oleh "
        "karena itu, data kendaraan (motor, mobil, truk, dan bus) dapat mencakup kendaraan yang hanya "
        "melintas dan tidak seluruhnya berhenti atau berkunjung ke kawasan"
        + (
            ", sedangkan kategori Orang merupakan representasi yang lebih dekat dengan aktivitas "
            "pengunjung aktual."
            if "Orang" in cats
            else "."
        ),
    )
    _body(
        doc,
        f"Data diagregasi untuk analisis tren secara {period.lower()} sehingga menghasilkan "
        f"{n_p} periode observasi.",
    )

    # Tabel 1: data per periode
    periode_tabel = {
        "Harian": "Hari",
        "Mingguan": "Minggu",
        "Bulanan": "Bulan",
    }

    # Pilih data sesuai periode dari filter dashboard
    if period == "Keseluruhan":
        t1_g = g.copy()
        t1_cap = "Data Volume Lalu Lintas Keseluruhan"
        t1_note = None

    else:
        # Tetap menggunakan agregasi sesuai filter yang dipilih
        if n_p <= 24 and n_p >= 1:
            t1_g = g.copy()
            t1_note = None
        else:
            # Jika observasi terlalu banyak, tampilkan 12 periode awal
            # dan 12 periode akhir tanpa mengubah frekuensi agregasi
            t1_g = pd.concat([g.head(12), g.tail(12)])
            t1_note = (n_p, 24)

        nama_periode = periode_tabel.get(period, period)
        t1_cap = f"Data Volume Lalu Lintas per {nama_periode}"

        if t1_note:
            t1_cap += " (12 Periode Awal dan 12 Periode Akhir)"

    _caption_table(doc, cnt, t1_cap)

    hdr = ["Waktu"] + cats + ["Total"]
    first_w = 3.4
    other = (TEXT_WIDTH_CM - first_w) / (len(hdr) - 1)

    rows = [
        [r["Label"]] + [_fmt_id(r[c]) for c in cats] + [_fmt_id(r["Total"])]
        for _, r in t1_g.iterrows()
    ]

    if t1_note:
        rows.insert(12, ["⋮"] + [""] * (len(hdr) - 1))

    _table(
        doc,
        hdr,
        rows,
        [first_w] + [other] * (len(hdr) - 1),
        ["left"] + ["right"] * (len(hdr) - 1),
        font_size=8,
    )

    if t1_note:
        _static_rows_note(doc, *t1_note)

    _h2(doc, "2.2 Metode Analisis")
    _body(
        doc,
        "Analisis data dilakukan menggunakan pendekatan statistik deskriptif, meliputi:",
        indent=False,
        keep_next=True,
    )
    metode = [
        (
            "Analisis Statistik Deskriptif",
            "Analisis Statistik Deskriptif digunakan untuk memberikan gambaran umum karakteristik volume lalu lintas melalui nilai total, rata-rata, "
            "minimum, median, dan maksimum, untuk mengetahui tingkat dan variasi volume selama periode "
            "pengamatan.",
        ),
        (
            "Analisis Tren",
            "Analisis ini digunakan untuk mengamati perubahan volume dari waktu ke waktu dalam bentuk grafik deret waktu (total "
            "seluruh lokasi dan per lokasi) guna mengidentifikasi pola kenaikan, penurunan, serta "
            "periode volume tinggi dan rendah.",
        ),
    ]
    if multi:
        metode.append(
            (
                "Perbandingan Antar Lokasi",
                "Perbandingan antar lokasi digunakan untuk membandingkan total volume setiap lokasi untuk mengidentifikasi lokasi dengan "
                "volume tertinggi maupun terendah (di antara lokasi aktif).",
            )
        )
    metode.append(
        (
            "Analisis Komposisi",
            "Analisis Komposisi digunakan untuk menghitung kontribusi masing-masing kategori terhadap total volume, secara "
            "keseluruhan maupun per lokasi, untuk mengidentifikasi kategori yang mendominasi.",
        )
    )
    metode.append(
        (
            "Analisis Kualitas Data",
            "Analisis Kualitas Data digunakan untuk memeriksa kelengkapan data, hari bernilai nol, dan kecocokan nilai Total dengan "
            "penjumlahan kategori pada setiap lokasi, untuk menilai kesesuaian hasil analisis.",
        )
    )
    for i, (jd, ds) in enumerate(metode):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.9)
        p.paragraph_format.first_line_indent = Cm(-0.9)
        p.paragraph_format.space_after = Pt(1)
        _keep_next(p)
        r = p.add_run(f"{LET[i]}.\t{jd}")
        _set_run_font(r, FONT_HEAD, 11, True, None, NAVY)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(0.9))
        _body(doc, ds, indent=False, after=8).paragraph_format.left_indent = Cm(0.9)

    # ======================================================================
    # BAB III
    # ======================================================================
    _h1(doc, 3, "Hasil dan Pembahasan")

    # ---------- 3.1 ----------
    _h2(doc, "3.1 Statistik Deskriptif")
    _kpi_cards(
        doc,
        [
            ("Total Keseluruhan", _fmt_id(S["total_all"]), None),
            ("Rata-rata / Hari", _fmt_id(S["avg_daily"]), None),
            ("Lokasi Aktif", f"{S['n_aktif']} / {S['n_lokasi']}", None),
            ("Total Hari", f"{S['n_hari']}", None),
            ("Lokasi Tersibuk", S["top_lok"], _fmt_id(S["top_val"])),
        ],
    )

    # Narasi gambaran umum
    txt = (
        f"Data CCTV yang dianalisis mencakup {S['n_hari']} hari pengamatan, "
        f"yaitu dari {a_tgl} hingga {b_tgl}. Selama periode tersebut, "
        f"tercatat total volume sebesar **{_fmt_id(S['total_all'])}** "
        f"dengan rata-rata **{_fmt_id(S['avg_daily'])}** objek per hari. "
        f"Angka tersebut menggambarkan akumulasi hitungan orang dan "
        f"kendaraan yang terdeteksi pada titik pemantauan. Dengan demikian, "
        f"data ini dapat digunakan untuk melihat banyaknya objek yang "
        f"tercatat melintas, tetapi tidak secara langsung menunjukkan "
        f"kecepatan kendaraan atau kondisi kemacetan. "
    )

    if multi:
        txt += (
            f"Di antara lokasi yang dianalisis, **{S['top_lok']}** mencatat "
            f"volume terbesar, yaitu **{_fmt_id(S['top_val'])}** atau sekitar "
            f"**{_p(S['share_top1'])}** dari total volume seluruh lokasi. "
            f"Hal ini menunjukkan bahwa lokasi tersebut memberikan kontribusi "
            f"terbesar terhadap keseluruhan hitungan objek pada periode "
            f"pengamatan. "
        )

    _body(doc, txt)

    if n_p >= 3:
        base_g, base_name = g, f"periode ({period.lower()})"
        unit_stat = {
            "Harian": "hari",
            "Mingguan": "minggu",
            "Bulanan": "bulan",
        }.get(period, "periode")
    else:
        base_g = d.groupby("Tanggal")[CATEGORIES + ["Total"]].sum().reset_index()
        base_name = "hari"
        unit_stat = "hari"

    _body(
        doc,
        f"Statistik deskriptif dihitung berdasarkan total volume setiap {unit_stat}, untuk memahami karakteristik volume yang tercatat. "
        f"Dengan demikian, terdapat **{len(base_g)} observasi** yang "
        f"dirangkum melalui nilai rata-rata, minimum, median, dan maksimum. "
        f"Ringkasan tersebut membantu menunjukkan tingkat volume yang "
        f"umum terjadi sekaligus memperlihatkan perbedaan jumlah objek "
        f"yang terpantau antarkategori. ",
        keep_next=True,
    )

    _caption_table(
        doc,
        cnt,
        f"Statistik Deskriptif Volume per {base_name.split(' ')[0].capitalize()}",
    )

    rows = []
    for c in cats + ["Total"]:
        s_ = base_g[c].astype(float)
        rows.append(
            [
                c,
                _fmt_id(d[c].sum() if c != "Total" else S["total_all"]),
                _fmt_id(s_.mean()),
                _fmt_id(s_.min()),
                _fmt_id(s_.median()),
                _fmt_id(s_.max()),
            ]
        )

    _table(
        doc,
        ["Kategori", "Total", "Rata-rata", "Minimum", "Median", "Maksimum"],
        rows,
        [2.6, 3.0, 2.6, 2.4, 2.4, 2.5],
        ["left"] + ["right"] * 5,
        font_size=8.5,
        bold_last_row=True,
    )

    # Statistik volume total per periode

    ts = base_g["Total"].astype(float)
    mean_, med_ = ts.mean(), ts.median()
    min_, max_ = float(ts.min()), float(ts.max())

    # Label waktu dengan volume terendah dan tertinggi
    if "Label" in base_g.columns:
        labels = base_g["Label"].astype(str)
    else:
        labels = (
            base_g["Tanggal"]
            .apply(lambda x: format_tanggal_id(pd.to_datetime(x)))
            .astype(str)
        )

    label_min = labels.loc[ts.idxmin()]
    label_max = labels.loc[ts.idxmax()]

    # ------------------------------------------------------------
    # 1. Gambaran umum volume berdasarkan rata-rata dan median
    # ------------------------------------------------------------
    if mean_ > med_ * 1.05:
        mm = (
            "Nilai rata-rata yang lebih tinggi daripada median memberi "
            "indikasi bahwa beberapa periode dengan volume tinggi mungkin "
            "menarik rata-rata ke atas. Artinya, tingkat volume pada "
            "sebagian periode dapat lebih ramai dibandingkan kondisi "
            "yang tergambar dari nilai tengahnya. Periode-periode tersebut "
            "dapat dilihat lebih lanjut pada analisis tren."
        )

    elif med_ > mean_ * 1.05:
        mm = (
            "Nilai median yang lebih tinggi daripada rata-rata memberi "
            "indikasi bahwa beberapa periode dengan volume rendah mungkin "
            "menarik rata-rata ke bawah. Dengan kata lain, nilai rata-rata "
            "belum sepenuhnya menggambarkan volume yang berada di sekitar "
            "nilai tengah pengamatan. Periode dengan volume rendah perlu "
            "dilihat kembali untuk mengetahui apakah memang lebih sepi "
            "atau terdapat data yang belum lengkap."
        )

    else:
        mm = (
            "Nilai rata-rata dan median yang relatif berdekatan menunjukkan "
            "bahwa gambaran volume secara umum tidak jauh berbeda jika "
            "dilihat dari kedua ukuran tersebut. Namun, kondisi ini tidak "
            "berarti jumlah orang dan kendaraan yang terpantau selalu sama "
            "pada setiap periode."
        )

    _body(
        doc,
        f"Berdasarkan hasil statistik deskriptif, rata-rata volume lalu "
        f"lintas mencapai **{_fmt_id(mean_)} objek per {unit_stat}**, "
        f"sedangkan median sebesar **{_fmt_id(med_)}**. Rata-rata "
        f"menunjukkan jumlah volume secara umum selama pengamatan, "
        f"sementara median menunjukkan nilai tengah dari periode yang "
        f"dianalisis. {mm}",
    )

    # ------------------------------------------------------------
    # 2. Perbedaan periode paling sepi dan paling ramai
    # ------------------------------------------------------------
    rentang = max_ - min_

    if min_ > 0:
        rasio = max_ / min_

        if rasio >= 2:
            perbandingan = (
                f"Volume pada periode tertinggi sekitar **{_n(rasio)} kali** "
                f"volume pada periode terendah. Perbedaan ini menunjukkan "
                f"adanya jarak yang cukup besar antara kondisi dengan "
                f"volume paling rendah dan paling tinggi."
            )
        else:
            perbandingan = (
                f"Volume pada periode tertinggi sekitar **{_n(rasio)} kali** "
                f"volume pada periode terendah. Meskipun selisihnya tetap "
                f"perlu diperhatikan, besar kecilnya perbedaan ini perlu "
                f"dibaca bersama pola volume pada periode lainnya."
            )

        tambahan_minimum = (
            f"Periode terendah tercatat pada **{label_min}**, dengan "
            f"volume **{_fmt_id(min_)} objek**. Periode tertinggi terjadi "
            f"pada **{label_max}**, dengan volume **{_fmt_id(max_)} objek**. "
            f"Selisih keduanya mencapai **{_fmt_id(rentang)} objek**. "
            f"{perbandingan}"
        )

    else:
        tambahan_minimum = (
            f"Periode terendah tercatat pada **{label_min}**, dengan "
            f"volume **nol**, sedangkan periode tertinggi terjadi pada "
            f"**{label_max}**, mencapai **{_fmt_id(max_)} objek**. "
            f"Selisih keduanya sebesar **{_fmt_id(rentang)} objek**. "
            f"Volume nol perlu diperiksa terlebih dahulu untuk memastikan "
            f"apakah memang tidak ada objek yang terpantau atau terdapat "
            f"data yang belum tercatat."
        )

    _body(
        doc,
        f"Perbedaan volume antarkurun pengamatan terlihat dari periode "
        f"terendah dan tertinggi. {tambahan_minimum} Temuan ini membantu "
        f"menunjukkan bahwa jumlah objek yang melintas pada titik CCTV "
        f"dapat berbeda dari satu periode ke periode lainnya.",
    )

    # ------------------------------------------------------------
    # 3. Pemeriksaan periode dengan volume sangat rendah
    # Menggunakan base_g agar konsisten dengan periode statistik tabel
    # ------------------------------------------------------------
    median_positif = float(ts[ts > 0].median()) if (ts > 0).any() else 0.0

    if len(ts) >= 7 and median_positif > 0:
        batas_rendah = 0.10 * median_positif
        mask_rendah = ts < batas_rendah
        periode_rendah = labels.loc[mask_rendah]
        jumlah_rendah = int(mask_rendah.sum())

        if jumlah_rendah > 0:
            contoh_rendah = _list_ex(periode_rendah.tolist(), n=3)

            _body(
                doc,
                f"Selain perbedaan antara periode terendah dan tertinggi, "
                f"terdapat **{jumlah_rendah} periode** dengan volume kurang "
                f"dari 10% median volume positif, antara lain "
                f"**{contoh_rendah}**. Volume yang sangat rendah dapat "
                f"mencerminkan kondisi lalu lintas yang memang lebih sepi, "
                f"tetapi juga perlu diperiksa kemungkinan data yang belum "
                f"lengkap atau gangguan pencatatan CCTV. Oleh karena itu, "
                f"periode tersebut sebaiknya dibandingkan dengan informasi "
                f"kualitas data sebelum digunakan untuk menarik kesimpulan "
                f"mengenai kondisi lalu lintas.",
            )

    # ---------- 3.2 ----------
    _h2(doc, "3.2 Analisis Tren")

    if n_p >= 2:
        g = S["g"]
        nilai = g["Total"].astype(float).to_numpy()
        label_periode = g["Label"].astype(str).tolist()
        n = len(nilai)

        valid = np.asarray(S["valid_mask"], dtype=bool)
        low = np.asarray(S["low_mask"], dtype=bool)

        # --------------------------------------------------
        # A. TREN TOTAL
        # --------------------------------------------------
        _figure(doc, _chart_trend_total(S))
        _caption_figure(doc, cnt, f"Line Chart Tren Volume Lalu Lintas ({period})")

        # Tren keseluruhan: membandingkan sepertiga awal dan akhir.
        if (
            S["trend_dir"] in ("naik", "turun", "stabil")
            and S["first_m"] is not None
            and S["last_m"] is not None
            and S["trend_pct"] is not None
        ):
            awal = float(S["first_m"])
            akhir = float(S["last_m"])
            perubahan = abs(float(S["trend_pct"]))

            if S["trend_dir"] == "naik":
                narasi = (
                    f"Secara keseluruhan, rata-rata jumlah objek pada "
                    f"sepertiga awal pengamatan sebesar {_fmt_id(awal)} "
                    f"objek, kemudian menjadi {_fmt_id(akhir)} objek "
                    f"pada sepertiga akhir. Perubahan ini menunjukkan "
                    f"kecenderungan peningkatan sekitar {_p(perubahan)}, "
                    f"sehingga pencatatan pada bagian akhir pengamatan "
                    f"relatif lebih tinggi. Temuan ini menggambarkan "
                    f"peningkatan volume objek yang terdeteksi, tetapi "
                    f"belum cukup untuk menyimpulkan bahwa kepadatan "
                    f"atau kemacetan lalu lintas ikut meningkat."
                )

            elif S["trend_dir"] == "turun":
                narasi = (
                    f"Rata-rata jumlah objek pada sepertiga awal "
                    f"pengamatan sebesar {_fmt_id(awal)} objek, sedangkan "
                    f"pada sepertiga akhir menjadi {_fmt_id(akhir)} objek. "
                    f"Perubahan tersebut menunjukkan kecenderungan "
                    f"penurunan sekitar {_p(perubahan)}, sehingga jumlah "
                    f"objek yang tercatat pada bagian akhir relatif lebih "
                    f"rendah. Penurunan ini perlu dibaca bersama kualitas "
                    f"data karena volume pencatatan yang lebih rendah "
                    f"belum tentu sepenuhnya mencerminkan perubahan "
                    f"aktivitas lalu lintas yang sebenarnya."
                )

            else:
                narasi = (
                    f"Rata-rata jumlah objek pada sepertiga awal "
                    f"pengamatan sebesar {_fmt_id(awal)} objek dan "
                    f"sepertiga akhir sebesar {_fmt_id(akhir)} objek. "
                    f"Perbedaan relatif sekitar {_p(perubahan)} membuat "
                    f"tren keseluruhan tergolong stabil berdasarkan "
                    f"ambang klasifikasi yang digunakan. Kondisi ini "
                    f"menunjukkan bahwa rata-rata kedua bagian pengamatan "
                    f"tidak berbeda jauh, meskipun volume pada periode "
                    f"tertentu masih dapat mengalami kenaikan atau "
                    f"penurunan."
                )

            _body(doc, narasi)

        # Periode dengan volume tertinggi dan terendah.
        pi = S["peak_idx"]
        puncak = float(g.loc[pi, "Total"])

        if S["valid_mask"].any():
            ti = S["trough_idx"]
            terendah = float(g.loc[ti, "Total"])
            selisih = puncak - terendah

            _body(
                doc,
                f"Pencatatan tertinggi terjadi pada {_lbl(S, pi)}, "
                f"sebanyak {_fmt_id(puncak)} objek, sedangkan nilai "
                f"terendah di antara periode yang memenuhi kriteria "
                f"penilaian tren terjadi pada {_lbl(S, ti)}, sebanyak "
                f"{_fmt_id(terendah)} objek. Selisih keduanya mencapai "
                f"{_fmt_id(selisih)} objek, yang menunjukkan adanya "
                f"perbedaan volume pencatatan yang cukup penting untuk "
                f"dicermati. Periode tertinggi dapat menjadi acuan untuk "
                f"meninjau waktu dengan volume terdeteksi besar, sedangkan "
                f"periode terendah perlu dibaca dengan memperhatikan "
                f"kelengkapan pencatatannya.",
            )
        else:
            _body(
                doc,
                f"Pencatatan tertinggi terjadi pada {_lbl(S, pi)}, "
                f"sebanyak {_fmt_id(puncak)} objek. Nilai tersebut "
                f"merupakan volume terbesar yang tercatat selama "
                f"pengamatan, tetapi belum dapat dibandingkan dengan "
                f"periode terendah yang memenuhi kriteria penilaian tren. "
                f"Pola pencatatan pada periode lain perlu diperiksa "
                f"sebelum menarik kesimpulan mengenai perubahan volume "
                f"secara keseluruhan.",
            )

        # Tingkat fluktuasi.
        if S["cv"] is not None:
            cv = float(S["cv"])

            if cv >= 0.5:
                narasi = (
                    f"Volume pencatatan menunjukkan variasi yang besar "
                    f"dibandingkan rata-ratanya, dengan koefisien variasi "
                    f"{_n(cv, 2)}. Hal ini mengindikasikan bahwa jumlah "
                    f"objek yang tercatat tidak merata antarperiode, "
                    f"sehingga nilai rata-rata saja kurang memadai untuk "
                    f"menggambarkan pola pengamatan. Periode dengan "
                    f"perubahan paling menonjol perlu diperiksa lebih "
                    f"lanjut untuk membedakan variasi yang wajar dari "
                    f"kemungkinan masalah pencatatan."
                )
            elif cv >= 0.25:
                narasi = (
                    f"Variasi volume pencatatan cukup terlihat "
                    f"dibandingkan rata-ratanya, dengan koefisien variasi "
                    f"{_n(cv, 2)}. Artinya, jumlah objek yang terdeteksi "
                    f"masih mengalami perubahan antarperiode meskipun "
                    f"besar variasinya tidak termasuk kategori tertinggi "
                    f"berdasarkan ambang yang digunakan. Informasi ini "
                    f"menunjukkan bahwa evaluasi pola waktu tetap penting "
                    f"agar perubahan pada periode tertentu tidak tertutup "
                    f"oleh nilai rata-rata keseluruhan."
                )
            else:
                narasi = (
                    f"Variasi volume pencatatan relatif kecil dibandingkan "
                    f"rata-ratanya, dengan koefisien variasi {_n(cv, 2)}. "
                    f"Hasil ini menunjukkan bahwa perubahan relatif "
                    f"antarperiode cenderung terbatas berdasarkan ambang "
                    f"klasifikasi yang digunakan. Kondisi tersebut tidak "
                    f"berarti volume selalu sama, sehingga periode dengan "
                    f"perubahan mencolok tetap perlu diperhatikan."
                )

            _body(doc, narasi)

        # Periode dengan volume sangat rendah.
        if S["low_mask"].any():
            indeks_rendah = np.flatnonzero(low)
            contoh_rendah = indeks_rendah[:3]

            narasi = (
                f"Terdapat {int(S['low_mask'].sum())} periode dengan "
                f"pencatatan di bawah 10% median periode yang memiliki "
                f"volume positif. Contoh periode tersebut meliputi "
                + _join_id(
                    [
                        f"{label_periode[j]} ({_fmt_id(nilai[j])} objek)"
                        for j in contoh_rendah
                    ]
                )
                + ". Volume yang sangat rendah dapat mencerminkan "
                "penurunan jumlah objek yang terdeteksi, tetapi juga "
                "perlu diperiksa dari sisi kelengkapan data dan kondisi "
                "kamera. Pemeriksaan ini penting terutama apabila "
                "pencatatan rendah berlangsung berturut-turut."
            )

            if S["tail_low"]:
                narasi += (
                    f" Sebanyak {S['tail_low']} periode dengan volume "
                    f"sangat rendah juga muncul berturut-turut pada "
                    f"bagian akhir pengamatan, sehingga kesinambungan "
                    f"pencatatan pada rentang tersebut layak diverifikasi."
                )

            _body(doc, narasi)

        # --------------------------------------------------
        # B. TREN PER LOKASI
        # --------------------------------------------------
        _figure(doc, _chart_trend_lokasi(S))
        _caption_figure(
            doc, cnt, "Line Chart Tren Volume per Lokasi (5 Lokasi Tertinggi)"
        )
        _table_tren_lokasi(doc, S)

        t5 = S["share_top5"]

        _body(
            doc,
            f"Lima lokasi dengan total pencatatan tertinggi, yaitu "
            f"{_join_id(S['top5'])}, menyumbang {_p(t5)} dari keseluruhan "
            f"objek yang tercatat. Besarnya kontribusi ini menunjukkan "
            f"bahwa pencatatan terkonsentrasi pada kelompok lokasi tersebut "
            f"dalam data yang dianalisis. Lokasi-lokasi ini layak menjadi "
            f"prioritas ketika meninjau pola volume, dengan tetap "
            f"mempertimbangkan durasi dan kelengkapan data setiap lokasi.",
        )

        ups = sorted(
            [
                (lokasi, hasil[0])
                for lokasi, hasil in S["loc_trend"].items()
                if hasil[1] == "naik"
            ],
            key=lambda x: -x[1],
        )

        dns = sorted(
            [
                (lokasi, hasil[0])
                for lokasi, hasil in S["loc_trend"].items()
                if hasil[1] == "turun"
            ],
            key=lambda x: x[1],
        )

        stb = [
            lokasi for lokasi, hasil in S["loc_trend"].items() if hasil[1] == "stabil"
        ]

        # Ringkas arah perubahan lokasi tanpa menguraikan tiap lokasi
        # dalam paragraf terpisah.
        bagian_tren = []

        if ups:
            bagian_tren.append(
                "kecenderungan meningkat pada "
                + _join_id(
                    [f"{lokasi} (sekitar {_p(abs(persen))})" for lokasi, persen in ups]
                )
            )

        if dns:
            bagian_tren.append(
                "kecenderungan menurun pada "
                + _join_id(
                    [f"{lokasi} (sekitar {_p(abs(persen))})" for lokasi, persen in dns]
                )
            )

        if stb:
            bagian_tren.append("perubahan relatif stabil pada " + _join_id(stb))

        if bagian_tren:
            narasi = (
                "Perbandingan antar-lokasi memperlihatkan "
                + "; ".join(bagian_tren)
                + ". Perbedaan arah tersebut menunjukkan bahwa perubahan "
                "volume tidak selalu berlangsung seragam pada setiap "
                "titik pemantauan. Lokasi dengan kecenderungan meningkat "
                "dapat ditinjau untuk memahami pola pertambahan objek "
                "yang tercatat, sedangkan lokasi yang menurun perlu "
                "diperiksa bersama kualitas datanya. Klasifikasi stabil "
                "juga tidak berarti jumlah objek selalu sama pada setiap "
                "periode."
            )
            _body(doc, narasi)

        _body(
            doc,
            "Secara keseluruhan, hasil tren membantu menunjukkan kapan "
            "volume pencatatan relatif tinggi, seberapa besar variasinya, "
            "dan lokasi mana yang memperlihatkan perubahan paling jelas. "
            "Informasi tersebut dapat digunakan untuk menentukan prioritas "
            "pemantauan dan pemeriksaan kualitas data. Kesimpulan tentang "
            "kemacetan atau kebutuhan rekayasa lalu lintas tetap memerlukan "
            "informasi pendukung seperti kecepatan kendaraan, kepadatan, "
            "kondisi jalan, dan validitas rekaman CCTV.",
        )

    else:
        _body(
            doc,
            "Data pada pengaturan saat ini hanya mencakup satu periode "
            "pengamatan, sehingga kecenderungan perubahan dari waktu ke "
            "waktu belum dapat dinilai. Hasil tersebut hanya memberikan "
            "gambaran volume objek yang tercatat pada periode tersedia. "
            "Perbandingan tren baru dapat dilakukan ketika data mencakup "
            "lebih dari satu periode.",
        )

    # ---------- 3.3 ----------
    if multi:
        _h2(doc, "3.3 Perbandingan Antar Lokasi")
        N = S["top_n"]
        _body(
            doc,
            f"Perbandingan total volume antar lokasi disajikan pada Gambar {cnt.gambar + 1} "
            f"(lokasi teratas) dan Gambar {cnt.gambar + 2} (lokasi terendah di antara lokasi aktif).",
            keep_next=True,
        )
        rank = S["by_lok"].head(N)
        _figure(
            doc,
            _chart_rank_bar(rank, f"Lokasi Teratas berdasarkan Total Volume (Top {N})"),
        )
        _caption_figure(doc, cnt, f"Bar Chart {N} Lokasi dengan Total Volume Tertinggi")
        t = (
            f"**{S['top_lok']}** merupakan lokasi dengan volume tertinggi, yaitu {_fmt_id(S['top_val'])} "
            f"({_p(S['share_top1'])} dari total)"
        )
        if len(S["by_lok"]) >= 3:
            t += f"; tiga lokasi teratas secara kumulatif menyumbang {_p(S['share_top3'])}"
        t += ". "
        if len(S["by_lok"]) >= 2 and S["by_lok"].iloc[1] > 0:
            t += (
                f"Volume lokasi teratas {_n(S['by_lok'].iloc[0] / S['by_lok'].iloc[1])} kali lipat "
                f"dibanding peringkat kedua ({S['by_lok'].index[1]}). "
            )
        _body(doc, t)

        avg_map = dict(zip(dq["Lokasi"], dq["rata_rata_harian"]))

        _caption_table(
            doc,
            cnt,
            f"Peringkat {N} Lokasi dengan Volume Tertinggi",
        )

        tot_l = float(S["by_lok"].sum())

        rows = [
            [
                i + 1,
                loc,
                _fmt_id(v),
                _p(v / tot_l * 100 if tot_l else 0),
                _fmt_id(avg_map.get(loc)),
            ]
            for i, (loc, v) in enumerate(rank.items())
        ]

        _table(
            doc,
            [
                "No",
                "Lokasi",
                "Total Volume",
                "Kontribusi",
                "Rata-rata/Hari",
            ],
            rows,
            [1.0, 5.0, 3.0, 2.5, 4.0],
            ["center", "left", "right", "right", "right"],
            font_size=8,
        )

        low = S["rank_low"]
        _figure(
            doc,
            _chart_rank_bar(
                low,
                f"Lokasi dengan Volume Terendah (di antara lokasi aktif, {len(low)} lokasi)",
            ),
        )
        _caption_figure(
            doc, cnt, f"Bar Chart {len(low)} Lokasi Aktif dengan Total Volume Terendah"
        )
        if len(low):
            kurang = [x for x in low.index if S["status"].get(x) != ST_OK]
            t = (
                f"Di antara lokasi aktif, volume terendah tercatat pada **{low.index[0]}** "
                f"({_fmt_id(low.iloc[0])}). "
            )
            if S["small_locs"]:
                t += (
                    f"Terdapat {len(S['small_locs'])} lokasi aktif yang kontribusinya kurang dari 1% dari "
                    f"total volume. "
                )
            if kurang:
                t += (
                    f"{len(kurang)} dari {len(low)} lokasi terendah memiliki catatan kualitas data "
                    f"(mis. sebagian data kosong), sehingga rendahnya volume pada lokasi tersebut perlu "
                    f"diverifikasi sebelum disimpulkan sebagai kondisi riil."
                )
            else:
                t += "Seluruh lokasi terendah memiliki data yang lengkap, sehingga volume rendah kemungkinan mencerminkan kondisi riil."
            _body(doc, t)
        if S["n_inactive"]:
            _body(
                doc,
                f"Sebanyak {S['n_inactive']} lokasi tidak disertakan pada peringkat terendah karena "
                f"berstatus tidak ada data / kamera tidak aktif.",
            )
        sec_no = 4
    else:
        sec_no = 3

        _body(
            doc,
            f"Komposisi jumlah orang dan kendaraan yang tercatat "
            f"selama pengamatan disajikan melalui pie chart pada "
            f"Gambar {cnt.gambar + 1}. Rincian jumlah dan persentase "
            f"setiap kategori dapat dilihat pada tabel berikutnya. "
            f"Perbandingan ini menunjukkan kategori yang paling "
            f"banyak tercatat dan seberapa besar kontribusinya "
            f"terhadap keseluruhan objek yang terdeteksi.",
            keep_next=True,
        )

    # ---------- 3.4 Komposisi ----------
    _h2(doc, f"3.{sec_no} Komposisi Berdasarkan Kategori")

    if len(cp) >= 1:
        # Komposisi keseluruhan
        _figure(doc, _chart_comp_pie(S))
        _caption_figure(
            doc,
            cnt,
            "Pie Chart Komposisi Volume Berdasarkan Kategori",
        )

        _caption_table(doc, cnt, "Komposisi Keseluruhan")

        rows = [[k, _fmt_id(v), _p(cp[k])] for k, v in S["comp_present"].items()]

        _figure(doc, _chart_comp_bar(S))
        _caption_figure(
            doc,
            cnt,
            "Bar Chart Distribusi per Kategori",
        )

        # --------------------------------------------------
        # A. CERITA KOMPOSISI KESELURUHAN
        # --------------------------------------------------
        if len(cp) >= 2:
            top_c = cp.index[0]
            top_p = float(cp.iloc[0])
            sec_c = cp.index[1]
            sec_p = float(cp.iloc[1])
            orang = float(S["share"]["Orang"])
            kendaraan_pribadi = float(S["share_priv"])

            # Menafsirkan kategori dominan berdasarkan hasil aktual.
            if kendaraan_pribadi >= 70:
                interpretasi = (
                    f"Gabungan sepeda motor dan mobil mencapai "
                    f"{_p(kendaraan_pribadi)} dari seluruh volume, "
                    f"sehingga pencatatan sangat didominasi oleh "
                    f"dua kategori kendaraan tersebut. Pola ini "
                    f"mengindikasikan bahwa pengelolaan arus kendaraan, "
                    f"akses keluar-masuk, dan kebutuhan parkir layak "
                    f"menjadi perhatian dalam perencanaan operasional. "
                    f"Besarnya volume kendaraan yang terdeteksi belum "
                    f"dapat langsung diartikan sebagai kemacetan, "
                    f"karena data ini tidak mengukur kecepatan maupun "
                    f"kepadatan lalu lintas."
                )

            elif kendaraan_pribadi >= 50:
                interpretasi = (
                    f"Sepeda motor dan mobil secara bersama-sama "
                    f"menyumbang {_p(kendaraan_pribadi)} dari seluruh "
                    f"volume, sehingga kendaraan pribadi merupakan "
                    f"kelompok utama dalam pencatatan. Kondisi ini "
                    f"mengindikasikan bahwa pengaturan pergerakan "
                    f"kendaraan dan area parkir patut menjadi bagian "
                    f"dari evaluasi lokasi. Volume pencatatan yang "
                    f"besar tetap perlu dibedakan dari kondisi macet, "
                    f"karena kepadatan dan kelancaran arus tidak "
                    f"dapat dipastikan hanya dari jumlah objek."
                )

            elif top_c == "Orang" and top_p >= 40:
                interpretasi = (
                    f"Kategori Orang menjadi komponen terbesar dengan "
                    f"proporsi {_p(top_p)}, menunjukkan bahwa orang "
                    f"yang terdeteksi memberikan kontribusi besar "
                    f"terhadap pencatatan keseluruhan. Temuan ini "
                    f"dapat menjadi alasan untuk memberi perhatian "
                    f"pada akses, kenyamanan, dan keselamatan pejalan "
                    f"kaki di lokasi yang dipantau. Angka tersebut "
                    f"tidak otomatis menggambarkan jumlah pengunjung "
                    f"unik atau tujuan kedatangan mereka, sehingga "
                    f"interpretasi aktivitas kawasan perlu "
                    f"mempertimbangkan konteks lokasi."
                )

            else:
                interpretasi = (
                    f"Komposisi pencatatan dipimpin oleh kategori "
                    f"{top_c} sebesar {_p(top_p)}, diikuti {sec_c} "
                    f"sebesar {_p(sec_p)}, tanpa dominasi gabungan "
                    f"sepeda motor dan mobil yang mencapai separuh "
                    f"volume. Pola ini menunjukkan bahwa gambaran "
                    f"pencatatan perlu mempertimbangkan lebih dari "
                    f"satu kategori, bukan hanya kategori terbesar. "
                    f"Perencanaan pemantauan sebaiknya memperhatikan "
                    f"kebutuhan tiap kategori yang terdeteksi serta "
                    f"perbedaan karakter antar-lokasi."
                )

            # Soroti porsi kategori terkecil.
            kategori_kecil = cp.index[-1]
            persen_kecil = float(cp.iloc[-1])

            interpretasi += (
                f" Pada sisi lain, kategori {kategori_kecil} memiliki "
                f"proporsi terkecil, yaitu "
                f"{_p(persen_kecil, 2 if persen_kecil < 1 else 1)}. "
                f"Porsi yang kecil menunjukkan kontribusi pencatatan "
                f"yang rendah dibandingkan kategori lain, tetapi "
                f"tidak dengan sendirinya berarti kategori tersebut "
                f"tidak penting bagi pengelolaan lokasi."
            )

            _body(doc, interpretasi)

            # Makna porsi orang dan kendaraan pribadi.
            if orang > 0:
                if kendaraan_pribadi >= 50 and orang >= 20:
                    narasi_orang = (
                        f"Walaupun pencatatan didominasi sepeda motor "
                        f"dan mobil, kategori Orang masih menyumbang "
                        f"{_p(orang)} dari keseluruhan volume. Hal ini "
                        f"menunjukkan bahwa aspek pergerakan kendaraan "
                        f"dan kebutuhan pejalan kaki sama-sama perlu "
                        f"diperhatikan. Evaluasi sebaiknya mencakup "
                        f"kesesuaian jalur kendaraan, akses pejalan "
                        f"kaki, dan keselamatan di sekitar titik "
                        f"pemantauan tanpa mengasumsikan bahwa seluruh "
                        f"kendaraan merupakan pengunjung lokasi."
                    )

                elif orang < 15:
                    narasi_orang = (
                        f"Kategori Orang menyumbang "
                        f"{_p(orang)} dari total pencatatan, jauh "
                        f"lebih kecil dibandingkan gabungan sepeda "
                        f"motor dan mobil yang mencapai "
                        f"{_p(kendaraan_pribadi)}. Perbedaan ini "
                        f"menunjukkan bahwa data yang tercatat lebih "
                        f"banyak berasal dari objek kendaraan daripada "
                        f"orang pada titik pemantauan tersebut. "
                        f"Sebelum mengartikan pola ini sebagai "
                        f"rendahnya aktivitas pejalan kaki, perlu "
                        f"diperiksa fungsi lokasi, cakupan pandang "
                        f"kamera, serta konsistensi pendeteksian "
                        f"kategori Orang."
                    )

                else:
                    narasi_orang = (
                        f"Kategori Orang menyumbang {_p(orang)} dari "
                        f"total pencatatan, sehingga keberadaannya "
                        f"tetap menjadi bagian yang perlu diperhatikan "
                        f"bersama kategori kendaraan. Komposisi ini "
                        f"dapat digunakan sebagai gambaran awal "
                        f"perbandingan objek orang dan kendaraan yang "
                        f"terdeteksi pada titik pemantauan. Angka "
                        f"tersebut tidak cukup untuk memastikan "
                        f"apakah orang yang terdeteksi merupakan "
                        f"pengunjung, pekerja, atau orang yang "
                        f"hanya melintas."
                    )

                _body(doc, narasi_orang)

            # Soroti kendaraan berat jika proporsinya relevan.
            truk = float(S["share"]["Truk"])
            bus = float(S["share"]["Bus"])

            kategori_khusus = []

            if truk >= 5:
                kategori_khusus.append(
                    f"Truk menyumbang {_p(truk)} dari volume, sehingga "
                    f"pergerakan kendaraan berat layak dicermati "
                    f"terutama pada lokasi dengan keterbatasan ruang "
                    f"atau interaksi dengan pejalan kaki"
                )
            elif truk >= 1:
                kategori_khusus.append(
                    f"Truk memiliki proporsi {_p(truk)}; meskipun "
                    f"lebih kecil dibandingkan kategori dominan, "
                    f"keberadaannya tetap dapat menjadi pertimbangan "
                    f"dalam evaluasi akses dan ruang gerak kendaraan"
                )

            if bus >= 2:
                kategori_khusus.append(
                    f"Bus menyumbang {_p(bus)}, sehingga pola "
                    f"pergerakannya dapat ditinjau lebih lanjut "
                    f"untuk memahami kebutuhan akses dan ruang "
                    f"berhenti di lokasi"
                )
            elif 0 < bus < 2:
                kategori_khusus.append(
                    f"Bus memiliki proporsi {_p(bus, 2)}, sehingga "
                    f"kontribusinya terhadap total pencatatan relatif "
                    f"kecil pada rentang data ini"
                )

            if kategori_khusus:
                _body(
                    doc,
                    ". ".join(kategori_khusus)
                    + ". Informasi komposisi ini dapat menjadi dasar "
                    "untuk menentukan hal yang perlu diperiksa lebih "
                    "lanjut, tetapi belum membuktikan adanya masalah "
                    "keselamatan atau hambatan arus kendaraan. "
                    "Penentuan tindakan operasional tetap perlu "
                    "mempertimbangkan kondisi lapangan dan pola "
                    "pencatatan pada periode terkait.",
                )

        else:
            # Jika hanya satu kategori memiliki volume positif.
            kategori_satu = cp.index[0]
            persen_satu = float(cp.iloc[0])

            _body(
                doc,
                f"Hanya kategori {kategori_satu} yang memiliki "
                f"volume positif dalam hasil komposisi, dengan "
                f"proporsi {_p(persen_satu)} dari kategori yang "
                f"tercatat. Kondisi ini membuat perbandingan "
                f"antar-kategori belum dapat memberikan gambaran "
                f"komposisi yang beragam. Jika kategori lain "
                f"seharusnya muncul pada lokasi tersebut, periksa "
                f"kembali kelengkapan data dan hasil pendeteksian "
                f"sebelum menarik kesimpulan substantif.",
            )

        # --------------------------------------------------
        # B. KONSISTENSI ANTARA TOTAL DAN KATEGORI
        # --------------------------------------------------
        if abs(S["comp_sum"] - S["total_all"]) > 0.5:
            _body(
                doc,
                f"Jumlah seluruh kategori yang digunakan untuk "
                f"menghitung komposisi adalah "
                f"{_fmt_id(S['comp_sum'])}, berbeda dari nilai "
                f"Total sumber sebesar {_fmt_id(S['total_all'])}. "
                f"Perbedaan ini berarti distribusi kategori belum "
                f"sepenuhnya sejalan dengan total yang tercatat "
                f"pada sumber data dan dapat berkaitan dengan "
                f"kategori kosong atau nilai Total yang tidak cocok. "
                f"Karena persentase komposisi dihitung berdasarkan "
                f"jumlah kategori, hasilnya perlu dibaca bersama "
                f"informasi kualitas data sebelum digunakan untuk "
                f"menentukan prioritas operasional.",
                size=10.5,
                italic=True,
            )

        # --------------------------------------------------
        # C. KOMPOSISI PER LOKASI
        # --------------------------------------------------
        if multi:
            _figure(doc, _chart_comp_stacked(S))
            _caption_figure(
                doc,
                cnt,
                "Stacked Bar Chart Komposisi Kategori per Lokasi "
                "(5 Lokasi Tertinggi)",
            )

            tab = (
                d[d["Lokasi"].isin(S["top5"])]
                .groupby("Lokasi")[CATEGORIES]
                .sum()
                .reindex(S["top5"])
            )

            dom_l = {}

            for loc, r in tab.iterrows():
                if r.sum() > 0:
                    dom_l[loc] = (
                        r.idxmax(),
                        r.max() / r.sum() * 100,
                    )

            if dom_l:
                beda = [
                    (lokasi, kategori, persen)
                    for lokasi, (kategori, persen) in dom_l.items()
                    if kategori != cp.index[0]
                ]

                rincian_lokasi = _join_id(
                    [
                        f"{kategori} dominan di {lokasi} " f"({ _p(persen, 0) })"
                        for lokasi, (kategori, persen) in dom_l.items()
                    ]
                )

                if beda:
                    narasi_lokasi = (
                        f"Komposisi per lokasi memperlihatkan bahwa "
                        f"kategori dominan tidak selalu sama dengan "
                        f"pola keseluruhan. Hasil per lokasi adalah "
                        f"{rincian_lokasi}, dan lokasi dengan pola "
                        f"berbeda dari kategori dominan keseluruhan "
                        f"meliputi "
                        + _join_id(
                            [f"{lokasi} ({kategori})" for lokasi, kategori, _ in beda]
                        )
                        + ". Hal ini mengindikasikan bahwa kebutuhan "
                        f"pemantauan dapat berbeda antar-lokasi, "
                        f"sehingga pola total seluruh area sebaiknya "
                        f"tidak dijadikan satu-satunya dasar penilaian "
                        f"setiap titik."
                    )
                else:
                    narasi_lokasi = (
                        f"Pada lima lokasi dengan volume tertinggi, "
                        f"kategori dominan per lokasi adalah "
                        f"{rincian_lokasi}. Kesamaan kategori dominan "
                        f"ini menunjukkan bahwa pola utama pencatatan "
                        f"di lokasi-lokasi tersebut sejalan dengan "
                        f"komposisi keseluruhan. Besarnya proporsi "
                        f"setiap kategori tetap perlu diperhatikan "
                        f"karena kategori yang sama dapat memiliki "
                        f"porsi berbeda di tiap lokasi."
                    )

                _body(doc, narasi_lokasi)

    else:
        _body(
            doc,
            "Tidak terdapat kategori dengan volume positif yang "
            "dapat dibandingkan dalam pengamatan ini. Kondisi "
            "tersebut membuat pola komposisi belum bisa dijelaskan "
            "berdasarkan distribusi kategori. Kelengkapan data dan "
            "hasil pendeteksian perlu diperiksa sebelum komposisi "
            "digunakan sebagai dasar evaluasi.",
        )

    # ---------- 3.5 Kualitas Data ----------
    sec_no += 1
    _h2(doc, f"3.{sec_no} Kualitas Data")

    cnts = [
        (ST_OK, S["n_ok"]),
        (ST_PARTIAL, S["n_partial"]),
        (ST_MISMATCH, S["n_mismatch"]),
        (ST_ZERO, S["n_zero"]),
        (ST_INACTIVE, S["n_inactive"]),
    ]

    total_lokasi = int(S["n_lokasi"])
    jumlah_bermasalah = sum(int(jumlah) for status, jumlah in cnts if status != ST_OK)

    ringkas = _join_id(
        [f'{int(n_)} lokasi berstatus "{status}"' for status, n_ in cnts if n_ > 0]
    )

    # --------------------------------------------------
    # A. GAMBARAN UMUM STATUS KUALITAS DATA
    # --------------------------------------------------
    if total_lokasi > 0:
        persentase_lengkap = S["n_ok"] / total_lokasi * 100
        persentase_bermasalah = jumlah_bermasalah / total_lokasi * 100

        pengantar_kualitas = (
            f"Dari {total_lokasi} lokasi yang dianalisis, "
            f"{ringkas}. Rincian status kualitas data setiap "
            f"lokasi disajikan pada Tabel {cnt.tabel + 1}. "
            f"Hasil pemeriksaan menunjukkan bahwa "
            f"{int(S['n_ok'])} lokasi atau "
            f"{_p(persentase_lengkap)} telah berstatus data "
            f"lengkap, sedangkan {jumlah_bermasalah} lokasi "
            f"atau {_p(persentase_bermasalah)} masih memerlukan "
            f"perhatian."
        )
    else:
        pengantar_kualitas = (
            "Tidak terdapat lokasi yang dapat dievaluasi "
            "pada pengaturan data saat ini."
        )

    _body(
        doc,
        pengantar_kualitas,
        keep_next=True,
    )

    # --------------------------------------------------
    # B. TABEL STATUS KUALITAS DATA
    # --------------------------------------------------
    _caption_table(
        doc,
        cnt,
        "Status Kualitas Data per Lokasi",
    )

    dqs = dq.copy()

    rows = [
        [
            r["Lokasi"],
            _fmt_id(r["jumlah_hari_tercatat"]),
            _fmt_id(r["total_keseluruhan"]),
            _fmt_id(r["rata_rata_harian"]),
            _fmt_id(r["hari_nol"]),
            _fmt_id(r["hari_data_kosong"]),
            _fmt_id(r["hari_total_tidak_cocok"]),
            r["status"],
        ]
        for _, r in dqs.iterrows()
    ]

    _table(
        doc,
        [
            "Lokasi",
            "Hari Tercatat",
            "Total Keseluruhan",
            "Rata-rata Harian",
            "Hari Nol",
            "Hari Data Kosong",
            "Hari Total Tdk Cocok",
            "Status",
        ],
        rows,
        [3.7, 1.3, 2.0, 1.8, 1.0, 1.3, 1.4, 3.0],
        [
            "left",
            "right",
            "right",
            "right",
            "right",
            "right",
            "right",
            "center",
        ],
        font_size=7.2,
        cell_styler=_status_styler(7),
    )

    # --------------------------------------------------
    # C. ANALISIS LOKASI YANG MEMERLUKAN PERHATIAN
    # --------------------------------------------------
    bermasalah = dqs[dqs["status"] != ST_OK].copy()

    if len(bermasalah):
        nm = bermasalah.sort_values(
            "total_keseluruhan",
            ascending=False,
            na_position="last",
        )["Lokasi"].tolist()

        daftar_lokasi_masalah = _list_ex(nm, 5)

        _body(
            doc,
            f"Hasil pemeriksaan menemukan {len(bermasalah)} "
            f"lokasi yang statusnya belum lengkap atau memerlukan "
            f"verifikasi, yaitu {daftar_lokasi_masalah}. "
            f"Lokasi tersebut perlu diperiksa sebelum hasil "
            f"pencatatannya digunakan untuk menarik kesimpulan "
            f"mengenai perubahan jumlah orang dan kendaraan.",
        )

        # Lokasi dengan volume tinggi dan status bermasalah
        top_aff = [loc for loc in S["top5"] if S["status"].get(loc) != ST_OK]

        if top_aff:
            _body(
                doc,
                f"Perhatian khusus diperlukan pada "
                f"{_join_id(top_aff)}, karena lokasi tersebut "
                f"termasuk dalam lima lokasi dengan total "
                f"pencatatan tertinggi, tetapi status kualitas "
                f"datanya belum lengkap atau masih perlu "
                f"diverifikasi. Masalah pencatatan pada lokasi "
                f"bervolume tinggi berpotensi memengaruhi "
                f"interpretasi total dan perbandingan antar lokasi. "
                f"Oleh karena itu, data dari lokasi tersebut "
                f"perlu diperiksa sebelum digunakan sebagai "
                f"dasar penentuan prioritas operasional.",
            )

        # ----------------------------------------------
        # D. RINCIAN MASALAH BERDASARKAN STATUS
        # ----------------------------------------------
        # 1. Sebagian data kosong
        partial = dqs[dqs["status"] == ST_PARTIAL]

        if not partial.empty:
            jumlah_hari_kosong = int(
                pd.to_numeric(
                    partial["hari_data_kosong"],
                    errors="coerce",
                )
                .fillna(0)
                .sum()
            )

            nama_partial = _list_ex(
                partial.sort_values(
                    "hari_data_kosong",
                    ascending=False,
                )["Lokasi"].tolist(),
                5,
            )

            _body(
                doc,
                f"Status sebagian data kosong ditemukan pada "
                f"{len(partial)} lokasi, yaitu {nama_partial}. "
                f"Secara keseluruhan, terdapat "
                f"{_fmt_id(jumlah_hari_kosong)} catatan hari "
                f"dengan data kosong pada kelompok lokasi ini. "
                f"Kekosongan nilai membuat jumlah objek pada "
                f"periode terkait tidak dapat dipastikan "
                f"sepenuhnya. Nilai kosong juga tidak boleh "
                f"langsung dianggap sebagai nol, karena keduanya "
                f"memiliki makna yang berbeda.",
            )

        # 2. Total tidak sesuai penjumlahan kategori
        mismatch = dqs[dqs["status"] == ST_MISMATCH]

        if not mismatch.empty:
            jumlah_tidak_cocok = int(
                pd.to_numeric(
                    mismatch["hari_total_tidak_cocok"],
                    errors="coerce",
                )
                .fillna(0)
                .sum()
            )

            nama_mismatch = _list_ex(
                mismatch.sort_values(
                    "hari_total_tidak_cocok",
                    ascending=False,
                )["Lokasi"].tolist(),
                5,
            )

            _body(
                doc,
                f"Ketidaksesuaian antara nilai Total dan "
                f"penjumlahan kategori ditemukan pada "
                f"{len(mismatch)} lokasi, yaitu {nama_mismatch}. "
                f"Tercatat {_fmt_id(jumlah_tidak_cocok)} "
                f"catatan hari dengan total yang tidak cocok "
                f"pada kelompok lokasi tersebut. Kondisi ini "
                f"menimbulkan ketidakpastian mengenai jumlah "
                f"objek yang benar. Pemeriksaan perlu diarahkan "
                f"pada kesesuaian Total dengan kategori Orang, "
                f"Motor, Mobil, Truk, dan Bus, serta pada "
                f"proses rekapitulasi sumber datanya.",
            )

        # 3. Hari bernilai nol
        zero = dqs[dqs["status"] == ST_ZERO]

        if not zero.empty:
            jumlah_hari_nol = int(
                pd.to_numeric(
                    zero["hari_nol"],
                    errors="coerce",
                )
                .fillna(0)
                .sum()
            )

            hari_tercatat_zero = int(
                pd.to_numeric(
                    zero["jumlah_hari_tercatat"],
                    errors="coerce",
                )
                .fillna(0)
                .sum()
            )

            nama_zero = _list_ex(
                zero.sort_values(
                    "hari_nol",
                    ascending=False,
                )["Lokasi"].tolist(),
                5,
            )

            _body(
                doc,
                f"Status perlu verifikasi nilai nol ditemukan "
                f"pada {len(zero)} lokasi, yaitu {nama_zero}. "
                f"Pada kelompok ini terdapat "
                f"{_fmt_id(jumlah_hari_nol)} hari dengan "
                f"nilai nol dari total "
                f"{_fmt_id(hari_tercatat_zero)} hari tercatat "
                f"pada lokasi-lokasi tersebut. Nilai nol dapat "
                f"mencerminkan tidak adanya objek yang terdeteksi, "
                f"tetapi juga dapat berkaitan dengan data yang "
                f"belum tercatat atau kendala perangkat. Karena "
                f"itu, hari bernilai nol perlu dibandingkan "
                f"dengan kondisi pencatatan dan log perangkat "
                f"sebelum ditafsirkan sebagai aktivitas yang "
                f"benar-benar tidak terjadi.",
            )

        # 4. Tidak ada data atau kamera tidak aktif
        inactive = dqs[dqs["status"] == ST_INACTIVE]

        if not inactive.empty:
            nama_inactive = _list_ex(
                inactive["Lokasi"].tolist(),
                5,
            )

            _body(
                doc,
                f"Status tidak ada data atau kamera tidak aktif "
                f"ditemukan pada {len(inactive)} lokasi, yaitu "
                f"{nama_inactive}. Data dari lokasi-lokasi ini "
                f"belum dapat memberikan gambaran yang dapat "
                f"diandalkan mengenai jumlah objek selama "
                f"pengamatan. Ketiadaan pencatatan tidak dapat "
                f"disimpulkan sebagai tidak adanya aktivitas "
                f"orang atau kendaraan. Status perangkat, "
                f"ketersediaan rekaman, dan proses pengambilan "
                f"data perlu diperiksa sebelum data lokasi "
                f"tersebut digunakan dalam perbandingan.",
            )

    else:
        _body(
            doc,
            "Seluruh lokasi berstatus data lengkap. Pemeriksaan "
            "tidak menemukan lokasi dengan status sebagian data "
            "kosong, total tidak cocok, perlu verifikasi nilai "
            "nol, atau tidak ada data/kamera tidak aktif. "
            "Hasil ini menunjukkan bahwa pemeriksaan kualitas "
            "yang diterapkan tidak mendeteksi masalah pada "
            "status lokasi, meskipun tidak berarti semua "
            "kemungkinan kesalahan pencatatan telah dipastikan "
            "tidak ada.",
        )

    # --------------------------------------------------
    # E. KESIAPAN DATA UNTUK DIGUNAKAN
    # --------------------------------------------------
    if jumlah_bermasalah > 0:
        _body(
            doc,
            "Berdasarkan hasil pemeriksaan, data dapat digunakan "
            "dengan mempertimbangkan status kualitas pada "
            "masing-masing lokasi. Lokasi dengan data lengkap "
            "dapat menjadi bagian dari analisis sesuai cakupan "
            "dan periode pengamatannya. Lokasi dengan data "
            "kosong, total tidak cocok, nilai nol yang perlu "
            "diverifikasi, atau tanpa data memerlukan pemeriksaan "
            "lebih lanjut. Setelah koreksi atau konfirmasi "
            "dilakukan, hasil rekapitulasi dan analisis perlu "
            "diperbarui agar keputusan yang diambil tidak "
            "bergantung pada data yang belum terverifikasi.",
        )
    else:
        _body(
            doc,
            "Berdasarkan status yang diperiksa, seluruh lokasi "
            "dapat melanjutkan ke tahap interpretasi hasil "
            "dengan tetap mengikuti batasan variabel CCTV. "
            "Data kualitas yang baik mendukung konsistensi "
            "perbandingan jumlah objek, tetapi tidak dengan "
            "sendirinya menjelaskan penyebab perubahan volume "
            "atau membuktikan kondisi kemacetan.",
        )

    # --------------------------------------------------
    # F. KETERANGAN STATUS KUALITAS DATA
    # --------------------------------------------------
    _callout(
        doc,
        "Keterangan Status Kualitas Data",
        [
            "**Data lengkap**: seluruh kategori yang diharapkan "
            "memiliki nilai valid dan Total sesuai dengan "
            "penjumlahan kategori.",
            "**Sebagian data kosong**: terdapat kategori yang "
            "kosong atau tidak bernilai numerik valid.",
            "**Total tidak cocok**: nilai Total tidak sesuai "
            "dengan penjumlahan kategori.",
            "**Perlu verifikasi nilai nol**: lebih dari 30% "
            "hari bernilai nol sehingga perlu diperiksa kembali.",
            "**Tidak ada data / kamera tidak aktif**: tidak ada "
            "nilai kategori valid atau seluruh Total bernilai nol.",
        ],
        accent=RED,
        bg=RED_LIGHT,
        title_color="B71C1C",
        body_size=9.5,
    )

    # ======================================================================
    # BAB IV
    # ======================================================================
    _h1(doc, 4, "Insight dan Rekomendasi")
    _h2(doc, "4.1 Insight")
    if not insights:
        _body(
            doc,
            "Data pada filter saat ini masih terbatas sehingga insight otomatis belum dapat "
            "dihasilkan. Perluas rentang tanggal, lokasi, atau ubah agregasi periode.",
        )
    else:
        for i, it in enumerate(insights):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.left_indent = Cm(0.9)
            p.paragraph_format.first_line_indent = Cm(-0.9)
            p.paragraph_format.tab_stops.add_tab_stop(Cm(0.9))
            _keep_next(p)
            r = p.add_run(f"{LET[i]}.\t{it['judul']}")
            _set_run_font(r, FONT_HEAD, 11, True, None, NAVY)
            r = p.add_run(f"   [{it['kelompok']}]")
            _set_run_font(r, FONT_HEAD, 8, True, None, ORANGE)
            _body(
                doc, it["insight"], indent=False, after=4
            ).paragraph_format.left_indent = Cm(0.9)

    _h2(doc, "4.2 Rekomendasi")
    _body(
        doc,
        "Rekomendasi difokuskan pada tindakan untuk menjaga kelancaran dan keselamatan lalu lintas, "
        "seperti penertiban parkir, pengaturan titik berhenti, pengelolaan akses kendaraan, serta "
        "perlindungan pejalan kaki. Urutan dimulai dari prioritas tinggi, kemudian sedang, dan rendah. "
        "Rekomendasi ditampilkan sesuai kondisi yang teridentifikasi, sehingga satu lokasi dapat memiliki "
        "lebih dari satu tindak lanjut. Jumlah objek yang tercatat CCTV tidak otomatis menunjukkan "
        "kemacetan; penertiban, penempatan petugas, atau perubahan pengaturan jalan perlu disesuaikan "
        "dengan masalah yang benar-benar terlihat saat pemeriksaan lapangan.",
        keep_next=True,
    )

    pal = {
        "Tinggi": (RED, RED_LIGHT),
        "Sedang": (ORANGE, ORANGE_LIGHT),
        "Rendah": (TEAL, TEAL_LIGHT),
    }
    priority_order = {"Tinggi": 0, "Sedang": 1, "Rendah": 2}
    group_order = {
        "Kualitas Data": 0,
        "Pola Waktu": 1,
        "Lokasi": 2,
        "Komposisi": 3,
        "Pejalan Kaki": 4,
        "Kendaraan": 5,
        "Angkutan Umum": 6,
        "Kendaraan Berat": 7,
        "Perencanaan": 8,
    }

    # Gabungkan rekomendasi dari insight utama dan bank rekomendasi tambahan.
    # Bank tambahan hanya aktif jika kondisi pemicunya ditemukan pada data.
    all_recs = []
    # Insight kualitas data tetap ditampilkan pada bagian insight, tetapi daftar
    # rekomendasi difokuskan pada tindakan pengelolaan lalu lintas.
    traffic_groups = {
        "Pola Waktu", "Lokasi", "Komposisi", "Pejalan Kaki",
        "Kendaraan", "Angkutan Umum", "Kendaraan Berat",
    }
    non_traffic_titles = {
        "Ada periode dengan volume sangat rendah yang perlu diverifikasi",
        "Pola bulanan dapat dijadikan dasar penjadwalan kegiatan",
        "Validasi dan pantau perubahan tren volume",
        "Jadikan periode ini sebagai baseline pemantauan",
        "Sejumlah lokasi aktif berkontribusi sangat kecil dan memiliki catatan kualitas data",
        "Sejumlah lokasi memiliki volume sangat kecil",
    }
    # Kriteria untuk insight otomatis ditulis eksplisit agar pembaca mengetahui
    # indikator, ambang pemicu, dan alasan rekomendasi dipilih. Nilai aktual
    # diambil dari statistik yang sama dengan yang membentuk insight.
    def _insight_trigger(it):
        title = it.get("judul", "")
        share = S.get("share", {})
        if title == "Tren volume menunjukkan peningkatan":
            return ("Perubahan rata-rata sepertiga akhir dibanding sepertiga awal > 0%; prioritas menjadi Tinggi jika kenaikan ≥ 50%, selain itu Sedang.", f"Kenaikan rata-rata: {_p(S.get('trend_pct', 0))}; awal {_fmt_id(S.get('first_m', 0))} dan akhir {_fmt_id(S.get('last_m', 0))} objek per periode")
        if title == "Tren volume menunjukkan penurunan":
            return ("Perubahan rata-rata sepertiga akhir dibanding sepertiga awal < 0%; perlu validasi kelengkapan data sebelum menyesuaikan layanan.", f"Penurunan rata-rata: {_p(abs(S.get('trend_pct', 0)))}; awal {_fmt_id(S.get('first_m', 0))} dan akhir {_fmt_id(S.get('last_m', 0))} objek per periode")
        if title == "Tren volume relatif stabil":
            return ("Arah tren tidak menunjukkan kenaikan atau penurunan yang berarti berdasarkan perbandingan rata-rata sepertiga awal dan akhir.", f"Perubahan rata-rata: {_p(abs(S.get('trend_pct', 0)))}; awal {_fmt_id(S.get('first_m', 0))} dan akhir {_fmt_id(S.get('last_m', 0))} objek per periode")
        if title == "Terdapat lonjakan volume ekstrem pada periode puncak":
            return ("Rasio volume periode puncak terhadap rata-rata ≥ 2,00 kali.", f"Rasio puncak: {_n(S.get('peak_ratio', 0), 2)} kali; volume puncak {_fmt_id(S['g'].loc[S['peak_idx'], 'Total'] if S.get('peak_idx') is not None else 0)} dan rata-rata {_fmt_id(S.get('mean_valid', 0))} objek")
        if title == "Periode puncak lebih ramai dibanding kondisi rata-rata":
            return ("Rasio volume periode puncak terhadap rata-rata ≥ 1,50 dan < 2,00 kali.", f"Rasio puncak: {_n(S.get('peak_ratio', 0), 2)} kali; rata-rata {_fmt_id(S.get('mean_valid', 0))} objek")
        if title == "Tidak ada lonjakan tajam; volume relatif merata antar periode":
            return ("Rasio volume periode puncak terhadap rata-rata < 1,50 kali.", f"Rasio puncak: {_n(S.get('peak_ratio', 0), 2)} kali; rata-rata {_fmt_id(S.get('mean_valid', 0))} objek")
        if title == "Volume sangat bergantung pada kendaraan pribadi":
            return ("Proporsi gabungan sepeda motor dan mobil ≥ 50% dari seluruh objek terpantau.", f"Motor + mobil: {_p(S.get('share_priv', 0))}; motor {_p(share.get('Motor', 0))} dan mobil {_p(share.get('Mobil', 0))}")
        if title == "Proporsi pejalan kaki masih rendah":
            return ("Proporsi pejalan kaki > 0% dan < 15%, bersamaan dengan proporsi kendaraan pribadi ≥ 50%.", f"Pejalan kaki: {_p(share.get('Orang', 0))}; kendaraan pribadi: {_p(S.get('share_priv', 0))}")
        if title == "Pejalan kaki mendominasi volume":
            return ("Pejalan kaki menjadi kategori terbesar dan proporsinya ≥ 40%.", f"Pejalan kaki: {_p(share.get('Orang', 0))}; kategori terbesar kedua: {S.get('comp_pct', pd.Series(dtype=float)).index[1] if len(S.get('comp_pct', [])) > 1 else '-'}")
        if title == "Komposisi relatif beragam tanpa satu kategori yang sangat dominan":
            return ("Tidak memenuhi kondisi dominasi kendaraan pribadi ≥ 50% maupun dominasi pejalan kaki ≥ 40%; tidak ada satu kategori yang mencapai 50%.", f"Dua kategori terbesar: {_p(S.get('comp_pct', pd.Series(dtype=float)).iloc[0] if len(S.get('comp_pct', [])) else 0)} dan {_p(S.get('comp_pct', pd.Series(dtype=float)).iloc[1] if len(S.get('comp_pct', [])) > 1 else 0)}")
        if title == "Kendaraan berat terpantau dalam komposisi lalu lintas":
            return ("Proporsi truk ≥ 5% dari total volume.", f"Proporsi truk: {_p(share.get('Truk', 0))}")
        if title == "Kendaraan berat (truk) berproporsi kecil namun berdampak pada ruang jalan":
            return ("Proporsi truk ≥ 1% dan < 5% dari total volume.", f"Proporsi truk: {_p(share.get('Truk', 0))}")
        if title == "Keramaian sangat terkonsentrasi pada beberapa lokasi":
            return ("Analisis multi-lokasi dengan minimal 3 lokasi; kontribusi gabungan 3 lokasi teratas ≥ 60%.", f"Kontribusi 3 lokasi teratas: {_p(S.get('share_top3', 0))}; lokasi: {_join_id(S.get('by_lok', pd.Series(dtype=float)).head(3).index.tolist())}")
        if title == "Keramaian cukup terkonsentrasi pada beberapa lokasi":
            return ("Analisis multi-lokasi dengan minimal 3 lokasi; kontribusi gabungan 3 lokasi teratas ≥ 40% dan < 60%.", f"Kontribusi 3 lokasi teratas: {_p(S.get('share_top3', 0))}")
        if title == "Sebaran volume antar lokasi relatif merata":
            return ("Analisis multi-lokasi dengan minimal 3 lokasi; kontribusi 3 lokasi teratas < 40%.", f"Kontribusi 3 lokasi teratas: {_p(S.get('share_top3', 0))}")
        return (f"Rekomendasi muncul karena kondisi pada indikator kelompok {it.get('kelompok', '-') } memenuhi aturan insight otomatis; rincian kondisi terukur ditampilkan pada hasil analisis.", re.sub(r"\*\*", "", it.get("insight", "")))

    for it in insights:
        if it["kelompok"] not in traffic_groups or it["judul"] in non_traffic_titles:
            continue
        trigger, actual = _insight_trigger(it)
        all_recs.append(
            {
                "rek_judul": it["rek_judul"],
                "rek": it["rek"],
                "kelompok": it["kelompok"],
                "prio": it["prio"],
                "evidence": None,
                "basis_prioritas": trigger,
                "nilai_analisis": actual,
            }
        )

    existing_keys = {
        re.sub(r"[^a-z0-9]+", " ", x["rek_judul"].lower()).strip() for x in all_recs
    }
    for extra in _build_policy_recommendations(S):
        if extra["kelompok"] not in traffic_groups:
            continue
        if extra.get("key") == "low_location_review":
            continue
        norm = re.sub(r"[^a-z0-9]+", " ", extra["rek_judul"].lower()).strip()
        if norm not in existing_keys:
            all_recs.append(extra)
            existing_keys.add(norm)

    # Urutkan seluruh rekomendasi yang sesuai dengan temuan berdasarkan urgensi,
    # lalu kelompok temuan. Tidak ada pemotongan jumlah rekomendasi agar kondisi
    # berbeda tetap memiliki tindak lanjut yang dapat dipertimbangkan.
    ranked_recs = sorted(
        all_recs,
        key=lambda item: (
            priority_order.get(item["prio"], 9),
            group_order.get(item["kelompok"], 9),
        ),
    )

    if ranked_recs:
        for it in ranked_recs:
            ac, bg = pal[it["prio"]]
            body = it["rek"]
            if it.get("evidence"):
                body = f"Temuan yang mendasari: {it['evidence']}\n\n{body}"

            # Judul langsung berupa tindakan; tidak memakai kode A, B, C, atau Ref.
            _callout(
                doc,
                it["rek_judul"],
                body,
                accent=ac,
                bg=bg,
                tag=f"PRIORITAS {it['prio'].upper()}",
                tag_color=ac if it["prio"] != "Sedang" else "B8730F",
            )

        _h2(doc, "4.3 Ringkasan Prioritas Tindak Lanjut")
        nt = sum(1 for x in ranked_recs if x["prio"] == "Tinggi")
        ns = sum(1 for x in ranked_recs if x["prio"] == "Sedang")
        nr = sum(1 for x in ranked_recs if x["prio"] == "Rendah")

        _body(
            doc,
            f"Sebanyak {len(ranked_recs)} rekomendasi dipilih dari temuan yang teridentifikasi, "
            f"terdiri atas {nt} prioritas tinggi, {ns} prioritas sedang, dan {nr} prioritas rendah. "
            f"Urutan pada tabel mengikuti tingkat urgensi yang sama dengan rekomendasi di atas. "
            f"Tindakan operasional atau perubahan fasilitas tetap perlu dikonfirmasi melalui "
            f"pemeriksaan kondisi lapangan sebelum diterapkan.",
            keep_next=True,
        )

        _body(
            doc,
            "Dasar prioritas mengikuti aturan pemicu yang digunakan pada analisis, bukan skor gabungan baru. "
            "Prioritas Tinggi ditetapkan pada kondisi yang menurut aturan memerlukan verifikasi atau perbaikan "
            "lebih dahulu; Sedang menunjukkan kondisi yang perlu dievaluasi atau dipantau; sedangkan Rendah "
            "menunjukkan tindak lanjut bertahap atau pemantauan. Batas indikator tidak disamaratakan karena "
            "setiap indikator memiliki satuan dan ambang analisis yang berbeda. Nilai aktual pada tabel "
            "merupakan hasil perhitungan berdasarkan data dan filter yang sedang digunakan.",
            size=10,
            italic=True,
        )

        _caption_table(doc, cnt, "Ringkasan Rekomendasi, Kriteria, Nilai Analisis, dan Prioritas")
        rows = [
            [
                it["rek_judul"],
                it["kelompok"],
                it.get("basis_prioritas", "Mengikuti kondisi temuan analisis"),
                it.get("nilai_analisis", it.get("evidence") or "Lihat uraian temuan"),
                it["prio"],
            ]
            for it in ranked_recs
        ]
        pf = {
            "Tinggi": ("FBE7E4", "B5391F"),
            "Sedang": ("FFF3D9", "8A6100"),
            "Rendah": ("E4F5E9", "1E7B3E"),
        }
        _table(
            doc,
            ["Rekomendasi", "Bidang", "Kriteria pemicu", "Hasil analisis", "Prioritas"],
            rows,
            [4.0, 2.0, 4.1, 3.3, 2.1],
            ["left", "center", "left", "center", "center"],
            font_size=7.2,
            cell_styler=lambda ri, ci, v: (
                {"bg": pf[v][0], "color": pf[v][1], "bold": True} if ci == 4 else None
            ),
        )
    else:
        _h2(doc, "4.3 Ringkasan Prioritas Tindak Lanjut")
        _body(
            doc,
            "Belum terdapat temuan yang cukup kuat untuk menetapkan tindakan pengelolaan lalu lintas "
            "tertentu dari data pada periode ini. Pertahankan pengaturan rutin dan lakukan pemantauan "
            "terhadap kelancaran arus, ketertiban parkir, akses keluar-masuk, serta keselamatan pejalan kaki. "
            "Tindakan khusus ditetapkan apabila pemeriksaan lapangan menunjukkan adanya hambatan atau risiko.",
        )

    # ======================================================================
    # BAB V
    # ======================================================================
    _h1(doc, 5, "Kesimpulan")
    _body(
        doc,
        "Berdasarkan hasil analisis data CCTV lalu lintas periode "
        f"{a_tgl} sampai {b_tgl}, dapat disimpulkan beberapa hal sesuai dengan tujuan analisis "
        "sebagai berikut:",
        indent=False,
    )
    ks = []
    ks.append(
        (
            "Karakteristik dan kondisi umum volume",
            f"Total volume tercatat sebesar {_fmt_id(S['total_all'])} selama {S['n_hari']} hari, dengan "
            f"rata-rata {_fmt_id(S['avg_daily'])} per hari"
            + (
                f" dari {S['n_aktif']} lokasi aktif (dari {S['n_lokasi']} lokasi)"
                if multi
                else ""
            )
            + f". Volume harian tertinggi terjadi pada {format_tanggal_id(S['day_max'])}.",
        )
    )
    if n_p >= 2:
        fr = {
            "naik": "cenderung meningkat",
            "turun": "cenderung menurun",
            "stabil": "relatif stabil",
            "na": "belum dapat ditentukan arahnya",
        }[S["trend_dir"]]
        ks.append(
            (
                "Tren dan pola volume",
                f"Secara {period.lower()}, volume {fr}; puncak terjadi pada {_lbl(S, S['peak_idx'])} "
                f"({_fmt_id(g.loc[S['peak_idx'], 'Total'])})"
                + (
                    f", dan terdapat {int(S['low_mask'].sum())} periode volume sangat rendah yang perlu "
                    f"diverifikasi."
                    if S["low_mask"].any()
                    else "."
                ),
            )
        )
    if multi:
        ks.append(
            (
                "Perbandingan antar lokasi",
                f"{S['top_lok']} merupakan lokasi tersibuk ({_p(S['share_top1'])} dari total); tiga lokasi "
                f"teratas menyumbang {_p(S['share_top3'])}. "
                + (
                    f"Sebanyak {len(S['small_locs'])} lokasi aktif berkontribusi kurang dari 1%."
                    if S["small_locs"]
                    else ""
                ),
            )
        )
    if len(cp):
        ks.append(
            (
                "Komposisi kategori",
                f"Kategori {cp.index[0]} paling besar ({_p(cp.iloc[0])})"
                + (
                    f", diikuti {cp.index[1]} ({_p(cp.iloc[1])}); motor dan mobil bersama-sama "
                    f"{_p(S['share_priv'])}."
                    if len(cp) > 1
                    else "."
                ),
            )
        )
    kk = []
    if S["n_inactive"]:
        kk.append(f"{S['n_inactive']} lokasi tanpa data/kamera tidak aktif")
    if S["n_partial"]:
        kk.append(f"{S['n_partial']} lokasi dengan sebagian data kosong")
    if S["n_mismatch"]:
        kk.append(f"{S['n_mismatch']} lokasi dengan Total tidak cocok")
    if S["n_zero"]:
        kk.append(f"{S['n_zero']} lokasi perlu verifikasi nilai nol")
    ks.append(
        (
            "Kualitas data",
            (
                (
                    "Ditemukan "
                    + _join_id(kk)
                    + ", sehingga hasil analisis agregat perlu dibaca dengan hati-hati "
                    "dan perbaikan pencatatan menjadi prioritas."
                )
                if kk
                else "Seluruh lokasi berstatus data lengkap sehingga hasil analisis dapat diandalkan."
            ),
        )
    )
    if insights:
        tinggi = [x["rek_judul"] for x in insights if x["prio"] == "Tinggi"]
        ks.append(
            (
                "Insight dan implikasi pengelolaan",
                (
                    (
                        f"Terdapat {len(tinggi)} rekomendasi berprioritas tinggi yang perlu ditindaklanjuti lebih "
                        f"dahulu, yaitu: "
                        + "; ".join(t_.lower() for t_ in tinggi)
                        + "."
                    )
                    if tinggi
                    else "Tidak ada rekomendasi berprioritas tinggi; pengelolaan dapat dilanjutkan dengan "
                    "pemantauan rutin dan perbaikan bertahap."
                ),
            )
        )
    for i, (jd, ds) in enumerate(ks, 1):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Cm(0.9)
        p.paragraph_format.first_line_indent = Cm(-0.9)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(0.9))
        _keep_next(p)
        r = p.add_run(f"{i}.\t{jd}")
        _set_run_font(r, FONT_HEAD, 11, True, None, NAVY)
        _body(doc, ds, indent=False, after=4).paragraph_format.left_indent = Cm(0.9)
    _body(
        doc,
        "Pemanfaatan data CCTV secara berkelanjutan dapat menjadi dasar yang kuat dalam pengambilan "
        "keputusan dan perencanaan kebijakan pengelolaan lalu lintas maupun ruang publik ke depan.",
        after=6,
    ).paragraph_format.space_before = Pt(8)

    # ======================================================================
    # LAMPIRAN
    # ======================================================================
    _h1(doc, 1, "Data Volume per Lokasi", label="LAMPIRAN")
    _caption_table(
        doc, cnt, "Total Volume per Lokasi dan Kategori (urut dari volume tertinggi)"
    )
    tab = (
        d.groupby("Lokasi")[cats + ["Total"]]
        .sum()
        .sort_values("Total", ascending=False)
    )
    rows = [
        [loc] + [_fmt_id(r_[c]) for c in cats] + [_fmt_id(r_["Total"])]
        for loc, r_ in tab.iterrows()
    ]
    rows.append(
        ["Total"]
        + [_fmt_id(tab[c].sum()) for c in cats]
        + [_fmt_id(tab["Total"].sum())]
    )
    fw = 4.4
    ow = (TEXT_WIDTH_CM - fw) / (len(cats) + 1)
    _table(
        doc,
        ["Lokasi"] + cats + ["Total"],
        rows,
        [fw] + [ow] * (len(cats) + 1),
        ["left"] + ["right"] * (len(cats) + 1),
        font_size=7.8,
        bold_last_row=True,
    )
    _body(
        doc,
        "Laporan ini dihasilkan secara otomatis oleh Dashboard Analisis CCTV pada "
        f"{format_tanggal_id(datetime.now(), with_time=True)}. Angka merepresentasikan hasil "
        "penghitungan pada masing-masing titik CCTV dan dapat dipengaruhi oleh kondisi teknis di "
        "lapangan.",
        indent=False,
        size=9,
        italic=True,
        color=GREY,
    )

    _force_update_fields_on_open(doc)
    _fix_package_order(doc)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ============================================================================
# Logo (tertanam, diambil dari sampul laporan pedoman)
# ============================================================================
_LOGOS = {}
for _line in """madiun=iVBORw0KGgoAAAANSUhEUgAAAHoAAACWCAMAAAAmAWwvAAABgFBMVEVVVx4PDg4aIBkbFhYQDg4VWE6upRPvqx3vcSnqKC0cZzpxcXHv19iHehfqqKr9/f0IlExDOyminJ8BbZrxXWAMNU0Eg7VkXV4BmdSDenteDw+GfYECxf9HPEHOvcEmZmZ5gR7sPkIAgz94gYGPgHykpKS7ydB+foH8+mLJwb4CwWIA//+qf3//AADxf4MAAAAveUr+9hH9/f0zg1AArFYAre306BEvRi0XOSgtNy0makbHxMUYRzErVjWKhYcbGRnRxxElJyTr5+ju2ROxpxIaWEQBt/gcGhqOhhZLWCQBtltLSEhraBrX1tcgHh6rqKh4dhkdGxvGuRFsaGkYUzVYVlcVKSqZlBV3dXa3treXlpcZGhr0hib4xxna0xD3lyMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABnolDNAAAAYHRSTlP/YPUhmP///////wL+//8B//7+/////////wX////+BP//////A///////AQYB/wD///3////////9//7///62//3+//////X////9//78/v/Q//7//v////7+Ef////+pmyH8AAAaLUlEQVR42tVcCXuiyrYtUdBrYsZOp/vMd773TRE4TQGCAcGIESPOZlD//794axdoHNJ9kz6nv/se57QxKKzaVXuvvXZVEXb1bzvY/ynox0fxmqfXfP5RvOYfH9O36YfizeP6dX1N9qWrPK5Or8QHm/f7MvRjv5+/ery+K1xd5aS7u9Z1/ip3h5+tO5zPtfrZLQp3dzl6beWuxGlxDl9qSTidw2f91mP6pkBffGxdvwIad3+89guPhFjIXfv9fL5QaN0VCrl8vi8AM+gCmkm/r08X7q4Lokk5n6DzBF24ur7Du1dB51uSsPkYNxT38+n1+o66Lde6Tu0TFvYf83f9ljCSTn+Hrwq0/iZ0DpfizCuh+7A0fYM+wI9rvF7fYRjzeMX/GXThLp9rFajDV6dFK/P91g506xqvr4PGyJK9+bvrrCkYTXFragV1YV50Lm573ad70xfp9HfC6kdpDxr9l38ldCHfwnXCavKxa9HhdCO/f92nkRPQ+X6/VRB+tDotoHHtCjq/gsYgvHascZNr+JpwzWs/l1mNxvdxpI4G6IJ/lxO+vDpdSNt3nUKLCMCbQosuv3vdWGOA+wSI7xf6aERmNW6KGIVrPQrve8z5sMfPrU+jLdI1PDwvoHN3rYJw7rTVfuEV0H3CykkI4Px1q9UqCEYo4NeCRI04KIjQzsG/+7CvlSdPu8JAox0Ygn6BrhZ+30cQ5NJLceYV0McHB8fHB1dXPx8cEI/hBb/TyYPvDq4OcFzRF46v6A2dX5/O/3x1IO5wcPyfefqGILSD7H4HB/+HOfzx+lsdYhy/AH2cG9eOat/gwE07+eMvQB9cFXratzrquS9ZfXzld7X7j/Xf//hY07r9L3Z4YaZpluSYv/fhSB1NG+a+AJ33G1pvon+LY3KvdaWr489A549ppH90q9/g0Ecw+yHNPPvQedHddVOvfhNs6UzTxrkN7E2rc2NNq02q3+iw4cFaJ/eS1ccFIB9a7otG679Hl59qWmMDm60C+rEwxCc/Oi+B6LbL32Kf/XLzTQx3Y1zYtTrXf/gcsm47jPG32G19AVur93Mb0JCcnRrG2dpD1u1q1WEDg72pa13JfXGIdPNHjPe9Vcg/Q/fraM6ZvzfOusw4V1TV428b13bEmy9iOz4R9UeR8dMOp3geSvYLyAspVA118Kb+rurcC2L+Uq/rZekjxfdf19B5DELnBd/mgRGqgK403wbdTFR1wF66SHcuQdXPHX581a9pD3tcojvADYK5GlpvjaS2sSwtY6fK94e7p91fb3p44QFOtvMtd9KGwfPS05uHGj0+MEqlYtLkeyNlHWqz3CZ0bqg1eHnXVwaqMZ+Xil8B3YxVo/RktN0dB9JdDO04vwmdtxraeLT1LTumUV4SsmW/mbu4p5ZKhmLG2yOuO7DR2qYUqavNVlGtpw2PaKCXpfnbR5r4gKnzomFwZrBNstDNula73oQ+uLquPecsx86gDQxYAFe1iRzfcNBNRhUVvdbmqppsYJOX9QqpaF5ZjcjuZdBcacNwnTo8KBUNdYqRtiz++sPCBdx20Gsq48HcSDZG0vxJe8htdzhc/H4FXWGcYrwNow3VYw6aYS2M1x8hblT2q25sqMnEg7ewZ1+xamt5uE4f9XV08cgIYseyESCBkQhSsi1Dff2hkA2MyeDC0IzgbsrIXkkfq7uKrWfoWQatyzRIwWlsY7AXmX8S9GvBjRTaHcS82Q5cRkHi2tYz9HAfuis+5mgncBLFBTST0zF4O7TuKGrInYFihkVAV307u1FjH3q4go4ZXR6FTA6VlW9+BXRzqs4Hk8QbucUirGbuGvryM9C6O2BeCm3N24gTneL8K6DleFlSFSlgcqyiB5XU1V60OutwvakoIEHVC6xkMJIlu1m1na8Z6yorleaqFA0cZ5DI3AtFD34GOnUzOzbigFJ0y2NmxHjEWeRWvwLaLYKPIgmsErZlJCL+WTdbBZcN9sXh+bEnTcOYPC7ehDZeCY2cXSwtPWkwlUJXjhDiIlheCq46NBNBT2QWDhSDTbw49kQGAUO8psON9b/U6qozxUUWM5SK7HpPRUVwGijl4z6bCSJ14IqOM/ESayEtQvQ87uRuQntwQWXBPtcUYw1tl3kSqJ7kIUTbIrg/w2YZh7tsAbbXm2GUKJNAmBFOqmto/JPYwPfZRbzd9WhPpCanHlOerbZcm1eMANDI3stlSUCbR8ThW+nj+p4Uku56qnfq6nLiDX4xSRAGCYsxECk0/q/I55P35/LNHrZ1IeGji/YaWieZ0LSmSF3yL8HTskIdvp+50qQJ9zehfANKPEZgcVjtsanXtrPgIuSLTzefbuj/C7aBDbiLm/c3Nx9uLgbqymo5og4EOXqTZP70FMtVsIR5pt3vJM1UKshMNQI15joPFMmY0m8el9eUYjCB/AnHzY18uiBPSJEX0gc0CGfPT4VfimSQqFPUIRZ6nHQD18uukAr31zsqBaEOq9tGsWQYbWjRgRc5I6YkZlVkLvSAGkrnNwL3E8HfnPuKMBz4wen56gO03kuttpF1PW5Dr4SSocZNncNqB/whXb2gzWzLC57mRrs6ClUIG91u2rqecfiAXXzI7p+hfLgQnJu2aX1WZpG1UsRPBkbPHKiJOrB0e2ILbXa4A001QJmUifH0pCace2hldcRdedQU0CG7WJu8AvnhRpYq//jHTptuzidWmnXAZ0U1dEdEzG1b56BxfYRKelsWkhiG7tR5qBaX6HHugcHMKIzDaUJ61L84/7QDfPNevhje1846u2364X0KrY8qJO3Aohh+58ScRrCiTDZuieFV9rCZMS8ZnskGrhuFagBpFpKKNz/9sAt8cfpweFjfA8ZnGTQG+wmCOGq6C5VVR4kaNfXN/MF2KJwv1GUQmr8MnLaqFp9WUtjcHWbZ6jW6Q1/+sAu8AQ24Im7gut6A28hJInk9kzhbTxLep77Mw3kwNZVQms5Ly6I6SOdWtqARQpOPWvdy8hLwMzSCaURUCq2SNOFyc1HEPDNpVu4V7rUzUxf+bLJgIA3iWJ0HiNcK34OmkK5pw8n5p5eAN62eVJuIgWgCxTPNWJzo7KGw2eGCzOy47XDbQl5nXqggd6ihL1d3oW8+TC61I1/+DPAmNOMI16nqDVw+UJ+eipDYuvmwU2kSmY3sSjJVpLjdZIsFcxzeTtZTKGtodLZ0ptWl958D3uhwuRKBCcGk8FmvCBaXy7buwKv6m9Ck1kY6gl/1ktB1lIFJVCI310VDBn3zScZgdS4+fRZ4A5okbRMs7nmuG8JlEa+CU7rSBnTKKNU2EeM0jMzplECb5Q2rhcXwr6FW8+XPDPIOtIPUx2xqAQeRl3iVW4JTslIzgwbHoLyeUJo0pt4krDR1mUcKs7KqwQR13Xw4n3Rq2l8m5y+59c26Oeux5h5lH6JIBQPu4obttMDuPE9opGRWtZjI9OpioiSyzbxpFHqKJSr+iSzLF9bwn1rv9OLTDzcbQBnuh/cfVn53c76CNopLtdKkzlQHvMkHC57S2Ti3AS3IDGgCO5yQfF0sQkoNSdimaPzLWa92eHjGLt7/cPP+/Pz8/Yc1PH6cX1idjiln+XRldbv4BGLgVagpBUXQgN7jVis6YxtkpjcrBqSjyszBL00mBBEusshqf/jw8OdLWX7/6b1sdi6Hlx3r4vy96GYw2+TyvqFpNYQ6/b7icNt9elqqC5dEAJCNZZByyorO2EaJa1epJp5OF9KAWWk6VIO07PJ9SZpcoNetTr3WoJWMbm8IO+VznBrWstWNf4LR37/PMpftusWnp8BoA3raBEGDHN1NOmMbotD2Zbfd5uZgGiqeApUWTuN0uhOkX3uYzT7Wjw43F1K6vb98/NgTLdHS195lp/PxYyro/WayLC0DsprJERLCU0TyzLzPivsMWohCOY4cBwqDB8agPamoShRGp+6qVFoftfrl+LJ+tNmGWn3c6ayNT+cI7FOXQ35DpzCDcS8oLZdQeSthuIaGHq2bNMcXsYpThXaFIOORpyRtt7yqV1b3vbSccrk8Mjvfr87V6NTJn0aTy/TMUVp9RIncDoSbBS4PIAyZ7a6F4Rq6X9O+d3SIKQwwtC/3yKtl12lWdSGQqpPT+k/dbq33sWOWf72l49cTx+zUf6pBK0zSUye/jvz6YQO5NIPGMAObN5NggjsXmS1b1ZTEC1ffraBTPepKgYHqUGmWBwRty23ucKdJ1akJM03LdEYnhCoO/CzTqpRDwNmpE8fqcFMiN9NHUzVoN7mH0s07DdSIV21J5A8xfXWwDW2zAFI9MCyLrHalOIgUL4zI08xf/+NWwGUgKdKvZtqE5zPojJPbEwENGp0PQNsDLwhoMkhvxpRNRP64plWnDPpQSGEIUqO0hPaE1aOI6p45Km3icevkdgsjPcwfnd1z1PNlAc3DoKROXXAaAhYyxElEsbkHnapwIvEADBQPGM00QpOX5pQAUujd42R0qf1Y3j9/W+aiw5WnkrrgZD3VBLwSqHQrgu5eP7sZQYsalCeLYkmNISvaBgkkeB01SbfkPYjbE7+r1Sb7bbp1xMz6KFqimuA05lDClgI+F8HlfP8iNFhlAGGkhGwyVYsBpRLmit4b7VuHKKEVuRc+oIB07RiSCNDNKZkeQaXMmYDetzqdLXLbLFIMtRJTpaOGMbfTueT9QT05FSS2NxTwMkFCVAE8zV0UzAMHGQRkNnera+jjbQ9PJ9Fl2eVKELvtaBCw9cqAv4fwp47gD773gfAyW7KbSbGowODBtGmj1CwVhRjec7MMWq82rditykloVu0m5246d0WDXd5DGAvozh60GGo7dsnBlBGK1qTanBql4oJvTEsf7LAZ520wKXwMxQe+pENLobZuCk1t7vnxJfIH6arbXe9LxRGCGDHtoq+5A1ILwnQJKJOk321zOA+pcgwTL7Y8V0wmMZNZcbrcs+vjt+Uh2FvTLnehRyK0oH7hVG3wWQS1YRgK4zQvUhUc/pw+vqOkiQqgbNH0qOophuRRCHJUe1CosZiZd3fMvh19j1yiacNt6NRoWqZCTIMcfpE8UrkYumqTp0QKrMKuVKDFknSWQlJQmvEwUtTVHMHeaN/CXbSepn0/Otk+nYrYtlo0FFeOFrAmMWVakE0UMdjWT1tSIav2dDf2xGxNHIXwSm9KCilcLb1tO/mtc6Y1HhDYWxF/WxZG201WLKUCRVVjMDePlXm2ULcjkCALDykaOYvjaIFKxRi4zanoAm9V+1CXb2YKdFxjpmlnmxGP7hY8wEdsiRzISOFF8F8rVEtieihVHVuKlHR5GY1rQ/6C/Rw2b8s0GaAqrJJNZcN1tlDMI60BT+ttnLw9maRd5DtsCUkYN5mYG4UweyoFU2F0eUeHU/VBy1x61WmDRWWUQJFMtUglXsRUfqYhaTkb2cusad0xSZLbPWRoI5KE8wTQ4BGuFEvohHY61XC5XX080xlIZYQ6SU4GkqcM1CBA1eTaVjoxo1vrPr/91TrUumhxzdxDrrpK2QmXqCzRc4kN530qPQWRkLaUPWrSZn19fZQup7oWd8FidtX1vNDkEa1Du4yJMkD0+SpTgakbWi2Fvk0TdXm1/m0zmkAJ5qHbVFTLRDqYz4OIby507c0q0FXeIPHCU9eZGrC0CSabMDVQ2+s1Sb8soG5POC3zg9As+h3KzPGz5Usi0MQeRWqEEgYVHMS8BxP2Vpt2pnGIhQwkrchnVCpglN1KODcG1vNWDSsVhiew+N5vaA1Ocu1kNOHrNcU2RAZHbRtOEiKopM0BnEVoeW8aZ712brchDZegolY4xeCUJ7E6L8291dKPoHPLhDqkxNUjaOQPAD/vjLCb0VwkyyhggZqYJCtJ1qbZl+bNOp9ZTjUVdU4xGVsUh240QM7JQnKFjTzjOxLu0btDZdCRLIuvd0XoLqc8jRGyE6ITUny23HS5X17nre3ZwgPUH+mapqUs1EAsHShTxEWgLlE3sTTpbOyQsSXE9EMLqWtsbu2Y8d1yESHNbFoTVRzd4XFSmSqrdWywQeZlm8suh+mAulYFjlGcT5HtxDJ0UVVT6JG9uRD9ESRK0MPNZW+Eg1xGw1OrDehwZK0lrM+Qyy+s+AhSEcGhj5gCywdh5IWOPIWINiJaGEQltLEsSj0HaGTN2eYacTMJm7TUM8iUKDyHlv1jWd9YOc/v7EHCYK8WsO2myRgfyZYXQR160zYKE2a51jTehIYonEmArj9DwwkMBPO0WKzIVWQO5gfzEpTRdL0oj6iuXe9CI292/eehbNK6XttgHrko4kWJgyA2t9fAtSFBb+7sQFWOmGQkvxGli8iApyyXlWcnRX+vFnyeoWkm53m3Aoog5rpgk6mZFjFiciOWn3dSmCQTJLycraFFrROQJlPkpli0N4JiMWo7Kzek/l7NC29u+pK2tqbIUjQYhFKyEMq2vQgpdVfY87aFCQweE3RvfRFqtlIpaFebkUL7SqI2d9uuS9VquoegSrv87vv7+80gF7rZVgxkqkiJklgxTqcLmmNBqU2rP5FCCqcqqn1/BX0/WfvYlObeGXzbqBgGc2CALNPeBc5XZNRYr+1tQpOPp3WAzWnhh6hXpQoVliLt4oho8afqpGmMiv2O1NOet9Kgo9N6gyKa1jpc3k4SaI9VX1FUNPwXdtkRq4jtffCQxCMsGq2EBQvLdiskFz0IDpRtETm6jZwJ6ActW/ZOoWE1WARVlhqCE6B3EJq0Ovi8ZP9QOH5hb2Fu5WhunC2nEHYwWCBn88FUGVS4Xka0qZAbuk075iyCblhVPZsTpSl0sA8oMOA0Ha0WwciRWX3Jyba3NaIQuJfSwI4X2foVKhBzajBrwB2kH7uNmFWjETclfzjrAvr+ctiSTDNbcDCCyDlBaWecTto0e/L0VFSeXRfucV/4+eUdlWjVONUqMmc4IiJhncpEhXSp7kqQd0WPT4a9s95DD9D1bv0B7zspZ/DFwtRJChohmYyQXlac6pbRL+6ohFZBfB1NMjmSDLwF3NSmchWZc2GBC9uhWIFjrXSmqOZLs2wOScrmYyNZIBuGBwYFqbHRRmLpouI5+MzGXTL7MnPyNouThJkWsxKWxEz1GM2zGiWkYGmcAd6toLV0Dyj3ErQSRY6L1NlGTI+e913p0vfbRm9D5zHa3ZX2rSLNmvFUXYiaQDLUxOGLYtEwKuRcYgqtRZlTHN+bKbSH4rLdpEOWq6uZr3VMI13+/Lntyjn4bV3KkghHQeCpxXkxCAyVtQZq6INOPYuvpg43oEXtImoNhTddFk2jGFxGInOFDQpAzjr+wiZtGJT5DBK3AgZ7goQGU8wrcDFvoSQV7kpjOjoPBN0dp0cLBQSShzqdjNhAFWtzRsjWG9V0E42cFb60NT1PUzP+ik6hMBZzo1gsLudz1aB4m0ooD2Y9HPe9Wq/VGjfwho6hCX2jqgupPQ0oqJZkv73qcN0l6fzlremCV86kdfKchMGAyGURhnFC/b9gnIs5lAaYqdXqNLrdRjqngiwp6JZ2Iy6XRSN2NvQcyL5xmvsy9F+py4fSyj1c1zwNwmjAfI5MAFJTA+bPxp1Tq2P1HqRWp9vpdKzTzngmJYY6bTPWjkN0UzHc3NinS/W97n7puQ94+Xq4RbC2ucntsnBXp3XKxrOH+/tao9GAQ7ZOhf21+/v6bMxOW45M7olyxXXsTVlBZaF08C+hcyD5xtbuYV2ESdmUrOHZ3xqNw8OjI+3duz8c/VlqWYfv3r070o6ODtGWvz1cWpKTpufNyx2qj/xXPLsnIqy7vX8Y9krW7B6wR8DE0cDr0UzyW128+Z8GnUITgN8bWjt7lXXnFIWClXvVw5K0U3wTW6+a0hi4sPUP2XF4hH9jyZf+fPiHP/x0tDr9TjtqNO7HkvMMrjtEA53cK5/TFNjPfe74w27jCLCNDOOP7w7Rx91Wy2+1/nYIvHd/TD84QkPeHdLCNt+0WRsXXvuI6DE9j9A9zeZvzE6tcSTufYSefUfjqh3ed7UhkPEfbSzHiSO07Y9pt7w7atSyTe64+HPIn3my6b/E7vyxJDzMma06+p12CJMeZmOr1anNJKl11/IlaVh76HRmZ90GmrT+4jidfJSGAvn4LQ/GFsYNkvhiO6V/mBpzqNUexpZkOqb/oDXOer0OrJ716hBodaslWZdnNS396lFXSjeD17efPHjd81y5TpeqWDRel3qH1NmHDx1forqLty5mw2F9VnuQ7u66Z/U65ELdmliO5I/PGsIDKQfpjn9PiSX35seBc3RhtwPDnXEDwDPgpoWya5dHOKThvYSh9iXHcUyH02RudST5dYBTrW5LY9H43Fc8iZw+IVD3nap/iB+mrW8QFI7R33sELYtf9GyTMxAxGj+ZMJmy+kzKf9VD0I8FanetY0rDzgsPCoxmAprvbZKWxkPJvBSXFvJf+fz1cc6nRYYz0OOJvv/Qzt/R4VZtf8/9iSmd0oMOD63cb3j0O18YI24bdcvcfZBDdy4b9dYL0LptWtTXMPk3PnWek+oIM/CTswNudceIrc7OIwQgXf8j+rox6+d++wPvBYt6r7YDbvND0Il0v/XgBNKMPyOl/ODn/uV9X/OcZr7Qud8Dh05q9Oq1mr/d1QK4d/ol93rbI6K5wliAzyjE1sV9p342m5T19ZSWaX0UwJ3XAL/66dSfc30B3p1ZZjmLb90Fk6zjvGxadVrSvwfwX3/nP24Ay2nMydszS/WVxAewZJE3ksW5197wDc/kHucKnRX4FsEQ8AOAGw9vAH7j48D5fOFUYGzGuYhjAWwV8sff8K9J5AspzkffWYeTCPy6eEDsv78d9Hdr8NqYaB19PeyKbnidV//W56/zKcmcIZk74t3DVwB/7aPfKbV3O63Lxr/KT7/7Xw45zlEyb5DJ/5qsf/c/WiLUGywv/Bv+XkrOqmn3fu7q3wAN7Acpf/VvgUYu/w3Iv/GPG3z3//Rv4/wvkbnjOtRlCP0AAAAASUVORK5CYII=
kominfo=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""".strip().splitlines():
    _k, _v = _line.split("=", 1)
    _LOGOS[_k.strip()] = _v.strip()


if __name__ == "__main__":
    # Uji cepat dari terminal:  python report.py data_cctv.xlsx [Harian|Mingguan|Bulanan|Keseluruhan]
    import sys
    import parser as dp

    if len(sys.argv) < 2:
        print("Pemakaian: python report.py <file.xlsx|csv> [periode]")
        raise SystemExit(1)
    path = sys.argv[1]
    per = sys.argv[2] if len(sys.argv) > 2 else "Harian"
    with open(path, "rb") as fh:
        bio = io.BytesIO(fh.read())
    bio.name = path
    _df = dp.load_and_clean(bio)
    _dq = dp.data_quality_report(_df)
    out = generate_docx_report(_df, _dq, {"periode_label": per, "sumber_data": path})
    with open("laporan_analisis_cctv.docx", "wb") as fh:
        fh.write(out.getvalue())
    print("Tersimpan: laporan_analisis_cctv.docx")
