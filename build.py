#!/usr/bin/env python3
"""Сборка курсовой: text/*.md -> build/*.docx по требованиям приказа РТА № 550.

Запуск: .venv/bin/python build.py [имя_файла_без_расширения]
Нужны pandoc, lxml и python-pptx (из него берётся генератор нативных диаграмм).

Метки в тексте (отдельным абзацем):
  CHART:имя   -> диаграмма Excel из assets/chart_имя.csv (первый столбец — категории,
                 остальные — ряды; данные вшиваются в docx и правятся в Word)
  SCHEME:имя  -> блок-схема из assets/schemes.json (блоки-таблицы со стрелками)
"""
import csv
import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent
TEXT = ROOT / "text"
ASSETS = ROOT / "assets"
BUILD = ROOT / "build"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS = {"w": W}
FONT = "Times New Roman"
CHARTS = []  # (номер, xml диаграммы, xlsx с данными)


def q(tag):
    return f"{{{W}}}{tag}"


def el(tag, **attrs):
    e = etree.Element(q(tag))
    for k, v in attrs.items():
        e.set(q(k), str(v))
    return e


def child(parent, tag, first=False):
    c = parent.find(q(tag))
    if c is None:
        c = el(tag)
        if first:
            parent.insert(0, c)
        else:
            parent.append(c)
    return c


def dump(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def set_ppr(ppr, jc=None, first_line=None, line=None, before=0, after=0,
            keep_next=False, page_break=False):
    for tag in ("jc", "ind", "spacing", "keepNext", "pageBreakBefore"):
        for old in ppr.findall(q(tag)):
            ppr.remove(old)
    if keep_next:
        ppr.append(el("keepNext"))
    if page_break:
        ppr.append(el("pageBreakBefore"))
    sp = el("spacing", before=before, after=after)
    if line:
        sp.set(q("line"), str(line))
        sp.set(q("lineRule"), "auto")
    ppr.append(sp)
    if first_line is not None:
        ppr.append(el("ind", left=0, right=0, firstLine=first_line))
    if jc:
        ppr.append(el("jc", val=jc))


def set_rpr(rpr, size=28, bold=False):
    for tag in ("rFonts", "sz", "szCs", "color", "i", "iCs", "b", "bCs",
                "caps", "smallCaps", "u"):
        for old in rpr.findall(q(tag)):
            rpr.remove(old)
    rpr.append(el("rFonts", ascii=FONT, hAnsi=FONT, cs=FONT, eastAsia=FONT))
    if bold:
        rpr.append(el("b"))
        rpr.append(el("bCs"))
    rpr.append(el("color", val="000000"))
    rpr.append(el("sz", val=size))
    rpr.append(el("szCs", val=size))


def make_par(text, size=28, bold=False, jc="center", line=240, after=0, keep_next=False):
    p = el("p")
    set_ppr(child(p, "pPr"), jc=jc, first_line=0, line=line, after=after, keep_next=keep_next)
    r = etree.SubElement(p, q("r"))
    set_rpr(child(r, "rPr"), size=size, bold=bold)
    t = etree.SubElement(r, q("t"))
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return p


def norm_text(root):
    """Короткое тире вместо длинного; неразрывные пробелы в числах, перед % и после №."""
    for t in root.iter(q("t")):
        if t.text:
            s = t.text.replace("—", "–")
            s = re.sub(r"(?<=\d) (?=\d{3}(?!\d))", " ", s)
            s = re.sub(r"(?<=\d) (?=%)", " ", s)
            s = re.sub(r"№ (?=\d)", "№ ", s)
            t.text = s


# стиль -> (кегль*2, полужирный, выравнивание, отступ первой строки, интервал, после, разрыв страницы)
STYLES = {
    "Normal": (28, False, "both", 709, 360, 0, False),
    "BodyText": (28, False, "both", 709, 360, 0, False),
    "FirstParagraph": (28, False, "both", 709, 360, 0, False),
    "Compact": (24, False, "left", 0, 240, 0, False),
    "Heading1": (28, True, "center", 0, 240, 360, True),
    "Heading2": (28, True, "center", 0, 240, 360, False),
    "Heading3": (28, True, "center", 0, 240, 360, False),
    "FootnoteText": (20, False, "both", 709, 240, 0, False),
    "ImageCaption": (28, False, "center", 0, 360, 240, False),
    "CaptionedFigure": (28, False, "center", 0, 360, 0, False),
    "TableCaption": (28, False, "center", 0, 360, 0, False),
}


def fix_styles(xml):
    root = etree.fromstring(xml)
    defaults = root.find(".//w:docDefaults/w:rPrDefault/w:rPr", NS)
    if defaults is not None:
        set_rpr(defaults)
        for lang in defaults.findall(q("lang")):
            defaults.remove(lang)
        defaults.append(el("lang", val="ru-RU", eastAsia="ru-RU", bidi="ar-SA"))
    for style in root.findall("w:style", NS):
        sid = style.get(q("styleId"))
        rpr = style.find(q("rPr"))
        if rpr is not None:
            for tag in ("rFonts", "color", "i", "iCs", "caps", "smallCaps"):
                for old in rpr.findall(q(tag)):
                    rpr.remove(old)
        if style.get(q("type")) == "paragraph" and sid in STYLES:
            size, bold, jc, ind, line, after, brk = STYLES[sid]
            heading = sid.startswith("Heading")
            set_ppr(child(style, "pPr"), jc=jc, first_line=ind, line=line,
                    before=360 if heading and not brk else 0,
                    after=after, keep_next=heading, page_break=brk)
            set_rpr(child(style, "rPr"), size=size, bold=bold)
    have = {s.get(q("styleId")) for s in root.findall("w:style", NS)}
    for level, left in ((1, 0), (2, 0), (3, 284)):
        sid = f"TOC{level}"
        if sid in have:
            continue
        st = el("style", type="paragraph", styleId=sid)
        st.append(el("name", val=f"toc {level}"))
        st.append(el("basedOn", val="Normal"))
        st.append(el("uiPriority", val=39))
        ppr = child(st, "pPr")
        set_ppr(ppr, jc="left", first_line=0, line=360)
        ppr.find(q("ind")).set(q("left"), str(left))
        root.append(st)
    return dump(root)


def para_text(p):
    return "".join(t.text or "" for t in p.iter(q("t")))


def build_chart(name):
    """Нативная диаграмма Word/Excel: xml диаграммы и вшитая книга с данными."""
    from pptx.chart.data import CategoryChartData
    from pptx.chart.xlsx import CategoryWorkbookWriter
    from pptx.chart.xmlwriter import ChartXmlWriter
    from pptx.enum.chart import XL_CHART_TYPE

    rows = list(csv.reader(open(ASSETS / f"chart_{name}.csv", encoding="utf8"), delimiter=";"))
    head, data = rows[0], rows[1:]
    cd = CategoryChartData(number_format="General")
    cd.categories = [r[0] for r in data]
    for i, series in enumerate(head[1:], start=1):
        cd.add_series(series, [float(r[i].replace(",", ".")) for r in data])
    root = etree.fromstring(ChartXmlWriter(XL_CHART_TYPE.COLUMN_CLUSTERED, cd).xml.encode("utf8"))
    cns = {"c": C, "a": A}
    chart = root.find("c:chart", cns)
    legend = etree.fromstring(
        f'<c:legend xmlns:c="{C}"><c:legendPos val="b"/><c:overlay val="0"/></c:legend>')
    chart.find("c:plotArea", cns).addnext(legend)
    atd = chart.find("c:autoTitleDeleted", cns)
    if atd is not None:
        atd.set("val", "1")
    for d in root.iter(f"{{{A}}}defRPr"):
        d.set("sz", "1100")
        for old in d.findall(f"{{{A}}}latin"):
            d.remove(old)
        etree.SubElement(d, f"{{{A}}}latin").set("typeface", FONT)
    ext = etree.SubElement(root, f"{{{C}}}externalData")
    ext.set(f"{{{R}}}id", "rId1")
    etree.SubElement(ext, f"{{{C}}}autoUpdate").set("val", "0")
    n = len(CHARTS) + 1
    CHARTS.append((n, dump(root), CategoryWorkbookWriter(cd).xlsx_blob))
    return n


def chart_run(n, width_cm=15.5, height_cm=8.0):
    cx, cy = int(width_cm * 360000), int(height_cm * 360000)
    xml = f'''<w:r xmlns:w="{W}" xmlns:wp="{WP}" xmlns:a="{A}" xmlns:c="{C}" xmlns:r="{R}">
<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">
<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:docPr id="{9000 + n}" name="Диаграмма {n}"/><wp:cNvGraphicFramePr/>
<a:graphic><a:graphicData uri="{C}"><c:chart r:id="rIdChartKR{n}"/></a:graphicData></a:graphic>
</wp:inline></w:drawing></w:r>'''
    return etree.fromstring(xml)


def toc_par():
    """Поле оглавления; Word заполняет его при открытии файла (или по F9)."""
    xml = f'''<w:p xmlns:w="{W}"><w:pPr><w:ind w:firstLine="0"/></w:pPr>
<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>Чтобы оглавление появилось, выделите эту строку и нажмите F9.</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'''
    return etree.fromstring(xml)


def scheme_elements(blocks):
    out = []
    for i, block in enumerate(blocks):
        tbl = el("tbl")
        tblpr = child(tbl, "tblPr")
        tblpr.append(el("tblW", w=4300, type="pct"))
        tblpr.append(el("jc", val="center"))
        borders = el("tblBorders")
        for side in ("top", "left", "bottom", "right"):
            borders.append(el(side, val="single", sz=8, space=0, color="000000"))
        tblpr.append(borders)
        tc = etree.SubElement(etree.SubElement(tbl, q("tr")), q("tc"))
        child(tc, "tcPr").append(el("tcW", w=4300, type="pct"))
        if block.get("title"):
            tc.append(make_par(block["title"], size=24, bold=True, keep_next=True))
        for line in block.get("lines", []):
            tc.append(make_par(line, size=24, keep_next=True))
        out.append(tbl)
        if i < len(blocks) - 1:
            out.append(make_par("↓", size=28, keep_next=True))
    return out


def fix_document(xml):
    root = etree.fromstring(xml)
    body = root.find(q("body"))
    norm_text(root)
    schemes = json.loads((ASSETS / "schemes.json").read_text(encoding="utf8"))
    paras = list(body.iter(q("p")))
    for i, p in enumerate(paras):
        text = para_text(p).strip()
        if re.fullmatch(r"(Продолжение таблицы|Таблица) \d+", text):
            set_ppr(child(p, "pPr", first=True), jc="right", first_line=0, line=360, keep_next=True)
            if i + 1 < len(paras):
                set_ppr(child(paras[i + 1], "pPr", first=True), jc="center",
                        first_line=0, line=360, after=120, keep_next=True)
        elif re.fullmatch(r"(Продолжение приложения|Приложение) \d+", text):
            set_ppr(child(p, "pPr", first=True), jc="right", first_line=0, line=360,
                    keep_next=True, page_break=True)
            if i + 1 < len(paras):
                set_ppr(child(paras[i + 1], "pPr", first=True), jc="center",
                        first_line=0, line=360, after=120, keep_next=True)
                for r in paras[i + 1].iter(q("r")):
                    set_rpr(child(r, "rPr", first=True), bold=True)
        elif re.match(r"Рис\. \d+\. ", text):
            set_ppr(child(p, "pPr", first=True), jc="center", first_line=0, line=360, after=240)
    for tbl in body.iter(q("tbl")):
        tblpr = child(tbl, "tblPr", first=True)
        for tag in ("tblBorders", "tblW"):
            for old in tblpr.findall(q(tag)):
                tblpr.remove(old)
        borders = el("tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            borders.append(el(side, val="single", sz=4, space=0, color="000000"))
        tblpr.append(borders)
        tblpr.append(el("tblW", w=5000, type="pct"))
        for row_i, row in enumerate(tbl.findall(q("tr"))):
            for p in row.iter(q("p")):
                set_ppr(child(p, "pPr", first=True), jc="center" if row_i == 0 else "left",
                        first_line=0, line=240)
                for r in p.iter(q("r")):
                    set_rpr(child(r, "rPr", first=True), size=24)
    for p in list(body.iter(q("p"))):
        text = para_text(p).strip()
        if text.startswith("CHART:"):
            n = build_chart(text.split(":", 1)[1].strip())
            for r in p.findall(q("r")):
                p.remove(r)
            set_ppr(child(p, "pPr", first=True), jc="center", first_line=0, line=240,
                    before=120, keep_next=True)
            p.append(chart_run(n))
        elif text.startswith("::"):
            kind, _, rest = text.partition(" ")
            for r in p.findall(q("r")):
                p.remove(r)
            ppr = child(p, "pPr", first=True)
            size, bold, jc = 28, False, "center"
            if kind == "::t18":
                size, bold = 36, True
            elif kind == "::r":
                jc = "left"
            elif kind == "::toc":
                bold = True
            set_ppr(ppr, jc=jc, first_line=0, line=276, after=240 if kind == "::toc" else 0,
                    page_break=(kind == "::toc"))
            if kind == "::r":
                ppr.find(q("ind")).set(q("left"), "4536")
            if kind != "::gap":
                r = etree.SubElement(p, q("r"))
                set_rpr(child(r, "rPr"), size=size, bold=bold)
                etree.SubElement(r, q("t")).text = rest.strip()
            if kind == "::toc":
                p.addnext(toc_par())
        elif text.startswith("SCHEME:"):
            for new in scheme_elements(schemes[text.split(":", 1)[1].strip()]):
                p.addprevious(new)
            p.getparent().remove(p)
    sect = body.find(q("sectPr"))
    if sect is None:
        sect = el("sectPr")
        body.append(sect)
    for tag in ("pgSz", "pgMar", "footnotePr", "footerReference"):
        for old in sect.findall(q(tag)):
            sect.remove(old)
    footer = el("footerReference", type="default")
    footer.set(f"{{{R}}}id", "rIdFooterKR")
    sect.insert(0, footer)
    fpr = el("footnotePr")
    fpr.append(el("numRestart", val="eachPage"))
    sect.append(fpr)
    sect.append(el("pgSz", w=11906, h=16838))
    sect.append(el("pgMar", top=1134, right=851, bottom=1134, left=1418,
                   header=709, footer=709, gutter=0))
    return dump(root)


def fix_footnotes(xml):
    root = etree.fromstring(xml)
    norm_text(root)
    return dump(root)


FOOTER = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="{W}"><w:p><w:pPr><w:jc w:val="center"/><w:ind w:firstLine="0"/></w:pPr>
<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>"""


def fix_rels(xml):
    root = etree.fromstring(xml)
    ns = root.nsmap[None]

    def add(rid, rtype, target):
        rel = etree.SubElement(root, f"{{{ns}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", f"{R}/{rtype}")
        rel.set("Target", target)

    add("rIdFooterKR", "footer", "footerKR.xml")
    for n, _, _ in CHARTS:
        add(f"rIdChartKR{n}", "chart", f"charts/chartKR{n}.xml")
    return dump(root)


def fix_types(xml):
    root = etree.fromstring(xml)
    ns = root.nsmap[None]

    def override(part, ctype):
        ov = etree.SubElement(root, f"{{{ns}}}Override")
        ov.set("PartName", part)
        ov.set("ContentType", ctype)

    override("/word/footerKR.xml",
             "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
    for n, _, _ in CHARTS:
        override(f"/word/charts/chartKR{n}.xml",
                 "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
    if CHARTS and not any(d.get("Extension") == "xlsx" for d in root):
        d = etree.Element(f"{{{ns}}}Default")
        d.set("Extension", "xlsx")
        d.set("ContentType",
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        root.insert(0, d)
    return dump(root)


def fix_settings(xml):
    root = etree.fromstring(xml)
    if root.find(q("autoHyphenation")) is None:
        root.insert(0, el("autoHyphenation", val="true"))
    return dump(root)


def chart_rels(n):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{R}/package" '
            f'Target="../embeddings/Microsoft_Excel_SheetKR{n}.xlsx"/></Relationships>')


def main():
    name = sys.argv[1] if len(sys.argv) > 1 else "КР_Жмура_черновик"
    BUILD.mkdir(exist_ok=True)
    files = sorted(p.name for p in TEXT.glob("*.md"))
    if not files:
        sys.exit("в text/ нет файлов .md")
    raw = BUILD / "_raw.docx"
    subprocess.run(["pandoc", *files, "-f", "markdown+smart", "-t", "docx",
                    "--resource-path=.", "-o", str(raw)], cwd=TEXT, check=True)
    with zipfile.ZipFile(raw) as zin:
        parts = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    raw.unlink()
    # документ первым: при его обработке собираются диаграммы
    parts["word/document.xml"] = fix_document(parts["word/document.xml"])
    parts["word/styles.xml"] = fix_styles(parts["word/styles.xml"])
    parts["word/_rels/document.xml.rels"] = fix_rels(parts["word/_rels/document.xml.rels"])
    parts["[Content_Types].xml"] = fix_types(parts["[Content_Types].xml"])
    parts["word/settings.xml"] = fix_settings(parts["word/settings.xml"])
    if "word/footnotes.xml" in parts:
        parts["word/footnotes.xml"] = fix_footnotes(parts["word/footnotes.xml"])
    parts["word/footerKR.xml"] = FOOTER.encode("utf8")
    for n, xml, blob in CHARTS:
        parts[f"word/charts/chartKR{n}.xml"] = xml
        parts[f"word/charts/_rels/chartKR{n}.xml.rels"] = chart_rels(n).encode("utf8")
        parts[f"word/embeddings/Microsoft_Excel_SheetKR{n}.xlsx"] = blob
    out = BUILD / f"{name}.docx"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("[Content_Types].xml", parts.pop("[Content_Types].xml"))
        for fname, data in parts.items():
            zout.writestr(fname, data)
    print(out)


if __name__ == "__main__":
    main()
