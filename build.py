#!/usr/bin/env python3
"""Сборка курсовой: text/*.md -> build/*.docx по требованиям приказа РТА № 550.

Запуск: python3 build.py [имя_файла_без_расширения]
Нужны pandoc и lxml.
"""
import re
import shutil
import subprocess
import sys
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent
TEXT = ROOT / "text"
BUILD = ROOT / "build"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W}
FONT = "Times New Roman"


def q(tag):
    return f"{{{W}}}{tag}"


def el(tag, **attrs):
    e = etree.Element(q(tag))
    for k, v in attrs.items():
        e.set(q(k), str(v))
    return e


def child(parent, tag, first=False):
    """Вернуть дочерний элемент, создав его при отсутствии."""
    c = parent.find(q(tag))
    if c is None:
        c = el(tag)
        if first:
            parent.insert(0, c)
        else:
            parent.append(c)
    return c


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
    "Figure": (28, False, "center", 0, 360, 0, False),
    "TableCaption": (28, False, "center", 0, 360, 0, False),
    "Bibliography": (28, False, "both", 709, 360, 0, False),
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
        stype = style.get(q("type"))
        rpr = style.find(q("rPr"))
        if rpr is not None:
            for tag in ("rFonts", "color", "i", "iCs", "caps", "smallCaps"):
                for old in rpr.findall(q(tag)):
                    rpr.remove(old)
        if stype == "paragraph" and sid in STYLES:
            size, bold, jc, ind, line, after, brk = STYLES[sid]
            set_ppr(child(style, "pPr"), jc=jc, first_line=ind, line=line,
                    before=360 if sid.startswith("Heading") and not brk else 0,
                    after=after, keep_next=sid.startswith("Heading"), page_break=brk)
            set_rpr(child(style, "rPr"), size=size, bold=bold)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def para_text(p):
    return "".join(t.text or "" for t in p.iter(q("t")))


def fix_document(xml):
    root = etree.fromstring(xml)
    body = root.find(q("body"))
    # неразрывные пробелы в числах и перед знаком процента и номера
    for part in (root,):
        for t in part.iter(q("t")):
            if t.text:
                s = re.sub(r"(?<=\d) (?=\d{3}(?!\d))", " ", t.text)
                s = re.sub(r"(?<=\d) (?=%)", " ", s)
                s = re.sub(r"№ (?=\d)", "№ ", s)
                t.text = s
    paras = [p for p in body.iter(q("p"))]
    for i, p in enumerate(paras):
        text = para_text(p).strip()
        if re.fullmatch(r"(Продолжение таблицы|Таблица) \d+", text):
            set_ppr(child(p, "pPr", first=True), jc="right", first_line=0, line=360, keep_next=True)
            if i + 1 < len(paras):
                set_ppr(child(paras[i + 1], "pPr", first=True), jc="center",
                        first_line=0, line=360, after=120, keep_next=True)
    for tbl in body.iter(q("tbl")):
        tblpr = child(tbl, "tblPr", first=True)
        for old in tblpr.findall(q("tblBorders")):
            tblpr.remove(old)
        borders = el("tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            borders.append(el(side, val="single", sz=4, space=0, color="000000"))
        tblpr.append(borders)
        for old in tblpr.findall(q("tblW")):
            tblpr.remove(old)
        tblpr.append(el("tblW", w=5000, type="pct"))
        for row_i, row in enumerate(tbl.findall(q("tr"))):
            for p in row.iter(q("p")):
                set_ppr(child(p, "pPr", first=True), jc="center" if row_i == 0 else "left",
                        first_line=0, line=240)
                for r in p.iter(q("r")):
                    set_rpr(child(r, "rPr", first=True), size=24)
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
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


FOOTER = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="{W}"><w:p><w:pPr><w:jc w:val="center"/><w:ind w:firstLine="0"/></w:pPr>
<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>"""


def fix_rels(xml):
    root = etree.fromstring(xml)
    ns = root.nsmap[None]
    rel = etree.SubElement(root, f"{{{ns}}}Relationship")
    rel.set("Id", "rIdFooterKR")
    rel.set("Type", f"{R}/footer")
    rel.set("Target", "footerKR.xml")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def fix_types(xml):
    root = etree.fromstring(xml)
    ns = root.nsmap[None]
    ov = etree.SubElement(root, f"{{{ns}}}Override")
    ov.set("PartName", "/word/footerKR.xml")
    ov.set("ContentType",
           "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def fix_settings(xml):
    root = etree.fromstring(xml)
    if root.find(q("autoHyphenation")) is None:
        root.insert(0, el("autoHyphenation", val="true"))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def main():
    name = sys.argv[1] if len(sys.argv) > 1 else "КР_Жмура_черновик"
    BUILD.mkdir(exist_ok=True)
    files = sorted(p.name for p in TEXT.glob("*.md"))
    if not files:
        sys.exit("в text/ нет файлов .md")
    raw = BUILD / "_raw.docx"
    subprocess.run(["pandoc", *files, "-f", "markdown+smart", "-t", "docx",
                    "--resource-path=.", "-o", str(raw)], cwd=TEXT, check=True)
    out = BUILD / f"{name}.docx"
    fixers = {
        "word/styles.xml": fix_styles,
        "word/document.xml": fix_document,
        "word/_rels/document.xml.rels": fix_rels,
        "[Content_Types].xml": fix_types,
        "word/settings.xml": fix_settings,
    }
    with zipfile.ZipFile(raw) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in fixers:
                data = fixers[item.filename](data)
            zout.writestr(item, data)
        zout.writestr("word/footerKR.xml", FOOTER)
    raw.unlink()
    print(out)


if __name__ == "__main__":
    main()
