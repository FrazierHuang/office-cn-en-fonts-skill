#!/usr/bin/env python3
"""Normalize Office OOXML fonts for Chinese/English mixed outputs."""

from __future__ import annotations

import argparse
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
SS_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ET.register_namespace("w", W_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("", SS_NS)
ET.register_namespace("r", R_NS)

CJK_RE = re.compile(r"([\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]+)")
CJK_CLASS = r"\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff"
WESTERN_CLASS = r"A-Za-z0-9"
CJK_WESTERN_SPACE_RE = re.compile(
    rf"(?<=[{CJK_CLASS}])\s+(?=[{WESTERN_CLASS}])|(?<=[{WESTERN_CLASS}])\s+(?=[{CJK_CLASS}])"
)


def qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def write_zip(src: Path, dst: Path, transform) -> None:
    with zipfile.ZipFile(src, "r") as zin:
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                new_data = transform(item.filename, data)
                info = zipfile.ZipInfo(item.filename, item.date_time)
                info.compress_type = zipfile.ZIP_DEFLATED
                info.external_attr = item.external_attr
                zout.writestr(info, new_data)


def parse_xml(data: bytes):
    try:
        return ET.fromstring(data)
    except ET.ParseError:
        return None


def xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def ensure_child(parent: ET.Element, tag: str, first: bool = False) -> ET.Element:
    child = parent.find(tag)
    if child is None:
        child = ET.Element(tag)
        if first:
            parent.insert(0, child)
        else:
            parent.append(child)
    return child


def normalize_text_spacing(text: str) -> str:
    return CJK_WESTERN_SPACE_RE.sub("", text)


def normalize_word(data: bytes) -> bytes:
    root = parse_xml(data)
    if root is None:
        return data

    for run in root.iter(qn(W_NS, "r")):
        rpr = run.find(qn(W_NS, "rPr"))
        if rpr is None:
            rpr = ET.Element(qn(W_NS, "rPr"))
            run.insert(0, rpr)
        set_word_run_props(rpr)

    for text_node in root.iter(qn(W_NS, "t")):
        if text_node.text:
            text_node.text = normalize_text_spacing(text_node.text)

    for rpr in root.iter(qn(W_NS, "rPr")):
        set_word_run_props(rpr)

    for tcpr in root.iter(qn(W_NS, "tcPr")):
        for shd in list(tcpr.findall(qn(W_NS, "shd"))):
            tcpr.remove(shd)

    for tblpr in root.iter(qn(W_NS, "tblPr")):
        for shd in list(tblpr.findall(qn(W_NS, "shd"))):
            tblpr.remove(shd)

    return xml_bytes(root)


def set_word_run_props(rpr: ET.Element) -> None:
    rfonts = ensure_child(rpr, qn(W_NS, "rFonts"), first=True)
    rfonts.set(qn(W_NS, "eastAsia"), "宋体")
    rfonts.set(qn(W_NS, "ascii"), "Times New Roman")
    rfonts.set(qn(W_NS, "hAnsi"), "Times New Roman")
    rfonts.set(qn(W_NS, "cs"), "Times New Roman")

    color = ensure_child(rpr, qn(W_NS, "color"))
    color.set(qn(W_NS, "val"), "000000")
    for attr in (qn(W_NS, "themeColor"), qn(W_NS, "themeShade"), qn(W_NS, "themeTint")):
        color.attrib.pop(attr, None)


def normalize_xlsx_styles(data: bytes) -> bytes:
    root = parse_xml(data)
    if root is None:
        return data

    fonts = root.find(qn(SS_NS, "fonts"))
    if fonts is not None:
        for font in fonts.findall(qn(SS_NS, "font")):
            name = ensure_child(font, qn(SS_NS, "name"), first=True)
            name.set("val", "Times New Roman")
            color = ensure_child(font, qn(SS_NS, "color"))
            color.attrib.clear()
            color.set("rgb", "FF000000")

    fills = root.find(qn(SS_NS, "fills"))
    if fills is not None:
        fills.clear()
        fills.set("count", "2")
        for pattern in ("none", "gray125"):
            fill = ET.SubElement(fills, qn(SS_NS, "fill"))
            ET.SubElement(fill, qn(SS_NS, "patternFill"), {"patternType": pattern})

    for xf_parent_name in ("cellXfs", "cellStyleXfs"):
        xf_parent = root.find(qn(SS_NS, xf_parent_name))
        if xf_parent is not None:
            for xf in xf_parent.findall(qn(SS_NS, "xf")):
                xf.set("fillId", "0")
                xf.attrib.pop("applyFill", None)

    return xml_bytes(root)


def rich_text_runs(text: str) -> list[ET.Element]:
    text = normalize_text_spacing(text)
    parts = [part for part in CJK_RE.split(text) if part]
    runs = []
    for part in parts:
        is_cjk = bool(CJK_RE.fullmatch(part))
        r = ET.Element(qn(SS_NS, "r"))
        rpr = ET.SubElement(r, qn(SS_NS, "rPr"))
        ET.SubElement(rpr, qn(SS_NS, "rFont"), {"val": "宋体" if is_cjk else "Times New Roman"})
        ET.SubElement(rpr, qn(SS_NS, "color"), {"rgb": "FF000000"})
        t = ET.SubElement(r, qn(SS_NS, "t"))
        if part[:1].isspace() or part[-1:].isspace():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = part
        runs.append(r)
    return runs


def normalize_shared_strings(data: bytes) -> bytes:
    root = parse_xml(data)
    if root is None:
        return data

    for si in root.findall(qn(SS_NS, "si")):
        text = "".join(t.text or "" for t in si.iter(qn(SS_NS, "t")))
        if not text:
            continue
        si.clear()
        for run in rich_text_runs(text):
            si.append(run)

    return xml_bytes(root)


def normalize_sheet(data: bytes) -> bytes:
    root = parse_xml(data)
    if root is None:
        return data

    changed = False
    for cell in root.iter(qn(SS_NS, "c")):
        is_elem = cell.find(qn(SS_NS, "is"))
        if is_elem is None:
            continue
        text = "".join(t.text or "" for t in is_elem.iter(qn(SS_NS, "t")))
        if not text:
            continue
        is_elem.clear()
        for run in rich_text_runs(text):
            is_elem.append(run)
        changed = True

    return xml_bytes(root) if changed else data


def normalize_pptx_xml(data: bytes) -> bytes:
    root = parse_xml(data)
    if root is None:
        return data

    for rpr in root.iter(qn(A_NS, "rPr")):
        set_drawing_fonts(rpr)

    for defrpr in root.iter(qn(A_NS, "defRPr")):
        set_drawing_fonts(defrpr)

    for text_node in root.iter(qn(A_NS, "t")):
        if text_node.text:
            text_node.text = normalize_text_spacing(text_node.text)

    for tag in ("latin", "ea", "cs"):
        for elem in root.iter(qn(A_NS, tag)):
            elem.set("typeface", "宋体" if tag == "ea" else "Times New Roman")

    return xml_bytes(root)


def set_drawing_fonts(rpr: ET.Element) -> None:
    latin = ensure_child(rpr, qn(A_NS, "latin"))
    latin.set("typeface", "Times New Roman")
    ea = ensure_child(rpr, qn(A_NS, "ea"))
    ea.set("typeface", "宋体")
    cs = ensure_child(rpr, qn(A_NS, "cs"))
    cs.set("typeface", "Times New Roman")


def transform_docx(name: str, data: bytes) -> bytes:
    if name.startswith("word/") and name.endswith(".xml"):
        return normalize_word(data)
    return data


def transform_xlsx(name: str, data: bytes) -> bytes:
    if name == "xl/styles.xml":
        return normalize_xlsx_styles(data)
    if name == "xl/sharedStrings.xml":
        return normalize_shared_strings(data)
    if name.startswith("xl/worksheets/") and name.endswith(".xml"):
        return normalize_sheet(data)
    return data


def transform_pptx(name: str, data: bytes) -> bytes:
    prefixes = ("ppt/slides/", "ppt/slideLayouts/", "ppt/slideMasters/", "ppt/theme/", "ppt/notesSlides/")
    if name.endswith(".xml") and name.startswith(prefixes):
        return normalize_pptx_xml(data)
    return data


def output_path(path: Path, in_place: bool) -> Path:
    if in_place:
        fd, tmp_name = tempfile.mkstemp(suffix=path.suffix)
        os.close(fd)
        Path(tmp_name).unlink(missing_ok=True)
        return Path(tmp_name)
    return path.with_name(f"{path.stem}_fontfixed{path.suffix}")


def normalize_file(path: Path, in_place: bool) -> Path:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        transform = transform_docx
    elif suffix == ".xlsx":
        transform = transform_xlsx
    elif suffix == ".pptx":
        transform = transform_pptx
    else:
        raise ValueError(f"Unsupported file type: {path}")

    dst = output_path(path, in_place)
    write_zip(path, dst, transform)
    if in_place:
        shutil.move(str(dst), str(path))
        return path
    return dst


def main() -> int:
    parser = argparse.ArgumentParser(description="Enforce Chinese/English fonts in Office OOXML files.")
    parser.add_argument("files", nargs="+", type=Path, help="Office files: .docx, .xlsx, .pptx")
    parser.add_argument("--in-place", action="store_true", help="Overwrite input files after successful rewrite")
    args = parser.parse_args()

    for file_path in args.files:
        result = normalize_file(file_path, args.in_place)
        print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
