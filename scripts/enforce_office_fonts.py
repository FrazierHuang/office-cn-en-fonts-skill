#!/usr/bin/env python3
"""Normalize and check Office OOXML fonts for Chinese-English mixed outputs."""

from __future__ import annotations

import argparse
import os
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass, field
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

SUPPORTED_SUFFIXES = {".docx", ".xlsx", ".pptx"}
CJK_RE = re.compile(r"([\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]+)")
CJK_CLASS = r"\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff"
WESTERN_CLASS = r"A-Za-z0-9"
CJK_WESTERN_SPACE_RE = re.compile(
    rf"(?<=[{CJK_CLASS}])\s+(?=[{WESTERN_CLASS}])|(?<=[{WESTERN_CLASS}])\s+(?=[{CJK_CLASS}])"
)


@dataclass
class CheckResult:
    path: Path
    issues: dict[str, int] = field(default_factory=dict)

    @property
    def ok(self) -> bool:
        return not self.issues

    def add(self, key: str, count: int = 1) -> None:
        if count:
            self.issues[key] = self.issues.get(key, 0) + count


def qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


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


def count_spacing_issues(text: str) -> int:
    return len(CJK_WESTERN_SPACE_RE.findall(text))


def is_cjk_char(char: str) -> bool:
    return bool(re.fullmatch(f"[{CJK_CLASS}]", char))


def is_western_char(char: str) -> bool:
    return bool(re.fullmatch(f"[{WESTERN_CLASS}]", char))


def is_cjk_western_boundary(left: str, right: str) -> bool:
    return (is_cjk_char(left) and is_western_char(right)) or (is_western_char(left) and is_cjk_char(right))


def first_nonspace(text: str) -> str | None:
    stripped = text.lstrip()
    return stripped[0] if stripped else None


def last_nonspace(text: str) -> str | None:
    stripped = text.rstrip()
    return stripped[-1] if stripped else None


def normalize_spacing_across_text_nodes(root: ET.Element, text_tag: str) -> None:
    last_sig_node: ET.Element | None = None
    last_sig_char: str | None = None
    space_nodes: list[ET.Element] = []

    for text_node in root.iter(text_tag):
        if text_node.text is None:
            continue

        text_node.text = normalize_text_spacing(text_node.text)
        current_first = first_nonspace(text_node.text)

        if current_first is None:
            if last_sig_node is not None and text_node.text:
                space_nodes.append(text_node)
            continue

        if last_sig_node is not None and last_sig_char is not None and is_cjk_western_boundary(last_sig_char, current_first):
            last_sig_node.text = (last_sig_node.text or "").rstrip()
            for space_node in space_nodes:
                space_node.text = ""
            text_node.text = text_node.text.lstrip()

        current_last = last_nonspace(text_node.text)
        if current_last is not None:
            last_sig_node = text_node
            last_sig_char = current_last
            space_nodes = []


def count_spacing_issues_across_text_nodes(root: ET.Element, text_tag: str) -> int:
    issues = 0
    last_sig_char: str | None = None
    pending_space = False

    for text_node in root.iter(text_tag):
        text = text_node.text
        if text is None:
            continue

        issues += count_spacing_issues(text)
        current_first = first_nonspace(text)

        if current_first is None:
            if last_sig_char is not None and text:
                pending_space = True
            continue

        if last_sig_char is not None and is_cjk_western_boundary(last_sig_char, current_first):
            if pending_space or text[:1].isspace():
                issues += 1

        current_last = last_nonspace(text)
        if current_last is not None:
            last_sig_char = current_last
            pending_space = bool(text[-1:].isspace())

    return issues


def is_black_color(value: str | None) -> bool:
    if value is None:
        return False
    value = value.upper().replace("#", "")
    return value in {"000000", "FF000000"}


def is_default_fill(fill: ET.Element) -> bool:
    pattern = fill.find(qn(SS_NS, "patternFill"))
    if pattern is None:
        return False
    children = list(pattern)
    return pattern.get("patternType") in {"none", "gray125"} and not children


def is_supported_office_file(path: Path) -> bool:
    return path.suffix.lower() in SUPPORTED_SUFFIXES and ".backup" not in path.stem


def walk_office_files(paths: list[Path], recursive: bool) -> list[Path]:
    files: list[Path] = []
    for path in paths:
        if path.is_dir():
            iterator = path.rglob("*") if recursive else path.iterdir()
            files.extend(p for p in iterator if p.is_file() and is_supported_office_file(p))
        elif path.is_file() and is_supported_office_file(path):
            files.append(path)
        else:
            raise ValueError(f"Unsupported path or file type: {path}")
    return sorted(dict.fromkeys(files))


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

    normalize_spacing_across_text_nodes(root, qn(W_NS, "t"))

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


def check_word_xml(data: bytes, result: CheckResult) -> None:
    root = parse_xml(data)
    if root is None:
        return

    result.add("spacing", count_spacing_issues_across_text_nodes(root, qn(W_NS, "t")))

    for rpr in root.iter(qn(W_NS, "rPr")):
        rfonts = rpr.find(qn(W_NS, "rFonts"))
        if rfonts is None:
            result.add("word_font_missing")
        else:
            if rfonts.get(qn(W_NS, "eastAsia")) not in {"宋体", "SimSun"}:
                result.add("word_chinese_font")
            for attr in ("ascii", "hAnsi", "cs"):
                if rfonts.get(qn(W_NS, attr)) != "Times New Roman":
                    result.add("word_english_font")
                    break
        color = rpr.find(qn(W_NS, "color"))
        if color is None or not is_black_color(color.get(qn(W_NS, "val"))):
            result.add("word_color")

    for tcpr in root.iter(qn(W_NS, "tcPr")):
        result.add("word_table_fill", len(tcpr.findall(qn(W_NS, "shd"))))
    for tblpr in root.iter(qn(W_NS, "tblPr")):
        result.add("word_table_fill", len(tblpr.findall(qn(W_NS, "shd"))))


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


def check_xlsx_styles(data: bytes, result: CheckResult) -> None:
    root = parse_xml(data)
    if root is None:
        return

    fonts = root.find(qn(SS_NS, "fonts"))
    if fonts is not None:
        for font in fonts.findall(qn(SS_NS, "font")):
            name = font.find(qn(SS_NS, "name"))
            if name is None or name.get("val") != "Times New Roman":
                result.add("xlsx_font")
            color = font.find(qn(SS_NS, "color"))
            if color is None or not is_black_color(color.get("rgb")):
                result.add("xlsx_color")

    fills = root.find(qn(SS_NS, "fills"))
    if fills is not None:
        for fill in fills.findall(qn(SS_NS, "fill")):
            if not is_default_fill(fill):
                result.add("xlsx_fill")

    for xf_parent_name in ("cellXfs", "cellStyleXfs"):
        xf_parent = root.find(qn(SS_NS, xf_parent_name))
        if xf_parent is not None:
            for xf in xf_parent.findall(qn(SS_NS, "xf")):
                if xf.get("fillId") not in {None, "0"}:
                    result.add("xlsx_fill")


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


def check_spreadsheet_strings(data: bytes, result: CheckResult) -> None:
    root = parse_xml(data)
    if root is None:
        return

    for text_node in root.iter(qn(SS_NS, "t")):
        if text_node.text:
            result.add("spacing", count_spacing_issues(text_node.text))

    for run in root.iter(qn(SS_NS, "r")):
        text = "".join(t.text or "" for t in run.iter(qn(SS_NS, "t")))
        if not text:
            continue
        rpr = run.find(qn(SS_NS, "rPr"))
        rfont = rpr.find(qn(SS_NS, "rFont")) if rpr is not None else None
        color = rpr.find(qn(SS_NS, "color")) if rpr is not None else None
        expected = "宋体" if CJK_RE.fullmatch(text) else "Times New Roman"
        if rfont is None or rfont.get("val") != expected:
            result.add("xlsx_rich_text_font")
        if color is None or not is_black_color(color.get("rgb")):
            result.add("xlsx_rich_text_color")


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

    normalize_spacing_across_text_nodes(root, qn(A_NS, "t"))

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


def check_pptx_xml(data: bytes, result: CheckResult) -> None:
    root = parse_xml(data)
    if root is None:
        return

    result.add("spacing", count_spacing_issues_across_text_nodes(root, qn(A_NS, "t")))

    for rpr in list(root.iter(qn(A_NS, "rPr"))) + list(root.iter(qn(A_NS, "defRPr"))):
        latin = rpr.find(qn(A_NS, "latin"))
        ea = rpr.find(qn(A_NS, "ea"))
        cs = rpr.find(qn(A_NS, "cs"))
        if latin is None or latin.get("typeface") != "Times New Roman":
            result.add("pptx_english_font")
        if ea is None or ea.get("typeface") not in {"宋体", "SimSun"}:
            result.add("pptx_chinese_font")
        if cs is None or cs.get("typeface") != "Times New Roman":
            result.add("pptx_complex_font")


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


def check_file(path: Path) -> CheckResult:
    result = CheckResult(path)
    with zipfile.ZipFile(path, "r") as archive:
        for name in archive.namelist():
            data = archive.read(name)
            suffix = path.suffix.lower()
            if suffix == ".docx" and name.startswith("word/") and name.endswith(".xml"):
                check_word_xml(data, result)
            elif suffix == ".xlsx" and name == "xl/styles.xml":
                check_xlsx_styles(data, result)
            elif suffix == ".xlsx" and (
                name == "xl/sharedStrings.xml" or name.startswith("xl/worksheets/") and name.endswith(".xml")
            ):
                check_spreadsheet_strings(data, result)
            elif suffix == ".pptx" and name.endswith(".xml") and name.startswith(
                ("ppt/slides/", "ppt/slideLayouts/", "ppt/slideMasters/", "ppt/theme/", "ppt/notesSlides/")
            ):
                check_pptx_xml(data, result)
    return result


def output_path(path: Path, in_place: bool) -> Path:
    if in_place:
        fd, tmp_name = tempfile.mkstemp(suffix=path.suffix)
        os.close(fd)
        Path(tmp_name).unlink(missing_ok=True)
        return Path(tmp_name)
    return path.with_name(f"{path.stem}_fontfixed{path.suffix}")


def backup_path(path: Path) -> Path:
    candidate = path.with_name(f"{path.stem}.backup{path.suffix}")
    if not candidate.exists():
        return candidate
    index = 2
    while True:
        candidate = path.with_name(f"{path.stem}.backup{index}{path.suffix}")
        if not candidate.exists():
            return candidate
        index += 1


def normalize_file(path: Path, in_place: bool, backup: bool) -> Path:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        transform = transform_docx
    elif suffix == ".xlsx":
        transform = transform_xlsx
    elif suffix == ".pptx":
        transform = transform_pptx
    else:
        raise ValueError(f"Unsupported file type: {path}")

    if backup and in_place:
        shutil.copy2(path, backup_path(path))

    dst = output_path(path, in_place)
    write_zip(path, dst, transform)
    if in_place:
        shutil.move(str(dst), str(path))
        return path
    return dst


def print_check_result(result: CheckResult) -> None:
    if result.ok:
        print(f"PASS {result.path}")
        return
    print(f"FAIL {result.path}")
    for key, count in sorted(result.issues.items()):
        print(f"  - {key}: {count}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Enforce or check Chinese/English fonts in Office OOXML files.")
    parser.add_argument("paths", nargs="+", type=Path, help="Office files or folders containing .docx, .xlsx, .pptx")
    parser.add_argument("--in-place", action="store_true", help="Overwrite input files after successful rewrite")
    parser.add_argument("--backup", action="store_true", help="Create .backup files before in-place rewrites")
    parser.add_argument("--check", action="store_true", help="Check files and return non-zero when issues are found")
    parser.add_argument("--recursive", action="store_true", help="Process Office files inside folders recursively")
    args = parser.parse_args()

    if args.backup and not args.in_place:
        parser.error("--backup requires --in-place")

    files = walk_office_files(args.paths, args.recursive)
    if args.check:
        results = [check_file(file_path) for file_path in files]
        for result in results:
            print_check_result(result)
        return 0 if all(result.ok for result in results) else 1

    for file_path in files:
        result = normalize_file(file_path, args.in_place, args.backup)
        print(f"OK   {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
