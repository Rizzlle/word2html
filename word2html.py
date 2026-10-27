"""Convert a text-focused DOCX document to a complete, readable HTML page.

Uses only the Python standard library. Word list labels are reconstructed from
the document's numbering definitions so legal clause references remain visible.
"""

from __future__ import annotations

import argparse
from collections import defaultdict
from dataclasses import dataclass
from html import escape
from pathlib import Path
import re
import sys
from urllib.parse import urlparse
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_R = "{http://schemas.openxmlformats.org/package/2006/relationships}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
URL_RE = re.compile(r"https?://[^\s<>\"']+", re.IGNORECASE)
HEADING_RE = re.compile(r"heading\s*([1-9])$", re.IGNORECASE)
ROMAN_HEADING_RE = re.compile(r"^(?:[IVXLCDM]+|[A-Z])\.\s+\S", re.IGNORECASE)
DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent / "outputs"


def val(element: ET.Element | None, name: str, default: str = "") -> str:
    child = element.find(W + name) if element is not None else None
    return child.get(W + "val", default) if child is not None else default


def enabled(element: ET.Element | None, name: str) -> bool:
    child = element.find(W + name) if element is not None else None
    return child is not None and child.get(W + "val", "1").lower() not in {"0", "false", "off"}


def readable_url(url: str) -> str | None:
    parsed = urlparse(url)
    if parsed.scheme.lower() in {"http", "https"} and parsed.netloc:
        return url
    if parsed.scheme.lower() == "mailto" and parsed.path:
        return url
    if url.startswith("#"):
        return url
    return None


def css_font_family(font: str) -> str:
    safe = (font.replace("\\", "\\\\").replace('"', '\\"')
            .replace("<", "\\3c ").replace(">", "\\3e ")
            .replace("\r", " ").replace("\n", " "))
    return f'"{safe}", Calibri, Arial, sans-serif'


def linkify(text: str) -> str:
    pieces = []
    cursor = 0
    for match in URL_RE.finditer(text):
        start, end = match.span()
        url = match.group().rstrip(".,;:!?)")
        if not url or readable_url(url) is None:
            continue
        pieces.append(escape(text[cursor:start]))
        pieces.append(f'<a href="{escape(url, quote=True)}">{escape(url)}</a>')
        cursor = start + len(url)
    pieces.append(escape(text[cursor:]))
    return "".join(pieces)


def alpha(value: int, upper: bool = False) -> str:
    result = ""
    while value > 0:
        value, rem = divmod(value - 1, 26)
        result = chr((65 if upper else 97) + rem) + result
    return result or "0"


def roman(value: int) -> str:
    if value <= 0 or value >= 4000:
        return str(value)
    parts = []
    for amount, symbol in ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
                           (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                           (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")):
        count, value = divmod(value, amount)
        parts.append(symbol * count)
    return "".join(parts)


def formatted_number(value: int, fmt: str) -> str:
    if fmt == "lowerLetter":
        return alpha(value)
    if fmt == "upperLetter":
        return alpha(value, upper=True)
    if fmt == "lowerRoman":
        return roman(value).lower()
    if fmt == "upperRoman":
        return roman(value)
    if fmt == "decimalZero":
        return f"{value:02d}"
    if fmt == "decimal":
        return str(value)
    raise ValueError(f"Format penomoran Word belum didukung: {fmt}")


@dataclass
class Level:
    fmt: str
    pattern: str
    start: int
    restart: int | None
    indent: int | None


@dataclass
class StyleInfo:
    name: str
    outline: int | None
    num_id: str
    level: int
    based_on: str


class Numbering:
    def __init__(self, root: ET.Element | None):
        self.abstract: dict[str, dict[int, Level]] = {}
        self.instances: dict[str, tuple[str, dict[int, Level]]] = {}
        self.counters: dict[str, dict[int, int]] = defaultdict(dict)
        if root is None:
            return
        for item in root.findall(W + "abstractNum"):
            self.abstract[item.get(W + "abstractNumId", "")] = {
                int(level.get(W + "ilvl", "0")): self._level(level)
                for level in item.findall(W + "lvl")
            }
        for item in root.findall(W + "num"):
            abstract_id = val(item, "abstractNumId")
            overrides: dict[int, Level] = {}
            for override in item.findall(W + "lvlOverride"):
                index = int(override.get(W + "ilvl", "0"))
                base = self.abstract.get(abstract_id, {}).get(index)
                definition = override.find(W + "lvl")
                current = self._level(definition) if definition is not None else base
                if current is None:
                    continue
                start_override = val(override, "startOverride")
                if start_override:
                    current = Level(current.fmt, current.pattern, int(start_override),
                                    current.restart, current.indent)
                overrides[index] = current
            self.instances[item.get(W + "numId", "")] = (abstract_id, overrides)

    @staticmethod
    def _level(element: ET.Element) -> Level:
        restart = val(element, "lvlRestart")
        ppr = element.find(W + "pPr")
        indent_element = ppr.find(W + "ind") if ppr is not None else None
        left = (indent_element.get(W + "left") or indent_element.get(W + "start")) if indent_element is not None else None
        return Level(val(element, "numFmt", "decimal"), val(element, "lvlText", "%1."),
                     int(val(element, "start", "1")), int(restart) if restart else None,
                     int(left) if left is not None else None)

    def level(self, num_id: str, index: int) -> Level | None:
        instance = self.instances.get(num_id)
        if instance is None:
            return None
        abstract_id, overrides = instance
        return overrides.get(index) or self.abstract.get(abstract_id, {}).get(index)

    def label(self, num_id: str, index: int) -> tuple[str, bool] | None:
        level = self.level(num_id, index)
        if level is None:
            return None
        state = self.counters[num_id]
        state[index] = state.get(index, level.start - 1) + 1
        for deeper in list(state):
            if deeper > index:
                deeper_level = self.level(num_id, deeper)
                if deeper_level is None or deeper_level.restart != 0:
                    del state[deeper]
        if level.fmt == "bullet":
            marker = level.pattern
            if any("\uf000" <= char <= "\uf8ff" for char in marker):
                marker = "•"
            return marker, True
        if level.fmt == "none":
            return "", False

        def replace(match: re.Match[str]) -> str:
            ref_index = int(match.group(1)) - 1
            ref_level = self.level(num_id, ref_index)
            if ref_level is None:
                return match.group()
            value = state.get(ref_index, ref_level.start)
            return formatted_number(value, ref_level.fmt)

        return re.sub(r"%([1-9])", replace, level.pattern), False


class Converter:
    def __init__(self, archive: ZipFile):
        self.archive = archive
        self.styles = self._styles()
        self.theme_fonts, self.base_font, self.base_size = self._font_config()
        self.relationships = self._relationships()
        self.numbering = Numbering(self._read_xml("word/numbering.xml", required=False))
        self.lines: list[str] = []
        self.list_stack: list[tuple[int, bool]] = []
        self.previous_center_heading = False
        self.page_title: str | None = None

    def _read_xml(self, name: str, required: bool = True) -> ET.Element | None:
        try:
            return ET.fromstring(self.archive.read(name))
        except KeyError:
            if required:
                raise ValueError(f"Bagian DOCX tidak ditemukan: {name}") from None
            return None
        except ET.ParseError as exc:
            raise ValueError(f"XML DOCX tidak valid pada {name}: {exc}") from exc

    def _styles(self) -> dict[str, StyleInfo]:
        root = self._read_xml("word/styles.xml", required=False)
        result: dict[str, StyleInfo] = {}
        if root is None:
            return result
        for style in root.findall(W + "style"):
            ident = style.get(W + "styleId", "")
            name = val(style, "name", ident)
            ppr = style.find(W + "pPr")
            outline = val(ppr, "outlineLvl")
            numpr = ppr.find(W + "numPr") if ppr is not None else None
            result[ident] = StyleInfo(
                name,
                int(outline) if outline.isdigit() else None,
                val(numpr, "numId"),
                int(val(numpr, "ilvl", "0")),
                val(style, "basedOn"),
            )
        return result

    def _font_config(self) -> tuple[dict[str, str], str, float]:
        theme: dict[str, str] = {}
        theme_root = self._read_xml("word/theme/theme1.xml", required=False)
        if theme_root is not None:
            for theme_name, key in (("majorFont", "major"), ("minorFont", "minor")):
                latin = theme_root.find(f".//{A}{theme_name}/{A}latin")
                if latin is not None and latin.get("typeface"):
                    theme[key] = latin.get("typeface", "")
        styles_root = self._read_xml("word/styles.xml", required=False)
        default_run = styles_root.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr") if styles_root is not None else None
        fonts = default_run.find(W + "rFonts") if default_run is not None else None
        face = self._font_from_element(fonts, theme) or "Calibri"
        size = val(default_run, "sz")
        return theme, face, int(size) / 2 if size.isdigit() else 12.0

    @staticmethod
    def _font_from_element(fonts: ET.Element | None, theme: dict[str, str]) -> str | None:
        if fonts is None:
            return None
        explicit = fonts.get(W + "ascii") or fonts.get(W + "hAnsi")
        if explicit:
            return explicit
        theme_name = fonts.get(W + "asciiTheme") or fonts.get(W + "hAnsiTheme", "")
        if theme_name.startswith("major"):
            return theme.get("major")
        if theme_name.startswith("minor"):
            return theme.get("minor")
        return None

    def _style_info(self, style_id: str) -> StyleInfo:
        seen = set()
        current = style_id
        while current and current not in seen:
            seen.add(current)
            info = self.styles.get(current)
            if info is None:
                break
            if info.num_id or info.outline is not None or not info.based_on:
                return info
            current = info.based_on
        return self.styles.get(style_id, StyleInfo(style_id, None, "", 0, ""))

    def _relationships(self) -> dict[str, str]:
        root = self._read_xml("word/_rels/document.xml.rels", required=False)
        if root is None:
            return {}
        return {
            item.get("Id", ""): item.get("Target", "")
            for item in root.findall(PKG_R + "Relationship")
        }

    @staticmethod
    def _paragraph_text(paragraph: ET.Element) -> str:
        return "".join(node.text or "" for node in paragraph.iter(W + "t")).strip()

    def _render_run(self, run: ET.Element, in_link: bool = False) -> str:
        chunks: list[str] = []
        for item in run:
            if item.tag == W + "t":
                chunks.append(escape(item.text or "") if in_link else linkify(item.text or ""))
            elif item.tag == W + "tab":
                chunks.append("&emsp;")
            elif item.tag in {W + "br", W + "cr"}:
                chunks.append("<br>")
            elif item.tag == W + "noBreakHyphen":
                chunks.append("&#8209;")
            elif item.tag == W + "softHyphen":
                chunks.append("&shy;")
        content = "".join(chunks)
        if not content:
            return ""
        props = run.find(W + "rPr")
        font = self._font_from_element(props.find(W + "rFonts") if props is not None else None,
                                       self.theme_fonts)
        size = val(props, "sz")
        styles = []
        if font and font != self.base_font:
            styles.append(f"font-family: {css_font_family(font)}")
        if size.isdigit() and int(size) / 2 != self.base_size:
            styles.append(f"font-size: {int(size) / 2:g}pt")
        if styles:
            content = f'<span style="{escape("; ".join(styles), quote=True)}">{content}</span>'
        if enabled(props, "b"):
            content = f"<strong>{content}</strong>"
        if enabled(props, "i"):
            content = f"<em>{content}</em>"
        if enabled(props, "u"):
            content = f"<u>{content}</u>"
        if enabled(props, "strike"):
            content = f"<s>{content}</s>"
        vertical = val(props, "vertAlign")
        if vertical == "superscript":
            content = f"<sup>{content}</sup>"
        elif vertical == "subscript":
            content = f"<sub>{content}</sub>"
        return content

    def _render_inline(self, paragraph: ET.Element) -> str:
        pieces: list[str] = []
        for item in paragraph:
            if item.tag == W + "r":
                pieces.append(self._render_run(item))
            elif item.tag == W + "hyperlink":
                href = self.relationships.get(item.get(R + "id", ""))
                if not href and item.get(W + "anchor"):
                    href = "#" + item.get(W + "anchor", "")
                inner = "".join(self._render_run(run, in_link=True) for run in item.iter(W + "r"))
                safe_href = readable_url(href) if href else None
                pieces.append(f'<a href="{escape(safe_href, quote=True)}">{inner}</a>' if safe_href else inner)
            elif item.tag in {W + "sdt", W + "ins"}:
                pieces.extend(self._render_run(run) for run in item.iter(W + "r"))
        return "".join(pieces)

    def _heading_level(self, paragraph: ET.Element, text: str, infer_manual: bool = True) -> int | None:
        ppr = paragraph.find(W + "pPr")
        style_id = val(ppr, "pStyle")
        info = self._style_info(style_id)
        style_name, outline = info.name, info.outline
        match = HEADING_RE.match(style_id) or HEADING_RE.match(style_name)
        if match:
            return min(int(match.group(1)), 6)
        if outline is not None and 0 <= outline <= 8:
            return min(outline + 1, 6)
        direct_outline = val(ppr, "outlineLvl")
        if direct_outline.isdigit() and int(direct_outline) <= 8:
            return min(int(direct_outline) + 1, 6)
        if not infer_manual:
            return None
        if not text or len(text) > 120 or text.endswith((".", ":", ";", "?", "!")):
            return None
        runs = list(paragraph.findall(W + "r"))
        weighted_length = sum(len(self._paragraph_text(run)) for run in runs)
        bold_length = sum(len(self._paragraph_text(run)) for run in runs
                          if enabled(run.find(W + "rPr"), "b"))
        fully_bold = weighted_length > 0 and bold_length / weighted_length >= 0.9
        alignment = val(ppr, "jc")
        if alignment == "center" and fully_bold:
            return 2 if self.previous_center_heading else 1
        if ROMAN_HEADING_RE.match(text) and fully_bold:
            return 3
        letters = [char for char in text if char.isalpha()]
        upper = bool(letters) and all(char.isupper() for char in letters)
        if upper and (fully_bold or len(text) <= 40 and len(text.split()) >= 2):
            return 2
        if fully_bold and len(text) <= 80:
            return 3
        return None

    def _close_lists(self) -> None:
        while self.list_stack:
            self.lines.append("</div></li>")
            self.lines.append("</ul>" if self.list_stack.pop()[1] else "</ol>")

    def _add_list_item(self, indent: int, marker: str, bullet: bool, html: str) -> None:
        while self.list_stack and (self.list_stack[-1][0] > indent or
                                   self.list_stack[-1][0] == indent and self.list_stack[-1][1] != bullet):
            self.lines.append("</div></li>")
            self.lines.append("</ul>" if self.list_stack.pop()[1] else "</ol>")
        if self.list_stack and self.list_stack[-1][0] == indent:
            self.lines.append("</div></li>")
        else:
            tag = "ul" if bullet else "ol"
            self.lines.append(f'<{tag} class="word-list" data-level="{len(self.list_stack)}">')
            self.list_stack.append((indent, bullet))
        self.lines.append(f'<li><span class="list-marker">{escape(marker)}</span><div class="list-content">{html}')

    def _add_paragraph(self, paragraph: ET.Element, in_table: bool = False) -> None:
        text = self._paragraph_text(paragraph)
        html = self._render_inline(paragraph)
        if not text and not html:
            self._close_lists()
            self.previous_center_heading = False
            return
        ppr = paragraph.find(W + "pPr")
        numpr = ppr.find(W + "numPr") if ppr is not None else None
        style = self._style_info(val(ppr, "pStyle"))
        num_id = val(numpr, "numId", style.num_id)
        if num_id and num_id != "0":
            level = int(val(numpr, "ilvl", str(style.level)))
            label = self.numbering.label(num_id, level)
            if label is None:
                raise ValueError(f"Definisi penomoran Word tidak ditemukan: numId={num_id}, level={level}")
            marker, bullet = label
            definition = self.numbering.level(num_id, level)
            direct_indent = ppr.find(W + "ind") if ppr is not None else None
            left = (direct_indent.get(W + "left") or direct_indent.get(W + "start")) if direct_indent is not None else None
            indent = int(left) if left is not None else (definition.indent if definition and definition.indent is not None else (level + 1) * 720)
            self._add_list_item(indent, marker, bullet, html)
            self.previous_center_heading = False
            return
        self._close_lists()
        heading = self._heading_level(paragraph, text, infer_manual=not in_table)
        alignment = ' class="word-center"' if val(ppr, "jc") == "center" else ""
        if heading:
            self.lines.append(f"<h{heading}{alignment}>{html}</h{heading}>")
            if self.page_title is None and heading == 1:
                self.page_title = text
        else:
            self.lines.append(f"<p{alignment}>{html}</p>")
        self.previous_center_heading = heading is not None and val(ppr, "jc") == "center"

    def _add_table(self, table: ET.Element) -> None:
        self._close_lists()
        self.lines.append("<table>")
        for row in table.findall(W + "tr"):
            self.lines.append("<tr>")
            for cell in row.findall(W + "tc"):
                self.lines.append("<td>")
                self._add_blocks(cell, in_table=True)
                self._close_lists()
                self.lines.append("</td>")
            self.lines.append("</tr>")
        self.lines.append("</table>")

    def _add_blocks(self, parent: ET.Element, in_table: bool = False) -> None:
        for item in parent:
            if item.tag == W + "p":
                self._add_paragraph(item, in_table=in_table)
            elif item.tag == W + "tbl":
                self._add_table(item)
            elif item.tag == W + "sdt":
                content = item.find(W + "sdtContent")
                if content is not None:
                    self._add_blocks(content, in_table=in_table)

    def convert(self, fallback_title: str) -> str:
        root = self._read_xml("word/document.xml")
        body = root.find(W + "body") if root is not None else None
        if body is None:
            raise ValueError("Isi dokumen Word tidak ditemukan")
        self._add_blocks(body)
        self._close_lists()
        title = self.page_title or fallback_title
        content = "\n".join(self.lines)
        return ("<!doctype html>\n<html lang=\"id\">\n<head>\n"
                "<meta charset=\"utf-8\">\n"
                "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n"
                f"<title>{escape(title)}</title>\n"
                "<style>\n"
                f"body {{ max-width: 75ch; margin: 2rem auto; padding: 0 1rem; "
                f"font-family: {css_font_family(self.base_font)}; font-size: {self.base_size:g}pt; "
                "line-height: 1.5; }\n"
                "h1, h2, h3, h4, h5, h6 { font-size: 1em; font-weight: normal; }\n"
                ".word-center { text-align: center; }\n"
                "table { border-collapse: collapse; }\n"
                "th, td { border: 1px solid #888; padding: .4rem; vertical-align: top; }\n"
                ".word-list { list-style: none; padding-left: 2rem; }\n"
                ".word-list > li { display: grid; grid-template-columns: max-content minmax(0, 1fr); gap: .4rem; }\n"
                ".list-marker { white-space: nowrap; }\n"
                ".list-content { min-width: 0; }\n"
                "</style>\n</head>\n<body>\n"
                f"{content}\n</body>\n</html>\n")


def convert_file(source: Path, destination: Path, force: bool = False) -> None:
    if source.suffix.lower() != ".docx":
        raise ValueError("File sumber harus berformat .docx")
    if not source.is_file():
        raise ValueError(f"File sumber tidak ditemukan: {source}")
    if destination.resolve() == source.resolve():
        raise ValueError("File hasil harus berbeda dari file sumber")
    if destination.exists() and not force:
        raise FileExistsError(f"File hasil sudah ada: {destination} (gunakan --force untuk menimpa)")
    with ZipFile(source) as archive:
        html = Converter(archive).convert(source.stem)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("w" if force else "x", encoding="utf-8", newline="\n") as output:
        output.write(html)


def run_wizard() -> int:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except ImportError as exc:
        print(f"Gagal membuka pemilih file: {exc}", file=sys.stderr)
        return 1
    try:
        root = tk.Tk()
    except tk.TclError as exc:
        print(f"Gagal membuka pemilih file: {exc}", file=sys.stderr)
        return 1

    root.withdraw()
    try:
        selected = filedialog.askopenfilenames(
            parent=root,
            title="Pilih satu atau beberapa dokumen Word",
            filetypes=[("Dokumen Word (.docx)", "*.docx")],
        )
        if not selected:
            return 0
        planned: list[tuple[Path, Path]] = []
        used_names: set[str] = set()
        for raw_source in selected:
            source = Path(raw_source)
            name = source.with_suffix(".html").name
            if name.casefold() in used_names:
                number = 2
                while True:
                    name = f"{source.stem} ({number}).html"
                    if name.casefold() not in used_names:
                        break
                    number += 1
            used_names.add(name.casefold())
            planned.append((source, DEFAULT_OUTPUT_DIR / name))

        existing = [target for _, target in planned if target.exists()]
        overwrite = False
        if existing:
            overwrite = messagebox.askyesno(
                "File hasil sudah ada",
                f"{len(existing)} file HTML sudah ada di folder outputs.\n"
                "Pilih Ya untuk menimpanya, atau Tidak untuk melewatinya.",
                parent=root,
            )
        converted: list[Path] = []
        skipped = 0
        errors: list[str] = []
        for source, destination in planned:
            if destination.exists() and not overwrite:
                skipped += 1
                continue
            try:
                convert_file(source, destination, force=overwrite and destination.exists())
            except (ValueError, FileExistsError, BadZipFile, OSError) as exc:
                errors.append(f"{source.name}: {exc}")
            else:
                converted.append(destination)

        summary = (f"{len(converted)} berhasil, {skipped} dilewati, {len(errors)} gagal.\n"
                   f"Folder hasil: {DEFAULT_OUTPUT_DIR}")
        if errors:
            details = "\n".join(errors[:5])
            if len(errors) > 5:
                details += f"\n... dan {len(errors) - 5} kesalahan lain"
            messagebox.showerror("Konversi selesai dengan kesalahan", f"{summary}\n\n{details}", parent=root)
            return 1
        if len(converted) == 1 and skipped == 0:
            summary = f"HTML tersimpan di:\n{converted[0]}"
        messagebox.showinfo("Konversi selesai", summary, parent=root)
        return 0
    finally:
        root.destroy()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Ubah dokumen Word .docx menjadi halaman HTML utuh.")
    parser.add_argument("source", nargs="?", type=Path, help="File .docx sumber; kosongkan untuk pemilih file")
    parser.add_argument("-o", "--output", type=Path, help="Lokasi file .html hasil")
    parser.add_argument("--force", action="store_true", help="Timpa file hasil yang sudah ada")
    args = parser.parse_args(argv)
    if args.source is None:
        if args.output is not None or args.force:
            parser.error("-o dan --force memerlukan file sumber")
        return run_wizard()
    destination = args.output if args.output else DEFAULT_OUTPUT_DIR / args.source.with_suffix(".html").name
    try:
        convert_file(args.source, destination, args.force)
    except (ValueError, FileExistsError, BadZipFile, OSError) as exc:
        print(f"Gagal: {exc}", file=sys.stderr)
        return 1
    print(f"Berhasil: {destination}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
