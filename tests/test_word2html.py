from contextlib import redirect_stdout
from io import StringIO
from shutil import copyfile
import tempfile
import unittest
from html.parser import HTMLParser
from pathlib import Path
from unittest.mock import patch
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from word2html import Numbering, convert_file, main


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
SAMPLE = Path(__file__).resolve().parents[1] / "sample" / "ISA ID AFTER - Konven & Syariah.docx"


class VisibleText(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.in_body = False
        self.in_marker = False
        self.text = []
        self.tags = []
        self.markers = []
        self.marker_depths = []
        self.li_depth = 0

    def handle_starttag(self, tag, attrs):
        self.tags.append(tag)
        if tag == "body":
            self.in_body = True
        if tag == "li":
            self.li_depth += 1
        if tag == "span" and ("class", "list-marker") in attrs:
            self.in_marker = True
            self.markers.append("")
            self.marker_depths.append(self.li_depth)

    def handle_endtag(self, tag):
        if tag == "span" and self.in_marker:
            self.in_marker = False
        if tag == "body":
            self.in_body = False
        if tag == "li":
            self.li_depth -= 1

    def handle_data(self, data):
        if self.in_marker:
            self.markers[-1] += data
        elif self.in_body:
            self.text.append(data)


class ConversionTests(unittest.TestCase):
    def test_sample_preserves_all_source_text_and_numbered_items(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "tnc.html"
            convert_file(SAMPLE, target)
            html = target.read_text(encoding="utf-8")
            parsed = VisibleText()
            parsed.feed(html)

            with ZipFile(SAMPLE) as archive:
                root = ET.fromstring(archive.read("word/document.xml"))
            source = "".join(
                "".join(
                    (item.text or "") if item.tag == W + "t" else "\u2003"
                    for item in paragraph.iter()
                    if item.tag in {W + "t", W + "tab"}
                )
                for paragraph in root.iter(W + "p")
                if "".join(item.text or "" for item in paragraph.iter(W + "t")).strip()
            )
            self.assertEqual(source, "".join(parsed.text).replace("\n", ""))
            self.assertEqual(1436, len(parsed.markers))
            self.assertEqual(["1.", "a.", "b."], parsed.markers[:3])
            self.assertEqual([1, 2, 2, 2], parsed.marker_depths[:4])
            self.assertEqual(1, parsed.marker_depths[7])
            nested_fourteen = [
                i for i in range(len(parsed.markers) - 3)
                if parsed.markers[i:i + 4] == ["14.", "d.", "e.", "f."]
            ]
            self.assertTrue(nested_fourteen)
            for index in nested_fourteen:
                self.assertEqual([1, 2, 2, 2], parsed.marker_depths[index:index + 4])
            self.assertEqual(2, parsed.tags.count("h1"))
            self.assertIn('<html lang="id">', html)
            self.assertIn('href="https://', html)
            self.assertIn('font-family: "Aptos", Calibri, Arial, sans-serif; font-size: 12pt', html)
            self.assertIn('<h1 class="word-center">', html)
            self.assertIn('font-family: &quot;Arial&quot;', html)

    def test_output_is_not_overwritten_without_force(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "tnc.html"
            target.write_text("existing", encoding="utf-8")
            with self.assertRaises(FileExistsError):
                convert_file(SAMPLE, target)
            self.assertEqual("existing", target.read_text(encoding="utf-8"))
            convert_file(SAMPLE, target, force=True)
            self.assertTrue(target.read_text(encoding="utf-8").startswith("<!doctype html>"))

    def test_cli_creates_a_separate_default_output_directory(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "input" / "terms.docx"
            source.parent.mkdir()
            copyfile(SAMPLE, source)
            output_dir = root / "outputs"
            with patch("word2html.DEFAULT_OUTPUT_DIR", output_dir), redirect_stdout(StringIO()):
                self.assertEqual(0, main([str(source)]))
            self.assertTrue((output_dir / "terms.html").is_file())
            self.assertFalse(source.with_suffix(".html").exists())

    def test_numbering_respects_starts_levels_and_restarts(self):
        xml = f'''<w:numbering xmlns:w="{W[1:-1]}">
          <w:abstractNum w:abstractNumId="1">
            <w:lvl w:ilvl="0"><w:start w:val="2"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
            <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl>
          </w:abstractNum>
          <w:num w:numId="5"><w:abstractNumId w:val="1"/></w:num>
        </w:numbering>'''
        numbering = Numbering(ET.fromstring(xml))
        self.assertEqual(("2.", False), numbering.label("5", 0))
        self.assertEqual(("a)", False), numbering.label("5", 1))
        self.assertEqual(("b)", False), numbering.label("5", 1))
        self.assertEqual(("3.", False), numbering.label("5", 0))
        self.assertEqual(("a)", False), numbering.label("5", 1))

    def test_heading_table_and_plain_url_in_another_document(self):
        xml = f'''<w:document xmlns:w="{W[1:-1]}"><w:body>
          <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Terms &amp; Conditions</w:t></w:r></w:p>
          <w:p><w:r><w:t>See https://example.com/rules.</w:t></w:r></w:p>
          <w:tbl><w:tr>
            <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
          </w:tr></w:tbl>
        </w:body></w:document>'''
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "other.docx"
            target = Path(folder) / "other.html"
            with ZipFile(source, "w") as archive:
                archive.writestr("word/document.xml", xml)
            convert_file(source, target)
            html = target.read_text(encoding="utf-8")
            self.assertIn("<h1>Terms &amp; Conditions</h1>", html)
            self.assertIn('<a href="https://example.com/rules">https://example.com/rules</a>.', html)
            self.assertIn("<table>\n<tr>\n<td>\n<p>A</p>\n</td>\n<td>\n<p>B</p>", html)


if __name__ == "__main__":
    unittest.main()
