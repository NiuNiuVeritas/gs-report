from __future__ import annotations

import argparse
import html
import re
from html.parser import HTMLParser
from pathlib import Path

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


class TextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        self.parts.append(data)


def normalized(value: str) -> str:
    value = html.unescape(value)
    value = value.replace("\uf06c", "")
    value = re.sub(r"\s+", "", value)
    return value


def visible_text(markdown: str) -> str:
    parser = TextExtractor()
    parser.feed(markdown)
    return normalized("\n".join(parser.parts))


def iter_blocks(document: Document):
    for child in document.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)


def paragraph_text(paragraph: Paragraph) -> str:
    return ("".join(run.text for run in paragraph.runs).strip() or paragraph.text.strip()).replace("\uf06c", "").strip()


def run_is_bold(run) -> bool:
    if run.bold is True:
        return True
    bold_nodes = run._r.xpath("./w:rPr/w:b")
    if not bold_nodes:
        return False
    value = bold_nodes[0].get(f"{{{NS['w']}}}val")
    return value not in {"0", "false", "False", "off"}


def paragraph_bold_segments(paragraph: Paragraph) -> list[str]:
    segments: list[str] = []
    current: list[str] = []
    for run in paragraph.runs:
        if run_is_bold(run):
            current.append(run.text)
        elif current:
            text = "".join(current).strip()
            if text:
                segments.append(text)
            current = []
    if current:
        text = "".join(current).strip()
        if text:
            segments.append(text)
    return segments


def row_unique_texts(row) -> list[str]:
    values: list[str] = []
    for cell in row.cells:
        text = cell.text.strip()
        if text and (not values or values[-1] != text):
            values.append(text)
    return values


def table_has_image(table: Table) -> bool:
    root = etree.fromstring(table._element.xml.encode("utf-8"))
    refs = root.xpath(".//a:blip/@r:embed | .//v:imagedata/@r:id", namespaces=NS)
    return bool(refs)


def table_title(table: Table) -> str:
    if not table.rows:
        return "图表"
    titles = [text for text in row_unique_texts(table.rows[0]) if not text.startswith("资料来源")]
    return titles[0] if titles else "图表"


def clean_text(value: str) -> str:
    value = value.replace("\uf06c", "").replace("\xa0", " ")
    value = re.sub(r"[ \t]+", " ", value)
    return value.strip()


def iter_table_paragraph_nodes(document: Document):
    for table in document.tables:
        root = etree.fromstring(table._element.xml.encode("utf-8"))
        for pnode in root.xpath(".//w:p", namespaces=NS):
            text = clean_text("".join(pnode.xpath(".//w:t/text()", namespaces=NS)))
            if not text:
                continue
            style = pnode.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
            numid = pnode.xpath("./w:pPr/w:numPr/w:numId/@w:val", namespaces=NS)
            yield text, bool(numid), style


def extract_core_summary_rows(document: Document) -> list[tuple[str, bool]]:
    paragraphs = list(iter_table_paragraph_nodes(document))
    start = next((index for index, item in enumerate(paragraphs) if item[0] == "核心观点"), None)
    if start is None:
        return []
    rows: list[tuple[str, bool]] = []
    in_core_body = False
    for text, is_bullet, style in paragraphs[start + 1 :]:
        if text.startswith("风险提示："):
            break
        if "23" in style:
            in_core_body = True
        if in_core_body:
            rows.append((text, is_bullet))
    return rows


def summary_block(markdown: str) -> str:
    start = markdown.find("报告摘要")
    if start == -1:
        return ""
    end = markdown.find('background:#343d5d;color:#ffffff', start)
    if end == -1:
        return markdown[start:]
    return markdown[start:end]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Verify GS report WeChat output against the source Word report.")
    parser.add_argument("--docx", required=True, type=Path)
    parser.add_argument("--markdown", required=True, type=Path)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    document = Document(args.docx)
    markdown = args.markdown.read_text(encoding="utf-8")
    text = visible_text(markdown)
    summary_raw = summary_block(markdown)
    summary_text = visible_text(summary_raw)

    missing_paragraphs: list[str] = []
    checked_paragraphs = 0
    bold_segments_checked = 0
    missing_bold_segments: list[str] = []
    summary_rows = extract_core_summary_rows(document)
    summary_text_rows = [row for row in summary_rows if not row[0].startswith("风险提示：")]
    missing_summary: list[str] = [row for row, _ in summary_text_rows if normalized(row) not in summary_text]
    expected_summary_bullets = sum(1 for _, is_bullet in summary_rows if is_bullet)
    actual_summary_bullets = len(re.findall(r'data-gs-summary-bullet="true"', summary_raw))
    missing_markers: list[str] = []
    checked_markers = 0
    started = False
    figure_count = 0
    table_count = 0

    for block in iter_blocks(document):
        if isinstance(block, Paragraph):
            source = paragraph_text(block)
            if not source:
                continue
            style = block.style.name
            if style == "国信研报正文-1.正文一级标题":
                started = True
            if started and style == "Normal" and source == "免责声明":
                break
            if started and style in {
                "国信研报正文-1.正文一级标题",
                "国信研报正文-2.正文二级标题",
                "国信研报正文-4.正文",
            }:
                checked_paragraphs += 1
                if normalized(source) not in text:
                    missing_paragraphs.append(source)
                if style == "国信研报正文-4.正文":
                    for segment in paragraph_bold_segments(block):
                        bold_segments_checked += 1
                        pattern = re.compile(r"<strong>\s*" + re.escape(html.escape(segment)) + r"\s*</strong>")
                        if not pattern.search(markdown):
                            missing_bold_segments.append(segment)
        elif started:
            title = table_title(block)
            if table_has_image(block):
                root = etree.fromstring(block._element.xml.encode("utf-8"))
                refs = root.xpath(".//a:blip/@r:embed | .//v:imagedata/@r:id", namespaces=NS)
                for _ in refs:
                    figure_count += 1
                    checked_markers += 1
                    marker = f"图{figure_count}：{title}"
                    if normalized(marker) not in text:
                        missing_markers.append(marker)
            else:
                if title and len(block.rows) > 1:
                    table_count += 1
                    checked_markers += 1
                    marker = f"表{table_count}：{title}"
                    if normalized(marker) not in text:
                        missing_markers.append(marker)

    unresolved = re.findall(r"\{\{[^}]+\}\}", markdown)
    img_refs = re.findall(r'<img src="([^"]+)"', markdown)
    missing_images = [ref for ref in img_refs if not (args.markdown.parent / ref).exists()]
    footer_checks = {
        "source_note": "注：本文选自国信证券于" in markdown,
        "analyst": "分析师：" in markdown and re.search(r"S\d{13}", markdown) is not None,
        "risk": "风险提示：" in markdown,
        "profile": "mp-common-profile" in markdown,
        "law": "law.png" in markdown,
    }

    print(f"paragraphs_checked={checked_paragraphs}")
    print(f"paragraphs_missing={len(missing_paragraphs)}")
    print(f"bold_segments_checked={bold_segments_checked}")
    print(f"bold_segments_missing={len(missing_bold_segments)}")
    print(f"summary_rows_checked={len(summary_text_rows)}")
    print(f"summary_rows_missing={len(missing_summary)}")
    print(f"summary_bullets_expected={expected_summary_bullets}")
    print(f"summary_bullets_actual={actual_summary_bullets}")
    print(f"markers_checked={checked_markers}")
    print(f"markers_missing={len(missing_markers)}")
    print(f"image_refs={len(img_refs)}")
    print(f"missing_images={len(missing_images)}")
    print(f"unresolved_placeholders={len(unresolved)}")
    for key, ok in footer_checks.items():
        print(f"footer_{key}={ok}")

    failed = (
        missing_paragraphs
        or missing_bold_segments
        or missing_summary
        or expected_summary_bullets != actual_summary_bullets
        or missing_markers
        or missing_images
        or unresolved
        or not all(footer_checks.values())
    )
    if failed:
        for item in missing_summary[:20]:
            print(f"MISSING_SUMMARY: {item[:200]}")
        for item in missing_bold_segments[:20]:
            print(f"MISSING_BOLD: {item[:200]}")
        for item in missing_paragraphs[:20]:
            print(f"MISSING_PARAGRAPH: {item[:200]}")
        for item in missing_markers[:20]:
            print(f"MISSING_MARKER: {item[:200]}")
        for item in missing_images:
            print(f"MISSING_IMAGE: {item}")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
