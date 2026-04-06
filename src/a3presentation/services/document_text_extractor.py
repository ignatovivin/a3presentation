from __future__ import annotations

import base64
import re
from io import BytesIO
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from pypdf import PdfReader

from a3presentation.domain.api import DocumentBlock
from a3presentation.domain.presentation import TableBlock

class DocumentTextExtractor:
    SUPPORTED_EXTENSIONS = {".txt", ".md", ".markdown", ".pdf", ".docx"}
    TITLE_STYLE_NAMES = {"title", "заголовок документа"}
    SUBTITLE_STYLE_NAMES = {"subtitle", "подзаголовок"}
    HEADING_STYLE_PREFIXES = ("heading", "заголовок")
    LIST_STYLE_MARKERS = ("list", "список", "bullet", "маркир", "number", "нумер")
    LIST_TEXT_PREFIXES = ("- ", "• ", "– ", "— ", "* ")
    OUTLINE_HEADING_PATTERN = re.compile(r"^(?P<number>\d+(?:\.\d+)*)[.)]?\s+.+$")
    QUARTER_HEADING_PATTERN = re.compile(r"^Q[1-4](?:\s+\d{4})?(?:\s*[-—–:].+)?$", re.IGNORECASE)
    STRUCTURED_LABEL_PATTERN = re.compile(r"^\s*[^:]{1,80}:\s+\S")
    CONTINUATION_START_PATTERN = re.compile(r"^(?:[a-zа-яё0-9(«\"'—–-]|и\b|или\b|по\b|для\b|который\b)")
    URL_ONLY_PATTERN = re.compile(r"^(?:https?://|www\.)\S+$", re.IGNORECASE)
    REFERENCE_LINE_PATTERN = re.compile(r"^(?:\[\d+\]\s*)+(?:https?://\S+)?$", re.IGNORECASE)
    TERMINAL_PUNCTUATION = (".", "!", "?", ";")

    def extract(self, filename: str, content: bytes) -> tuple[str, list[TableBlock], list[DocumentBlock]]:
        extension = Path(filename).suffix.lower()
        if extension not in self.SUPPORTED_EXTENSIONS:
            raise ValueError(f"Unsupported document type: {extension or 'unknown'}")

        if extension in {".txt", ".md", ".markdown"}:
            text = content.decode("utf-8", errors="ignore").strip()
            blocks = self._extract_plain_text_blocks(text, markdown=extension in {".md", ".markdown"})
            return text, [], blocks
        if extension == ".pdf":
            text = self._extract_pdf(content)
            return text, [], [DocumentBlock(kind="paragraph", text=text, items=[])] if text else []
        if extension == ".docx":
            return self._extract_docx(content)

        raise ValueError(f"Unsupported document type: {extension}")

    def _extract_pdf(self, content: bytes) -> str:
        reader = PdfReader(BytesIO(content))
        pages: list[str] = []
        for page in reader.pages:
            extracted = page.extract_text() or ""
            if extracted.strip():
                pages.append(extracted.strip())
        return "\n\n".join(pages).strip()

    def _extract_docx(self, content: bytes) -> tuple[str, list[TableBlock], list[DocumentBlock]]:
        document = Document(BytesIO(content))
        tables: list[TableBlock] = []
        blocks: list[DocumentBlock] = []

        list_buffer: list[str] = []
        for block in self._iter_docx_blocks(document):
            if isinstance(block, Table):
                table_block = self._normalize_table(block)
                if table_block is None:
                    continue
                if list_buffer:
                    blocks.append(DocumentBlock(kind="list", items=list_buffer.copy()))
                    list_buffer = []
                tables.append(table_block)
                blocks.append(
                    DocumentBlock(
                        kind="table",
                        items=[],
                        table=table_block,
                    )
                )
                continue

            paragraph = block
            text = paragraph.text.strip()
            hyperlinks = self._extract_hyperlinks(paragraph)
            inline_images = self._extract_inline_images(paragraph)
            if not text:
                if list_buffer:
                    blocks.append(DocumentBlock(kind="list", items=list_buffer.copy()))
                    list_buffer = []
                blocks.extend(inline_images)
                continue

            paragraph_kind, level = self._classify_paragraph(paragraph, text)
            style_name = paragraph.style.name if paragraph.style is not None else None
            style_id = paragraph.style.style_id if paragraph.style is not None else None
            if paragraph_kind == "list":
                list_buffer.append(text)
                blocks.extend(inline_images)
                continue

            if list_buffer:
                blocks.append(DocumentBlock(kind="list", items=list_buffer.copy()))
                list_buffer = []

            blocks.append(
                DocumentBlock(
                    kind=paragraph_kind,
                    text=text,
                    level=level,
                    style_name=style_name,
                    style_id=style_id,
                    hyperlinks=hyperlinks,
                    run_count=len(paragraph.runs),
                )
            )
            blocks.extend(inline_images)

        if list_buffer:
            blocks.append(DocumentBlock(kind="list", items=list_buffer.copy()))

        blocks = self._merge_narrative_paragraph_blocks(blocks)
        return self._blocks_to_text(blocks), tables, blocks

    def _extract_plain_text_blocks(self, text: str, *, markdown: bool) -> list[DocumentBlock]:
        lines = [line.rstrip() for line in text.replace("\r", "").split("\n")]
        blocks: list[DocumentBlock] = []
        list_buffer: list[str] = []

        def flush_list() -> None:
            nonlocal list_buffer
            if list_buffer:
                blocks.append(DocumentBlock(kind="list", items=list_buffer.copy()))
                list_buffer = []

        for raw_line in lines:
            line = raw_line.strip()
            if not line:
                flush_list()
                continue

            heading_level = self._plain_text_heading_level(line, markdown)
            if heading_level is not None:
                flush_list()
                normalized = re.sub(r"^\s*#+\s*", "", line).strip() if markdown else line
                kind = "heading" if heading_level <= 1 else "subheading"
                blocks.append(DocumentBlock(kind=kind, text=normalized, level=heading_level))
                continue

            list_item = self._plain_text_list_item(line, markdown)
            if list_item is not None:
                list_buffer.append(list_item)
                continue

            flush_list()
            blocks.append(DocumentBlock(kind="paragraph", text=line))

        flush_list()
        return blocks

    def _plain_text_heading_level(self, line: str, markdown: bool) -> int | None:
        if markdown:
            match = re.match(r"^(#{1,6})\s+(.+)$", line)
            if match:
                return len(match.group(1))
            if re.match(r"^\s*\d+[.)]\s+.+$", line):
                return None
        normalized = line.strip()
        if len(normalized) <= 90 and normalized == normalized.upper() and any(char.isalpha() for char in normalized):
            return 1
        outline_level = self._outline_heading_level(normalized)
        if outline_level is not None:
            return outline_level
        return None

    def _plain_text_list_item(self, line: str, markdown: bool) -> str | None:
        if markdown:
            match = re.match(r"^\s*(?:[-*+]|\d+[.)])\s+(.+)$", line)
            if match:
                return match.group(1).strip()
        else:
            match = re.match(r"^\s*(?:[-*•]|\d+[.)])\s+(.+)$", line)
            if match:
                return match.group(1).strip()
        return None

    def _iter_docx_blocks(self, document):
        for child in document.element.body.iterchildren():
            tag = child.tag.rsplit("}", 1)[-1]
            if tag == "p":
                yield Paragraph(child, document)
            elif tag == "tbl":
                yield Table(child, document)

    def _normalize_table(self, table: Table) -> TableBlock | None:
        rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
        fill_rows = [[self._table_cell_fill_color(cell) for cell in row.cells] for row in table.rows]
        normalized_pairs = [
            (row, fill_row)
            for row, fill_row in zip(rows, fill_rows, strict=False)
            if any(cell for cell in row)
        ]
        normalized_rows = [row for row, _ in normalized_pairs]
        if not normalized_rows:
            return None

        headers = normalized_rows[0]
        body_rows = normalized_rows[1:] if len(normalized_rows) > 1 else []
        header_fill_colors = normalized_pairs[0][1] if normalized_pairs else []
        body_fill_colors = [fill_row for _, fill_row in normalized_pairs[1:]] if len(normalized_pairs) > 1 else []
        return TableBlock(
            headers=headers,
            header_fill_colors=header_fill_colors,
            rows=body_rows,
            row_fill_colors=body_fill_colors,
        )

    def _classify_paragraph(self, paragraph: Paragraph, text: str) -> tuple[str, int | None]:
        style_name = (paragraph.style.name if paragraph.style is not None else "").strip().lower()
        style_id = (paragraph.style.style_id if paragraph.style is not None else "").strip().lower()
        if style_name in self.TITLE_STYLE_NAMES or style_id == "title":
            return "title", 0
        if style_name in self.SUBTITLE_STYLE_NAMES or style_id == "subtitle":
            return "subheading", 2
        if any(style_name.startswith(prefix) for prefix in self.HEADING_STYLE_PREFIXES) or style_id.startswith("heading"):
            level = self._extract_heading_level(style_name or style_id)
            return ("heading" if level <= 1 else "subheading"), level
        if self._looks_like_list_paragraph(paragraph, text, style_name):
            return "list", None
        heading_level = self._infer_heading_level_from_text(text)
        if heading_level is not None:
            return ("heading" if heading_level <= 1 else "subheading"), heading_level
        return "paragraph", None

    def _extract_heading_level(self, style_name: str) -> int:
        digits = "".join(character for character in style_name if character.isdigit())
        if not digits:
            return 1
        return int(digits)

    def _outline_heading_level(self, text: str) -> int | None:
        match = self.OUTLINE_HEADING_PATTERN.match(text.strip())
        if match is None:
            return None
        return max(1, min(len(match.group("number").split(".")), 3))

    def _infer_heading_level_from_text(self, text: str) -> int | None:
        normalized = text.strip()
        if not normalized or len(normalized) > 90:
            return None
        if normalized.endswith((".", "!", "?")):
            return None

        outline_level = self._outline_heading_level(normalized)
        if outline_level is not None:
            return outline_level
        if self.QUARTER_HEADING_PATTERN.match(normalized):
            return 2
        return None

    def _looks_like_list_paragraph(self, paragraph: Paragraph, text: str, style_name: str) -> bool:
        stripped = text.strip()
        if not stripped:
            return False
        if any(marker in style_name for marker in self.LIST_STYLE_MARKERS):
            return True
        if self._has_numbering(paragraph):
            return True
        if stripped.startswith(self.LIST_TEXT_PREFIXES):
            return True
        if self._has_hanging_list_indent(paragraph):
            return True
        return False

    def _has_numbering(self, paragraph: Paragraph) -> bool:
        paragraph_properties = paragraph._p.pPr
        if paragraph_properties is not None and paragraph_properties.numPr is not None:
            return True

        style = paragraph.style
        style_element = getattr(style, "_element", None)
        if style_element is None:
            return False

        paragraph_properties = style_element.find(qn("w:pPr"))
        if paragraph_properties is None:
            return False
        numbering = paragraph_properties.find(qn("w:numPr"))
        return numbering is not None

    def _has_hanging_list_indent(self, paragraph: Paragraph) -> bool:
        paragraph_format = paragraph.paragraph_format
        left_indent = paragraph_format.left_indent
        first_line_indent = paragraph_format.first_line_indent
        if left_indent is None or first_line_indent is None:
            return False
        try:
            return left_indent.pt >= 12 and first_line_indent.pt <= -6
        except AttributeError:
            return False

    def _blocks_to_text(self, blocks: list[DocumentBlock]) -> str:
        lines: list[str] = []
        for block in blocks:
            if block.kind == "title" and block.text:
                lines.append(block.text)
                lines.append("")
                continue
            if block.kind in {"heading", "subheading"} and block.text:
                lines.append(block.text)
                lines.append("")
                continue
            if block.kind == "list":
                for item in block.items:
                    lines.append(f"- {item}")
                lines.append("")
                continue
            if block.kind == "table" and block.table:
                if block.table.headers:
                    lines.append(" | ".join(block.table.headers))
                for row in block.table.rows:
                    lines.append(" | ".join(row))
                lines.append("")
                continue
            if block.kind == "image":
                caption = block.text or "Изображение"
                lines.append(f"[Image] {caption}")
                lines.append("")
                continue
            if block.text:
                lines.append(block.text)
                lines.append("")
        return "\n".join(lines).strip()

    def _extract_hyperlinks(self, paragraph: Paragraph) -> list[str]:
        urls: list[str] = []
        seen: set[str] = set()
        for hyperlink in paragraph._p.xpath('.//*[local-name()="hyperlink"]'):
            relationship_id = hyperlink.get(qn("r:id"))
            if not relationship_id:
                continue
            relationship = paragraph.part.rels.get(relationship_id)
            target = getattr(relationship, "target_ref", None)
            if target and target not in seen:
                urls.append(target)
                seen.add(target)
        return urls

    def _extract_inline_images(self, paragraph: Paragraph) -> list[DocumentBlock]:
        blocks: list[DocumentBlock] = []
        for index, drawing in enumerate(paragraph._p.xpath('.//*[local-name()="drawing"]'), start=1):
            blips = drawing.xpath('.//*[local-name()="blip"]')
            if not blips:
                continue
            embed_id = blips[0].get(qn("r:embed"))
            if not embed_id:
                continue
            image_part = paragraph.part.related_parts.get(embed_id)
            if image_part is None:
                continue
            filename = Path(getattr(image_part, "partname", f"image_{index}").__str__()).name
            blocks.append(
                DocumentBlock(
                    kind="image",
                    text=(paragraph.text or "").strip() or filename,
                    image_name=filename,
                    image_content_type=getattr(image_part, "content_type", None),
                    image_base64=base64.b64encode(image_part.blob).decode("ascii"),
                    run_count=len(paragraph.runs),
                )
            )
        return blocks

    def _merge_narrative_paragraph_blocks(self, blocks: list[DocumentBlock]) -> list[DocumentBlock]:
        if not blocks:
            return blocks

        merged: list[DocumentBlock] = []
        for block in blocks:
            if (
                merged
                and self._should_merge_paragraph_blocks(merged[-1], block)
            ):
                previous = merged[-1]
                merged[-1] = DocumentBlock(
                    kind=previous.kind,
                    text=self._merge_paragraph_text(previous.text or "", block.text or ""),
                    level=previous.level,
                    style_name=previous.style_name,
                    style_id=previous.style_id,
                    items=previous.items.copy(),
                    table=previous.table,
                    hyperlinks=[*previous.hyperlinks, *[link for link in block.hyperlinks if link not in previous.hyperlinks]],
                    run_count=(previous.run_count or 0) + (block.run_count or 0),
                    image_name=previous.image_name,
                    image_content_type=previous.image_content_type,
                    image_base64=previous.image_base64,
                )
                continue
            merged.append(block)
        return merged

    def _should_merge_paragraph_blocks(self, previous: DocumentBlock, current: DocumentBlock) -> bool:
        if previous.kind != "paragraph" or current.kind != "paragraph":
            return False
        previous_text = (previous.text or "").strip()
        current_text = (current.text or "").strip()
        if not previous_text or not current_text:
            return False
        if previous.table is not None or current.table is not None:
            return False
        if previous.image_base64 or current.image_base64:
            return False
        if previous.hyperlinks or current.hyperlinks:
            return False
        if self._looks_like_structured_label(previous_text) or self._looks_like_structured_label(current_text):
            return False
        if self._looks_like_outline_heading(previous_text) or self._looks_like_outline_heading(current_text):
            return False
        if self._looks_like_list_text(previous_text) or self._looks_like_list_text(current_text):
            return False
        if self._looks_like_reference_line(previous_text) or self._looks_like_reference_line(current_text):
            return False
        if (previous.style_name or "").strip().lower() not in {"", "normal"}:
            return False
        if (current.style_name or "").strip().lower() not in {"", "normal"}:
            return False

        normalized_previous = re.sub(r"\s+", " ", previous_text)
        normalized_current = re.sub(r"\s+", " ", current_text)
        if len(normalized_previous) > 220 or len(normalized_current) > 220:
            return False
        if normalized_previous.endswith(self.TERMINAL_PUNCTUATION):
            return False
        if normalized_previous.endswith(":"):
            return True
        if normalized_previous.endswith(("—", "–", "-", ",")):
            return True
        if self.CONTINUATION_START_PATTERN.match(normalized_current):
            return True
        return False

    def _merge_paragraph_text(self, previous: str, current: str) -> str:
        left = re.sub(r"\s+", " ", previous.strip())
        right = re.sub(r"\s+", " ", current.strip())
        if not left:
            return right
        if not right:
            return left
        return f"{left} {right}".strip()

    def _looks_like_structured_label(self, text: str) -> bool:
        normalized = re.sub(r"\s+", " ", text.strip())
        if not normalized:
            return False
        if len(normalized) > 140:
            return False
        match = self.STRUCTURED_LABEL_PATTERN.match(normalized)
        if not match:
            return False
        label = normalized.split(":", 1)[0].strip()
        if len(label) > 48:
            return False
        if len(label.split()) > 6:
            return False
        if any(char in label for char in ".?!;"):
            return False
        return True

    def _looks_like_outline_heading(self, text: str) -> bool:
        normalized = re.sub(r"\s+", " ", text.strip())
        return bool(self.OUTLINE_HEADING_PATTERN.match(normalized) or self.QUARTER_HEADING_PATTERN.match(normalized))

    def _looks_like_list_text(self, text: str) -> bool:
        normalized = text.strip()
        return normalized.startswith(self.LIST_TEXT_PREFIXES)

    def _looks_like_reference_line(self, text: str) -> bool:
        normalized = re.sub(r"\s+", " ", text.strip())
        if not normalized:
            return False
        return bool(self.URL_ONLY_PATTERN.match(normalized) or self.REFERENCE_LINE_PATTERN.match(normalized))

    def _table_cell_fill_color(self, cell) -> str | None:
        tc_pr = getattr(cell._tc, "tcPr", None)
        if tc_pr is None:
            return None

        shading = tc_pr.find(qn("w:shd"))
        if shading is None:
            return None

        fill = shading.get(qn("w:fill")) or shading.get("fill")
        if not fill:
            return None
        normalized = fill.strip().lstrip("#").upper()
        if normalized in {"AUTO", "NIL", "NONE", "FFFFFF"}:
            return None
        if re.fullmatch(r"[0-9A-F]{6}", normalized) is None:
            return None
        return normalized
