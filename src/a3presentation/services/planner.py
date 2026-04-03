from __future__ import annotations

import re
from dataclasses import dataclass, field

from a3presentation.domain.api import DocumentBlock
from a3presentation.domain.presentation import PresentationPlan, SlideKind, SlideSpec, TableBlock
from a3presentation.domain.semantic import DocumentKind, SemanticDocument, SemanticImage, SemanticSection
from a3presentation.services.semantic_normalizer import SemanticDocumentNormalizer


@dataclass
class Section:
    title: str
    level: int = 1
    subtitle: str | None = None
    paragraphs: list[str] = field(default_factory=list)
    bullet_lists: list[list[str]] = field(default_factory=list)
    tables: list[TableBlock] = field(default_factory=list)
    images: list[SemanticImage] = field(default_factory=list)


class TextToPlanService:
    EMAIL_PATTERN = re.compile(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}")
    PHONE_PATTERN = re.compile(r"(\+?\d[\d\s().-]{7,}\d)")
    HEADING_PATTERN = re.compile(r"^(\d+(\.\d+)*[.)]?\s+.+|[А-ЯA-Z].{0,90})$")
    CARD_BULLET_MAX_CHARS = 100
    CARD_BULLET_COUNT = 3
    LIST_BATCH_SIZE = 5
    LIST_SLIDE_MAX_WEIGHT = 9.0
    LIST_BULLET_MAX_CHARS = 220
    TEXT_SLIDE_MAX_WEIGHT = 7.5
    TEXT_SLIDE_MAX_CHARS = 420
    TEXT_PRIMARY_MAX_CHARS = 320
    TEXT_TAIL_MERGE_THRESHOLD = 120
    COVER_META_MAX_LINES = 2
    COVER_META_MAX_LINE_CHARS = 72
    STRUCTURED_SUMMARY_MAX_ITEMS = 5
    RESUME_SUMMARY_MAX_ITEMS = 6

    def __init__(self) -> None:
        self.normalizer = SemanticDocumentNormalizer()

    def build_plan(
        self,
        template_id: str,
        raw_text: str,
        title: str | None = None,
        tables: list[TableBlock] | None = None,
        blocks: list[DocumentBlock] | None = None,
    ) -> PresentationPlan:
        semantic_document = self.normalizer.normalize(
            raw_text=raw_text,
            blocks=blocks or [],
            tables=tables or [],
            title=title,
        )
        sections = [self._section_from_semantic(section) for section in semantic_document.sections]
        plan_title = semantic_document.title
        document_kind = semantic_document.kind.value

        cover_title, cover_meta = self._build_cover(plan_title, sections, blocks or [])
        slides: list[SlideSpec] = [
            SlideSpec(
                kind=SlideKind.TITLE,
                title=cover_title,
                subtitle="",
                notes=cover_meta,
                preferred_layout_key="cover",
            )
        ]

        content_slides: list[SlideSpec] = []
        contact_slides: list[SlideSpec] = []
        blocks_have_tables = any(block.kind == "table" and block.table is not None for block in blocks or [])

        if document_kind == DocumentKind.RESUME.value:
            content_slides = self._build_resume_fallback_slides(
                plan_title=plan_title,
                semantic_document=semantic_document,
            )
        elif document_kind in {DocumentKind.FORM.value, DocumentKind.TABLE_HEAVY.value, DocumentKind.UNKNOWN.value}:
            content_slides = self._build_safe_fallback_slides(
                plan_title=plan_title,
                blocks=blocks or [],
                sections=sections,
                loose_tables=tables or [],
                semantic_document=semantic_document,
            )
        else:
            for index, section in enumerate(sections):
                if index == 0 and self._is_cover_section(section, cover_title):
                    continue
                for slide in self._section_to_slides(section):
                    if slide.preferred_layout_key == "contacts":
                        contact_slides.append(slide)
                    else:
                        content_slides.append(slide)

        content_slides = self._compress_slides(content_slides)
        slides.extend(content_slides)

        if not blocks_have_tables:
            table_title_base = sections[-1].title if sections else plan_title
            for index, table in enumerate(tables or [], start=1):
                if not table.headers and not table.rows:
                    continue
                slides.append(
                    SlideSpec(
                        kind=SlideKind.TABLE,
                        title=f"{table_title_base} {index}",
                        subtitle="Ключевые данные из документа",
                        table=table,
                        preferred_layout_key="table",
                    )
                )

        slides.extend(contact_slides[:1])
        if len(slides) == 1 and (blocks or tables):
            slides.extend(
                self._build_safe_fallback_slides(
                    plan_title=plan_title,
                    blocks=blocks or [],
                    sections=sections,
                    loose_tables=tables or [],
                    semantic_document=semantic_document,
                )
            )
        slides = self._enforce_hard_safety_rules(slides, plan_title, semantic_document, blocks or [], tables or [], sections)
        return PresentationPlan(template_id=template_id, title=plan_title, slides=slides)

    def _section_from_semantic(self, section: SemanticSection) -> Section:
        return Section(
            title=section.title,
            level=section.level,
            subtitle=section.subtitle,
            paragraphs=section.paragraphs.copy(),
            bullet_lists=[section.bullets.copy()] if section.bullets else [],
            tables=section.tables.copy(),
            images=section.images.copy(),
        )

    def _classify_document(
        self,
        blocks: list[DocumentBlock],
        sections: list[Section],
        raw_text: str,
        tables: list[TableBlock],
    ) -> str:
        heading_count = sum(1 for block in blocks if block.kind in {"heading", "subheading"})
        paragraphs = [self._normalize_line(block.text or "") for block in blocks if block.kind == "paragraph" and (block.text or "").strip()]
        table_count = len([block for block in blocks if block.kind == "table" and block.table is not None]) or len(tables)
        list_count = sum(1 for block in blocks if block.kind == "list" and block.items)
        short_labels = sum(1 for paragraph in paragraphs if self._looks_like_structured_label(paragraph))
        long_paragraphs = sum(1 for paragraph in paragraphs if len(paragraph) >= 180)
        section_count = len(sections)
        resume_markers = sum(1 for paragraph in paragraphs if self._looks_like_resume_label(paragraph))
        contact_markers = sum(
            1
            for paragraph in paragraphs
            if self.EMAIL_PATTERN.search(paragraph) or self.PHONE_PATTERN.search(paragraph)
        )

        if table_count >= 3 and short_labels >= 2:
            return "form"
        if resume_markers >= 3 and contact_markers >= 1 and table_count <= 1:
            return "resume"
        if table_count >= 3 and heading_count <= 1 and section_count <= 2:
            return "table_heavy"
        if heading_count >= 2 and (long_paragraphs >= 2 or list_count >= 1):
            return "report"
        if table_count == 0 and heading_count >= 2 and len(raw_text.strip()) >= 600:
            return "report"
        if short_labels >= 4 and long_paragraphs <= 2:
            return "form"
        if section_count <= 1 and table_count >= 1:
            return "unknown"
        return "mixed"

    def _build_cover(self, plan_title: str, sections: list[Section], blocks: list[DocumentBlock]) -> tuple[str, str]:
        explicit_title = next((block.text for block in blocks if block.kind == "title" and block.text), None)
        leading_lines = self._leading_cover_lines(blocks)
        if not sections:
            if leading_lines:
                cover_title, meta_parts = self._cover_title_and_meta_from_lines(leading_lines)
                cover_meta = self._compact_cover_meta(meta_parts, fallback=(plan_title if not explicit_title else explicit_title))
                return cover_title[:120], cover_meta[:320]
            return (explicit_title or plan_title), (plan_title if not explicit_title else explicit_title)

        if leading_lines:
            cover_title, meta_parts = self._cover_title_and_meta_from_lines(leading_lines)
            cover_meta = self._compact_cover_meta(meta_parts, fallback=plan_title)
            return cover_title[:120], cover_meta[:320]

        first = sections[0]
        cover_title = explicit_title or first.title or plan_title
        meta_parts: list[str] = []
        if first.subtitle and first.subtitle != cover_title and self._is_cover_meta_candidate(first.subtitle):
            meta_parts.append(first.subtitle)

        cover_meta = self._compact_cover_meta(meta_parts, fallback="")
        return cover_title[:120], cover_meta[:320]

    def _is_cover_section(self, section: Section, cover_title: str) -> bool:
        if section.tables:
            return False
        if len(section.paragraphs) > 2:
            return False
        if sum(len(items) for items in section.bullet_lists) > 3:
            return False
        if section.title == cover_title:
            return True
        total_items = len(section.paragraphs) + sum(len(items) for items in section.bullet_lists)
        return section.level <= 1 and total_items <= 5

    def _detect_title(self, blocks: list[DocumentBlock], raw_text: str, sections: list[Section]) -> str:
        for block in blocks:
            if block.kind == "title" and block.text:
                return block.text[:120]
        leading_lines = self._leading_cover_lines(blocks)
        if leading_lines:
            cover_title, _ = self._cover_title_and_meta_from_lines(leading_lines)
            if cover_title:
                return cover_title[:120]
        for block in blocks:
            if block.kind in {"heading", "subheading", "paragraph"} and block.text:
                return block.text[:120]

        lines = [line.strip() for line in raw_text.replace("\r", "").split("\n") if line.strip()]
        if lines:
            return lines[0][:120]
        if sections:
            return sections[0].title[:120]
        return "Generated Presentation"

    def _build_sections(self, blocks: list[DocumentBlock], raw_text: str) -> list[Section]:
        if blocks:
            sections = self._build_sections_from_blocks(self._trim_leading_cover_blocks(blocks))
            if sections:
                return sections
        return self._build_sections_from_text(raw_text)

    def _build_sections_from_blocks(self, blocks: list[DocumentBlock]) -> list[Section]:
        sections: list[Section] = []
        current: Section | None = None

        def flush() -> None:
            nonlocal current
            if current is None:
                return
            if not current.paragraphs and not current.bullet_lists and not current.tables:
                if current.level <= 1 and current.subtitle:
                    current.paragraphs.append(current.subtitle[:1200])
                else:
                    current = None
                    return
            if not current.subtitle and current.paragraphs:
                current.subtitle = current.paragraphs[0][:120]
            sections.append(current)
            current = None

        for block in blocks:
            if block.kind == "title":
                continue

            if block.kind in {"heading", "subheading"} and block.text:
                block_level = block.level or (1 if block.kind == "heading" else 2)
                if current is not None and block_level <= current.level:
                    flush()
                if current is None:
                    current = Section(title=block.text[:120], level=block_level)
                else:
                    if not current.subtitle:
                        current.subtitle = block.text[:120]
                    else:
                        flush()
                        current = Section(title=block.text[:120], level=block_level)
                continue

            if current is None:
                if not block.text and not block.items:
                    continue
                seed = block.text or (block.items[0] if block.items else "Раздел")
                current = Section(title=seed[:120], level=1)

            if block.kind == "table" and block.table is not None:
                current.tables.append(block.table)
                continue

            if block.kind == "image" and block.image_base64:
                current.images.append(
                    SemanticImage(
                        name=block.image_name,
                        alt_text=block.text,
                        content_type=block.image_content_type,
                        image_base64=block.image_base64,
                    )
                )
                continue

            if block.kind == "list" and block.items:
                current.bullet_lists.append([item.strip()[:220] for item in block.items if item.strip()])
                continue

            if block.text:
                text = block.text.strip()
                if text:
                    current.paragraphs.append(text[:1200])

        flush()
        return [section for section in sections if section.title]

    def _build_sections_from_text(self, raw_text: str) -> list[Section]:
        lines = [self._normalize_line(line) for line in raw_text.replace("\r", "").split("\n") if self._normalize_line(line)]
        if not lines:
            return []

        sections: list[Section] = []
        current: Section | None = None

        def flush() -> None:
            nonlocal current
            if current is None:
                return
            if not current.paragraphs and not current.bullet_lists:
                if current.level <= 1 and current.subtitle:
                    current.paragraphs.append(current.subtitle[:1200])
                else:
                    current = None
                    return
            if not current.subtitle and current.paragraphs:
                current.subtitle = current.paragraphs[0][:120]
            sections.append(current)
            current = None

        for line in lines:
            if self._is_heading(line):
                flush()
                current = Section(title=line[:120], level=1)
                continue

            if current is None:
                current = Section(title=line[:120], level=1)
                continue

            normalized_bullet = re.sub(r"^\s*(?:[-*•]|\d+[.)])\s*", "", line).strip()
            if normalized_bullet != line or self._looks_like_short_item(line):
                if not current.bullet_lists:
                    current.bullet_lists.append([])
                current.bullet_lists[-1].append(normalized_bullet[:220])
            else:
                current.paragraphs.append(line[:1200])

        flush()
        return sections

    def _section_to_slides(self, section: Section) -> list[SlideSpec]:
        section_lines = [*section.paragraphs, *[item for bullet_list in section.bullet_lists for item in bullet_list]]
        slides: list[SlideSpec] = []

        if self._looks_like_contacts(section.title, section_lines):
            slides.append(self._build_contact_slide(section))
            slides.extend(self._build_table_slides(section))
            return slides

        if section.paragraphs or section.bullet_lists:
            if self._fits_single_slide(section):
                slides.append(self._build_single_slide(section))
            else:
                slides.extend(self._split_large_section(section))

        slides.extend(self._build_table_slides(section))
        for index, image in enumerate(section.images, start=1):
            slide = self._image_slide(image, section.title, index)
            if slide is not None:
                slides.append(slide)
        return slides

    def _build_table_slides(self, section: Section) -> list[SlideSpec]:
        slides: list[SlideSpec] = []
        table_index = 0
        for table in section.tables:
            if not table.headers and not table.rows:
                continue
            chunks = self._split_table_for_slides(table)
            for chunk_index, chunk in enumerate(chunks, start=1):
                table_index += 1
                title_base = section.title or "Таблица"
                use_suffix = len(section.tables) > 1 or len(chunks) > 1
                title = title_base if not use_suffix else f"{title_base} ({table_index if len(section.tables) > 1 else chunk_index})"
                slides.append(
                    SlideSpec(
                        kind=SlideKind.TABLE,
                        title=title[:120],
                        subtitle=(section.subtitle or "")[:120],
                        table=chunk,
                        preferred_layout_key="table",
                    )
                )
        return slides

    def _split_table_for_slides(self, table: TableBlock) -> list[TableBlock]:
        if not table.rows:
            return [table]

        col_count = len(table.headers) if table.headers else max((len(row) for row in table.rows), default=0)
        row_count = len(table.rows)
        max_cell_length = max((len(cell or "") for row in table.rows for cell in row), default=0)
        avg_row_length = (
            sum(sum(len(cell or "") for cell in row) for row in table.rows) / max(1, row_count)
            if row_count
            else 0.0
        )

        if col_count == 2 and row_count <= 6:
            return [table]
        if col_count == 3 and row_count <= 5:
            return [table]
        if col_count == 3 and row_count <= 7 and max_cell_length <= 70 and avg_row_length <= 85:
            return [table]
        if col_count == 3 and row_count <= 9 and max_cell_length <= 130 and avg_row_length <= 120:
            return [table]

        base_capacity = 8
        if col_count >= 3:
            base_capacity = 6
        if col_count >= 4:
            base_capacity = 5

        if max_cell_length >= 120:
            base_capacity -= 2
        elif max_cell_length >= 80:
            base_capacity -= 1

        base_capacity = max(base_capacity, 1)
        chunks: list[list[list[str]]] = []
        current_chunk: list[list[str]] = []
        current_weight = 0.0

        for row in table.rows:
            row_weight = self._estimate_table_row_weight(row, col_count)
            if current_chunk and current_weight + row_weight > base_capacity:
                chunks.append(current_chunk)
                current_chunk = []
                current_weight = 0.0
            current_chunk.append(row)
            current_weight += row_weight

        if current_chunk:
            chunks.append(current_chunk)

        if len(chunks) >= 2 and len(chunks[-1]) == 1:
            trailing_row = chunks[-1][0]
            trailing_weight = self._estimate_table_row_weight(trailing_row, col_count)
            previous_weight = sum(self._estimate_table_row_weight(row, col_count) for row in chunks[-2])
            previous_row_count = len(chunks[-2])
            if previous_weight + trailing_weight <= base_capacity * 1.55 or previous_row_count <= max(base_capacity - 1, 1):
                chunks[-2].append(trailing_row)
                chunks.pop()

        return [TableBlock(headers=table.headers, rows=chunk) for chunk in chunks]

    def _estimate_table_row_weight(self, row: list[str], col_count: int) -> float:
        max_length = max((len(cell or "") for cell in row), default=0)
        avg_length = sum(len(cell or "") for cell in row) / max(1, len(row))

        weight = 1.0
        if col_count >= 3:
            weight += 0.2
        if avg_length >= 40:
            weight += 0.35
        if avg_length >= 70:
            weight += 0.35
        if max_length >= 120:
            weight += 0.6
        if max_length >= 180:
            weight += 0.8
        return weight

    def _fits_single_slide(self, section: Section) -> bool:
        paragraph_chars = sum(len(item) for item in section.paragraphs)
        bullet_count = sum(len(items) for items in section.bullet_lists)
        max_bullet_len = max((len(item) for items in section.bullet_lists for item in items), default=0)

        if paragraph_chars <= self.TEXT_SLIDE_MAX_CHARS and bullet_count == 0:
            return True
        if (
            bullet_count <= self.CARD_BULLET_COUNT
            and paragraph_chars <= 160
            and max_bullet_len <= self.CARD_BULLET_MAX_CHARS
        ):
            return True
        if bullet_count <= self.LIST_BATCH_SIZE and paragraph_chars <= 260:
            return True
        return False

    def _build_single_slide(self, section: Section) -> SlideSpec:
        bullets = [item for bullet_list in section.bullet_lists for item in bullet_list]
        if bullets and self._should_use_cards_layout(section.title, bullets):
            return SlideSpec(
                kind=SlideKind.BULLETS,
                title=section.title,
                bullets=bullets,
                preferred_layout_key="cards_3",
            )

        if bullets:
            return SlideSpec(
                kind=SlideKind.BULLETS,
                title=section.title,
                bullets=bullets[: self.LIST_BATCH_SIZE],
                preferred_layout_key="list_full_width",
            )

        text = " ".join(section.paragraphs).strip()
        if len(text) <= self.TEXT_SLIDE_MAX_CHARS:
            primary_text, secondary_text = self._split_text_for_slide(text)
            subtitle = self._normalize_subtitle(section.subtitle or "", primary_text)
            return SlideSpec(
                kind=SlideKind.TEXT,
                title=section.title,
                subtitle=subtitle,
                text=primary_text,
                notes=secondary_text,
                preferred_layout_key="text_full_width",
            )

        sentences = self._sentence_chunks(text)
        return SlideSpec(
            kind=SlideKind.BULLETS,
            title=section.title,
            bullets=sentences[: self.LIST_BATCH_SIZE],
            preferred_layout_key="list_full_width",
        )

    def _split_large_section(self, section: Section) -> list[SlideSpec]:
        slides: list[SlideSpec] = []
        bullets = [item for bullet_list in section.bullet_lists for item in bullet_list]
        if bullets:
            batches = self._chunk_bullets_for_slides(bullets)
            for index, batch in enumerate(batches):
                slide_title = section.title if index == 0 else f"{section.title} ({index + 1})"
                slides.append(
                    SlideSpec(
                        kind=SlideKind.BULLETS,
                        title=slide_title,
                        bullets=batch,
                        preferred_layout_key="list_full_width",
                    )
                )
            return slides

        text = " ".join(section.paragraphs).strip()
        sentences = self._sentence_chunks(text)
        batches = self._chunk_text_for_slides(sentences)
        for index, batch in enumerate(batches):
            slide_title = section.title if index == 0 else f"{section.title} ({index + 1})"
            if len(batch) <= 2 and len(" ".join(batch)) <= self.TEXT_SLIDE_MAX_CHARS:
                merged = " ".join(batch)
                primary_text, secondary_text = self._split_text_for_slide(merged)
                subtitle = self._normalize_subtitle(section.subtitle or "", primary_text)
                slides.append(
                    SlideSpec(
                        kind=SlideKind.TEXT,
                        title=slide_title,
                        subtitle=subtitle,
                        text=primary_text,
                        notes=secondary_text,
                        preferred_layout_key="text_full_width",
                    )
                )
            else:
                slides.append(
                    SlideSpec(
                        kind=SlideKind.BULLETS,
                        title=slide_title,
                        bullets=batch[: self.LIST_BATCH_SIZE],
                        preferred_layout_key="list_full_width",
                    )
                )
        return slides

    def _chunk_bullets_for_slides(self, bullets: list[str]) -> list[list[str]]:
        batches: list[list[str]] = []
        current_batch: list[str] = []
        current_weight = 0.0

        for bullet in bullets:
            bullet_weight = self._estimate_bullet_weight(bullet)
            if current_batch and (
                len(current_batch) >= self.LIST_BATCH_SIZE or current_weight + bullet_weight > self.LIST_SLIDE_MAX_WEIGHT
            ):
                batches.append(current_batch)
                current_batch = []
                current_weight = 0.0

            current_batch.append(bullet)
            current_weight += bullet_weight

        if current_batch:
            batches.append(current_batch)

        return batches

    def _chunk_text_for_slides(self, chunks: list[str]) -> list[list[str]]:
        batches: list[list[str]] = []
        current_batch: list[str] = []
        current_weight = 0.0

        for chunk in chunks:
            chunk_weight = self._estimate_text_chunk_weight(chunk)
            if current_batch and (
                len(current_batch) >= self.LIST_BATCH_SIZE
                or current_weight + chunk_weight > self.TEXT_SLIDE_MAX_WEIGHT
            ):
                batches.append(current_batch)
                current_batch = []
                current_weight = 0.0

            current_batch.append(chunk)
            current_weight += chunk_weight

        if current_batch:
            batches.append(current_batch)

        return batches

    def _estimate_text_chunk_weight(self, chunk: str) -> float:
        normalized = (chunk or "").strip()
        length = len(normalized)
        weight = 1.0
        if length >= 120:
            weight += 0.6
        if length >= 200:
            weight += 0.9
        if length >= 280:
            weight += 1.2
        if normalized.count(",") >= 2:
            weight += 0.2
        if any(marker in normalized for marker in (";", ":", "—")):
            weight += 0.25
        return weight

    def _estimate_bullet_weight(self, bullet: str) -> float:
        length = len((bullet or "").strip())
        weight = 1.0
        if length >= 80:
            weight += 0.8
        if length >= 140:
            weight += 1.0
        if length >= 220:
            weight += 1.2
        if ":" in bullet:
            weight += 0.3
        if ";" in bullet:
            weight += 0.2
        return weight

    def _split_text_for_slide(self, text: str) -> tuple[str, str]:
        normalized = text.strip()
        if len(normalized) <= self.TEXT_PRIMARY_MAX_CHARS:
            return normalized, ""

        split_at = normalized.rfind(". ", 0, self.TEXT_PRIMARY_MAX_CHARS)
        if split_at == -1:
            split_at = normalized.rfind("; ", 0, self.TEXT_PRIMARY_MAX_CHARS)
        if split_at == -1:
            split_at = normalized.rfind(", ", 0, self.TEXT_PRIMARY_MAX_CHARS)
        if split_at == -1 or split_at < int(self.TEXT_PRIMARY_MAX_CHARS * 0.55):
            split_at = self.TEXT_PRIMARY_MAX_CHARS
        else:
            split_at += 1

        primary = normalized[:split_at].strip()
        secondary = normalized[split_at:].strip()

        if len(secondary) <= self.TEXT_TAIL_MERGE_THRESHOLD:
            return normalized, ""

        return primary, secondary

    def _normalize_subtitle(self, subtitle: str, main_text: str) -> str:
        normalized_subtitle = subtitle.strip()
        normalized_text = main_text.strip()
        if not normalized_subtitle:
            return ""
        if normalized_text.startswith(normalized_subtitle):
            return ""
        if len(normalized_subtitle) >= 24 and normalized_text.startswith(normalized_subtitle[:-1]):
            return ""
        return normalized_subtitle[:120]

    def _should_use_cards_layout(self, title: str, bullets: list[str]) -> bool:
        if not bullets or len(bullets) > self.CARD_BULLET_COUNT:
            return False
        normalized_title = title.strip().lower()
        if re.match(r"^(q\d+|question\b|\d+(\.\d+)*)", normalized_title):
            return False
        if any(len(item) > 55 for item in bullets):
            return False
        if any(any(marker in item for marker in (":", "—", ";", ".")) for item in bullets):
            return False
        return all(len(item) <= self.CARD_BULLET_MAX_CHARS for item in bullets)

    def _leading_cover_lines(self, blocks: list[DocumentBlock]) -> list[str]:
        lines: list[str] = []
        for block in blocks:
            if block.kind in {"heading", "subheading", "table", "list"}:
                break
            if block.kind not in {"title", "subtitle", "paragraph"}:
                continue
            text = (block.text or "").strip()
            if text:
                lines.append(text)
        return lines[:6]

    def _trim_leading_cover_blocks(self, blocks: list[DocumentBlock]) -> list[DocumentBlock]:
        first_structured_index = next(
            (index for index, block in enumerate(blocks) if block.kind in {"heading", "subheading"}),
            None,
        )
        if first_structured_index is None or first_structured_index <= 0:
            return blocks

        leading_blocks = blocks[:first_structured_index]
        if len(leading_blocks) > 4:
            return blocks
        if any(block.kind in {"list", "table", "image"} for block in leading_blocks):
            return blocks

        candidate_lines = [
            (block.text or "").strip()
            for block in leading_blocks
            if block.kind in {"title", "subtitle", "paragraph"} and (block.text or "").strip()
        ]
        if not candidate_lines:
            return blocks
        if any(len(line) > 140 for line in candidate_lines):
            return blocks

        return blocks[first_structured_index:]

    def _cover_title_and_meta_from_lines(self, lines: list[str]) -> tuple[str, list[str]]:
        normalized = [line.strip() for line in lines if line.strip()]
        if not normalized:
            return "", []
        if len(normalized) >= 2 and self._looks_like_cover_prefix(normalized[0]):
            title = f"{normalized[0]} {normalized[1]}".strip()
            meta = normalized[2:]
            return title, meta
        return normalized[0], normalized[1:]

    def _looks_like_cover_prefix(self, text: str) -> bool:
        compact = text.strip()
        if not compact:
            return False
        if len(compact) > 8:
            return False
        if compact[0].isdigit():
            return False
        return len(compact.split()) <= 2

    def _compact_cover_meta(self, parts: list[str], fallback: str) -> str:
        compact_parts: list[str] = []
        seen: set[str] = set()

        for part in parts:
            normalized = self._normalize_line(part)
            if not normalized or normalized in seen:
                continue
            if not self._is_cover_meta_candidate(normalized):
                continue
            compact_parts.append(normalized[: self.COVER_META_MAX_LINE_CHARS])
            seen.add(normalized)
            if len(compact_parts) >= self.COVER_META_MAX_LINES:
                break

        if compact_parts:
            return "\n".join(compact_parts)

        fallback_line = self._normalize_line(fallback)[: self.COVER_META_MAX_LINE_CHARS]
        return fallback_line

    def _is_cover_meta_candidate(self, text: str) -> bool:
        normalized = self._normalize_line(text)
        if not normalized:
            return False
        if len(normalized) > self.COVER_META_MAX_LINE_CHARS:
            return False
        if len(normalized.split()) > 10:
            return False
        sentence_like_punctuation = normalized.count(".") + normalized.count(";") + normalized.count(":")
        if sentence_like_punctuation >= 2:
            return False
        return True

    def _build_contact_slide(self, section: Section) -> SlideSpec:
        all_lines = [section.title, *(section.paragraphs or []), *[item for group in section.bullet_lists for item in group]]
        return SlideSpec(
            kind=SlideKind.TEXT,
            title=section.title,
            subtitle=section.subtitle or (section.paragraphs[0] if section.paragraphs else ""),
            text="\n".join(section.paragraphs[1:3]) if len(section.paragraphs) > 1 else "",
            left_bullets=[self._first_phone(all_lines)],
            right_bullets=[self._first_email(all_lines)],
            preferred_layout_key="contacts",
        )

    def _build_safe_fallback_slides(
        self,
        plan_title: str,
        blocks: list[DocumentBlock],
        sections: list[Section],
        loose_tables: list[TableBlock],
        semantic_document: SemanticDocument,
    ) -> list[SlideSpec]:
        slides: list[SlideSpec] = []
        if semantic_document.kind == DocumentKind.TABLE_HEAVY:
            summary_items = self._table_heavy_summary_items(blocks, sections, semantic_document)
        else:
            summary_items = self._structured_summary_items(blocks, sections, semantic_document)
        if summary_items:
            slides.append(
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title=f"{plan_title} - обзор",
                    bullets=summary_items[: self.LIST_BATCH_SIZE],
                    preferred_layout_key="list_full_width",
                )
            )

        narrative_entries = self._structured_text_entries(blocks, semantic_document)
        for title, text in narrative_entries:
            primary_text, secondary_text = self._split_text_for_slide(text)
            slides.append(
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title=title[:120],
                    text=primary_text,
                    notes=secondary_text,
                    preferred_layout_key="text_full_width",
                )
            )

        table_slides = self._structured_table_slides(blocks, semantic_document)
        if not table_slides:
            for index, table in enumerate(loose_tables, start=1):
                if not table.headers and not table.rows:
                    continue
                table_slides.append(
                    SlideSpec(
                        kind=SlideKind.TABLE,
                        title=f"{plan_title} ({index})",
                        subtitle="Ключевые данные из документа",
                        table=table,
                        preferred_layout_key="table",
                    )
                )

        slides.extend(table_slides)
        slides.extend(self._image_slides_from_semantic(semantic_document))

        if not slides:
            fallback_text = raw_text = "\n".join(
                self._normalize_line(block.text or "") for block in blocks if (block.text or "").strip()
            ).strip()
            if fallback_text:
                primary_text, secondary_text = self._split_text_for_slide(fallback_text[:2400])
                slides.append(
                    SlideSpec(
                        kind=SlideKind.TEXT,
                        title=plan_title,
                        text=primary_text,
                        notes=secondary_text,
                        preferred_layout_key="text_full_width",
                    )
                )

        return slides

    def _build_resume_fallback_slides(
        self,
        plan_title: str,
        semantic_document: SemanticDocument,
    ) -> list[SlideSpec]:
        slides: list[SlideSpec] = []
        summary_items = self._resume_summary_items(semantic_document)
        if summary_items:
            slides.append(
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title=f"{plan_title} - профиль",
                    bullets=summary_items[: self.RESUME_SUMMARY_MAX_ITEMS],
                    preferred_layout_key="list_full_width",
                )
            )

        current_title = "Опыт и квалификация"
        seen_texts: set[str] = set()
        for section in semantic_document.sections:
            if self._looks_like_resume_label(section.title):
                current_title = section.title[:120]
            text = " ".join(section.paragraphs or [])
            if len(text) < 60 or text in seen_texts:
                continue
            seen_texts.add(text)
            primary_text, secondary_text = self._split_text_for_slide(text[:1800])
            slides.append(
                SlideSpec(
                    kind=SlideKind.TEXT,
                    title=current_title,
                    text=primary_text,
                    notes=secondary_text,
                    preferred_layout_key="text_full_width",
                )
            )
            if len(slides) >= 4:
                break

        slides.extend(self._image_slides_from_semantic(semantic_document))

        return slides

    def _structured_summary_items(self, blocks: list[DocumentBlock], sections: list[Section], semantic_document: SemanticDocument) -> list[str]:
        items: list[str] = []
        seen: set[str] = set()

        for section in sections:
            candidate = self._normalize_line(section.title)
            if candidate and candidate != self._normalize_line(semantic_document.title) and candidate not in seen:
                items.append(candidate[:120])
                seen.add(candidate)
            if len(items) >= self.STRUCTURED_SUMMARY_MAX_ITEMS:
                return items

        for fact in semantic_document.facts:
            line = f"{fact.label}: {fact.value}"[:120]
            if line not in seen:
                items.append(line)
                seen.add(line)
            if len(items) >= self.STRUCTURED_SUMMARY_MAX_ITEMS:
                return items

        for block in blocks:
            text = self._normalize_line(block.text or "")
            if not text or text in seen:
                continue
            if self._looks_like_structured_label(text):
                items.append(text[:120])
                seen.add(text)
            if len(items) >= self.STRUCTURED_SUMMARY_MAX_ITEMS:
                break

        return items

    def _table_heavy_summary_items(self, blocks: list[DocumentBlock], sections: list[Section], semantic_document: SemanticDocument) -> list[str]:
        items = self._structured_summary_items(blocks, sections, semantic_document)
        table_count = semantic_document.stats.table_count
        if table_count:
            items.append(f"Таблиц в документе: {table_count}")
        return items[: self.STRUCTURED_SUMMARY_MAX_ITEMS]

    def _resume_summary_items(self, semantic_document: SemanticDocument) -> list[str]:
        items: list[str] = []
        seen: set[str] = set()

        for contact in semantic_document.contacts:
            if contact not in seen:
                items.append(contact[:120])
                seen.add(contact)
            if len(items) >= self.RESUME_SUMMARY_MAX_ITEMS:
                break

        if not items:
            for section in semantic_document.sections[: self.RESUME_SUMMARY_MAX_ITEMS]:
                title = self._normalize_line(section.title)
                if title and title not in seen:
                    items.append(title[:120])
                    seen.add(title)
        return items

    def _structured_text_entries(self, blocks: list[DocumentBlock], semantic_document: SemanticDocument) -> list[tuple[str, str]]:
        entries: list[tuple[str, str]] = []
        pending_title = semantic_document.title

        for block in blocks:
            if block.kind == "table":
                continue
            if block.kind == "list" and block.items:
                items = [self._normalize_line(item) for item in block.items if self._normalize_line(item)]
                if items:
                    entries.append((pending_title[:120], "; ".join(items)[:1800]))
                continue

            text = self._normalize_line(block.text or "")
            if not text:
                continue

            if self._looks_like_structured_label(text):
                pending_title = text[:120]
                continue

            if len(text) < 100:
                continue

            entries.append((pending_title[:120], text[:1800]))

        deduped: list[tuple[str, str]] = []
        seen_texts: set[str] = set()
        for title, text in entries:
            if text in seen_texts:
                continue
            deduped.append((title, text))
            seen_texts.add(text)
            if len(deduped) >= 3:
                break

        if not deduped:
            for section in semantic_document.sections:
                text = " ".join(section.paragraphs or [])
                if len(text) < 80 or text in seen_texts:
                    continue
                deduped.append((section.title[:120], text[:1800]))
                seen_texts.add(text)
                if len(deduped) >= 3:
                    break
        return deduped

    def _structured_table_slides(self, blocks: list[DocumentBlock], semantic_document: SemanticDocument) -> list[SlideSpec]:
        slides: list[SlideSpec] = []
        pending_title = semantic_document.title
        table_index = 0

        for block in blocks:
            text = self._normalize_line(block.text or "")
            if text and self._looks_like_structured_label(text):
                pending_title = text[:120]
                continue

            if block.kind != "table" or block.table is None:
                continue

            table_index += 1
            chunks = self._split_table_for_slides(block.table)
            for chunk_index, chunk in enumerate(chunks, start=1):
                use_suffix = len(chunks) > 1 or table_index > 1
                title = pending_title if not use_suffix else f"{pending_title} ({table_index if len(chunks) == 1 else chunk_index})"
                slides.append(
                    SlideSpec(
                        kind=SlideKind.TABLE,
                        title=title[:120],
                        subtitle="Ключевые данные из документа",
                        table=chunk,
                        preferred_layout_key="table",
                    )
                )

        return slides

    def _image_slides_from_semantic(self, semantic_document: SemanticDocument) -> list[SlideSpec]:
        slides: list[SlideSpec] = []
        for index, image in enumerate(semantic_document.images, start=1):
            slide = self._image_slide(image, semantic_document.title, index)
            if slide is not None:
                slides.append(slide)
        return slides[:3]

    def _image_slide(self, image: SemanticImage, plan_title: str, index: int) -> SlideSpec | None:
        if not image.image_base64:
            return None
        caption = self._normalize_line(image.alt_text or image.name or f"Изображение {index}")
        body = caption if len(caption) <= 240 else caption[:240]
        return SlideSpec(
            kind=SlideKind.IMAGE,
            title=f"{plan_title} - иллюстрация {index}",
            text=body,
            preferred_layout_key="image_text",
            image_base64=image.image_base64,
            image_content_type=image.content_type,
        )

    def _enforce_hard_safety_rules(
        self,
        slides: list[SlideSpec],
        plan_title: str,
        semantic_document: SemanticDocument,
        blocks: list[DocumentBlock],
        tables: list[TableBlock],
        sections: list[Section],
    ) -> list[SlideSpec]:
        normalized_slides: list[SlideSpec] = []
        for index, slide in enumerate(slides):
            if slide.kind != SlideKind.TITLE and self._is_empty_slide(slide):
                continue
            if not (slide.title or "").strip():
                slide.title = plan_title if index == 0 else f"{plan_title} ({index})"
            if slide.kind == SlideKind.BULLETS and len(slide.bullets) > self.LIST_BATCH_SIZE:
                slide.bullets = slide.bullets[: self.LIST_BATCH_SIZE]
            normalized_slides.append(slide)

        if len(normalized_slides) == 1 and semantic_document.stats.character_count > 0:
            normalized_slides.extend(
                self._build_safe_fallback_slides(
                    plan_title=plan_title,
                    blocks=blocks,
                    sections=sections,
                    loose_tables=tables,
                    semantic_document=semantic_document,
                )
            )

        appendix_items = self._appendix_items(semantic_document)
        has_appendix = any(slide.title and "Приложение" in slide.title for slide in normalized_slides)
        if appendix_items and not has_appendix:
            normalized_slides.append(
                SlideSpec(
                    kind=SlideKind.BULLETS,
                    title=f"{plan_title} - Приложение",
                    bullets=appendix_items[: self.LIST_BATCH_SIZE],
                    preferred_layout_key="list_full_width",
                )
            )

        return normalized_slides

    def _appendix_items(self, semantic_document: SemanticDocument) -> list[str]:
        items: list[str] = []
        for fact in semantic_document.facts[:3]:
            items.append(f"{fact.label}: {fact.value}"[:120])
        for contact in semantic_document.contacts[:2]:
            if contact not in items:
                items.append(contact[:120])
        for date in semantic_document.dates[:2]:
            line = f"Дата: {date}"
            if line not in items:
                items.append(line)
        for signature in semantic_document.signatures[:1]:
            if signature not in items:
                items.append(signature[:120])
        return items[: self.LIST_BATCH_SIZE]

    def _is_empty_slide(self, slide: SlideSpec) -> bool:
        return not any(
            [
                (slide.text or "").strip(),
                (slide.notes or "").strip(),
                slide.bullets,
                slide.left_bullets,
                slide.right_bullets,
                slide.table is not None,
                (slide.image_base64 or "").strip(),
            ]
        )

    def _compress_slides(self, slides: list[SlideSpec]) -> list[SlideSpec]:
        return slides

    def _chunk_items(self, items: list[str], size: int) -> list[list[str]]:
        return [items[index : index + size] for index in range(0, len(items), size)]

    def _sentence_chunks(self, text: str) -> list[str]:
        parts = [part.strip() for part in re.split(r"(?<=[.!?;])\s+", text) if part.strip()]
        if not parts:
            return [text[:220]] if text else []
        normalized_parts: list[str] = []
        buffer = ""
        for part in parts:
            candidate = f"{buffer} {part}".strip() if buffer else part
            if len(candidate) <= self.LIST_BULLET_MAX_CHARS:
                buffer = candidate
                continue
            if buffer:
                normalized_parts.append(buffer)
            buffer = part
        if buffer:
            normalized_parts.append(buffer)
        return normalized_parts

    def _is_heading(self, line: str) -> bool:
        return len(line) <= 90 and bool(self.HEADING_PATTERN.match(line.strip()))

    def _looks_like_short_item(self, line: str) -> bool:
        return len(line) <= 120 and not line.endswith(".")

    def _looks_like_structured_label(self, text: str) -> bool:
        normalized = self._normalize_line(text)
        if not normalized:
            return False
        if len(normalized) > 140:
            return False
        letters = [char for char in normalized if char.isalpha()]
        uppercase_ratio = (
            sum(1 for char in letters if char.isupper()) / len(letters)
            if letters
            else 0.0
        )
        has_form_marker = any(marker in normalized.lower() for marker in ("фио", "дата", "подпись", "контакт", "образование", "опыт"))
        if uppercase_ratio >= 0.75:
            return True
        if has_form_marker and len(normalized) <= 120:
            return True
        if normalized.endswith(":") and len(normalized) <= 120:
            return True
        return False

    def _looks_like_resume_label(self, text: str) -> bool:
        normalized = self._normalize_line(text).lower()
        if not normalized:
            return False
        markers = (
            "опыт работы",
            "education",
            "образование",
            "skills",
            "навыки",
            "summary",
            "о себе",
            "контакты",
            "достижения",
            "experience",
        )
        return any(marker in normalized for marker in markers)

    def _normalize_line(self, line: str) -> str:
        return re.sub(r"\s+", " ", line.strip())

    def _looks_like_contacts(self, heading: str, lines: list[str]) -> bool:
        joined = "\n".join([heading, *lines]).strip()
        normalized_heading = heading.lower().strip(" :.-")
        if normalized_heading in {"контакты", "contacts", "contact"}:
            return True
        return bool(self.EMAIL_PATTERN.search(joined) and self.PHONE_PATTERN.search(joined))

    def _first_phone(self, lines: list[str]) -> str:
        joined = "\n".join(lines)
        match = self.PHONE_PATTERN.search(joined)
        return match.group(1).strip() if match else ""

    def _first_email(self, lines: list[str]) -> str:
        joined = "\n".join(lines)
        match = self.EMAIL_PATTERN.search(joined)
        return match.group(0).strip() if match else ""
