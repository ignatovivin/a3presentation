from __future__ import annotations

import math
import re
from dataclasses import dataclass, field

from a3presentation.domain.api import ChartOverride, DocumentBlock
from a3presentation.domain.presentation import (
    PresentationPlan,
    SlideContentBlock,
    SlideContentBlockKind,
    SlideKind,
    SlideSpec,
    TableBlock,
)
from a3presentation.domain.semantic import DocumentKind, SemanticDocument, SemanticImage, SemanticSection
from a3presentation.services.layout_capacity import (
    DENSE_TEXT_FULL_WIDTH_PROFILE,
    LIST_FULL_WIDTH_PROFILE,
    LayoutCapacityProfile,
    TEXT_FULL_WIDTH_PROFILE,
)
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
    content_blocks: list["SectionContentBlock"] = field(default_factory=list)


@dataclass(frozen=True)
class ContinuationUnit:
    kind: str
    text: str


@dataclass(frozen=True)
class SectionContentBlock:
    kind: str
    text: str = ""
    items: tuple[str, ...] = ()
    table: TableBlock | None = None


class TextToPlanService:
    EMU_PER_PT = 12700
    DEFAULT_TEXT_MARGIN_X_EMU = 91440
    TOP_LEVEL_HEADING_PATTERN = re.compile(r"^\d+\.\s+\S")
    EMAIL_PATTERN = re.compile(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}")
    PHONE_PATTERN = re.compile(r"(\+?\d[\d\s().-]{7,}\d)")
    HEADING_PATTERN = re.compile(r"^(\d+(\.\d+)*[.)]?\s+.+|[А-ЯA-Z].{0,90})$")
    URL_ONLY_PATTERN = re.compile(r"^(?:https?://|www\.)\S+$", re.IGNORECASE)
    REFERENCE_LINE_PATTERN = re.compile(r"^(?:\[\d+\]\s*)+(?:https?://\S+)?$", re.IGNORECASE)
    QUESTION_HEADING_PATTERN = re.compile(r"^Q\d+\s*[:.)—–-]\s+\S", re.IGNORECASE)
    APPENDIX_HEADING_PATTERN = re.compile(
        r"\b(?:источники?|references?|reference|литература|bibliography|appendix|приложение)\b",
        re.IGNORECASE,
    )
    CARD_BULLET_MAX_CHARS = 100
    CARD_BULLET_COUNT = 3
    LIST_BATCH_SIZE = LIST_FULL_WIDTH_PROFILE.max_items
    LIST_SLIDE_MAX_WEIGHT = LIST_FULL_WIDTH_PROFILE.max_weight
    LIST_BULLET_MAX_CHARS = 220
    TEXT_SLIDE_MAX_WEIGHT = TEXT_FULL_WIDTH_PROFILE.max_weight
    TEXT_SLIDE_MAX_CHARS = TEXT_FULL_WIDTH_PROFILE.max_chars
    TEXT_PRIMARY_MAX_CHARS = TEXT_FULL_WIDTH_PROFILE.max_primary_chars
    TEXT_TAIL_MERGE_THRESHOLD = 120
    COVER_META_MAX_LINES = 3
    COVER_META_MAX_LINE_CHARS = 72
    STRUCTURED_SUMMARY_MAX_ITEMS = 5
    RESUME_SUMMARY_MAX_ITEMS = 6
    DENSE_TEXT_INITIAL_BATCH_MAX_CHARS = 900

    def __init__(self) -> None:
        self.normalizer = SemanticDocumentNormalizer()
        self.list_profile = LIST_FULL_WIDTH_PROFILE
        self.text_profile = TEXT_FULL_WIDTH_PROFILE

    def _text_slide_char_budget(self) -> int:
        return self.text_profile.max_chars

    def _text_primary_char_budget(self) -> int:
        return self.text_profile.max_primary_chars

    def _continuation_primary_char_budget(self, subtitle: str = "") -> int:
        if subtitle.strip():
            return min(self.text_profile.max_chars - 120, 380)
        return min(self.text_profile.max_chars - 60, 460)

    def _bullet_slide_item_budget(self) -> int:
        return self.list_profile.max_items

    def _list_slide_weight_budget(self) -> float:
        return self.list_profile.max_weight

    def _text_slide_weight_budget(self) -> float:
        return self.text_profile.max_weight

    def build_plan(
        self,
        template_id: str,
        raw_text: str,
        title: str | None = None,
        tables: list[TableBlock] | None = None,
        blocks: list[DocumentBlock] | None = None,
        chart_overrides: list[ChartOverride] | None = None,
    ) -> PresentationPlan:
        semantic_document = self.normalizer.normalize(
            raw_text=raw_text,
            blocks=blocks or [],
            tables=tables or [],
            title=title,
        )
        sections = self._build_sections(blocks or [], raw_text)
        if not sections:
            sections = [self._section_from_semantic(section) for section in semantic_document.sections]
        plan_title = semantic_document.title
        document_kind = semantic_document.kind.value

        cover_title, cover_meta = self._build_cover(plan_title, sections, blocks or [])
        sections = self._detach_cover_title_from_first_section(sections, cover_title)
        slides: list[SlideSpec] = [
            SlideSpec(
                kind=SlideKind.TITLE,
                title=cover_title,
                subtitle="",
                notes=cover_meta,
                preferred_layout_key="cover",
            )
        ]
        chart_override_map = {override.table_id: override for override in chart_overrides or []}
        table_sequence = 0

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
        content_slides, table_sequence = self._apply_chart_overrides_to_slide_list(
            content_slides,
            chart_override_map,
            table_sequence,
        )
        contact_slides, table_sequence = self._apply_chart_overrides_to_slide_list(
            contact_slides,
            chart_override_map,
            table_sequence,
        )
        slides.extend(content_slides)

        if not blocks_have_tables:
            table_title_base = sections[-1].title if sections else plan_title
            for index, table in enumerate(tables or [], start=1):
                if not table.headers and not table.rows:
                    continue
                table_slide = SlideSpec(
                    kind=SlideKind.TABLE,
                    title=f"{table_title_base} {index}",
                    subtitle="Ключевые данные из документа",
                    table=table,
                    preferred_layout_key="table",
                )
                table_sequence += 1
                slides.append(self._apply_chart_override(table_slide, f"table_{table_sequence}", chart_override_map))

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
        trailing_slides, table_sequence = self._apply_chart_overrides_to_slide_list(
            slides[1:],
            chart_override_map,
            table_sequence,
        )
        slides = [slides[0], *trailing_slides]
        slides = self._enforce_hard_safety_rules(slides, plan_title, semantic_document, blocks or [], tables or [], sections)
        slides = self._backfill_structural_subtitles(slides)
        slides = self._preserve_missing_subheadings(slides, blocks or [])
        return PresentationPlan(template_id=template_id, title=plan_title, slides=slides)

    def _backfill_structural_subtitles(self, slides: list[SlideSpec]) -> list[SlideSpec]:
        current_top_level_title = ""
        for slide in slides:
            title = (slide.title or "").strip()
            if not title or slide.kind == SlideKind.TITLE:
                continue
            if self.TOP_LEVEL_HEADING_PATTERN.match(title):
                current_top_level_title = title
                continue
            if (slide.subtitle or "").strip():
                continue
            if current_top_level_title and current_top_level_title != title:
                slide.subtitle = current_top_level_title
        return slides

    def _preserve_missing_subheadings(self, slides: list[SlideSpec], blocks: list[DocumentBlock]) -> list[SlideSpec]:
        if not slides or not blocks:
            return slides

        def slide_payload(slide: SlideSpec) -> str:
            parts = [slide.title or "", slide.subtitle or "", slide.text or "", slide.notes or "", *slide.bullets]
            if slide.table is not None:
                parts.extend(slide.table.headers)
                for row in slide.table.rows:
                    parts.extend(row)
            return self._normalize_line(" ".join(part for part in parts if part))

        payloads = [slide_payload(slide) for slide in slides]
        combined_payload = " ".join(payloads)

        for block_index, block in enumerate(blocks):
            if block.kind != "subheading" or not (block.text or "").strip():
                continue
            heading = self._normalize_line(block.text or "")
            if len(heading) < 8 or heading in combined_payload:
                continue

            target_index = self._find_following_payload_slide_index(blocks, block_index + 1, payloads)
            if target_index is None:
                continue
            target = slides[target_index]
            if target.kind == SlideKind.BULLETS:
                if heading not in [self._normalize_line(item) for item in target.bullets]:
                    target.bullets = [heading, *target.bullets]
                    target.content_blocks = [self._list_block(target.bullets)]
            elif not (target.subtitle or "").strip() or self.TOP_LEVEL_HEADING_PATTERN.match(target.subtitle or ""):
                target.subtitle = heading
            elif heading not in self._normalize_line(target.subtitle or ""):
                target.subtitle = f"{target.subtitle} · {heading}"[:120]
            payloads[target_index] = slide_payload(target)
            combined_payload = " ".join(payloads)
        return slides

    def _find_following_payload_slide_index(
        self,
        blocks: list[DocumentBlock],
        start_index: int,
        slide_payloads: list[str],
    ) -> int | None:
        for block in blocks[start_index:]:
            probes: list[str] = []
            if block.text:
                probes.append(self._normalize_line(block.text))
            if block.items:
                probes.extend(self._normalize_line(item) for item in block.items if item)
            if block.table is not None:
                probes.extend(self._normalize_line(value) for value in block.table.headers if value)
                for row in block.table.rows:
                    probes.extend(self._normalize_line(value) for value in row if value)
            probes = [probe for probe in probes if len(probe) >= 8]
            if not probes:
                continue
            for slide_index, payload in enumerate(slide_payloads):
                if any(probe in payload for probe in probes):
                    return slide_index
        return None

    def _apply_chart_override(
        self,
        slide: SlideSpec,
        table_id: str,
        chart_override_map: dict[str, ChartOverride],
    ) -> SlideSpec:
        override = chart_override_map.get(table_id)
        slide.source_table_id = table_id
        if override is None or override.mode != "chart" or override.selected_chart is None:
            return slide

        return SlideSpec(
            kind=SlideKind.CHART,
            title=override.selected_chart.title or slide.title,
            subtitle="",
            text=None,
            bullets=[],
            content_blocks=[],
            left_bullets=[],
            right_bullets=[],
            chart=override.selected_chart.model_copy(deep=True),
            source_table_id=table_id,
            notes=slide.notes,
            preferred_layout_key=slide.preferred_layout_key or "table",
        )

    def _apply_chart_overrides_to_slide_list(
        self,
        slides: list[SlideSpec],
        chart_override_map: dict[str, ChartOverride],
        start_index: int,
    ) -> tuple[list[SlideSpec], int]:
        updated_slides: list[SlideSpec] = []
        table_sequence = start_index
        for slide in slides:
            if slide.table is None:
                updated_slides.append(slide)
                continue
            if slide.source_table_id:
                table_id = slide.source_table_id
                matched = re.fullmatch(r"table_(\d+)", table_id)
                if matched:
                    table_sequence = max(table_sequence, int(matched.group(1)))
            else:
                table_sequence += 1
                table_id = f"table_{table_sequence}"
            updated_slides.append(self._apply_chart_override(slide, table_id, chart_override_map))
        return updated_slides, table_sequence

    def _section_from_semantic(self, section: SemanticSection) -> Section:
        content_blocks: list[SectionContentBlock] = [
            SectionContentBlock(kind=self._classify_content_block_kind(paragraph), text=paragraph)
            for paragraph in section.paragraphs
        ]
        if section.bullets:
            content_blocks.append(SectionContentBlock(kind="list", items=tuple(section.bullets)))
        return Section(
            title=section.title,
            level=section.level,
            subtitle=section.subtitle,
            paragraphs=section.paragraphs.copy(),
            bullet_lists=[section.bullets.copy()] if section.bullets else [],
            tables=section.tables.copy(),
            images=section.images.copy(),
            content_blocks=content_blocks,
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
        if section.images:
            return False
        if len(section.paragraphs) > 2:
            return False
        if sum(len(items) for items in section.bullet_lists) > 3:
            return False
        if section.title == cover_title:
            return True
        if re.match(r"^\d+(\.\d+)*[.)]?\s+", section.title or ""):
            return False
        total_items = len(section.paragraphs) + sum(len(items) for items in section.bullet_lists)
        return section.level <= 1 and total_items <= 2

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
                current.subtitle = self._derive_default_subtitle(current.paragraphs)
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

            if block.text and self._looks_like_question_heading(block.text) and current.content_blocks:
                flush()
                current = Section(title=block.text.strip()[:120], level=3)
                continue

            if block.kind == "table" and block.table is not None:
                current.tables.append(block.table)
                current.content_blocks.append(
                    SectionContentBlock(
                        kind="table",
                        table=block.table,
                    )
                )
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
                items = [item.strip()[:220] for item in block.items if item.strip()]
                if self._is_reference_list(items):
                    continue
                current.bullet_lists.append(items)
                current.content_blocks.append(SectionContentBlock(kind="list", items=tuple(items)))
                continue

            if block.text:
                text = block.text.strip()
                if text:
                    if self._should_skip_reference_tail_text(text):
                        continue
                    current.paragraphs.append(text[:1200])
                    current.content_blocks.append(
                        SectionContentBlock(kind=self._classify_content_block_kind(text[:1200]), text=text[:1200])
                    )

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
                current.subtitle = self._derive_default_subtitle(current.paragraphs)
            sections.append(current)
            current = None

        for line in lines:
            if self._is_heading(line) or self._looks_like_question_heading(line):
                flush()
                current = Section(title=line[:120], level=3 if self._looks_like_question_heading(line) else 1)
                continue

            if current is None:
                current = Section(title=line[:120], level=1)
                continue

            normalized_bullet = re.sub(r"^\s*(?:[-*•]|\d+[.)])\s*", "", line).strip()
            if normalized_bullet != line or self._looks_like_short_item(line):
                if not current.bullet_lists:
                    current.bullet_lists.append([])
                current.bullet_lists[-1].append(normalized_bullet[:220])
                if current.content_blocks and current.content_blocks[-1].kind == "list":
                    current.content_blocks[-1] = SectionContentBlock(
                        kind="list",
                        items=(*current.content_blocks[-1].items, normalized_bullet[:220]),
                    )
                else:
                    current.content_blocks.append(SectionContentBlock(kind="list", items=(normalized_bullet[:220],)))
            else:
                current.paragraphs.append(line[:1200])
                current.content_blocks.append(
                    SectionContentBlock(kind=self._classify_content_block_kind(line[:1200]), text=line[:1200])
                )

        flush()
        return sections

    def _section_to_slides(self, section: Section) -> list[SlideSpec]:
        if self._is_appendix_like_section(section):
            return []

        section_lines = [*section.paragraphs, *[item for bullet_list in section.bullet_lists for item in bullet_list]]
        slides: list[SlideSpec] = []

        if self._looks_like_contacts(section.title, section_lines):
            slides.append(self._build_contact_slide(section))
            slides.extend(self._build_table_slides(section))
            return slides

        if any(block.kind == "table" and block.table is not None for block in section.content_blocks):
            slides.extend(self._build_ordered_section_slides(section))
        elif section.paragraphs or section.bullet_lists:
            if self._fits_single_slide(section):
                slides.append(self._build_single_slide(section))
            else:
                slides.extend(self._split_large_section(section))

        elif section.tables:
            slides.extend(self._build_table_slides(section))

        for index, image in enumerate(section.images, start=1):
            slide = self._image_slide(image, section.title, index)
            if slide is not None:
                slides.append(slide)
        return slides

    def _build_ordered_section_slides(self, section: Section) -> list[SlideSpec]:
        slides: list[SlideSpec] = []
        buffered_blocks: list[SectionContentBlock] = []
        use_subtitle = True

        def flush_buffer() -> None:
            nonlocal buffered_blocks, use_subtitle
            if not buffered_blocks:
                return
            fragment = self._section_fragment_from_content_blocks(
                title=section.title,
                subtitle=section.subtitle if use_subtitle else "",
                level=section.level,
                content_blocks=buffered_blocks,
            )
            if fragment.paragraphs or fragment.bullet_lists:
                if self._fits_single_slide(fragment):
                    slides.append(self._build_single_slide(fragment))
                else:
                    slides.extend(self._split_large_section(fragment))
                use_subtitle = False
            buffered_blocks = []

        index = 0
        while index < len(section.content_blocks):
            block = section.content_blocks[index]
            if block.kind == "table" and block.table is not None:
                table_subtitle = section.subtitle if use_subtitle else ""
                if self._can_promote_pre_table_buffer_to_subtitle(buffered_blocks, table_subtitle):
                    table_subtitle = buffered_blocks[0].text.strip()
                    buffered_blocks = []
                else:
                    flush_buffer()
                    table_subtitle = section.subtitle if use_subtitle else ""
                slides.extend(
                    self._build_table_slides_from_table(
                        title=section.title,
                        subtitle=table_subtitle,
                        table=block.table,
                    )
                )
                use_subtitle = False
                index += 1
                continue
            buffered_blocks.append(block)
            index += 1

        flush_buffer()
        return slides

    def _can_promote_pre_table_buffer_to_subtitle(self, blocks: list[SectionContentBlock], existing_subtitle: str) -> bool:
        if len(blocks) != 1:
            return False
        block = blocks[0]
        if block.kind not in {"paragraph", "callout"}:
            return False
        text = block.text.strip()
        if not text or len(text) > 160:
            return False
        return not existing_subtitle.strip() or self._normalize_line(existing_subtitle) == self._normalize_line(text)

    def _section_fragment_from_content_blocks(
        self,
        *,
        title: str,
        subtitle: str,
        level: int,
        content_blocks: list[SectionContentBlock],
    ) -> Section:
        paragraphs = [
            block.text for block in content_blocks if block.kind in {"paragraph", "callout", "qa_item"} and block.text.strip()
        ]
        bullet_lists = [list(block.items) for block in content_blocks if block.kind == "list" and block.items]
        return Section(
            title=title,
            level=level,
            subtitle=subtitle,
            paragraphs=paragraphs,
            bullet_lists=bullet_lists,
            content_blocks=content_blocks.copy(),
        )

    def _detach_cover_title_from_first_section(self, sections: list[Section], cover_title: str) -> list[Section]:
        if not sections:
            return sections

        first = sections[0]
        if not first.title or not cover_title:
            return sections
        if self._normalize_line(first.title) != self._normalize_line(cover_title):
            return sections
        if not (first.subtitle or "").strip():
            return sections

        adjusted_first = Section(
            title=first.subtitle.strip()[:120],
            level=max(first.level + 1, 2),
            subtitle=(first.paragraphs[0][:120] if first.paragraphs else ""),
            paragraphs=first.paragraphs.copy(),
            bullet_lists=[items.copy() for items in first.bullet_lists],
            tables=first.tables.copy(),
            images=first.images.copy(),
            content_blocks=first.content_blocks.copy(),
        )
        return [adjusted_first, *sections[1:]]

    def _is_appendix_like_section(self, section: Section) -> bool:
        title = self._normalize_line(section.title or "")
        if not title:
            return False
        lowered = title.lower()
        if "какие источники данных" in lowered:
            return True
        return bool(self.APPENDIX_HEADING_PATTERN.search(title))

    def _build_table_slides(self, section: Section) -> list[SlideSpec]:
        slides: list[SlideSpec] = []
        table_index = 0
        for table in section.tables:
            if not table.headers and not table.rows:
                continue
            table_index += 1
            slides.extend(
                self._build_table_slides_from_table(
                    title=section.title,
                    subtitle=section.subtitle or "",
                    table=table,
                    table_index=table_index if len(section.tables) > 1 else None,
                )
            )
        return slides

    def _build_table_slides_from_table(
        self,
        *,
        title: str,
        subtitle: str,
        table: TableBlock,
        table_index: int | None = None,
        supporting_text: str = "",
    ) -> list[SlideSpec]:
        chunks = self._split_table_for_slides(table)
        title_base = title or "Таблица"
        slides: list[SlideSpec] = []
        use_suffix = table_index is not None or len(chunks) > 1
        for chunk_index, chunk in enumerate(chunks, start=1):
            suffix = table_index if table_index is not None else chunk_index
            slide_title = title_base if not use_suffix else f"{title_base} ({suffix})"
            slides.append(
                SlideSpec(
                    kind=SlideKind.TABLE,
                    title=slide_title[:120],
                    subtitle=(subtitle or "")[:120],
                    table=chunk,
                    text=supporting_text if len(chunks) == 1 else "",
                    content_blocks=self._paragraph_blocks_from_parts(supporting_text) if supporting_text and len(chunks) == 1 else [],
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
        if col_count == 2 and row_count <= 10 and max_cell_length <= 70 and avg_row_length <= 55:
            return [table]
        if col_count == 3 and row_count <= 5:
            return [table]
        if col_count == 3 and row_count <= 7 and max_cell_length <= 70 and avg_row_length <= 85:
            return [table]
        if col_count == 3 and row_count <= 9 and max_cell_length <= 130 and avg_row_length <= 120:
            return [table]
        if col_count == 3 and row_count <= 7 and max_cell_length <= 155 and avg_row_length <= 95:
            return [table]
        if 4 <= col_count <= 5 and row_count <= 7 and max_cell_length <= 120 and avg_row_length <= 170:
            return [table]

        base_capacity = 10
        if col_count >= 3:
            base_capacity = 8
        if col_count >= 4:
            base_capacity = 7
        if col_count >= 5:
            base_capacity = 6

        if max_cell_length >= 180:
            base_capacity -= 2
        elif max_cell_length >= 120:
            base_capacity -= 1
        if avg_row_length >= 130:
            base_capacity -= 1
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

        chunked_fill_colors: list[list[list[str | None]]] = []
        row_offset = 0
        for chunk in chunks:
            chunk_length = len(chunk)
            chunked_fill_colors.append(table.row_fill_colors[row_offset : row_offset + chunk_length])
            row_offset += chunk_length

        return [
            TableBlock(
                headers=table.headers,
                header_fill_colors=table.header_fill_colors.copy(),
                rows=chunk,
                row_fill_colors=chunk_fill_colors,
            )
            for chunk, chunk_fill_colors in zip(chunks, chunked_fill_colors, strict=False)
        ]

    def _estimate_table_row_weight(self, row: list[str], col_count: int) -> float:
        max_length = max((len(cell or "") for cell in row), default=0)
        avg_length = sum(len(cell or "") for cell in row) / max(1, len(row))

        weight = 1.0
        if col_count >= 4:
            weight += 0.15
        if col_count >= 5:
            weight += 0.1
        if avg_length >= 55:
            weight += 0.2
        if avg_length >= 90:
            weight += 0.25
        if max_length >= 140:
            weight += 0.35
        if max_length >= 220:
            weight += 0.45
        return weight

    def _fits_single_slide(self, section: Section) -> bool:
        units = self._section_continuation_units(section)
        paragraph_chars = sum(len(unit.text) for unit in units if unit.kind == "paragraph")
        bullet_count = sum(1 for unit in units if unit.kind == "bullet")
        max_bullet_len = max((len(unit.text) for unit in units if unit.kind == "bullet"), default=0)
        total_chars = sum(len(unit.text) for unit in units)

        if paragraph_chars <= self._text_slide_char_budget() and bullet_count == 0:
            return True
        if (
            bullet_count <= self.CARD_BULLET_COUNT
            and paragraph_chars <= 160
            and max_bullet_len <= self.CARD_BULLET_MAX_CHARS
        ):
            return True
        if bullet_count > 0 and bullet_count <= self.list_profile.max_items and total_chars <= self.list_profile.max_chars:
            return True
        return False

    def _build_single_slide(self, section: Section) -> SlideSpec:
        units = self._section_continuation_units(section)
        text = " ".join(unit.text for unit in units if unit.kind == "paragraph").strip()
        bullets = [unit.text for unit in units if unit.kind == "bullet"]
        if bullets and not text and self._should_use_cards_layout(section.title, bullets):
            return SlideSpec(
                kind=SlideKind.BULLETS,
                title=section.title,
                bullets=bullets,
                content_blocks=self._content_blocks_from_units(units),
                preferred_layout_key="cards_3",
            )

        if bullets:
            return self._build_continuation_slide(section.title, section.subtitle or "", units)

        if len(text) <= self._text_slide_char_budget():
            primary_text, secondary_text = self._split_text_for_slide(text)
            content_blocks = [
                SlideContentBlock(kind=SlideContentBlockKind(self._classify_content_block_kind(block.text)), text=block.text)
                for block in section.content_blocks
                if block.kind in {"paragraph", "callout", "qa_item"} and block.text.strip()
            ] or self._paragraph_blocks_from_parts(primary_text, secondary_text)
            slide_title, slide_subtitle = self._normalize_sparse_text_header(
                section.title,
                self._normalize_subtitle(section.subtitle or "", primary_text),
                content_blocks,
                fallback_text=primary_text,
            )
            subtitle = self._sanitize_content_subtitle(
                slide_subtitle,
                content_blocks,
                fallback_text=primary_text,
            )
            return SlideSpec(
                kind=SlideKind.TEXT,
                title=slide_title,
                subtitle=subtitle,
                text=primary_text,
                notes=secondary_text,
                content_blocks=content_blocks,
                preferred_layout_key="text_full_width",
            )

        sentences = self._sentence_chunks(text)
        return SlideSpec(
            kind=SlideKind.BULLETS,
            title=section.title,
            bullets=sentences[: self._bullet_slide_item_budget()],
            content_blocks=[self._list_block(sentences[: self._bullet_slide_item_budget()])],
            preferred_layout_key="list_full_width",
        )

    def _split_large_section(self, section: Section) -> list[SlideSpec]:
        units = self._section_continuation_units(section)
        if not units:
            return []

        if all(unit.kind == "paragraph" for unit in units):
            slides: list[SlideSpec] = []
            text_units = self._refine_text_continuation_units([unit.text for unit in units])
            batches, profile = self._choose_text_continuation_batches(
                text_units,
                title=section.title,
                subtitle=section.subtitle or "",
            )
            for index, batch in enumerate(batches):
                slide_title = section.title if index == 0 else f"{section.title} ({index + 1})"
                content_blocks = self._paragraph_blocks_from_parts(*batch)
                merged = " ".join(batch)
                primary_text, secondary_text = self._split_text_for_slide(
                    merged,
                    primary_budget=self._continuation_primary_char_budget(section.subtitle or ""),
                )
                subtitle = self._sanitize_content_subtitle(
                    self._normalize_subtitle(section.subtitle or "", primary_text),
                    content_blocks,
                    fallback_text=primary_text,
                )
                slides.append(
                    SlideSpec(
                        kind=SlideKind.TEXT,
                        title=slide_title,
                        subtitle=subtitle,
                        text=primary_text,
                        notes=secondary_text,
                        content_blocks=content_blocks,
                        preferred_layout_key=profile.layout_key,
                    )
                )
            return slides

        slides = []
        batches = self._chunk_continuation_units_for_slides(units)
        for index, batch in enumerate(batches):
            slide_title = section.title if index == 0 else f"{section.title} ({index + 1})"
            slide_subtitle = section.subtitle if index == 0 else ""
            slides.append(self._build_continuation_slide(slide_title, slide_subtitle, batch))
        return slides

    def _section_continuation_units(self, section: Section) -> list[ContinuationUnit]:
        if section.content_blocks:
            units: list[ContinuationUnit] = []
            for block in section.content_blocks:
                if block.kind in {"paragraph", "callout", "qa_item"} and block.text.strip():
                    units.extend(
                        ContinuationUnit(kind="paragraph", text=chunk)
                        for chunk in self._paragraph_chunks_for_text_flow(block.text.strip())
                        if chunk.strip()
                    )
                elif block.kind == "list":
                    units.extend(
                        ContinuationUnit(kind="bullet", text=item.strip())
                        for item in block.items
                        if item.strip()
                    )
            return units

        units = [
            ContinuationUnit(kind="paragraph", text=chunk)
            for paragraph in section.paragraphs
            for chunk in self._paragraph_chunks_for_text_flow(self._normalize_line(paragraph))
            if chunk.strip()
        ]
        units.extend(
            ContinuationUnit(kind="bullet", text=item)
            for bullet_list in section.bullet_lists
            for item in bullet_list
            if item.strip()
        )
        return units

    def _chunk_continuation_units_for_slides(self, units: list[ContinuationUnit]) -> list[list[ContinuationUnit]]:
        batches: list[list[ContinuationUnit]] = []
        current_batch: list[ContinuationUnit] = []
        current_weight = 0.0
        current_chars = 0

        for unit in units:
            unit_weight = self._continuation_unit_weight(unit)
            unit_chars = len(unit.text)
            if current_batch and (
                len(current_batch) >= self._bullet_slide_item_budget()
                or current_weight + unit_weight > (self._list_slide_weight_budget() + 1.25)
                or current_chars + unit_chars > self.list_profile.max_chars
            ):
                batches.append(current_batch)
                current_batch = []
                current_weight = 0.0
                current_chars = 0

            current_batch.append(unit)
            current_weight += unit_weight
            current_chars += unit_chars

        if current_batch:
            batches.append(current_batch)

        return batches

    def _chunk_bullets_for_slides(self, bullets: list[str]) -> list[list[str]]:
        batches: list[list[str]] = []
        current_batch: list[str] = []
        current_weight = 0.0

        for bullet in bullets:
            bullet_weight = self._estimate_bullet_weight(bullet)
            if current_batch and (
                len(current_batch) >= self._bullet_slide_item_budget() or current_weight + bullet_weight > self._list_slide_weight_budget()
            ):
                batches.append(current_batch)
                current_batch = []
                current_weight = 0.0

            current_batch.append(bullet)
            current_weight += bullet_weight

        if current_batch:
            batches.append(current_batch)

        return batches

    def _choose_text_continuation_batches(
        self,
        chunks: list[str],
        *,
        title: str = "",
        subtitle: str = "",
    ) -> tuple[list[list[str]], LayoutCapacityProfile]:
        regular_batches = self._chunk_text_for_slides(
            chunks,
            profile=self.text_profile,
            title=title,
            subtitle=subtitle,
        )
        dense_batches = self._chunk_text_for_slides(
            chunks,
            profile=DENSE_TEXT_FULL_WIDTH_PROFILE,
            title=title,
            subtitle=subtitle,
        )
        if self._should_use_dense_text_batches(regular_batches, dense_batches):
            return dense_batches, DENSE_TEXT_FULL_WIDTH_PROFILE
        return regular_batches, self.text_profile

    def _should_use_dense_text_batches(self, regular_batches: list[list[str]], dense_batches: list[list[str]]) -> bool:
        if len(dense_batches) >= len(regular_batches):
            return False
        if len(dense_batches) <= 1:
            return False

        dense_char_counts = [sum(len(chunk) for chunk in batch) for batch in dense_batches if batch]
        if not dense_char_counts:
            return False

        minimum_tail_chars = max(
            int(
                self.DENSE_TEXT_INITIAL_BATCH_MAX_CHARS
                * (
                    DENSE_TEXT_FULL_WIDTH_PROFILE.target_fill_ratio
                    - DENSE_TEXT_FULL_WIDTH_PROFILE.continuation_balance_tolerance
                    - 0.01
                )
            ),
            340,
        )
        if dense_char_counts[-1] < minimum_tail_chars:
            return False

        maximum_dense_chars = max(dense_char_counts)
        if maximum_dense_chars > int(self.DENSE_TEXT_INITIAL_BATCH_MAX_CHARS * DENSE_TEXT_FULL_WIDTH_PROFILE.max_fill_ratio):
            return False

        return True

    def _chunk_text_for_slides(
        self,
        chunks: list[str],
        *,
        profile: LayoutCapacityProfile | None = None,
        title: str = "",
        subtitle: str = "",
    ) -> list[list[str]]:
        profile = profile or self.text_profile
        batches: list[list[str]] = []
        current_batch: list[str] = []
        current_weight = 0.0
        current_chars = 0

        for chunk in chunks:
            chunk_weight = self._estimate_text_chunk_weight(chunk)
            if (
                current_batch
                and len(current_batch) > 1
                and self._looks_like_inline_section_heading(current_batch[-1])
                and not self._looks_like_inline_section_heading(chunk)
                and current_chars >= int(profile.max_chars * 0.45)
            ):
                trailing_heading = current_batch.pop()
                batches.append(current_batch)
                current_batch = [trailing_heading]
                current_weight = self._estimate_text_chunk_weight(trailing_heading)
                current_chars = len(trailing_heading)

            if current_batch and (
                len(current_batch) >= profile.max_items
                or current_weight + chunk_weight > profile.max_weight
                or current_chars + len(chunk) > profile.max_chars
            ):
                if (
                    len(current_batch) > 1
                    and self._looks_like_inline_section_heading(current_batch[-1])
                    and not self._looks_like_inline_section_heading(chunk)
                ):
                    trailing_heading = current_batch.pop()
                    batches.append(current_batch)
                    current_batch = [trailing_heading]
                    current_weight = self._estimate_text_chunk_weight(trailing_heading)
                    current_chars = len(trailing_heading)
                else:
                    batches.append(current_batch)
                    current_batch = []
                    current_weight = 0.0
                    current_chars = 0

            current_batch.append(chunk)
            current_weight += chunk_weight
            current_chars += len(chunk)

        if current_batch:
            batches.append(current_batch)

        return self._rebalance_text_batches(batches, profile=profile)

    def _refine_text_continuation_units(self, chunks: list[str]) -> list[str]:
        refined: list[str] = []
        for chunk in chunks:
            refined.extend(
                part
                for part in self._paragraph_chunks_for_text_flow(
                    chunk,
                    max_chars=self._text_primary_char_budget(),
                )
                if part.strip()
            )
        return refined or [chunk for chunk in chunks if chunk.strip()]

    def _rebalance_text_batches(
        self,
        batches: list[list[str]],
        *,
        profile: LayoutCapacityProfile | None = None,
    ) -> list[list[str]]:
        profile = profile or self.text_profile
        if len(batches) <= 1:
            return batches

        min_tail_chars = max(int(profile.target_fill_ratio * profile.max_chars) - 160, 180)
        rebalanced = [batch.copy() for batch in batches if batch]
        if len(rebalanced) <= 1:
            return rebalanced

        for index in range(len(rebalanced) - 1):
            current = rebalanced[index]
            next_batch = rebalanced[index + 1]
            if len(current) <= 1 or not next_batch:
                continue
            trailing_chunk = current[-1].strip()
            if not self._looks_like_inline_section_heading(trailing_chunk):
                continue
            if self._looks_like_inline_section_heading(next_batch[0]):
                continue
            projected_next = [trailing_chunk, *next_batch]
            projected_chars = sum(len(chunk) for chunk in projected_next)
            projected_weight = sum(self._estimate_text_chunk_weight(chunk) for chunk in projected_next)
            if projected_chars > profile.max_chars or projected_weight > profile.max_weight:
                continue
            next_batch.insert(0, current.pop())

        for index in range(len(rebalanced) - 1, 0, -1):
            current = rebalanced[index]
            if not current:
                continue
            current_chars = sum(len(chunk) for chunk in current)
            if current_chars >= min_tail_chars:
                continue

            previous = rebalanced[index - 1]
            while len(previous) > 1 and current_chars < min_tail_chars:
                candidate = previous[-1]
                candidate_chars = len(candidate)
                current_weight = sum(self._estimate_text_chunk_weight(chunk) for chunk in current)
                candidate_weight = self._estimate_text_chunk_weight(candidate)
                projected_chars = current_chars + candidate_chars
                projected_weight = current_weight + candidate_weight
                if projected_chars > profile.max_chars or projected_weight > profile.max_weight:
                    break
                current.insert(0, previous.pop())
                current_chars = projected_chars

        return [batch for batch in rebalanced if batch]

    def _paragraphs_as_bullets(self, paragraphs: list[str]) -> list[str]:
        items: list[str] = []
        for paragraph in paragraphs:
            normalized = self._normalize_line(paragraph)
            if not normalized:
                continue
            items.extend(self._sentence_chunks(normalized))
        return items

    def _paragraph_chunks_for_text_flow(self, text: str, *, max_chars: int | None = None) -> list[str]:
        normalized = self._normalize_line(text)
        if not normalized:
            return []
        chunk_limit = max_chars or self.text_profile.max_chars
        if len(normalized) <= chunk_limit:
            return [normalized]

        parts = [part.strip() for part in re.split(r"(?<=[.!?;:])\s+", normalized) if part.strip()]
        if len(parts) <= 1:
            return self._hard_wrap_text_chunks(normalized, chunk_limit)

        chunks: list[str] = []
        current = ""
        for part in parts:
            candidate = f"{current} {part}".strip() if current else part
            if len(candidate) <= chunk_limit:
                current = candidate
                continue
            if current:
                chunks.append(current)
            if len(part) <= chunk_limit:
                current = part
            else:
                hard_chunks = self._hard_wrap_text_chunks(part, chunk_limit)
                chunks.extend(hard_chunks[:-1])
                current = hard_chunks[-1]
        if current:
            chunks.append(current)
        return chunks

    def _merge_heading_like_text_chunks(self, chunks: list[str]) -> list[str]:
        return [chunk.strip() for chunk in chunks if chunk.strip()]

    def _looks_like_inline_section_heading(self, text: str) -> bool:
        normalized = self._normalize_line(text)
        if not normalized:
            return False
        if len(normalized) > 80:
            return False
        if normalized.endswith((".", "!", "?", ";", ":")):
            return False
        if self._outline_heading_level(normalized) is not None:
            return False
        if self._looks_like_structured_label(normalized):
            return False
        words = normalized.split()
        if not 2 <= len(words) <= 8:
            return False
        if any(word.startswith(("http://", "https://", "www.")) for word in words):
            return False
        letters = [char for char in normalized if char.isalpha()]
        if not letters:
            return False
        uppercase_ratio = sum(1 for char in letters if char.isupper()) / len(letters)
        return uppercase_ratio <= 0.45

    def _looks_like_question_heading(self, text: str) -> bool:
        normalized = self._normalize_line(text)
        if not normalized or len(normalized) > 140:
            return False
        lowered = normalized.lower()
        return bool(self.QUESTION_HEADING_PATTERN.match(normalized)) or lowered.startswith(
            ("вопрос:", "question:", "q:")
        )

    def _hard_wrap_text_chunks(self, text: str, max_chars: int) -> list[str]:
        chunks: list[str] = []
        remaining = text.strip()
        while remaining:
            if len(remaining) <= max_chars:
                chunks.append(remaining)
                break
            split_at = remaining.rfind(" ", 0, max_chars)
            if split_at < int(max_chars * 0.5):
                split_at = max_chars
            chunks.append(remaining[:split_at].strip())
            remaining = remaining[split_at:].strip()
        return chunks

    def _should_skip_reference_tail_text(self, text: str) -> bool:
        normalized = self._normalize_line(text)
        if not normalized:
            return True
        if self.URL_ONLY_PATTERN.match(normalized):
            return True
        if self.REFERENCE_LINE_PATTERN.match(normalized):
            return True
        return False

    def _is_reference_list(self, items: list[str]) -> bool:
        return bool(items) and all(self._should_skip_reference_tail_text(item) for item in items)

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

    def _split_text_for_slide(self, text: str, *, primary_budget: int | None = None) -> tuple[str, str]:
        normalized = text.strip()
        effective_primary_budget = min(primary_budget or self._text_primary_char_budget(), self.text_profile.max_chars)
        if len(normalized) <= effective_primary_budget:
            return normalized, ""

        split_at = normalized.rfind(". ", 0, effective_primary_budget)
        if split_at == -1:
            split_at = normalized.rfind("; ", 0, effective_primary_budget)
        if split_at == -1:
            split_at = normalized.rfind(", ", 0, effective_primary_budget)
        if split_at == -1 or split_at < int(effective_primary_budget * 0.55):
            split_at = effective_primary_budget
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

    def _derive_default_subtitle(self, paragraphs: list[str]) -> str:
        if not paragraphs:
            return ""
        candidate = self._normalize_line(paragraphs[0])[:120]
        if self._looks_like_explicit_subtitle(candidate):
            return candidate
        return ""

    def _looks_like_explicit_subtitle(self, text: str) -> bool:
        normalized = self._normalize_line(text)
        if not normalized:
            return False
        if len(normalized) > 110:
            return False
        if len(normalized.split()) > 14:
            return False
        sentence_like_punctuation = normalized.count(".") + normalized.count(";") + normalized.count("?") + normalized.count("!")
        if sentence_like_punctuation >= 2:
            return False
        return True

    def _sanitize_content_subtitle(
        self,
        subtitle: str,
        content_blocks: list[SlideContentBlock],
        fallback_text: str = "",
    ) -> str:
        normalized_subtitle = (subtitle or "").strip()
        if not normalized_subtitle:
            return ""
        if not self._looks_like_explicit_subtitle(normalized_subtitle):
            return ""

        first_block_text = next(
            (
                (block.text or "").strip()
                for block in content_blocks
                if block.kind in {
                    SlideContentBlockKind.PARAGRAPH,
                    SlideContentBlockKind.CALLOUT,
                    SlideContentBlockKind.QA_ITEM,
                }
                and (block.text or "").strip()
            ),
            "",
        )
        lead_text = first_block_text or (fallback_text or "").strip()
        if lead_text:
            if lead_text.startswith(normalized_subtitle):
                return ""
            if len(normalized_subtitle) >= 24 and lead_text.startswith(normalized_subtitle[:-1]):
                return ""
        return normalized_subtitle[:120]

    def _normalize_sparse_text_header(
        self,
        title: str,
        subtitle: str,
        content_blocks: list[SlideContentBlock],
        *,
        fallback_text: str = "",
    ) -> tuple[str, str]:
        normalized_title = (title or "").strip()[:120]
        normalized_subtitle = (subtitle or "").strip()[:120]
        return normalized_title, normalized_subtitle

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
        if len(leading_blocks) > 6:
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
            content_blocks=self._paragraph_blocks_from_parts(
                "\n".join(section.paragraphs[1:3]) if len(section.paragraphs) > 1 else ""
            ),
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
                    bullets=summary_items[: self._bullet_slide_item_budget()],
                    content_blocks=[self._list_block(summary_items[: self._bullet_slide_item_budget()])],
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
                    content_blocks=self._paragraph_blocks_from_parts(primary_text, secondary_text),
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
                        content_blocks=self._paragraph_blocks_from_parts(primary_text, secondary_text),
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
                    content_blocks=[self._list_block(summary_items[: self.RESUME_SUMMARY_MAX_ITEMS])],
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
                    content_blocks=self._paragraph_blocks_from_parts(primary_text, secondary_text),
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
            if slide.kind == SlideKind.BULLETS and len(slide.bullets) > self._bullet_slide_item_budget():
                slide.bullets = slide.bullets[: self._bullet_slide_item_budget()]
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
                    bullets=appendix_items[: self._bullet_slide_item_budget()],
                    content_blocks=[self._list_block(appendix_items[: self._bullet_slide_item_budget()])],
                    preferred_layout_key="list_full_width",
                )
            )

        return normalized_slides

    def _appendix_items(self, semantic_document: SemanticDocument) -> list[str]:
        if semantic_document.kind == DocumentKind.REPORT:
            return []

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
        return items[: self._bullet_slide_item_budget()]

    def _is_empty_slide(self, slide: SlideSpec) -> bool:
        return not any(
            [
                (slide.text or "").strip(),
                (slide.notes or "").strip(),
                slide.bullets,
                slide.left_bullets,
                slide.right_bullets,
                slide.table is not None,
                slide.chart is not None,
                (slide.image_base64 or "").strip(),
            ]
        )

    def _compress_slides(self, slides: list[SlideSpec]) -> list[SlideSpec]:
        if not slides:
            return slides

        compressed: list[SlideSpec] = []
        for slide in slides:
            if compressed and self._can_merge_slide_pair(compressed[-1], slide):
                compressed[-1] = self._merge_slide_pair(compressed[-1], slide)
                continue
            compressed.append(slide)

        compressed = self._rebalance_continuation_groups(compressed)
        return self._renumber_continuation_titles(compressed)

    def _can_merge_slide_pair(self, previous: SlideSpec, current: SlideSpec) -> bool:
        if previous.kind in {SlideKind.TABLE, SlideKind.CHART, SlideKind.IMAGE}:
            return False
        if current.kind in {SlideKind.TABLE, SlideKind.CHART, SlideKind.IMAGE}:
            return False
        if self._base_slide_title(previous.title) != self._base_slide_title(current.title):
            return False

        previous_bullets = self._slide_as_mergeable_bullets(previous)
        current_bullets = self._slide_as_mergeable_bullets(current)
        if previous_bullets is None or current_bullets is None:
            return False

        merged = previous_bullets + current_bullets
        if len(merged) > self._bullet_slide_item_budget():
            return False
        merged_weight = sum(self._estimate_bullet_weight(item) for item in merged)
        return merged_weight <= self._list_slide_weight_budget()

    def _merge_slide_pair(self, previous: SlideSpec, current: SlideSpec) -> SlideSpec:
        merged_bullets = (self._slide_as_mergeable_bullets(previous) or []) + (self._slide_as_mergeable_bullets(current) or [])
        return SlideSpec(
            kind=SlideKind.BULLETS,
            title=self._base_slide_title(previous.title),
            subtitle=previous.subtitle or current.subtitle,
            bullets=merged_bullets,
            content_blocks=[self._list_block(merged_bullets)],
            preferred_layout_key="list_full_width",
        )

    def _slide_as_mergeable_bullets(self, slide: SlideSpec) -> list[str] | None:
        if slide.kind == SlideKind.BULLETS:
            return [item for item in slide.bullets if item.strip()]
        if slide.kind == SlideKind.TEXT:
            text_parts = [part.strip() for part in (slide.text or "", slide.notes or "") if part and part.strip()]
            if not text_parts:
                return []
            total_chars = sum(len(part) for part in text_parts)
            if total_chars > 180:
                return None
            if slide.content_blocks and len(slide.content_blocks) > 1:
                return None
            return self._sentence_chunks(" ".join(text_parts))
        return None

    def _base_slide_title(self, title: str | None) -> str:
        normalized = (title or "").strip()
        return re.sub(r"\s+\(\d+\)$", "", normalized)

    def _renumber_continuation_titles(self, slides: list[SlideSpec]) -> list[SlideSpec]:
        groups: dict[str, list[int]] = {}
        order: list[str] = []
        for index, slide in enumerate(slides):
            if slide.kind in {SlideKind.TABLE, SlideKind.CHART, SlideKind.IMAGE, SlideKind.TITLE}:
                continue
            base_title = self._base_slide_title(slide.title)
            if not base_title:
                continue
            if base_title not in groups:
                groups[base_title] = []
                order.append(base_title)
            groups[base_title].append(index)

        for base_title in order:
            indexes = groups[base_title]
            if len(indexes) <= 1:
                slides[indexes[0]].title = base_title
                continue
            for position, slide_index in enumerate(indexes, start=1):
                slides[slide_index].title = base_title if position == 1 else f"{base_title} ({position})"

        return slides

    def _rebalance_continuation_groups(self, slides: list[SlideSpec]) -> list[SlideSpec]:
        if len(slides) < 2:
            return slides

        rebalanced = [slide.model_copy(deep=True) for slide in slides]
        grouped_ranges: list[tuple[int, int]] = []
        group_start = 0

        for index in range(1, len(rebalanced) + 1):
            group_closed = index == len(rebalanced)
            if not group_closed:
                same_title = self._base_slide_title(rebalanced[group_start].title) == self._base_slide_title(
                    rebalanced[index].title
                )
                compatible_kinds = (
                    rebalanced[group_start].kind in {SlideKind.TEXT, SlideKind.BULLETS}
                    and rebalanced[index].kind in {SlideKind.TEXT, SlideKind.BULLETS}
                )
                group_closed = not (same_title and compatible_kinds)
            if group_closed:
                if index - group_start > 1:
                    grouped_ranges.append((group_start, index))
                group_start = index

        for start, end in reversed(grouped_ranges):
            group = rebalanced[start:end]
            rebuilt = self._rebalance_single_continuation_group(group)
            rebalanced[start:end] = rebuilt

        return rebalanced

    def _rebalance_single_continuation_group(self, slides: list[SlideSpec]) -> list[SlideSpec]:
        if len(slides) <= 1:
            return slides

        all_text = all(slide.kind == SlideKind.TEXT for slide in slides)
        units: list[ContinuationUnit] = []
        subtitle = ""
        base_title = self._base_slide_title(slides[0].title)

        for slide in slides:
            if not subtitle and (slide.subtitle or "").strip():
                subtitle = (slide.subtitle or "").strip()
            units.extend(self._slide_to_continuation_units(slide))

        units = [unit for unit in units if unit.text.strip()]
        if not units:
            return slides

        total_weight = sum(self._continuation_unit_weight(unit) for unit in units)
        total_chars = sum(len(unit.text) for unit in units)
        if all_text:
            batches, profile = self._choose_text_continuation_batches(
                self._refine_text_continuation_units([unit.text for unit in units if unit.text.strip()]),
                title=base_title,
                subtitle=subtitle,
            )
            rebuilt: list[SlideSpec] = []
            for index, batch in enumerate(batches):
                if not batch:
                    continue
                slide_title = base_title if index == 0 else f"{base_title} ({index + 1})"
                slide_subtitle = subtitle if index == 0 else ""
                content_blocks = self._paragraph_blocks_from_parts(*batch)
                merged = " ".join(batch)
                primary_text, secondary_text = self._split_text_for_slide(merged)
                rebuilt.append(
                    SlideSpec(
                        kind=SlideKind.TEXT,
                        title=slide_title,
                        subtitle=slide_subtitle,
                        text=primary_text,
                        notes=secondary_text,
                        content_blocks=content_blocks,
                        preferred_layout_key=profile.layout_key,
                    )
                )
            return self._compact_continuation_slides(rebuilt or slides, base_title, subtitle)

        else:
            min_by_items = math.ceil(len(units) / max(self._bullet_slide_item_budget(), 1))
            min_by_chars = math.ceil(total_chars / max(self.list_profile.max_chars, 1))
            min_by_weight = math.ceil(total_weight / max(self._list_slide_weight_budget() + 1.75, 1.0))
            max_chars = self.list_profile.max_chars
            max_weight = self._list_slide_weight_budget() + 1.25
        target_slide_count = min(len(slides), max(1, min_by_items, min_by_chars, min_by_weight))
        buckets = self._pack_continuation_units(units, target_slide_count, max_chars=max_chars, max_weight=max_weight)

        rebuilt: list[SlideSpec] = []
        for index, bucket in enumerate(buckets):
            if not bucket:
                continue
            slide_title = base_title if index == 0 else f"{base_title} ({index + 1})"
            slide_subtitle = subtitle if index == 0 else ""
            rebuilt.append(self._build_continuation_slide(slide_title, slide_subtitle, bucket))

        return self._compact_continuation_slides(rebuilt or slides, base_title, subtitle)

    def _slide_to_continuation_units(self, slide: SlideSpec) -> list[ContinuationUnit]:
        if slide.content_blocks:
            units: list[ContinuationUnit] = []
            for block in slide.content_blocks:
                if block.kind in {
                    SlideContentBlockKind.PARAGRAPH,
                    SlideContentBlockKind.CALLOUT,
                    SlideContentBlockKind.QA_ITEM,
                }:
                    text = (block.text or "").strip()
                    if text:
                        units.extend(
                            ContinuationUnit(kind="paragraph", text=chunk)
                            for chunk in self._paragraph_chunks_for_text_flow(
                                text,
                                max_chars=self._text_primary_char_budget(),
                            )
                            if chunk.strip()
                        )
                    continue
                if block.kind == SlideContentBlockKind.BULLET_LIST:
                    units.extend(ContinuationUnit(kind="bullet", text=item.strip()) for item in block.items if item.strip())
            if units:
                return units

        if slide.kind == SlideKind.BULLETS:
            return [ContinuationUnit(kind="bullet", text=item.strip()) for item in slide.bullets if item.strip()]

        if slide.kind == SlideKind.TEXT:
            text_parts = [part.strip() for part in (slide.text or "", slide.notes or "") if part and part.strip()]
            units: list[ContinuationUnit] = []
            for part in text_parts:
                units.extend(
                    ContinuationUnit(kind="paragraph", text=chunk)
                    for chunk in self._paragraph_chunks_for_text_flow(part)
                    if chunk.strip()
                )
            return units

        return []

    def _slide_payload_chars(self, slide: SlideSpec) -> int:
        return sum(
            len(part)
            for part in [slide.text or "", slide.notes or "", *slide.bullets]
            if part and part.strip()
        )

    def _continuation_units_render_limits(self, units: list[ContinuationUnit]) -> tuple[int, float, int | None]:
        if not units:
            return self.text_profile.max_chars, self._text_slide_weight_budget(), None

        content_blocks = self._content_blocks_from_units(units)
        total_chars = sum(len(unit.text) for unit in units)
        has_bullets = any(unit.kind == "bullet" for unit in units)
        if not has_bullets:
            return self.text_profile.max_chars, self._text_slide_weight_budget(), None

        preferred_layout_key = self._preferred_textual_layout_key(content_blocks, total_chars=total_chars)
        paragraph_block_count = sum(
            1
            for block in content_blocks
            if block.kind in {
                SlideContentBlockKind.PARAGRAPH,
                SlideContentBlockKind.CALLOUT,
                SlideContentBlockKind.QA_ITEM,
            }
            and (block.text or "").strip()
        )
        bullet_item_count = sum(
            len(block.items) for block in content_blocks if block.kind == SlideContentBlockKind.BULLET_LIST
        )
        if paragraph_block_count >= bullet_item_count and paragraph_block_count > 0:
            if preferred_layout_key == "dense_text_full_width":
                return DENSE_TEXT_FULL_WIDTH_PROFILE.max_chars, DENSE_TEXT_FULL_WIDTH_PROFILE.max_weight, None
            if preferred_layout_key == "text_full_width":
                return self.text_profile.max_chars, self._text_slide_weight_budget(), None
            return (
                self.list_profile.max_chars,
                self._list_slide_weight_budget() + 1.25,
                self._bullet_slide_item_budget(),
            )

        return (
            self.list_profile.max_chars,
            self._list_slide_weight_budget() + 1.25,
            self._bullet_slide_item_budget(),
        )

    def _continuation_bucket_min_payload_chars(self, units: list[ContinuationUnit]) -> int:
        max_chars, _, _ = self._continuation_units_render_limits(units)
        if max_chars == DENSE_TEXT_FULL_WIDTH_PROFILE.max_chars:
            return max(int(DENSE_TEXT_FULL_WIDTH_PROFILE.target_fill_ratio * max_chars) - 260, 220)
        if max_chars <= self.text_profile.max_chars:
            return max(int(self.text_profile.target_fill_ratio * max_chars) - 240, 150)
        return max(int(self.list_profile.target_fill_ratio * max_chars) - 260, 160)

    def _continuation_units_fit_single_slide(self, units: list[ContinuationUnit]) -> bool:
        if not units:
            return False
        max_chars, max_weight, max_items = self._continuation_units_render_limits(units)
        total_chars = sum(len(unit.text) for unit in units)
        total_weight = sum(self._continuation_unit_weight(unit) for unit in units)
        if max_items is not None and len(units) > max_items:
            return False
        return total_chars <= max_chars and total_weight <= max_weight

    def _compact_continuation_slides(
        self,
        slides: list[SlideSpec],
        base_title: str,
        subtitle: str,
    ) -> list[SlideSpec]:
        if len(slides) <= 1:
            return slides

        compacted = slides.copy()
        min_payload_chars = 180
        index = 0
        while index < len(compacted):
            slide = compacted[index]
            payload_chars = self._slide_payload_chars(slide)
            if payload_chars >= min_payload_chars:
                index += 1
                continue

            neighbor_index = index + 1 if index + 1 < len(compacted) else index - 1
            if neighbor_index < 0 or neighbor_index >= len(compacted):
                index += 1
                continue

            merged_units = self._slide_to_continuation_units(compacted[min(index, neighbor_index)]) + self._slide_to_continuation_units(compacted[max(index, neighbor_index)])
            if not self._continuation_units_fit_single_slide(merged_units):
                index += 1
                continue

            slide_title = base_title if min(index, neighbor_index) == 0 else compacted[min(index, neighbor_index)].title
            slide_subtitle = subtitle if min(index, neighbor_index) == 0 else ""
            merged_slide = self._build_continuation_slide(slide_title, slide_subtitle, merged_units)
            compacted[min(index, neighbor_index)] = merged_slide
            del compacted[max(index, neighbor_index)]
            index = max(min(index, neighbor_index) - 1, 0)

        for idx, slide in enumerate(compacted):
            slide.title = base_title if idx == 0 else f"{base_title} ({idx + 1})"
            slide.subtitle = subtitle if idx == 0 else ""
        return compacted

    def _continuation_unit_weight(self, unit: ContinuationUnit) -> float:
        if unit.kind == "paragraph":
            return self._estimate_text_chunk_weight(unit.text)
        return self._estimate_bullet_weight(unit.text)

    def _pack_continuation_units(
        self,
        units: list[ContinuationUnit],
        slide_count: int,
        *,
        max_chars: int,
        max_weight: float,
    ) -> list[list[ContinuationUnit]]:
        if slide_count <= 1 or len(units) <= 1:
            return [units]

        total_weight = sum(self._continuation_unit_weight(unit) for unit in units)
        total_chars = sum(len(unit.text) for unit in units)
        buckets: list[list[ContinuationUnit]] = []
        current: list[ContinuationUnit] = []
        current_weight = 0.0
        current_chars = 0
        remaining_weight = total_weight
        remaining_chars = total_chars

        for index, unit in enumerate(units):
            remaining_units = len(units) - index
            remaining_buckets = slide_count - len(buckets)
            unit_weight = self._continuation_unit_weight(unit)
            unit_chars = len(unit.text)
            target_weight = remaining_weight / max(remaining_buckets, 1)
            target_chars = remaining_chars / max(remaining_buckets, 1)
            projected_weight = current_weight + unit_weight
            projected_chars = current_chars + unit_chars
            projected_bucket = [*current, unit]
            enough_units_left = remaining_units > remaining_buckets
            projected_overflow = (
                projected_weight > max_weight
                or projected_chars > max_chars
                or not self._continuation_units_fit_single_slide(projected_bucket)
            )
            should_wrap = (
                bool(current)
                and (
                    projected_overflow
                    or (
                        enough_units_left
                        and (
                            (current_weight >= target_weight * 0.92)
                            or (current_chars >= target_chars * 0.92)
                        )
                    )
                )
            )

            if should_wrap:
                buckets.append(current)
                current = []
                current_weight = 0.0
                current_chars = 0
                remaining_buckets = slide_count - len(buckets)
                target_weight = remaining_weight / max(remaining_buckets, 1)
                target_chars = remaining_chars / max(remaining_buckets, 1)
                projected_weight = unit_weight
                projected_chars = unit_chars

            current.append(unit)
            current_weight = projected_weight
            current_chars = projected_chars
            remaining_weight -= unit_weight
            remaining_chars -= unit_chars

        if current:
            buckets.append(current)

        return self._rebalance_continuation_buckets(buckets, max_chars=max_chars, max_weight=max_weight)

    def _rebalance_continuation_buckets(
        self,
        buckets: list[list[ContinuationUnit]],
        *,
        max_chars: int,
        max_weight: float,
    ) -> list[list[ContinuationUnit]]:
        if len(buckets) <= 1:
            return buckets

        rebalanced = [bucket.copy() for bucket in buckets if bucket]
        if len(rebalanced) <= 1:
            return rebalanced

        for index in range(len(rebalanced) - 1, 0, -1):
            current = rebalanced[index]
            previous = rebalanced[index - 1]
            current_chars = sum(len(unit.text) for unit in current)
            min_tail_chars = self._continuation_bucket_min_payload_chars(current)
            if current_chars >= min_tail_chars:
                continue

            current_weight = sum(self._continuation_unit_weight(unit) for unit in current)
            while len(previous) > 1 and current_chars < min_tail_chars:
                candidate = previous[-1]
                candidate_chars = len(candidate.text)
                candidate_weight = self._continuation_unit_weight(candidate)
                projected_bucket = [*([*current]), candidate]
                if not self._continuation_units_fit_single_slide(projected_bucket):
                    break
                current.insert(0, previous.pop())
                current_chars += candidate_chars
                current_weight += candidate_weight

        for index in range(len(rebalanced) - 1):
            current = rebalanced[index]
            next_bucket = rebalanced[index + 1]
            current_chars = sum(len(unit.text) for unit in current)
            min_current_chars = self._continuation_bucket_min_payload_chars(current)
            while len(next_bucket) > 1 and current_chars < min_current_chars:
                candidate = next_bucket[0]
                projected_bucket = [*current, candidate]
                remainder_bucket = next_bucket[1:]
                if not self._continuation_units_fit_single_slide(projected_bucket):
                    break
                if remainder_bucket and sum(len(unit.text) for unit in remainder_bucket) < self._continuation_bucket_min_payload_chars(
                    remainder_bucket
                ):
                    break
                current.append(next_bucket.pop(0))
                current_chars += len(candidate.text)

        for index in range(len(rebalanced) - 1):
            current = rebalanced[index]
            next_bucket = rebalanced[index + 1]
            self._smooth_continuation_bucket_pair(current, next_bucket)

        return [bucket for bucket in rebalanced if bucket]

    def _smooth_continuation_bucket_pair(
        self,
        current: list[ContinuationUnit],
        next_bucket: list[ContinuationUnit],
    ) -> None:
        if not current or len(next_bucket) <= 1:
            return

        current_min_chars = self._continuation_bucket_min_payload_chars(current)
        next_min_chars = self._continuation_bucket_min_payload_chars(next_bucket)
        current_chars = sum(len(unit.text) for unit in current)
        next_chars = sum(len(unit.text) for unit in next_bucket)

        while len(next_bucket) > 1 and next_chars > current_chars + 80:
            candidate = next_bucket[0]
            candidate_chars = len(candidate.text)
            projected_current = [*current, candidate]
            projected_next = next_bucket[1:]
            if not self._continuation_units_fit_single_slide(projected_current):
                break
            if projected_next and sum(len(unit.text) for unit in projected_next) < next_min_chars:
                break
            current_delta = abs(next_chars - current_chars)
            projected_delta = abs((next_chars - candidate_chars) - (current_chars + candidate_chars))
            if projected_delta > current_delta:
                break
            current.append(next_bucket.pop(0))
            current_chars += candidate_chars
            next_chars -= candidate_chars
            current_min_chars = self._continuation_bucket_min_payload_chars(current)
            next_min_chars = self._continuation_bucket_min_payload_chars(next_bucket) if next_bucket else 0

        while len(current) > 1 and current_chars > next_chars + 80:
            candidate = current[-1]
            candidate_chars = len(candidate.text)
            projected_current = current[:-1]
            projected_next = [candidate, *next_bucket]
            if not projected_current or not self._continuation_units_fit_single_slide(projected_next):
                break
            if sum(len(unit.text) for unit in projected_current) < current_min_chars:
                break
            current_delta = abs(next_chars - current_chars)
            projected_delta = abs((next_chars + candidate_chars) - (current_chars - candidate_chars))
            if projected_delta > current_delta:
                break
            next_bucket.insert(0, current.pop())
            current_chars -= candidate_chars
            next_chars += candidate_chars
            current_min_chars = self._continuation_bucket_min_payload_chars(current)
            next_min_chars = self._continuation_bucket_min_payload_chars(next_bucket)

    def _build_continuation_slide(
        self,
        title: str,
        subtitle: str,
        units: list[ContinuationUnit],
    ) -> SlideSpec:
        content_blocks = self._content_blocks_from_units(units)
        normalized_subtitle = self._sanitize_content_subtitle(subtitle, content_blocks)
        has_bullets = any(unit.kind == "bullet" for unit in units)
        total_chars = sum(len(unit.text) for unit in units)
        paragraphs = [unit.text for unit in units]
        preferred_layout_key = self._preferred_textual_layout_key(content_blocks, total_chars=total_chars)
        paragraph_block_count = sum(
            1
            for block in content_blocks
            if block.kind in {
                SlideContentBlockKind.PARAGRAPH,
                SlideContentBlockKind.CALLOUT,
                SlideContentBlockKind.QA_ITEM,
            }
            and (block.text or "").strip()
        )
        bullet_item_count = sum(len(block.items) for block in content_blocks if block.kind == SlideContentBlockKind.BULLET_LIST)

        if not has_bullets:
            merged = " ".join(paragraphs)
            primary_text, secondary_text = self._split_text_for_slide(
                merged,
                primary_budget=self._continuation_primary_char_budget(normalized_subtitle),
            )
            return SlideSpec(
                kind=SlideKind.TEXT,
                title=title,
                subtitle=normalized_subtitle,
                text=primary_text,
                notes=secondary_text,
                content_blocks=content_blocks,
                preferred_layout_key="text_full_width",
            )

        if has_bullets and paragraph_block_count >= bullet_item_count and paragraph_block_count > 0:
            merged = " ".join(paragraphs)
            primary_text, secondary_text = self._split_text_for_slide(
                merged,
                primary_budget=self._continuation_primary_char_budget(normalized_subtitle),
            )
            return SlideSpec(
                kind=SlideKind.TEXT,
                title=title,
                subtitle=normalized_subtitle,
                text=primary_text,
                notes=secondary_text,
                content_blocks=content_blocks,
                preferred_layout_key=preferred_layout_key,
            )

        return SlideSpec(
            kind=SlideKind.BULLETS,
            title=title,
            bullets=[unit.text for unit in units],
            content_blocks=content_blocks,
            preferred_layout_key="list_full_width",
        )

    def _paragraph_blocks_from_parts(self, *parts: str) -> list[SlideContentBlock]:
        return [
            SlideContentBlock(kind=SlideContentBlockKind(self._classify_content_block_kind(part)), text=part.strip())
            for part in parts
            if part and part.strip()
        ]

    def _list_block(self, items: list[str]) -> SlideContentBlock:
        return SlideContentBlock(
            kind=SlideContentBlockKind.BULLET_LIST,
            items=[item.strip() for item in items if item and item.strip()],
        )

    def _content_blocks_from_units(self, units: list[ContinuationUnit]) -> list[SlideContentBlock]:
        blocks: list[SlideContentBlock] = []
        bullet_buffer: list[str] = []

        def flush_bullets() -> None:
            nonlocal bullet_buffer
            if bullet_buffer:
                blocks.append(self._list_block(bullet_buffer))
                bullet_buffer = []

        for unit in units:
            if unit.kind == "paragraph":
                flush_bullets()
                text = unit.text.strip()
                if text:
                    blocks.append(SlideContentBlock(kind=SlideContentBlockKind(self._classify_content_block_kind(text)), text=text))
                continue
            if unit.kind == "bullet":
                text = unit.text.strip()
                if text:
                    bullet_buffer.append(text)

        flush_bullets()
        return blocks

    def _preferred_textual_layout_key(
        self,
        content_blocks: list[SlideContentBlock],
        *,
        total_chars: int,
    ) -> str:
        bullet_items = sum(len(block.items) for block in content_blocks if block.kind == SlideContentBlockKind.BULLET_LIST)
        paragraph_like_chars = sum(
            len((block.text or "").strip())
            for block in content_blocks
            if block.kind in {
                SlideContentBlockKind.PARAGRAPH,
                SlideContentBlockKind.CALLOUT,
                SlideContentBlockKind.QA_ITEM,
            }
        )
        if bullet_items:
            if bullet_items <= 1 and paragraph_like_chars >= 180:
                if (
                    total_chars > self.text_profile.max_chars
                    and total_chars <= min(DENSE_TEXT_FULL_WIDTH_PROFILE.max_chars, self.text_profile.max_chars + 60)
                ):
                    return "dense_text_full_width"
                return "text_full_width"
            if bullet_items <= 2 and paragraph_like_chars >= 260 and total_chars <= DENSE_TEXT_FULL_WIDTH_PROFILE.max_chars:
                if total_chars > self.text_profile.max_chars and total_chars <= (self.text_profile.max_chars + 60):
                    return "dense_text_full_width"
                return "text_full_width"
            return "list_full_width"

        paragraph_like_blocks = [
            block
            for block in content_blocks
            if block.kind in {
                SlideContentBlockKind.PARAGRAPH,
                SlideContentBlockKind.CALLOUT,
                SlideContentBlockKind.QA_ITEM,
            }
            and (block.text or "").strip()
        ]
        if not paragraph_like_blocks:
            return "text_full_width"

        if len(paragraph_like_blocks) <= 2 and total_chars <= self._text_slide_char_budget():
            return "text_full_width"

        emphasized_blocks = sum(
            1 for block in paragraph_like_blocks if block.kind in {SlideContentBlockKind.CALLOUT, SlideContentBlockKind.QA_ITEM}
        )
        if emphasized_blocks and len(paragraph_like_blocks) <= 4 and total_chars <= self.text_profile.max_chars:
            return "text_full_width"

        return "list_full_width"

    def _classify_content_block_kind(self, text: str) -> str:
        normalized = self._normalize_line(text)
        lowered = normalized.lower()
        if not normalized:
            return "paragraph"
        if self._looks_like_inline_section_heading(normalized):
            return "callout"
        if normalized.endswith("?") or lowered.startswith(("q:", "q.", "вопрос:", "question:")):
            return "qa_item"
        if any(lowered.startswith(prefix) for prefix in ("важно:", "итог:", "вывод:", "риск:", "note:", "warning:")):
            return "callout"
        return "paragraph"

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

    def _outline_heading_level(self, text: str) -> int | None:
        match = re.match(r"^(?P<number>\d+(?:\.\d+)*)[.)]?\s+.+$", text.strip())
        if match is None:
            return None
        return max(1, min(len(match.group("number").split(".")), 3))

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
