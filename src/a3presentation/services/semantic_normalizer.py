from __future__ import annotations

import re

from a3presentation.domain.api import DocumentBlock
from a3presentation.domain.presentation import TableBlock
from a3presentation.domain.semantic import (
    DocumentKind,
    DocumentStats,
    SemanticDocument,
    SemanticFact,
    SemanticImage,
    SemanticSection,
)


class SemanticDocumentNormalizer:
    EMAIL_PATTERN = re.compile(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}")
    PHONE_PATTERN = re.compile(r"(\+?\d[\d\s().-]{7,}\d)")
    DATE_PATTERN = re.compile(r"\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b")
    LABEL_VALUE_PATTERN = re.compile(r"^\s*([^:]{2,80}):\s*(.+)\s*$")

    def normalize(
        self,
        *,
        raw_text: str,
        blocks: list[DocumentBlock],
        tables: list[TableBlock],
        title: str | None = None,
    ) -> SemanticDocument:
        effective_blocks = self._trim_leading_cover_blocks(blocks)
        sections = self._build_sections(effective_blocks)
        document_title = title or self._detect_title(blocks, raw_text, sections)
        facts = self._extract_facts(blocks)
        contacts = self._extract_contacts(blocks)
        dates = self._extract_dates(blocks)
        signatures = self._extract_signatures(blocks)
        images = self._extract_images(blocks)
        summary_lines = self._summary_lines(blocks, sections, document_title)
        stats = self._build_stats(raw_text, blocks, tables, facts, contacts, dates, signatures, images)
        kind = self._classify_document(blocks, sections, raw_text, tables, contacts)

        return SemanticDocument(
            title=document_title,
            kind=kind,
            summary_lines=summary_lines,
            facts=facts,
            contacts=contacts,
            dates=dates,
            signatures=signatures,
            images=images,
            sections=sections,
            loose_tables=tables,
            stats=stats,
        )

    def _build_sections(self, blocks: list[DocumentBlock]) -> list[SemanticSection]:
        sections: list[SemanticSection] = []
        current: SemanticSection | None = None

        def flush() -> None:
            nonlocal current
            if current is None:
                return
            if current.title:
                sections.append(current)
            current = None

        for block in blocks:
            if block.kind == "title":
                continue

            if block.kind in {"heading", "subheading"} and (block.text or "").strip():
                level = block.level or (1 if block.kind == "heading" else 2)
                if current is not None and level <= current.level:
                    flush()
                if current is None:
                    current = SemanticSection(title=(block.text or "").strip()[:120], level=level)
                elif not current.subtitle:
                    current.subtitle = (block.text or "").strip()[:120]
                else:
                    flush()
                    current = SemanticSection(title=(block.text or "").strip()[:120], level=level)
                continue

            if current is None:
                seed = (block.text or "").strip() or (block.items[0].strip() if block.items else "Раздел")
                if not seed and block.table is None:
                    continue
                current = SemanticSection(title=seed[:120] or "Раздел", level=1)

            if block.kind == "table" and block.table is not None:
                current.tables.append(block.table)
                continue

            if block.kind == "list" and block.items:
                current.bullets.extend(item.strip()[:220] for item in block.items if item.strip())
                continue

            if block.kind == "image":
                current.images.append(
                    SemanticImage(
                        name=block.image_name,
                        alt_text=block.text,
                        content_type=block.image_content_type,
                        image_base64=block.image_base64,
                    )
                )
                continue

            if (block.text or "").strip():
                current.paragraphs.append((block.text or "").strip()[:1600])
                fact = self._fact_from_text(block.text or "")
                if fact is not None:
                    current.facts.append(fact)

        flush()
        return sections

    def _detect_title(self, blocks: list[DocumentBlock], raw_text: str, sections: list[SemanticSection]) -> str:
        for block in blocks:
            if block.kind == "title" and (block.text or "").strip():
                return (block.text or "").strip()[:120]
        leading_lines = self._leading_cover_lines(blocks)
        if leading_lines:
            return self._cover_title_from_lines(leading_lines)[:120]
        for block in blocks:
            if block.kind in {"heading", "subheading", "paragraph"} and (block.text or "").strip():
                return (block.text or "").strip()[:120]
        lines = [line.strip() for line in raw_text.replace("\r", "").split("\n") if line.strip()]
        if lines:
            return lines[0][:120]
        if sections:
            return sections[0].title[:120]
        return "Generated Presentation"

    def _extract_facts(self, blocks: list[DocumentBlock]) -> list[SemanticFact]:
        facts: list[SemanticFact] = []
        seen: set[tuple[str, str]] = set()
        for block in blocks:
            if block.kind != "paragraph" or not (block.text or "").strip():
                continue
            fact = self._fact_from_text(block.text or "")
            if fact is None:
                continue
            key = (fact.label, fact.value)
            if key in seen:
                continue
            seen.add(key)
            facts.append(fact)
        return facts

    def _fact_from_text(self, text: str) -> SemanticFact | None:
        normalized = self._normalize_line(text)
        if not normalized:
            return None
        match = self.LABEL_VALUE_PATTERN.match(normalized)
        if match:
            label = match.group(1).strip()
            value = match.group(2).strip()
            # Only keep concise label:value facts. Long narrative sentences with a colon
            # belong to report prose and pollute appendix/fallback layers.
            if len(label) > 48 or len(value) > 160:
                return None
            if len(label.split()) > 6:
                return None
            if any(char in label for char in ".?!;"):
                return None
            if "http" in label.lower() or "www." in label.lower():
                return None
            if value.lower().startswith("//") or value.lower().startswith("http") or "www." in value.lower():
                return None
            return SemanticFact(label=label[:80], value=value[:240], confidence=0.9, source_text=normalized)
        return None

    def _extract_contacts(self, blocks: list[DocumentBlock]) -> list[str]:
        contacts: list[str] = []
        seen: set[str] = set()
        for block in blocks:
            candidates = []
            text = self._normalize_line(block.text or "")
            if not text:
                continue
            email = self.EMAIL_PATTERN.search(text)
            phone = self.PHONE_PATTERN.search(text)
            if email:
                candidates.append(email.group(0).strip())
            if phone:
                phone_candidate = phone.group(1).strip()
                digits_only = re.sub(r"\D", "", phone_candidate)
                # Avoid interpreting ranges like "2023-2024" as phone numbers.
                if len(digits_only) >= 10:
                    candidates.append(phone_candidate)
            for candidate in candidates:
                if candidate not in seen:
                    contacts.append(candidate)
                    seen.add(candidate)
        return contacts

    def _extract_dates(self, blocks: list[DocumentBlock]) -> list[str]:
        dates: list[str] = []
        seen: set[str] = set()
        for block in blocks:
            text = self._normalize_line(block.text or "")
            for match in self.DATE_PATTERN.findall(text):
                if match not in seen:
                    dates.append(match)
                    seen.add(match)
        return dates

    def _extract_signatures(self, blocks: list[DocumentBlock]) -> list[str]:
        signatures: list[str] = []
        for block in blocks:
            text = self._normalize_line(block.text or "")
            if not text:
                continue
            lowered = text.lower()
            if "подпись" in lowered or "signature" in lowered:
                signatures.append(text[:240])
        return signatures[:6]

    def _extract_images(self, blocks: list[DocumentBlock]) -> list[SemanticImage]:
        images: list[SemanticImage] = []
        for block in blocks:
            if block.kind != "image":
                continue
            images.append(
                SemanticImage(
                    name=block.image_name,
                    alt_text=block.text,
                    content_type=block.image_content_type,
                    image_base64=block.image_base64,
                )
            )
        return images

    def _summary_lines(self, blocks: list[DocumentBlock], sections: list[SemanticSection], title: str) -> list[str]:
        lines: list[str] = []
        seen: set[str] = set()
        for section in sections:
            candidate = self._normalize_line(section.title)
            if candidate and candidate != self._normalize_line(title) and candidate not in seen:
                lines.append(candidate[:120])
                seen.add(candidate)
            if len(lines) >= 6:
                return lines
        for block in blocks:
            text = self._normalize_line(block.text or "")
            if text and len(text) <= 120 and text not in seen:
                lines.append(text[:120])
                seen.add(text)
            if len(lines) >= 6:
                break
        return lines

    def _classify_document(
        self,
        blocks: list[DocumentBlock],
        sections: list[SemanticSection],
        raw_text: str,
        tables: list[TableBlock],
        contacts: list[str],
    ) -> DocumentKind:
        heading_count = sum(1 for block in blocks if block.kind in {"heading", "subheading"})
        paragraphs = [self._normalize_line(block.text or "") for block in blocks if block.kind == "paragraph" and (block.text or "").strip()]
        table_count = len([block for block in blocks if block.kind == "table" and block.table is not None]) or len(tables)
        list_count = sum(1 for block in blocks if block.kind == "list" and block.items)
        short_labels = sum(1 for paragraph in paragraphs if self._looks_like_structured_label(paragraph))
        long_paragraphs = sum(1 for paragraph in paragraphs if len(paragraph) >= 180)
        resume_markers = sum(1 for paragraph in paragraphs if self._looks_like_resume_label(paragraph))
        short_paragraphs = sum(1 for paragraph in paragraphs if len(paragraph) <= 80)
        narrative_density = long_paragraphs / max(1, len(paragraphs))
        heading_density = heading_count / max(1, len(blocks))

        if heading_count >= 4 and (long_paragraphs >= 3 or narrative_density >= 0.2):
            return DocumentKind.REPORT
        if table_count >= 3 and short_labels >= 2 and short_paragraphs >= len(paragraphs) * 0.45 and heading_density < 0.08:
            return DocumentKind.FORM
        if resume_markers >= 3 and contacts and table_count <= 1:
            return DocumentKind.RESUME
        if table_count >= 3 and heading_count <= 1 and len(sections) <= 2:
            return DocumentKind.TABLE_HEAVY
        if heading_count >= 2 and (long_paragraphs >= 2 or list_count >= 1):
            return DocumentKind.REPORT
        if table_count == 0 and heading_count >= 2 and len(raw_text.strip()) >= 600:
            return DocumentKind.REPORT
        if short_labels >= 4 and long_paragraphs <= 2:
            return DocumentKind.FORM
        if len(sections) <= 1 and table_count >= 1:
            return DocumentKind.UNKNOWN
        return DocumentKind.MIXED

    def _build_stats(
        self,
        raw_text: str,
        blocks: list[DocumentBlock],
        tables: list[TableBlock],
        facts: list[SemanticFact],
        contacts: list[str],
        dates: list[str],
        signatures: list[str],
        images: list[SemanticImage],
    ) -> DocumentStats:
        return DocumentStats(
            paragraph_count=sum(1 for block in blocks if block.kind == "paragraph"),
            heading_count=sum(1 for block in blocks if block.kind in {"heading", "subheading", "title"}),
            list_count=sum(1 for block in blocks if block.kind == "list"),
            table_count=len([block for block in blocks if block.kind == "table" and block.table is not None]) or len(tables),
            image_count=len(images),
            character_count=len(raw_text),
            fact_count=len(facts),
            contact_count=len(contacts),
            date_count=len(dates),
            signature_count=len(signatures),
        )

    def _looks_like_structured_label(self, text: str) -> bool:
        normalized = self._normalize_line(text)
        if not normalized or len(normalized) > 140:
            return False
        letters = [char for char in normalized if char.isalpha()]
        uppercase_ratio = (
            sum(1 for char in letters if char.isupper()) / len(letters)
            if letters
            else 0.0
        )
        if uppercase_ratio >= 0.75:
            return True
        lowered = normalized.lower()
        if any(marker in lowered for marker in ("фио", "дата", "подпись", "контакт", "образование", "опыт")):
            return True
        return normalized.endswith(":")

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

    def _leading_cover_lines(self, blocks: list[DocumentBlock]) -> list[str]:
        lines: list[str] = []
        for block in blocks:
            if block.kind in {"heading", "subheading", "table", "list"}:
                break
            text = self._normalize_line(block.text or "")
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

        candidate_lines = [self._normalize_line(block.text or "") for block in leading_blocks if (block.text or "").strip()]
        if not candidate_lines or any(len(line) > 140 for line in candidate_lines):
            return blocks
        return blocks[first_structured_index:]

    def _cover_title_from_lines(self, lines: list[str]) -> str:
        normalized = [line for line in lines if line]
        if not normalized:
            return ""
        if len(normalized) >= 2 and self._looks_like_cover_prefix(normalized[0]):
            return f"{normalized[0]} {normalized[1]}".strip()
        return normalized[0]

    def _looks_like_cover_prefix(self, text: str) -> bool:
        compact = text.strip()
        if not compact or len(compact) > 8 or compact[0].isdigit():
            return False
        return len(compact.split()) <= 2
