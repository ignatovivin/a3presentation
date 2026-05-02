from __future__ import annotations

from a3presentation.domain.presentation import SlideSpec


def slide_content_chunks(slide: SlideSpec) -> list[tuple[str, str]]:
    chunks: list[tuple[str, str]] = []
    if slide.content_blocks:
        for block in slide.content_blocks:
            if block.text and block.text.strip():
                chunks.append(("paragraph", block.text.strip()))
            for item in block.items:
                if item and item.strip():
                    chunks.append(("bullet", item.strip()))
    if not chunks:
        for part in [slide.text or "", slide.notes or ""]:
            if part.strip():
                chunks.append(("paragraph", part.strip()))
        chunks.extend(("bullet", item.strip()) for item in slide.bullets if item.strip())
        chunks.extend(("bullet", item.strip()) for item in slide.left_bullets if item.strip())
        chunks.extend(("bullet", item.strip()) for item in slide.right_bullets if item.strip())
    return chunks


def slide_text_demand_chars(slide: SlideSpec) -> int:
    return sum(len(text) for _kind, text in slide_content_chunks(slide))
