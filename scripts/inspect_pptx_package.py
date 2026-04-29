from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path

from pptx import Presentation

from a3presentation.services.pptx_generator import PptxGenerator


def main() -> int:
    parser = argparse.ArgumentParser(description="Inspect generated PPTX package integrity.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--expected-slides", type=int)
    args = parser.parse_args()

    result = inspect_pptx(args.pptx, expected_slides=args.expected_slides)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if not result["violations"] else 1


def inspect_pptx(path: Path, *, expected_slides: int | None = None) -> dict:
    violations: list[str] = []
    slide_count = None
    generator = PptxGenerator()

    try:
        with zipfile.ZipFile(path) as archive:
            bad_entry = archive.testzip()
            if bad_entry is not None:
                violations.append(f"corrupted_archive_entry:{bad_entry}")
            names = archive.namelist()
            duplicates = sorted({name for name in names if names.count(name) > 1})
            violations.extend(f"duplicate_package_entry:{name}" for name in duplicates[:20])
            violations.extend(generator._package_validation_violations(archive))
    except Exception as exc:
        violations.append(f"zip_open_error:{type(exc).__name__}:{exc}")

    try:
        presentation = Presentation(str(path))
        slide_count = len(presentation.slides)
        if expected_slides is not None and slide_count != expected_slides:
            violations.append(f"slide_count_mismatch:expected={expected_slides}:actual={slide_count}")
    except Exception as exc:
        violations.append(f"python_pptx_open_error:{type(exc).__name__}:{exc}")

    return {
        "path": str(path),
        "exists": path.exists(),
        "size": path.stat().st_size if path.exists() else 0,
        "slide_count": slide_count,
        "violations": violations,
    }


if __name__ == "__main__":
    raise SystemExit(main())
