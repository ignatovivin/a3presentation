from __future__ import annotations

import base64
import subprocess
import tempfile
from pathlib import Path

from a3presentation.domain.template import TemplateManifest


class TemplatePreviewRenderer:
    def attach_previews(self, manifest: TemplateManifest, template_path: Path) -> None:
        if not self._should_render(manifest):
            return
        try:
            previews = self._render_powerpoint_previews(template_path)
        except Exception:
            return

        for layout in manifest.layouts:
            preview = previews.get(f"layout_{layout.slide_master_index}_{layout.slide_layout_index}")
            if preview:
                layout.preview_image_base64 = preview
                layout.preview_image_content_type = "image/png"

        for prototype in manifest.prototype_slides:
            preview = previews.get(f"slide_{prototype.source_slide_index}")
            if preview:
                prototype.preview_image_base64 = preview
                prototype.preview_image_content_type = "image/png"

    def _should_render(self, manifest: TemplateManifest) -> bool:
        return any(not layout.preview_image_base64 for layout in manifest.layouts) or any(
            not prototype.preview_image_base64 for prototype in manifest.prototype_slides
        )

    def _render_powerpoint_previews(self, template_path: Path) -> dict[str, str]:
        with tempfile.TemporaryDirectory(prefix="a3-template-previews-") as temp_dir:
            temp_path = Path(temp_dir)
            script_path = temp_path / "render_previews.ps1"
            script_path.write_text(self._powershell_script(), encoding="utf-8")
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(script_path),
                    "-PptxPath",
                    str(template_path.resolve()),
                    "-OutDir",
                    str(temp_path),
                ],
                check=True,
                capture_output=True,
                timeout=90,
            )
            previews: dict[str, str] = {}
            for image_path in temp_path.glob("*.png"):
                previews[image_path.stem] = base64.b64encode(image_path.read_bytes()).decode("ascii")
            return previews

    def _powershell_script(self) -> str:
        return r'''
param(
    [Parameter(Mandatory=$true)][string]$PptxPath,
    [Parameter(Mandatory=$true)][string]$OutDir,
    [int]$Width = 320,
    [int]$Height = 180
)

$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$ppt = $null
$presentation = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $presentation = $ppt.Presentations.Open($PptxPath, 0, 0, 0)

    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        $slide = $presentation.Slides.Item($i)
        $slide.Export((Join-Path $OutDir ("slide_{0}.png" -f ($i - 1))), "PNG", $Width, $Height)
    }

    for ($designIndex = 1; $designIndex -le $presentation.Designs.Count; $designIndex++) {
        $master = $presentation.Designs.Item($designIndex).SlideMaster
        for ($layoutIndex = 1; $layoutIndex -le $master.CustomLayouts.Count; $layoutIndex++) {
            $layout = $master.CustomLayouts.Item($layoutIndex)
            $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, $layout)
            $slide.Export((Join-Path $OutDir ("layout_{0}_{1}.png" -f ($designIndex - 1), ($layoutIndex - 1))), "PNG", $Width, $Height)
            $slide.Delete()
        }
    }
} finally {
    if ($presentation) {
        try { $presentation.Close() } catch {}
    }
    if ($ppt) {
        try { $ppt.Quit() } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
'''
