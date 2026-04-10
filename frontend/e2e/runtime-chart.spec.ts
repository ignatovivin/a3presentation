import { execFileSync } from "node:child_process";
import path from "node:path";

import { expect, test } from "@playwright/test";

const runtimeExtractResponse = {
  file_name: "runtime-chart.docx",
  text: "A3\nRuntime chart parity\n1. Каналы\nРост каналов нужно проверить через живой backend runtime.",
  tables: [
    {
      headers: ["Канал", "Q1", "Q2", "Q3"],
      rows: [
        ["SMB", "120", "140", "180"],
        ["Enterprise", "220", "260", "310"],
      ],
    },
  ],
  blocks: [
    { kind: "title", text: "A3", items: [] },
    { kind: "heading", text: "1. Каналы", level: 1, items: [] },
    { kind: "paragraph", text: "Рост каналов нужно проверить через живой backend runtime.", items: [] },
    {
      kind: "table",
      text: null,
      items: [],
      table: {
        headers: ["Канал", "Q1", "Q2", "Q3"],
        rows: [
          ["SMB", "120", "140", "180"],
          ["Enterprise", "220", "260", "310"],
        ],
      },
    },
  ],
  chart_assessments: [
    {
      table_id: "table_1",
      chartable: true,
      classification: "multi_series_category",
      confidence: "high",
      reasons: ["runtime smoke chart candidate"],
      warnings: [],
      candidate_specs: [
        {
          chart_id: "chart_runtime_column",
          source_table_id: "table_1",
          chart_type: "column",
          title: "Runtime chart parity",
          categories: ["Q1", "Q2", "Q3"],
          series: [
            { name: "SMB", values: [120, 140, 180], unit: null, axis: "primary", hidden: false },
            { name: "Enterprise", values: [220, 260, 310], unit: null, axis: "primary", hidden: false },
          ],
          x_axis_title: null,
          y_axis_title: null,
          legend_visible: true,
          data_labels_visible: false,
          value_format: "number",
          stacking: "none",
          confidence: "high",
          warnings: [],
          transpose_allowed: true,
        },
      ],
      structured_table: null,
    },
  ],
};

const runtimeMarketShareExtractResponse = {
  file_name: "runtime-market-share.docx",
  text: "A3\nMarket share runtime parity\n1. Рынок\nПроверка безопасного сравнения объема рынка и доли А3 GMV.",
  tables: [
    {
      headers: ["Рынок / сегмент", "Объем рынка (2024)", "Доля А3 GMV (2025)"],
      rows: [
        ["ЖКХ", "8496000000000", "2.12"],
        ["Налоги", "55600000000000", "0.02"],
        ["Образование", "851000000000", "0.06"],
      ],
    },
  ],
  blocks: [
    { kind: "title", text: "A3", items: [] },
    { kind: "heading", text: "1. Рынок", level: 1, items: [] },
    { kind: "paragraph", text: "Проверка безопасного сравнения объема рынка и доли А3 GMV.", items: [] },
    {
      kind: "table",
      text: null,
      items: [],
      table: {
        headers: ["Рынок / сегмент", "Объем рынка (2024)", "Доля А3 GMV (2025)"],
        rows: [
          ["ЖКХ", "8496000000000", "2.12"],
          ["Налоги", "55600000000000", "0.02"],
          ["Образование", "851000000000", "0.06"],
        ],
      },
    },
  ],
  chart_assessments: [
    {
      table_id: "table_1",
      chartable: true,
      classification: "multi_series_category",
      confidence: "low",
      reasons: ["mixed units detected", "runtime market/share chart candidate"],
      warnings: ["mixed-unit chart uses a secondary value axis"],
      candidate_specs: [
        {
          chart_id: "runtime_market_share_combo_1",
          source_table_id: "table_1",
          chart_type: "combo",
          variant_label: "Комбинированный: столбцы Объем рынка (2024); линия Доля А3 GMV (2025)",
          title: "Рынок / сегмент · Объем рынка (2024) / Доля А3 GMV (2025)",
          categories: ["ЖКХ", "Налоги", "Образование"],
          series: [
            { name: "Объем рынка (2024)", values: [8_496_000_000_000, 55_600_000_000_000, 851_000_000_000], unit: null, axis: "primary", hidden: false },
            { name: "Доля А3 GMV (2025)", values: [2.12, 0.02, 0.06], unit: "%", axis: "secondary", hidden: false },
          ],
          x_axis_title: "Рынок / сегмент",
          y_axis_title: null,
          legend_visible: true,
          data_labels_visible: false,
          value_format: "number",
          stacking: "none",
          confidence: "low",
          warnings: ["mixed-unit chart uses a secondary value axis"],
          transpose_allowed: false,
        },
        {
          chart_id: "runtime_market_share_column_1",
          source_table_id: "table_1",
          chart_type: "column",
          variant_label: "Единичный: Объем рынка (2024)",
          title: "Рынок / сегмент · Объем рынка (2024) / Доля А3 GMV (2025)",
          categories: ["ЖКХ", "Налоги", "Образование"],
          series: [
            { name: "Объем рынка (2024)", values: [8_496_000_000_000, 55_600_000_000_000, 851_000_000_000], unit: null, axis: "primary", hidden: false },
          ],
          x_axis_title: "Рынок / сегмент",
          y_axis_title: "Объем рынка (2024)",
          legend_visible: false,
          data_labels_visible: false,
          value_format: "number",
          stacking: "none",
          confidence: "low",
          warnings: ["mixed-unit chart uses a secondary value axis"],
          transpose_allowed: false,
        },
        {
          chart_id: "runtime_market_share_column_2",
          source_table_id: "table_1",
          chart_type: "column",
          variant_label: "Единичный: Доля А3 GMV (2025)",
          title: "Рынок / сегмент · Объем рынка (2024) / Доля А3 GMV (2025)",
          categories: ["ЖКХ", "Налоги", "Образование"],
          series: [
            { name: "Доля А3 GMV (2025)", values: [2.12, 0.02, 0.06], unit: "%", axis: "primary", hidden: false },
          ],
          x_axis_title: "Рынок / сегмент",
          y_axis_title: "Доля А3 GMV (2025)",
          legend_visible: false,
          data_labels_visible: false,
          value_format: "percent",
          stacking: "none",
          confidence: "low",
          warnings: ["mixed-unit chart uses a secondary value axis"],
          transpose_allowed: false,
        },
      ],
      structured_table: null,
    },
  ],
};

function auditGeneratedMarketShareDeck(outputPath: string, plan: unknown) {
  const repoRoot = path.resolve(process.cwd(), "..");
  const pythonExecutable = path.join(repoRoot, ".venv", process.platform === "win32" ? "Scripts\\python.exe" : "bin/python");
  const auditScript = `
import json
import os
import sys
from pathlib import Path

from a3presentation.domain.presentation import PresentationPlan
from a3presentation.services.deck_audit import audit_generated_presentation, find_capacity_violations

plan = PresentationPlan.model_validate(json.loads(os.environ["PLAN_JSON"]))
audits = audit_generated_presentation(Path(sys.argv[1]), plan)
violations = find_capacity_violations(audits)
if violations:
    raise AssertionError([violation.__dict__ for violation in violations])

chart_audit = next((audit for audit in audits if audit.kind == "chart"), None)
if chart_audit is None:
    raise AssertionError("chart audit not found")

assert chart_audit.expected_chart_type == "combo", chart_audit
assert chart_audit.rendered_chart_type == "combo", chart_audit
assert chart_audit.expected_chart_secondary_value_axis is True, chart_audit
assert chart_audit.rendered_chart_secondary_value_axis is True, chart_audit
assert chart_audit.expected_chart_secondary_value_axis_number_format == '0"%"', chart_audit
assert chart_audit.rendered_chart_secondary_value_axis_number_format == '0"%"', chart_audit
assert chart_audit.rendered_chart_bar_series_count == 1, chart_audit
assert chart_audit.rendered_chart_line_series_count == 1, chart_audit
`;
  execFileSync(pythonExecutable, ["-c", auditScript, outputPath], {
    cwd: repoRoot,
    env: {
      ...process.env,
      PLAN_JSON: JSON.stringify(plan),
    },
    encoding: "utf-8",
  });
}

test("@smoke runtime chart UI flow reaches real backend generate and download", async ({ page }) => {
  await page.route("**/api/documents/extract-text", async (route) => {
    await route.fulfill({ json: runtimeExtractResponse });
  });

  const generateResponses: Array<{ file_name: string; download_url: string; output_path: string }> = [];
  page.on("response", async (response) => {
    if (!response.url().includes("/api/presentations/generate") || !response.ok()) {
      return;
    }
    generateResponses.push(await response.json());
  });

  await page.goto("/");

  await page.getByTestId("upload-document-input").setInputFiles({
    name: "runtime-chart.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from("runtime-chart-docx-is-mocked-at-extraction-layer"),
  });

  await page.getByTestId("open-structure-drawer").click();
  await page.getByTestId("mode-chart-table_1").click();
  await page.getByTestId("save-structure-choices").click();
  await page.getByTestId("generate-presentation").click();

  await expect(page.getByTestId("generation-success")).toBeVisible();
  await expect.poll(() => generateResponses.length).toBe(1);

  const generated = generateResponses[0];
  expect(generated.file_name).toMatch(/\.pptx$/);
  expect(generated.download_url).toContain("/presentations/files/");
  await expect(page.getByTestId("generated-file-name")).toHaveText(generated.file_name);

  const downloadResponse = await page.request.get(`/api${generated.download_url}`);
  expect(downloadResponse.ok()).toBeTruthy();
  expect(downloadResponse.headers()["content-type"]).toContain("presentation");
});

test("@smoke runtime market/share combo comparison reaches real backend generate", async ({ page }) => {
  await page.route("**/api/documents/extract-text", async (route) => {
    await route.fulfill({ json: runtimeMarketShareExtractResponse });
  });

  const planResponses: any[] = [];
  const generateResponses: Array<{ file_name: string; download_url: string; output_path: string }> = [];
  page.on("response", async (response) => {
    if (response.url().includes("/api/plans/from-text") && response.ok()) {
      planResponses.push(await response.json());
    }
    if (response.url().includes("/api/presentations/generate") && response.ok()) {
      generateResponses.push(await response.json());
    }
  });

  await page.goto("/");

  await page.getByTestId("upload-document-input").setInputFiles({
    name: "runtime-market-share.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from("runtime-market-share-docx-is-mocked-at-extraction-layer"),
  });

  await page.getByTestId("open-structure-drawer").click();
  await page.getByTestId("mode-chart-table_1").click();
  await expect(page.getByTestId("chart-type-table_1")).toHaveCount(0);
  await expect(page.getByTestId("chart-variant-table_1")).toContainText("Сравнение: Объем рынка (2024), Доля А3 GMV (2025)");
  await expect(page.getByTestId("chart-variant-table_1")).toHaveValue("runtime_market_share_combo_1");
  await page.getByTestId("save-structure-choices").click();
  await page.getByTestId("generate-presentation").click();

  await expect(page.getByTestId("generation-success")).toBeVisible();
  await expect.poll(() => planResponses.length).toBe(1);
  await expect.poll(() => generateResponses.length).toBe(1);

  const chartSlide = planResponses[0].slides.find((slide: any) => slide.kind === "chart");
  expect(chartSlide.chart.chart_type).toBe("combo");
  expect(chartSlide.chart.series).toEqual([
    expect.objectContaining({ name: "Объем рынка (2024)", axis: "primary" }),
    expect.objectContaining({ name: "Доля А3 GMV (2025)", unit: "%", axis: "secondary" }),
  ]);

  const generated = generateResponses[0];
  expect(generated.file_name).toMatch(/\.pptx$/);
  auditGeneratedMarketShareDeck(generated.output_path, planResponses[0]);
  const downloadResponse = await page.request.get(`/api${generated.download_url}`);
  expect(downloadResponse.ok()).toBeTruthy();
  expect(downloadResponse.headers()["content-type"]).toContain("presentation");
});
