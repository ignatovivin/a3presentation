import { expect, type Page, test } from "@playwright/test";

const templatesResponse = [
  {
    template_id: "corp_light_v1",
    display_name: "Light Theme",
    description: "Corporate template",
  },
];

function chartSpec(tableId: string, chartType: string, title: string, overrides = {}) {
  return {
    chart_id: `chart_${tableId}_${chartType}`,
    source_table_id: tableId,
    chart_type: chartType,
    title,
    categories: ["Q1", "Q2", "Q3", "Q4"],
    series: [
      { name: "SMB", values: [120, -40, 160, 190], unit: null, axis: "primary", hidden: false },
      { name: "Enterprise", values: [220, 260, 210, 320], unit: null, axis: "primary", hidden: false },
    ],
    x_axis_title: null,
    y_axis_title: null,
    legend_visible: true,
    data_labels_visible: false,
    value_format: "number",
    stacking: chartType.startsWith("stacked") ? "stacked" : "none",
    confidence: "high",
    warnings: [],
    transpose_allowed: true,
    ...overrides,
  };
}

const chartAssessments = [
  {
    table_id: "table_column",
    candidate_specs: [chartSpec("table_column", "column", "Column preview")],
  },
  {
    table_id: "table_bar",
    candidate_specs: [chartSpec("table_bar", "bar", "Bar preview")],
  },
  {
    table_id: "table_line",
    candidate_specs: [
      chartSpec("table_line", "line", "Line preview", {
        categories: ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август"],
        series: [
          { name: "SMB", values: [55_600_000, 55_900_000, 56_100_000, 56_500_000, 56_800_000, 57_200_000, 57_400_000, 57_900_000], unit: null, axis: "primary", hidden: false },
          { name: "Enterprise", values: [2_100_000, 2_200_000, 2_000_000, 2_450_000, 2_700_000, 2_550_000, 2_900_000, 3_100_000], unit: null, axis: "primary", hidden: false },
        ],
        value_format: "currency",
      }),
    ],
  },
  {
    table_id: "table_stacked_column",
    candidate_specs: [chartSpec("table_stacked_column", "stacked_column", "Stacked column preview")],
  },
  {
    table_id: "table_stacked_bar",
    candidate_specs: [chartSpec("table_stacked_bar", "stacked_bar", "Stacked bar preview")],
  },
  {
    table_id: "table_pie",
    candidate_specs: [
      chartSpec("table_pie", "pie", "Pie preview", {
        categories: ["SMB", "Enterprise", "Партнеры"],
        series: [{ name: "Доля", values: [40, 35, 25], unit: null, axis: "primary", hidden: false }],
        value_format: "percent",
      }),
    ],
  },
  {
    table_id: "table_combo",
    candidate_specs: [
      chartSpec("table_combo", "combo", "Combo preview", {
        categories: ["Q1", "Q2", "Q3", "Q4"],
        series: [
          { name: "План", values: [120_000_000, 140_000_000, 160_000_000, 190_000_000], unit: "RUB", axis: "primary", hidden: false },
          { name: "Факт", values: [100_000_000, 150_000_000, 170_000_000, 200_000_000], unit: "RUB", axis: "primary", hidden: false },
          { name: "Маржа", values: [18, 22, 24, 27], unit: "%", axis: "primary", hidden: false },
        ],
      }),
    ],
  },
].map((assessment) => ({
  chartable: true,
  classification: "multi_series_category",
  confidence: "high",
  reasons: ["chart preview fixture"],
  warnings: [],
  structured_table: null,
  ...assessment,
}));

const extractResponse = {
  file_name: "chart-preview.docx",
  text: "Графики\nТест preview всех типов.",
  tables: chartAssessments.map((assessment) => ({
    headers: ["Категория", "SMB", "Enterprise"],
    header_fill_colors: [null, null, null],
    rows: [
      ["Q1", "120", "220"],
      ["Q2", "140", "260"],
    ],
    row_fill_colors: [
      [null, null, null],
      [null, null, null],
    ],
    table_id: assessment.table_id,
  })),
  blocks: [
    { kind: "title", text: "Графики", items: [] },
    { kind: "paragraph", text: "Тест preview всех типов.", items: [] },
  ],
  chart_assessments: chartAssessments,
};

test.beforeEach(async ({ page }) => {
  await page.route("**/api/templates", async (route) => {
    await route.fulfill({ json: templatesResponse });
  });

  await page.route("**/api/documents/extract-text", async (route) => {
    await route.fulfill({ json: extractResponse });
  });
});

async function openStructureDrawer(page: Page) {
  await page.goto("/");
  await page.getByTestId("upload-document-input").setInputFiles({
    name: "chart-preview.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from("mock-docx"),
  });
  await page.getByTestId("open-structure-drawer").click();
  await expect(page.getByTestId("structure-drawer")).toBeVisible();
}

test("@smoke chart preview renders every supported chart layout", async ({ page }) => {
  await openStructureDrawer(page);

  const columnCard = page.getByTestId("assessment-card-table_column");
  await page.getByTestId("mode-chart-table_column").click();
  await expect(columnCard.locator(".chart-preview-bar")).toHaveCount(8);
  const columnBarHeights = await columnCard.locator(".chart-preview-bar").evaluateAll((items) =>
    items.map((item) => (item as HTMLElement).getBoundingClientRect().height),
  );
  expect(columnBarHeights.every((height) => height > 14)).toBeTruthy();
  expect(new Set(columnBarHeights.map((height) => Math.round(height))).size).toBeGreaterThan(1);

  const barCard = page.getByTestId("assessment-card-table_bar");
  await page.getByTestId("mode-chart-table_bar").click();
  await expect(barCard.locator(".chart-preview-horizontal-bar")).toHaveCount(8);

  const lineCard = page.getByTestId("assessment-card-table_line");
  await page.getByTestId("mode-chart-table_line").click();
  await expect(lineCard.locator(".chart-preview-line-svg path.chart-preview-line")).toHaveCount(2);
  await expect(lineCard.locator(".chart-preview-line-marker")).toHaveCount(16);
  await expect(lineCard.locator(".chart-preview-line-label")).toHaveCount(8);
  await expect(lineCard.locator(".chart-preview-line").first()).toHaveCSS("stroke-width", "1.8px");
  const markerStyles = await lineCard.locator(".chart-preview-line-marker").evaluateAll((items) =>
    items.map((item) => (item as HTMLElement).getAttribute("style") ?? ""),
  );
  expect(markerStyles.every((style) => !style.includes("NaN"))).toBeTruthy();

  const stackedColumnCard = page.getByTestId("assessment-card-table_stacked_column");
  await page.getByTestId("mode-chart-table_stacked_column").click();
  await expect(stackedColumnCard.locator(".chart-preview-stack-segment")).toHaveCount(8);

  const stackedBarCard = page.getByTestId("assessment-card-table_stacked_bar");
  await page.getByTestId("mode-chart-table_stacked_bar").click();
  await expect(stackedBarCard.locator(".chart-preview-horizontal-segment")).toHaveCount(8);

  const pieCard = page.getByTestId("assessment-card-table_pie");
  await page.getByTestId("mode-chart-table_pie").click();
  await expect(pieCard.locator(".chart-preview-pie-svg path")).toHaveCount(3);
  await expect(pieCard.locator(".chart-preview-pie-item")).toHaveCount(3);

  const comboCard = page.getByTestId("assessment-card-table_combo");
  await page.getByTestId("mode-chart-table_combo").click();
  await expect(comboCard.locator(".chart-preview-bar")).toHaveCount(8);
  await expect(comboCard.locator(".chart-preview-line-svg path.chart-preview-line")).toHaveCount(1);
  await expect(comboCard.locator(".chart-preview-line-marker")).toHaveCount(4);
  await expect(comboCard.locator(".chart-preview-axis")).toHaveCount(2);
  await expect(comboCard.locator(".chart-preview-line-label").first()).toContainText("%");
  await expect(comboCard.locator(".chart-preview-bar-value").first()).toContainText("₽");
  await expect(comboCard.locator(".chart-preview-axis").first()).not.toContainText("%");
  await expect(comboCard.locator(".chart-preview-axis").nth(1)).toContainText("%");
});

test("@smoke combo preview falls back to column when line series is hidden", async ({ page }) => {
  await openStructureDrawer(page);

  const comboCard = page.getByTestId("assessment-card-table_combo");
  await page.getByTestId("mode-chart-table_combo").click();
  await expect(comboCard.locator(".chart-preview-line-svg path.chart-preview-line")).toHaveCount(1);

  await page.getByTestId("series-toggle-table_combo-Маржа").click();

  await expect(comboCard.locator(".chart-preview-line-svg path.chart-preview-line")).toHaveCount(0);
  await expect(comboCard.locator(".chart-preview-line-marker")).toHaveCount(0);
  await expect(comboCard.locator(".chart-preview-bar")).toHaveCount(8);
  await expect(comboCard.locator(".chart-preview-legend")).not.toContainText("Маржа");
});

test("@visual chart preview cards stay visually stable", async ({ page }) => {
  await openStructureDrawer(page);

  const columnCard = page.getByTestId("assessment-card-table_column");
  await page.getByTestId("mode-chart-table_column").click();
  await expect(columnCard).toHaveScreenshot("chart-preview-column-negative.png");

  const lineCard = page.getByTestId("assessment-card-table_line");
  await page.getByTestId("mode-chart-table_line").click();
  await expect(lineCard).toHaveScreenshot("chart-preview-line-dense.png");

  const comboCard = page.getByTestId("assessment-card-table_combo");
  await page.getByTestId("mode-chart-table_combo").click();
  await expect(comboCard).toHaveScreenshot("chart-preview-combo-mixed-units.png");

  await page.getByTestId("series-toggle-table_combo-Маржа").click();
  await expect(comboCard).toHaveScreenshot("chart-preview-combo-fallback-hidden-line.png");
});
