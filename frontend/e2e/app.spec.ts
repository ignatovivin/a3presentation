import { expect, test } from "@playwright/test";

const templatesResponse = [
  {
    template_id: "corp_light_v1",
    display_name: "Light Theme",
    description: "Corporate template",
  },
];

const extractResponse = {
  file_name: "sample.docx",
  text: "А3\nБизнес-стратегия 2026\n1. Рост\nКомпания растет за счет новых сегментов.",
  tables: [
    {
      headers: ["Канал", "Q1", "Q2"],
      rows: [
        ["SMB", "120", "140"],
        ["Enterprise", "220", "260"],
      ],
    },
  ],
  blocks: [
    { kind: "title", text: "А3", items: [] },
    { kind: "heading", text: "1. Рост", level: 1, items: [] },
    { kind: "paragraph", text: "Компания растет за счет новых сегментов.", items: [] },
  ],
  chart_assessments: [
    {
      table_id: "table_1",
      chartable: true,
      classification: "multi_series_category",
      confidence: "high",
      reasons: ["detected multiple numeric series with categorical labels"],
      warnings: [],
      candidate_specs: [
        {
          chart_id: "chart_1",
          source_table_id: "table_1",
          chart_type: "column",
          title: "Рост по каналам",
          categories: ["Q1", "Q2"],
          series: [
            { name: "SMB", values: [120, 140], unit: null, axis: "primary", hidden: false },
            { name: "Enterprise", values: [220, 260], unit: null, axis: "primary", hidden: false },
          ],
          x_axis_title: null,
          y_axis_title: null,
          legend_visible: true,
          data_labels_visible: false,
          value_format: "number",
          stacking: "none",
          confidence: "high",
          warnings: [],
        },
        {
          chart_id: "chart_2",
          source_table_id: "table_1",
          chart_type: "line",
          title: "Рост по каналам",
          categories: ["Q1", "Q2"],
          series: [
            { name: "SMB", values: [120, 140], unit: null, axis: "primary", hidden: false },
            { name: "Enterprise", values: [220, 260], unit: null, axis: "primary", hidden: false },
          ],
          x_axis_title: null,
          y_axis_title: null,
          legend_visible: true,
          data_labels_visible: false,
          value_format: "number",
          stacking: "none",
          confidence: "high",
          warnings: [],
        },
      ],
      structured_table: {
        table_id: "table_1",
        header_rows: [0],
        label_columns: [0],
        numeric_columns: [1, 2],
        time_columns: [],
        data_start_row: 1,
        summary_rows: [],
        warnings: [],
        cells: [
          [
            { text: "Канал", normalized_text: "канал", value_type: "text", unit: null, annotation: null, is_header_like: true },
            { text: "Q1", normalized_text: "q1", value_type: "text", unit: null, annotation: null, is_header_like: true },
            { text: "Q2", normalized_text: "q2", value_type: "text", unit: null, annotation: null, is_header_like: true },
          ],
          [
            { text: "SMB", normalized_text: "smb", value_type: "text", unit: null, annotation: null, is_header_like: false },
            { text: "120", normalized_text: "120", value_type: "number", parsed_value: 120, unit: null, annotation: null, is_header_like: false },
            { text: "140", normalized_text: "140", value_type: "number", parsed_value: 140, unit: null, annotation: null, is_header_like: false },
          ],
          [
            { text: "Enterprise", normalized_text: "enterprise", value_type: "text", unit: null, annotation: null, is_header_like: false },
            { text: "220", normalized_text: "220", value_type: "number", parsed_value: 220, unit: null, annotation: null, is_header_like: false },
            { text: "260", normalized_text: "260", value_type: "number", parsed_value: 260, unit: null, annotation: null, is_header_like: false },
          ],
        ],
      },
    },
  ],
};

const planResponse = {
  template_id: "corp_light_v1",
  title: "A3 Presentation",
  slides: [
    { kind: "title", title: "A3 Presentation", bullets: [], left_bullets: [], right_bullets: [], preferred_layout_key: "cover" },
    { kind: "text", title: "1. Рост", text: "Компания растет за счет новых сегментов.", bullets: [], left_bullets: [], right_bullets: [], preferred_layout_key: "text_full_width" },
    {
      kind: "chart",
      title: "Рост по каналам",
      bullets: [],
      left_bullets: [],
      right_bullets: [],
      preferred_layout_key: "table",
      chart: extractResponse.chart_assessments[0].candidate_specs[0],
      source_table_id: "table_1",
    },
  ],
};

const generationResponse = {
  output_path: "/tmp/A3_Presentation.pptx",
  file_name: "A3_Presentation.pptx",
  download_url: "/presentations/files/A3_Presentation.pptx",
};

test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    (window as Window & { __lastOpenedUrl?: string }).__lastOpenedUrl = "";
    window.open = ((url?: string | URL | undefined) => {
      (window as Window & { __lastOpenedUrl?: string }).__lastOpenedUrl = url?.toString() ?? "";
      return null;
    }) as typeof window.open;
  });

  await page.route("**/api/templates", async (route) => {
    await route.fulfill({ json: templatesResponse });
  });

  await page.route("**/api/documents/extract-text", async (route) => {
    await route.fulfill({ json: extractResponse });
  });

  await page.route("**/api/plans/from-text", async (route) => {
    await route.fulfill({ json: planResponse });
  });

  await page.route("**/api/presentations/generate", async (route) => {
    await route.fulfill({ json: generationResponse });
  });

  await page.route("**/api/presentations/files/*", async (route) => {
    await route.fulfill({
      status: 200,
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      body: "mock-pptx",
    });
  });
});

test("@smoke user can upload document inspect structure and generate presentation", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByTestId("app-shell")).toBeVisible();

  await page.getByTestId("upload-document-input").setInputFiles({
    name: "sample.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from("mock-docx"),
  });

  await expect(page.getByTestId("raw-text-input")).toHaveValue(/Бизнес-стратегия 2026/);
  await expect(page.getByTestId("open-structure-drawer")).toBeVisible();

  await page.getByTestId("open-structure-drawer").click();
  await expect(page.getByTestId("structure-drawer")).toBeVisible();
  await expect(page.getByTestId("assessment-card-table_1")).toBeVisible();

  await page.getByTestId("mode-chart-table_1").click();
  await page.getByTestId("chart-type-table_1").selectOption("chart_2");
  await page.getByTestId("series-toggle-table_1-Enterprise").click();
  await page.getByTestId("save-structure-choices").click();

  await page.getByTestId("generate-presentation").click();

  await expect(page.getByTestId("generation-success")).toBeVisible();
  await expect(page.getByTestId("generated-file-name")).toHaveText("A3_Presentation.pptx");

  await page.getByTestId("download-presentation").click();
  await expect.poll(async () => {
    return page.evaluate(() => (window as Window & { __lastOpenedUrl?: string }).__lastOpenedUrl ?? "");
  }).toContain("/api/presentations/files/A3_Presentation.pptx");
});

test("@visual main screen stays visually stable", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-shell")).toHaveScreenshot("app-shell.png");
});

test("@visual structure drawer stays visually stable", async ({ page }) => {
  await page.goto("/");

  await page.getByTestId("upload-document-input").setInputFiles({
    name: "sample.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from("mock-docx"),
  });
  await page.getByTestId("open-structure-drawer").click();

  await expect(page.getByTestId("structure-drawer")).toHaveScreenshot("structure-drawer.png");
});
