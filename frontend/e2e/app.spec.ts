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
    {
      headers: ["Этап", "Комментарий"],
      rows: [
        ["Discovery", "Интервью и сбор требований"],
        ["Prototype", "Проверка решения"],
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
          chart_id: "chart_1_column_compare",
          source_table_id: "table_1",
          chart_type: "column",
          variant_label: "Сравнение: SMB, Enterprise",
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
          transpose_allowed: true,
        },
        {
          chart_id: "chart_1_column_1",
          source_table_id: "table_1",
          chart_type: "column",
          variant_label: "Единичный: SMB",
          title: "Рост по каналам",
          categories: ["Q1", "Q2"],
          series: [{ name: "SMB", values: [120, 140], unit: null, axis: "primary", hidden: false }],
          x_axis_title: null,
          y_axis_title: "SMB",
          legend_visible: false,
          data_labels_visible: false,
          value_format: "number",
          stacking: "none",
          confidence: "high",
          warnings: [],
          transpose_allowed: false,
        },
        {
          chart_id: "chart_1_column_2",
          source_table_id: "table_1",
          chart_type: "column",
          variant_label: "Единичный: Enterprise",
          title: "Рост по каналам",
          categories: ["Q1", "Q2"],
          series: [{ name: "Enterprise", values: [220, 260], unit: null, axis: "primary", hidden: false }],
          x_axis_title: null,
          y_axis_title: "Enterprise",
          legend_visible: false,
          data_labels_visible: false,
          value_format: "number",
          stacking: "none",
          confidence: "high",
          warnings: [],
          transpose_allowed: false,
        },
        {
          chart_id: "chart_1_line_compare",
          source_table_id: "table_1",
          chart_type: "line",
          variant_label: "Сравнение: SMB, Enterprise",
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
          transpose_allowed: true,
        },
        {
          chart_id: "chart_1_line_1",
          source_table_id: "table_1",
          chart_type: "line",
          variant_label: "Единичный: SMB",
          title: "Рост по каналам",
          categories: ["Q1", "Q2"],
          series: [{ name: "SMB", values: [120, 140], unit: null, axis: "primary", hidden: false }],
          x_axis_title: null,
          y_axis_title: "SMB",
          legend_visible: false,
          data_labels_visible: false,
          value_format: "number",
          stacking: "none",
          confidence: "high",
          warnings: [],
          transpose_allowed: false,
        },
        {
          chart_id: "chart_1_line_2",
          source_table_id: "table_1",
          chart_type: "line",
          variant_label: "Единичный: Enterprise",
          title: "Рост по каналам",
          categories: ["Q1", "Q2"],
          series: [{ name: "Enterprise", values: [220, 260], unit: null, axis: "primary", hidden: false }],
          x_axis_title: null,
          y_axis_title: "Enterprise",
          legend_visible: false,
          data_labels_visible: false,
          value_format: "number",
          stacking: "none",
          confidence: "high",
          warnings: [],
          transpose_allowed: false,
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
    {
      table_id: "table_2",
      chartable: false,
      classification: "text_dominant",
      confidence: "none",
      reasons: ["table has no numeric columns"],
      warnings: [],
      candidate_specs: [],
      structured_table: {
        table_id: "table_2",
        header_rows: [0],
        label_columns: [0, 1],
        numeric_columns: [],
        time_columns: [],
        data_start_row: 1,
        summary_rows: [],
        warnings: [],
        cells: [
          [
            { text: "Этап", normalized_text: "этап", value_type: "text", unit: null, annotation: null, is_header_like: true },
            { text: "Комментарий", normalized_text: "комментарий", value_type: "text", unit: null, annotation: null, is_header_like: true },
          ],
          [
            { text: "Discovery", normalized_text: "discovery", value_type: "text", unit: null, annotation: null, is_header_like: false },
            { text: "Интервью и сбор требований", normalized_text: "интервью и сбор требований", value_type: "text", unit: null, annotation: null, is_header_like: false },
          ],
        ],
      },
    },
    {
      table_id: "table_3",
      chartable: true,
      classification: "time_series",
      confidence: "low",
      reasons: ["mixed units detected", "generated 10 chart candidates"],
      warnings: ["mixed-unit chart uses a secondary value axis"],
      candidate_specs: [
        {
          chart_id: "chart_3_column_1",
          source_table_id: "table_3",
          chart_type: "column",
          variant_label: "Сравнение: Выручка, Затраты",
          title: "Выручка и маржа",
          categories: ["Q1", "Q2"],
          series: [
            { name: "Выручка", values: [120000000, 150000000], unit: "RUB", axis: "primary", hidden: false },
            { name: "Затраты", values: [80000000, 90000000], unit: "RUB", axis: "primary", hidden: false },
          ],
          x_axis_title: null,
          y_axis_title: null,
          legend_visible: true,
          data_labels_visible: false,
          value_format: "currency",
          stacking: "none",
          confidence: "low",
          warnings: ["mixed-unit chart uses a secondary value axis"],
          transpose_allowed: true,
        },
        {
          chart_id: "chart_3_column_2",
          source_table_id: "table_3",
          chart_type: "column",
          variant_label: "Единичный: Выручка",
          title: "Выручка и маржа",
          categories: ["Q1", "Q2"],
          series: [
            { name: "Выручка", values: [120000000, 150000000], unit: "RUB", axis: "primary", hidden: false },
          ],
          x_axis_title: null,
          y_axis_title: "Выручка",
          legend_visible: false,
          data_labels_visible: false,
          value_format: "currency",
          stacking: "none",
          confidence: "low",
          warnings: ["mixed-unit chart uses a secondary value axis"],
          transpose_allowed: false,
        },
        {
          chart_id: "chart_3_column_3",
          source_table_id: "table_3",
          chart_type: "column",
          variant_label: "Единичный: Затраты",
          title: "Выручка и маржа",
          categories: ["Q1", "Q2"],
          series: [
            { name: "Затраты", values: [80000000, 90000000], unit: "RUB", axis: "primary", hidden: false },
          ],
          x_axis_title: null,
          y_axis_title: "Затраты",
          legend_visible: false,
          data_labels_visible: false,
          value_format: "currency",
          stacking: "none",
          confidence: "low",
          warnings: ["mixed-unit chart uses a secondary value axis"],
          transpose_allowed: false,
        },
        {
          chart_id: "chart_3_column_4",
          source_table_id: "table_3",
          chart_type: "column",
          variant_label: "Единичный: Маржа",
          title: "Выручка и маржа",
          categories: ["Q1", "Q2"],
          series: [
            { name: "Маржа", values: [18, 22], unit: "%", axis: "primary", hidden: false },
          ],
          x_axis_title: null,
          y_axis_title: "Маржа",
          legend_visible: false,
          data_labels_visible: false,
          value_format: "percent",
          stacking: "none",
          confidence: "low",
          warnings: ["mixed-unit chart uses a secondary value axis"],
          transpose_allowed: false,
        },
        {
          chart_id: "chart_3_line_1",
          source_table_id: "table_3",
          chart_type: "line",
          variant_label: "Сравнение: Выручка, Затраты",
          title: "Выручка и маржа",
          categories: ["Q1", "Q2"],
          series: [
            { name: "Выручка", values: [120000000, 150000000], unit: "RUB", axis: "primary", hidden: false },
            { name: "Затраты", values: [80000000, 90000000], unit: "RUB", axis: "primary", hidden: false },
          ],
          x_axis_title: null,
          y_axis_title: null,
          legend_visible: true,
          data_labels_visible: false,
          value_format: "currency",
          stacking: "none",
          confidence: "low",
          warnings: ["mixed-unit chart uses a secondary value axis"],
          transpose_allowed: true,
        },
        {
          chart_id: "chart_3_line_2",
          source_table_id: "table_3",
          chart_type: "line",
          variant_label: "Единичный: Выручка",
          title: "Выручка и маржа",
          categories: ["Q1", "Q2"],
          series: [
            { name: "Выручка", values: [120000000, 150000000], unit: "RUB", axis: "primary", hidden: false },
          ],
          x_axis_title: null,
          y_axis_title: "Выручка",
          legend_visible: false,
          data_labels_visible: false,
          value_format: "currency",
          stacking: "none",
          confidence: "low",
          warnings: ["mixed-unit chart uses a secondary value axis"],
          transpose_allowed: false,
        },
        {
          chart_id: "chart_3_line_3",
          source_table_id: "table_3",
          chart_type: "line",
          variant_label: "Единичный: Затраты",
          title: "Выручка и маржа",
          categories: ["Q1", "Q2"],
          series: [
            { name: "Затраты", values: [80000000, 90000000], unit: "RUB", axis: "primary", hidden: false },
          ],
          x_axis_title: null,
          y_axis_title: "Затраты",
          legend_visible: false,
          data_labels_visible: false,
          value_format: "currency",
          stacking: "none",
          confidence: "low",
          warnings: ["mixed-unit chart uses a secondary value axis"],
          transpose_allowed: false,
        },
        {
          chart_id: "chart_3_line_4",
          source_table_id: "table_3",
          chart_type: "line",
          variant_label: "Единичный: Маржа",
          title: "Выручка и маржа",
          categories: ["Q1", "Q2"],
          series: [
            { name: "Маржа", values: [18, 22], unit: "%", axis: "primary", hidden: false },
          ],
          x_axis_title: null,
          y_axis_title: "Маржа",
          legend_visible: false,
          data_labels_visible: false,
          value_format: "percent",
          stacking: "none",
          confidence: "low",
          warnings: ["mixed-unit chart uses a secondary value axis"],
          transpose_allowed: false,
        },
        {
          chart_id: "chart_3_combo_1",
          source_table_id: "table_3",
          chart_type: "combo",
          variant_label: "Комбинированный: столбцы Выручка, Затраты; линия Маржа",
          title: "Выручка и маржа",
          categories: ["Q1", "Q2"],
          series: [
            { name: "Выручка", values: [120000000, 150000000], unit: "RUB", axis: "primary", hidden: false },
            { name: "Затраты", values: [80000000, 90000000], unit: "RUB", axis: "primary", hidden: false },
            { name: "Маржа", values: [18, 22], unit: "%", axis: "secondary", hidden: false },
          ],
          x_axis_title: null,
          y_axis_title: null,
          legend_visible: true,
          data_labels_visible: false,
          value_format: "number",
          stacking: "none",
          confidence: "low",
          warnings: ["mixed-unit chart uses a secondary value axis"],
          transpose_allowed: false,
        },
      ],
      structured_table: {
        table_id: "table_3",
        header_rows: [0],
        label_columns: [0],
        numeric_columns: [1, 2],
        time_columns: [],
        data_start_row: 1,
        summary_rows: [],
        warnings: [],
        cells: [
          [
            { text: "Квартал", normalized_text: "квартал", value_type: "text", unit: null, annotation: null, is_header_like: true },
            { text: "Выручка", normalized_text: "выручка", value_type: "text", unit: null, annotation: null, is_header_like: true },
            { text: "Затраты", normalized_text: "затраты", value_type: "text", unit: null, annotation: null, is_header_like: true },
            { text: "Маржа", normalized_text: "маржа", value_type: "text", unit: null, annotation: null, is_header_like: true },
          ],
          [
            { text: "Q1", normalized_text: "q1", value_type: "text", unit: null, annotation: null, is_header_like: false },
            { text: "120 млн руб", normalized_text: "120 млн руб", value_type: "currency", parsed_value: 120000000, unit: "RUB", annotation: null, is_header_like: false },
            { text: "80 млн руб", normalized_text: "80 млн руб", value_type: "currency", parsed_value: 80000000, unit: "RUB", annotation: null, is_header_like: false },
            { text: "18%", normalized_text: "18%", value_type: "percent", parsed_value: 18, unit: "%", annotation: null, is_header_like: false },
          ],
          [
            { text: "Q2", normalized_text: "q2", value_type: "text", unit: null, annotation: null, is_header_like: false },
            { text: "150 млн руб", normalized_text: "150 млн руб", value_type: "currency", parsed_value: 150000000, unit: "RUB", annotation: null, is_header_like: false },
            { text: "90 млн руб", normalized_text: "90 млн руб", value_type: "currency", parsed_value: 90000000, unit: "RUB", annotation: null, is_header_like: false },
            { text: "22%", normalized_text: "22%", value_type: "percent", parsed_value: 22, unit: "%", annotation: null, is_header_like: false },
          ],
        ],
      },
    },
    {
      table_id: "table_4",
      chartable: true,
      classification: "multi_series_category",
      confidence: "low",
      reasons: ["mixed units detected", "generated 6 chart candidates"],
      warnings: ["mixed-unit chart uses a secondary value axis"],
      candidate_specs: [
        {
          chart_id: "chart_4_combo_1",
          source_table_id: "table_4",
          chart_type: "combo",
          variant_label: "Комбинированный: столбцы Объем рынка (2024); линия Доля А3 GMV (2025)",
          title: "Рынок / сегмент · Объем рынка (2024) / Доля А3 GMV (2025)",
          categories: ["ЖКХ", "Налоги", "Образование"],
          series: [
            { name: "Объем рынка (2024)", values: [8496000000000, 55600000000000, 851000000000], unit: null, axis: "primary", hidden: false },
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
          chart_id: "chart_4_combo_2",
          source_table_id: "table_4",
          chart_type: "combo",
          variant_label: "Комбинированный: столбцы Доля А3 GMV (2025); линия Объем рынка (2024)",
          title: "Рынок / сегмент · Объем рынка (2024) / Доля А3 GMV (2025)",
          categories: ["ЖКХ", "Налоги", "Образование"],
          series: [
            { name: "Доля А3 GMV (2025)", values: [2.12, 0.02, 0.06], unit: "%", axis: "primary", hidden: false },
            { name: "Объем рынка (2024)", values: [8496000000000, 55600000000000, 851000000000], unit: null, axis: "secondary", hidden: false },
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
          chart_id: "chart_4_column_1",
          source_table_id: "table_4",
          chart_type: "column",
          variant_label: "Единичный: Объем рынка (2024)",
          title: "Рынок / сегмент · Объем рынка (2024) / Доля А3 GMV (2025)",
          categories: ["ЖКХ", "Налоги", "Образование"],
          series: [
            { name: "Объем рынка (2024)", values: [8496000000000, 55600000000000, 851000000000], unit: null, axis: "primary", hidden: false },
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
          chart_id: "chart_4_column_2",
          source_table_id: "table_4",
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
        {
          chart_id: "chart_4_line_1",
          source_table_id: "table_4",
          chart_type: "line",
          variant_label: "Единичный: Объем рынка (2024)",
          title: "Рынок / сегмент · Объем рынка (2024) / Доля А3 GMV (2025)",
          categories: ["ЖКХ", "Налоги", "Образование"],
          series: [
            { name: "Объем рынка (2024)", values: [8496000000000, 55600000000000, 851000000000], unit: null, axis: "primary", hidden: false },
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
          chart_id: "chart_4_line_2",
          source_table_id: "table_4",
          chart_type: "line",
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

const planResponse = {
  template_id: "corp_light_v1",
  title: "A3 Presentation",
  slides: [
    { kind: "title", title: "A3 Presentation", bullets: [], left_bullets: [], right_bullets: [], preferred_layout_key: "cover" },
    {
      kind: "text",
      title: "1. Рост",
      text: "Компания растет за счет новых сегментов. Партнерская сеть ускоряет подключение клиентов. Автоматизация снижает стоимость сопровождения.",
      bullets: [],
      content_blocks: [],
      left_bullets: [],
      right_bullets: [],
      preferred_layout_key: "text_full_width",
    },
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

let lastPlanPayload: any = null;
let lastGeneratePayload: any = null;

test.beforeEach(async ({ page }) => {
  lastPlanPayload = null;
  lastGeneratePayload = null;

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
    lastPlanPayload = await route.request().postDataJSON();
    await route.fulfill({ json: planResponse });
  });

  await page.route("**/api/presentations/generate", async (route) => {
    lastGeneratePayload = await route.request().postDataJSON();
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

  await expect(page.getByTestId("raw-text-input")).toHaveValue("");
  await expect(page.getByTestId("attached-document")).toContainText("sample.docx");
  await expect(page.getByTestId("open-structure-drawer")).toBeVisible();

  await page.getByTestId("open-structure-drawer").click();
  await expect(page.getByTestId("structure-drawer")).toBeVisible();
  await expect(page.getByTestId("assessment-card-table_1")).toBeVisible();
  await expect(page.getByTestId("assessment-card-table_1")).toContainText("Рост по каналам");
  await expect(page.getByTestId("assessment-card-table_1")).toContainText("Эту таблицу можно использовать для графика.");
  await expect(page.locator(".drawer-switch")).toContainText("Показать все таблицы");
  await expect(page.getByRole("checkbox", { name: "Показать все таблицы" })).not.toBeChecked();
  await expect(page.getByTestId("assessment-card-table_2")).toHaveCount(0);
  await page.locator(".drawer-switch").click();
  await expect(page.getByRole("checkbox", { name: "Показать все таблицы" })).toBeChecked();
  await expect(page.getByTestId("assessment-card-table_2")).toBeVisible();

  await page.getByTestId("mode-chart-table_1").click();
  await expect(page.getByTestId("chart-type-table_1").locator("option")).toHaveCount(2);
  await expect(page.getByTestId("chart-type-table_1")).toBeVisible();
  await expect(page.getByTestId("chart-orientation-table_1")).toHaveCount(0);
  await page.getByTestId("chart-type-table_1").selectOption("line");
  await expect(page.getByTestId("chart-variant-table_1").locator("option")).toHaveCount(3);
  await expect(page.getByTestId("chart-variant-table_1")).toContainText("Сравнение: SMB, Enterprise");
  await expect(page.getByTestId("chart-variant-table_1")).toContainText("Единичный: SMB");
  await expect(page.getByTestId("chart-variant-table_1")).toContainText("Единичный: Enterprise");
  await page.getByTestId("chart-variant-table_1").selectOption("chart_1_line_compare");
  await expect(page.locator(".chart-preview-line-svg path.chart-preview-line")).toHaveCount(2);
  await expect(page.locator(".chart-preview-line-marker")).toHaveCount(4);
  await expect(page.locator(".chart-preview-line-label")).toHaveCount(2);
  await page.getByTestId("series-toggle-table_1-Enterprise").click();
  await expect(page.locator(".chart-preview-line-svg path.chart-preview-line")).toHaveCount(1);
  await expect(page.locator(".chart-preview-line-marker")).toHaveCount(2);
  await expect(page.locator(".chart-preview-line-label")).toHaveCount(2);
  await page.getByTestId("save-structure-choices").click();

  await page.getByTestId("open-structure-drawer").click();
  await page.getByTestId("mode-chart-table_3").click();
  await expect(page.getByTestId("chart-type-table_3").locator("option")).toHaveCount(2);
  await expect(page.getByTestId("chart-type-table_3")).toContainText("Вертикальные столбцы");
  await expect(page.getByTestId("chart-type-table_3")).toContainText("Линейный график");
  await expect(page.getByTestId("chart-type-table_3")).not.toContainText("Комбинированный график");
  await expect(page.getByTestId("chart-type-table_3")).toHaveValue("column");
  await expect(page.getByTestId("chart-variant-table_3").locator("option")).toHaveCount(5);
  await expect(page.getByTestId("chart-variant-table_3")).toHaveValue("chart_3_column_1");
  await expect(page.getByTestId("chart-variant-table_3")).toContainText("Сравнение: Выручка, Затраты");
  await expect(page.getByTestId("chart-variant-table_3")).toContainText("Сравнение: Выручка, Затраты, Маржа");
  await expect(page.getByTestId("chart-variant-table_3")).toContainText("Единичный: Выручка");
  await expect(page.getByTestId("chart-variant-table_3")).toContainText("Единичный: Затраты");
  await expect(page.getByTestId("chart-variant-table_3")).toContainText("Единичный: Маржа");

  await page.getByTestId("mode-chart-table_4").click();
  await expect(page.getByTestId("chart-type-table_4").locator("option")).toHaveCount(2);
  await expect(page.getByTestId("chart-type-table_4")).toContainText("Вертикальные столбцы");
  await expect(page.getByTestId("chart-type-table_4")).toContainText("Линейный график");
  await expect(page.getByTestId("chart-type-table_4")).not.toContainText("Комбинированный график");
  await expect(page.getByTestId("chart-variant-table_4").locator("option")).toHaveCount(4);
  await expect(page.getByTestId("chart-variant-table_4")).toHaveValue("chart_4_combo_1");
  await expect(page.getByTestId("chart-variant-table_4")).toContainText("Сравнение: Объем рынка (2024), Доля А3 GMV (2025)");
  await expect(page.getByTestId("chart-variant-table_4")).toContainText("Единичный: Объем рынка (2024)");
  await expect(page.getByTestId("chart-variant-table_4")).toContainText("Единичный: Доля А3 GMV (2025)");
  await page.getByTestId("save-structure-choices").click();

  await page.getByTestId("open-structure-drawer").click();
  await page.getByTestId("drawer-tab-text").click();
  await expect(page.getByTestId("slide-review-panel")).toBeVisible();
  await expect(page.getByTestId("drawer-tab-text")).toHaveAttribute("aria-selected", "true");
  await expect(page.getByTestId("card-slide-choice-1")).toContainText("1. Рост");
  await page.getByTestId("card-slide-choice-1").click();
  await page.getByRole("button", { name: "Сбросить выбор" }).click();
  await expect(page.getByTestId("card-slide-choice-1")).toBeVisible();
  await page.getByTestId("card-slide-choice-1").click();
  await expect(page.getByTestId("save-structure-choices")).toHaveText("Сохранить");
  await page.getByTestId("save-structure-choices").click();
  await expect(page.getByTestId("structure-drawer")).toHaveCount(0);
  await expect(page.getByTestId("generate-presentation")).toHaveText("Сгенерировать");
  await page.getByTestId("generate-presentation").click({ force: true });

  await expect(page.getByTestId("generation-success")).toBeVisible();
  const chartOverride = lastPlanPayload.chart_overrides.find((override: any) => override.table_id === "table_1");
  const tableOverride = lastPlanPayload.chart_overrides.find((override: any) => override.table_id === "table_2");
  const mixedChartOverride = lastPlanPayload.chart_overrides.find((override: any) => override.table_id === "table_3");
  const marketShareOverride = lastPlanPayload.chart_overrides.find((override: any) => override.table_id === "table_4");
  expect(chartOverride.mode).toBe("chart");
  expect(tableOverride.mode).toBe("table");
  expect(mixedChartOverride.mode).toBe("chart");
  expect(marketShareOverride.mode).toBe("chart");
  expect(chartOverride.selected_chart.chart_type).toBe("line");
  expect(mixedChartOverride.selected_chart.chart_id).toBe("chart_3_column_1:default");
  expect(mixedChartOverride.selected_chart.variant_label).toBe("Сравнение: Выручка, Затраты");
  expect(marketShareOverride.selected_chart.chart_id).toBe("chart_4_combo_1:default");
  expect(marketShareOverride.selected_chart.chart_type).toBe("combo");
  expect(chartOverride.selected_chart.series).toEqual([
    expect.objectContaining({ name: "SMB", hidden: false }),
    expect.objectContaining({ name: "Enterprise", hidden: true }),
  ]);
  expect(lastGeneratePayload.slides[1]).toEqual(
    expect.objectContaining({
      kind: "bullets",
      preferred_layout_key: "cards_3",
      bullets: [
        "Компания растет за счет новых сегментов.",
        "Партнерская сеть ускоряет подключение клиентов.",
        "Автоматизация снижает стоимость сопровождения.",
      ],
    }),
  );
  await expect(page.getByTestId("generated-file-name")).toHaveText("A3_Presentation.pptx");

  await page.getByTestId("download-presentation").click();
  await expect.poll(async () => {
    return page.evaluate(() => (window as Window & { __lastOpenedUrl?: string }).__lastOpenedUrl ?? "");
  }).toContain("/api/presentations/files/A3_Presentation.pptx");
});

test("attached document can be removed before replacement", async ({ page }) => {
  await page.goto("/");

  const emptyMetrics = await page.locator(".composer-card").evaluate((element) => {
    const rect = element.getBoundingClientRect();
    const action = element.querySelector('[data-testid="upload-document-trigger"]')?.getBoundingClientRect();
    return { composerHeight: rect.height, composerTop: rect.top, actionHeight: action?.height ?? 0 };
  });

  await page.getByTestId("upload-document-input").setInputFiles({
    name: "sample.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from("mock-docx"),
  });

  await expect(page.getByTestId("raw-text-input")).toHaveValue("");
  await expect(page.getByTestId("attached-document")).toContainText("sample.docx");
  const attachedMetrics = await page.locator(".composer-card").evaluate((element) => {
    const rect = element.getBoundingClientRect();
    const action = element.querySelector('[data-testid="attached-document"]')?.getBoundingClientRect();
    return { composerHeight: rect.height, composerTop: rect.top, actionHeight: action?.height ?? 0 };
  });
  expect(attachedMetrics).toEqual(emptyMetrics);

  await page.getByTestId("remove-attached-document").click();

  await expect(page.getByTestId("attached-document")).toHaveCount(0);
  await expect(page.getByTestId("upload-document-trigger")).toBeVisible();
  await expect(page.getByTestId("open-structure-drawer")).toHaveCount(0);
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
