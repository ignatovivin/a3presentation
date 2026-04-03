import { ChangeEvent, useEffect, useState, useTransition } from "react";

import { buildDownloadUrl, buildPlan, extractTextFromDocument, fetchTemplates, generatePresentation } from "@/api";
import type { DocumentBlock, GeneratePresentationResponse, PresentationPlan, TableBlock, TemplateSummary } from "@/types";

const PRIMARY_TEMPLATE_ID = "corp_light_v1";

const initialText = `Вставьте текст или загрузите документ в формате docx для презентации`;

export function App() {
  const [templates, setTemplates] = useState<TemplateSummary[]>([]);
  const [selectedTemplateId, setSelectedTemplateId] = useState(PRIMARY_TEMPLATE_ID);
  const [rawText, setRawText] = useState("");
  const [documentTables, setDocumentTables] = useState<TableBlock[]>([]);
  const [documentBlocks, setDocumentBlocks] = useState<DocumentBlock[]>([]);
  const [generationResult, setGenerationResult] = useState<GeneratePresentationResponse | null>(null);
  const [error, setError] = useState("");
  const [showLoadingNotice, setShowLoadingNotice] = useState(true);
  const [isPending, startTransition] = useTransition();

  useEffect(() => {
    startTransition(() => {
      fetchTemplates()
        .then((items) => {
          setTemplates(items);
          setShowLoadingNotice(false);
          if (items.some((item) => item.template_id === PRIMARY_TEMPLATE_ID)) {
            setSelectedTemplateId(PRIMARY_TEMPLATE_ID);
            return;
          }
          if (items[0]?.template_id) {
            setSelectedTemplateId(items[0].template_id);
          }
        })
        .catch((err: Error) => {
          setShowLoadingNotice(false);
          setError(err.message);
        });
    });
  }, []);

  function handleTextFileUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    setError("");
    setGenerationResult(null);
    startTransition(() => {
      extractTextFromDocument(file)
        .then((result) => {
          setRawText(result.text);
          setDocumentTables(result.tables);
          setDocumentBlocks(result.blocks);
        })
        .catch((err: Error) => {
          setError(err.message);
        });
    });

    event.target.value = "";
  }

  function handleGenerate() {
    const normalizedText = rawText.trim();
    if (!normalizedText) {
      setError("Введите текст или загрузите документ.");
      return;
    }

    setError("");
    setGenerationResult(null);
    startTransition(() => {
      buildPlan({
        template_id: selectedTemplateId || PRIMARY_TEMPLATE_ID,
        raw_text: normalizedText,
        title: "A3 Presentation",
        tables: documentTables,
        blocks: documentBlocks,
      })
        .then((plan: PresentationPlan) => generatePresentation(plan))
        .then((result) => setGenerationResult(result))
        .catch((err: Error) => setError(err.message));
    });
  }

  return (
    <main className="app-shell">
      <section className="hero-block" data-node-id="634:1739">
        <h1 className="hero-title" data-node-id="633:1698">
          A3 Presentation
        </h1>
        <p className="hero-description" data-node-id="633:3191">
          Превращайте документ в готовую презентацию в корпоративном стиле
          <br />
          Сервис извлекает структуру документа, собирает план слайдов и генерирует без ручной верстки.
        </p>
      </section>

      <div className="composer-stack">
        {generationResult ? (
          <div className="status-panel status-success">
            <button type="button" className="status-close" onClick={() => setGenerationResult(null)} aria-label="Закрыть">
              ×
            </button>
            <div className="status-title">Презентация готова</div>
            <div className="status-text">{generationResult.file_name}</div>
            <button
              type="button"
              className="primary-button status-download"
              onClick={() => window.open(buildDownloadUrl(generationResult.download_url), "_blank", "noopener,noreferrer")}
            >
              Скачать презентацию
            </button>
          </div>
        ) : null}

        {error ? (
          <div className="status-panel status-error">
            <button type="button" className="status-close" onClick={() => setError("")} aria-label="Закрыть">
              ×
            </button>
            <div className="status-title">Ошибка</div>
            <div className="status-text">{error}</div>
          </div>
        ) : null}

        {templates.length === 0 && !error && showLoadingNotice ? (
          <div className="status-panel status-muted">
            <button type="button" className="status-close" onClick={() => setShowLoadingNotice(false)} aria-label="Скрыть">
              ×
            </button>
            Загрузка конфигурации шаблона...
          </div>
        ) : null}

        <section className="composer-card" data-node-id="633:1701">
          <div className="composer-inner" data-node-id="634:1782">
            <div className="textarea-wrap" data-node-id="633:3164">
              <textarea
                value={rawText}
                onChange={(event) => setRawText(event.target.value)}
                placeholder={initialText}
                className="main-textarea"
                rows={10}
              />
            </div>

            <div className="actions-row" data-node-id="634:1772">
              <label className="secondary-button" data-node-id="633:2618">
                <input
                  type="file"
                  accept=".txt,.md,.markdown,.pdf,.docx,text/plain,text/markdown,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                  className="sr-only"
                  onChange={handleTextFileUpload}
                />
                <span>Загрузить документ</span>
              </label>

              <button
                type="button"
                className="primary-button"
                data-node-id="634:1743"
                onClick={handleGenerate}
                disabled={isPending}
              >
                {isPending ? "Генерация..." : "Сгенерировать"}
              </button>
            </div>
          </div>
        </section>
      </div>
    </main>
  );
}
