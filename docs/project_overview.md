# Обзор проекта A3 Presentation

## Назначение

Этот проект представляет собой специализированный сервис, который превращает структурированные бизнес-документы в брендированные PowerPoint-презентации.

Целевой workflow:

1. Пользователь загружает `docx` или другой текстовый документ.
2. Система читает структуру документа.
3. Система понимает, где в документе находятся:
   - title
   - section headings
   - subheadings
   - paragraphs
   - bullet lists
   - tables
   - images
4. Planner превращает эту структуру в slide plan.
5. PowerPoint generator рендерит план в брендированный `.pptx` с использованием корпоративного шаблона.

Это не универсальный инструмент «для любых презентаций».
Проект намеренно сфокусирован на структурированных бизнес-колодах, ограниченном наборе slide types и template-aware rendering pipeline.
Текущий встроенный корпоративный шаблон важен, но его нельзя считать вечной константой, потому что пользователи и компании смогут загружать свои шаблоны.

## Технологический стек

### Backend

- Python
- FastAPI
- Pydantic
- `python-pptx`
- `python-docx`
- `pypdf`

### Frontend

- React 19
- Vite
- TypeScript
- `shadcn/ui`
- Tailwind-подход через `shadcn/ui`

### Хранилище

- локальная файловая система
- шаблоны в `storage/templates`
- сгенерированные презентации в `storage/outputs`
- production outputs на Timeweb в `data/outputs`

## Основная структура проекта

```text
src/a3presentation/
  api/                  FastAPI routes
  domain/               Pydantic-модели для API, шаблонов и presentation plan
  services/             Основная бизнес-логика
    chart_render_contract.py
    chart_style.py
    deck_audit.py
    document_text_extractor.py
    layout_capacity.py
    planner.py
    pptx_generator.py
    semantic_normalizer.py
    table_chart_analyzer.py
    template_analyzer.py
    template_registry.py
  settings.py           Пути проекта

frontend/               React UI
docs/                   Внутренняя документация
storage/
  templates/            PowerPoint templates + manifests
  outputs/              Generated presentations
```

## Сквозной pipeline

Текущая карта backend/frontend-процесса выглядит так:

1. UI flow начинается в [App.tsx](../frontend/src/App.tsx) и вызывает [api.ts](../frontend/src/api.ts).
2. Template registry и template analysis обрабатываются в [template_registry.py](../src/a3presentation/services/template_registry.py) и [template_analyzer.py](../src/a3presentation/services/template_analyzer.py).
3. Исходные документы извлекаются [document_text_extractor.py](../src/a3presentation/services/document_text_extractor.py) в text, ordered document blocks, tables, hyperlinks и image blocks.
4. Извлечённые таблицы анализируются в [table_chart_analyzer.py](../src/a3presentation/services/table_chart_analyzer.py); chartable candidates могут вернуться в pipeline как chart overrides.
5. [planner.py](../src/a3presentation/services/planner.py) строит `PresentationPlan` на основе source blocks, tables, chart overrides и [semantic_normalizer.py](../src/a3presentation/services/semantic_normalizer.py).
6. [pptx_generator.py](../src/a3presentation/services/pptx_generator.py) рендерит план против активного `TemplateManifest` и исходного `.pptx`.
7. [deck_audit.py](../src/a3presentation/services/deck_audit.py) и [layout_capacity.py](../src/a3presentation/services/layout_capacity.py) держат planner/generator согласованными по capacity, geometry и mixed-content order.
8. [run_quality_contracts.py](../scripts/run_quality_contracts.py) является выделенным quality gate для generated decks.
9. Production delivery обеспечивается GitHub Actions, [deploy_server.sh](../scripts/deploy_server.sh) и [docker-compose.server.yml](../docker-compose.server.yml).
10. Для второго шага review source of truth должен проходить через backend render pipeline из [slide-review-render-contract.md](slide-review-render-contract.md): `plan -> pptx -> slide previews`, а не через локальную HTML-аппроксимацию.

## 1. Извлечение документа

Главный файл:
- [document_text_extractor.py](../src/a3presentation/services/document_text_extractor.py)

Extractor читает входные документы и превращает их в структурированные blocks.

Для `docx` он сохраняет порядок документа и выдаёт:

- `title`
- `heading`
- `subheading`
- `paragraph`
- `list`
- `table`
- `image`

Ключевой момент в том, что extractor не читает документ как просто plain text.
Он старается сохранить семантику документа:

- определение heading по Word styles
- определение list через:
  - list styles
  - numbering properties (`numPr`)
  - numbering, унаследованный от style
  - hanging indent, типичный для Word lists
  - явные list-like text prefixes
- извлечение таблиц как структурированных headers и rows
- сохранение image blocks для semantic planning

## 2. Планирование слайдов

Главный файл:
- [planner.py](../src/a3presentation/services/planner.py)

Связанные файлы:
- [semantic_normalizer.py](../src/a3presentation/services/semantic_normalizer.py)
- [layout_capacity.py](../src/a3presentation/services/layout_capacity.py)
- [table_chart_analyzer.py](../src/a3presentation/services/table_chart_analyzer.py)

Planner превращает document blocks в `PresentationPlan`.

Ключевая идея:
- не `block -> slide`
- а `section -> 1..N slides`

Sections строятся в основном из headings и subheadings.
Planner старается держать связанный контент вместе на одном слайде, если он помещается.

Поддерживаемые логические slide outcomes:

- `cover`
- `text_full_width`
- `list_full_width`
- `table`
- `chart`
- `image_text`
- `cards_3`
- `contacts`

API contract для этого этапа описан в [presentation.py](../src/a3presentation/domain/presentation.py), а frontend mirror лежит в [types.ts](../frontend/src/types.ts).

Что уже реализовано в planner:

- первая страница документа становится cover slide
- cover title строится из leading lines до первого настоящего section
- обычные lists отправляются в `list_full_width`
- blue-card layouts не используются для обычных bullet lists
- `cards_3` разрешён только для очень короткого label-like content
- длинные sections делятся только когда это действительно нужно
- крошечные text tails не выделяются в отдельные бессмысленные fragments
- порядок mixed `paragraph -> list -> paragraph` сохраняется на основном пути
- top-level headings не теряются
- компактные таблицы остаются на одном слайде, большие таблицы пагинируются
- chart overrides могут заменять table slides на chart slides
- semantic image blocks могут становиться image slides
- первая настоящая numbered section защищена от cover heuristics
- dense narrative continuation path теперь может выбирать `dense_text_full_width`, когда это уменьшает число continuation slides без нарушения quality bounds

## 3. Разрешение шаблонов

Главные файлы:
- [template_registry.py](../src/a3presentation/services/template_registry.py)
- [template_analyzer.py](../src/a3presentation/services/template_analyzer.py)
- [manifest.json](../storage/templates/corp_light_v1/manifest.json)

Сейчас проект поставляется с основным встроенным шаблоном:

- `corp_light_v1`

Но архитектура движется в сторону template-aware generation, а не жёсткой привязки к одному шаблону навсегда.
Пользователи и компании смогут загружать свои шаблоны, а analyzer/manifest metadata должны управлять и generator, и audit.

Система выбирает не физические layout'ы напрямую, а логические slide types, после чего резолвит их против активного template manifest.

## 4. Генерация PowerPoint

Главный файл:
- [pptx_generator.py](../src/a3presentation/services/pptx_generator.py)

Связанные файлы:
- [chart_render_contract.py](../src/a3presentation/services/chart_render_contract.py)
- [deck_audit.py](../src/a3presentation/services/deck_audit.py)
- [chart_style.py](../src/a3presentation/services/chart_style.py)
- [layout_capacity.py](../src/a3presentation/services/layout_capacity.py)

Generator рендерит `PresentationPlan` в реальный `.pptx`.

Важное архитектурное решение:

- проект ушёл от unsafe manual slide cloning как основного пути
- стабильный режим теперь основан на layout-based generation через PowerPoint masters и layouts

Это важно, потому что `.pptx` является Open XML package, и наивное клонирование слайдов может повредить:

- relationships
- media references
- layout links
- внутреннюю структуру пакета

После сохранения generator валидирует output:

- ZIP integrity
- отсутствие duplicate entries
- повторное открытие файла через `python-pptx`

Если выходной файл невалиден, он удаляется, а не возвращается молча.

## 5. Графики

Текущее поведение chart pipeline:

- chartable tables могут превращаться в реальные chart slides
- analyzer предлагает pie для composition tables и stacked charts для безопасных same-unit multi-series tables
- mixed-unit default candidates ограничены `column/line`; combo остаётся поддержанным только для explicit specs и legacy plans
- transpose разрешён только для chart specs, которые остаются семантически корректными после обмена series и categories
- unsafe mixed-unit tables и ordinal/status `1..N` tables считаются not chartable
- summary rows отфильтровываются до построения chart series
- column, bar, line, stacked, pie и explicit combo покрыты generator XML tests
- rendered charts валидируются в regression, API, generator, deck-audit и frontend smoke tests
- chart slides используют тот же layout-quality contract, что и остальная колода
- chart audit валидирует rendered type, series count, combo structure, размеры title/subtitle и compact value-axis number formats
- generator и deck audit теперь используют общий backend chart render contract для visible-series и combo-fallback semantics
- mixed-unit combo scenarios теперь могут использовать secondary value axis вместо принудительного single-axis render
- dense text continuation flow теперь не только объявлен в planner, но и реально доходит до runtime `.pptx` через `dense_text_full_width`

## 6. Frontend

Главные frontend-файлы:
- [App.tsx](../frontend/src/App.tsx)
- [api.ts](../frontend/src/api.ts)
- [chart-preview.tsx](../frontend/src/components/chart-preview.tsx)

Ответственность frontend:

- загрузка исходного документа
- загрузка/обновление шаблона при необходимости
- запуск построения плана
- запуск генерации PPTX
- скачивание результата
- review извлечённой структуры перед генерацией
- переключение chartable tables между `table` и `chart`
- preview поддерживаемых chart layouts до генерации
- сохранение выбранного chart type и hidden series в `chart_overrides`
- display backend-rendered previews на втором шаге после генерации

## Текущие слои проверки

- unit и regression backend tests
- API contract tests
- planner/generator compatibility tests
- deck-level quality-contract tests
- отдельный runner `quality-contracts`
- frontend `playwright` smoke и visual checks
- chart preview smoke matrix
- свежая runtime-генерация `.pptx` обязательна после значимых правок planner/generator/audit и уже используется для dense-text и template-aware verification
- `deck_audit` теперь контролирует не только overlap, но и аномально большой `title/subtitle -> body` gap, чтобы пустые провалы в layout не проходили как зелёный runtime

## Текущее production runtime

- host nginx на Ubuntu завершает HTTPS для `a3presentation.ru`
- host nginx проксирует в docker nginx на `127.0.0.1:8080`
- docker nginx маршрутизирует `/` на frontend и `/api/*` на backend
- backend читает bundled templates из `/app/storage/templates`
- runtime outputs сохраняются в `data/outputs`
- push в `dev` может авто-деплоиться через GitHub Actions после прохождения всех checks

## Текущие рабочие правила

Проект должен развиваться при следующих ограничениях:

- исправления должны целиться в общий механизм, а не в один документ или один слайд
- любые значимые изменения planner/generator/audit должны проверяться по классам документов, а не по одному regression case
- активный шаблон нельзя считать навсегда фиксированным
- если следующий implementation step безопасен и прямо следует из текущей задачи, его нужно выполнять без лишних остановок

## Что ещё улучшать

Это уже не core blockers, но это даст следующий прирост качества:

- дальше расширять deck-audit для более тонких layout-specific geometry rules
- расширять visual snapshots для frontend и generated-slide scenarios
- вводить template-specific typography rules на уровне layout
- расширять `TemplateManifest.component_styles` дальше до полного component grammar layer; runtime-охват уже есть для `cards/text/table/chart/image/cover/list_with_icons/contacts`, включая geometry/behavior contracts для `table/chart/image/cover/list_with_icons/contacts`, и следующий шаг - переносить туда остальные PowerPoint-компоненты и layout-specific rules
- добавлять export previews
- расширять parity и quality checks вокруг secondary value axis и mixed-unit chart scenarios

## Краткое резюме

Проект A3 Presentation это document-to-presentation engine, построенный вокруг брендированного PowerPoint-стиля.

Его основная архитектура:

`docx -> structured blocks -> slide plan -> branded pptx`

Главная ценность выполненной работы не просто в том, что проект «генерирует слайды», а в том, что он уже понимает:

- что такое title
- что такое section
- что такое paragraph
- что такое настоящий list
- что такое table
- какой corporate slide type должен использоваться в каждом случае

Поэтому текущий результат уже существенно ближе к рабочему продукту, чем исходный MVP.
