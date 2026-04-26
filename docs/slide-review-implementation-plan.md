# Slide Review Implementation Plan

## Назначение

Этот документ фиксирует план внедрения второго шага с точным фирменным preview и возможностью редактирования.

Он опирается на:

- [slide-review-render-contract.md](slide-review-render-contract.md)
- [new-tooling-rules.md](new-tooling-rules.md)
- [frontend-visual-contracts.md](frontend-visual-contracts.md)
- [quality-contracts.md](quality-contracts.md)

## Целевой результат

После нажатия `Сгенерировать` пользователь должен попадать на второй шаг, где:

1. основной frame показывает backend-rendered slide preview в фирменном стиле
2. слева отображаются thumbnails тех же preview
3. пользователь может менять тип контента и допустимые свойства слайда
4. после изменения запускается controlled regenerate cycle
5. обновлённый backend-rendered result возвращается в тот же экран

## Этапы

### Этап 1. Stabilize rendered preview contract

Статус:

- в работе, базовый контракт уже начат

Задачи:

1. закрепить `preview_fidelity` в API contract
2. добавить `preview_source`
3. гарантировать, что второй шаг открывается на `slides_preview`, а не на local HTML preview
4. показывать явный degraded-state для `fallback_rendered`

Критерий готовности:

- review-step image-first
- UI явно знает, точный это render или fallback

### Этап 2. Harden PowerPoint export pipeline

Задачи:

1. сделать `powerpoint_rendered` основным и проверяемым runtime path
2. добавить диагностику причин fallback
3. ввести contract test на число preview-файлов и режим preview source
4. добавить runtime-проверку на свежем generated `.pptx`

Критерий готовности:

- на рабочей Windows-среде backend стабильно отдаёт `powerpoint_rendered`

### Этап 3. Introduce editable overlay

Задачи:

1. добавить editor overlay поверх image-preview
2. выделить editable-зоны слайда
3. реализовать selection model
4. реализовать базовые edit operations:
   - change slide mode
   - table -> chart
   - chart variant change
   - text/cards/list mode switch

Инструменты:

- `react-konva`
- `ResizeObserver`

Критерий готовности:

- пользователь редактирует содержимое через overlay/controls, не теряя backend-rendered preview

### Этап 4. Add regenerate orchestration

Задачи:

1. debounce regenerate flow
2. добавить `AbortController`
3. защититься от race conditions
4. ввести request/revision id для regenerate cycle

Критерий готовности:

- быстрые изменения не приводят к визуальному прыганию и устаревшим preview

### Этап 5. Persist editable review state

Задачи:

1. сохранять review state
2. сохранять edit state
3. при reload возвращать пользователя в тот же review-step
4. после восстановления запрашивать актуальные preview при необходимости

Критерий готовности:

- hard refresh не возвращает пользователя к шагу 1 и не теряет edits

### Этап 6. Quality and regression coverage

Задачи:

1. backend API tests для `preview_fidelity` / `preview_source`
2. frontend smoke на generated review-screen
3. visual snapshots второго шага
4. runtime generation check на локальном backend
5. включить этот flow в quality review перед переносом в `test`

Критерий готовности:

- pipeline стабильно проверяется автоматикой

## MVP-порядок выполнения

1. Stabilize rendered preview contract
2. Harden PowerPoint export pipeline
3. Add regenerate orchestration
4. Introduce editable overlay
5. Persist editable review state
6. Expand quality coverage

## Ближайшие задачи

### P0

1. добавить `preview_source` в backend/frontend contract
2. возвращать причину fallback из preview service
3. убрать оставшиеся main-preview зависимости от local React slide rendering

### P1

1. ввести request cancellation для regenerate flow
2. сделать editor overlay scaffold
3. подготовить mapping `overlay edits -> plan changes`

### P2

1. drag/resize editable zones
2. richer inline editing
3. расширенный visual regression layer

## Текущий execution plan для editable second step

Цель текущего цикла:

- второй шаг должен работать как сгенерированная фирменная презентация с редактируемыми блоками, а не как набор полей ввода рядом со слайдом

Порядок реализации:

1. описать editable block model на уровне frontend: `title`, `subtitle`, `body`, `table`, `chart`, `cards`
2. связать клики по overlay-зонам с редактированием соответствующего блока прямо поверх preview
3. оставить backend-rendered PNG preview source of truth, а frontend-редактор использовать только как слой правки
4. любые изменения блока маппить обратно в `PresentationPlan`
5. после изменения запускать controlled regenerate cycle через существующий `AbortController`/request-id flow
6. расширить smoke-тесты на редактирование блоков и переключение типа контента по всем слайдам
7. новые rich-text/canvas библиотеки подключать только после того, как базовый block model стабилен:
   - `Tiptap/ProseMirror` для rich-text внутри блоков
   - `React Konva` для drag/resize зон, если потребуется геометрическое редактирование

Правило инструментария:

- новый инструмент не должен заменять `PptxGenerator -> SlidePreviewService`
- если задача решается через существующий React overlay и `PresentationPlan`, зависимость не добавляется

## Архитектурная корректировка после анализа Gamma

Gamma работает как web-editor с карточками/блоками и темами, а не как редактор уже экспортированного PPTX.
Для нашего продукта это означает изменение порядка второго шага:

1. пользователь загружает документ
2. backend строит `PresentationPlan`
3. frontend открывает второй экран как editable deck model:
   - слайды уже заполнены данными из документа
   - стиль применяется из активного `TemplateManifest`/design tokens
   - блоки `title`, `text`, `cards`, `table`, `chart`, `image` редактируются прямо на слайде
4. пользователь меняет контент и типы блоков
5. только после подтверждения backend генерирует финальный `.pptx`
6. PowerPoint-rendered preview остаётся проверочным/export-preview слоем, а не единственным editor canvas

Решение по стеку:

- не переписывать backend с FastAPI/Python: текущий `PptxGenerator` нужен как export engine
- не переходить на NextJS/NestJS только из-за Gamma: текущий Vite/React подходит
- добавить rich-text editor слой для блоков:
  - P0: `Tiptap/ProseMirror` для inline rich-text
  - P1: `Yjs` только если потребуется совместное редактирование
  - P2: `React Konva` или аналог только если потребуется drag/resize геометрии

## Правило выполнения

Любая реализация этого плана должна соблюдать:

1. source of truth остаётся на стороне backend render pipeline
2. frontend overlay не подменяет фирменный рендер
3. новые инструменты применяются по правилам из [new-tooling-rules.md](new-tooling-rules.md)
4. изменения подтверждаются через backend tests, frontend smoke/visual и свежую runtime-генерацию
