# New Tooling Rules

## Назначение

Этот документ фиксирует правила применения новых инструментов для второго шага review/editable preview.

Его нужно использовать вместе с:

- [analysis_rule.md](analysis_rule.md)
- [frontend-visual-contracts.md](frontend-visual-contracts.md)
- [slide-review-render-contract.md](slide-review-render-contract.md)

## Новые инструменты

### 1. PowerPoint COM Export

Использование:

- точный экспорт preview слайдов из реально сгенерированного `.pptx`
- основной source of truth для фирменного вида второго шага

Техническая документация:

- `Presentation.Export`
  - `https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.export`
- `Presentation.ExportAsFixedFormat2`
  - `https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.exportasfixedformat2`

Обязательные правила:

1. для review-step целевым режимом считается только `powerpoint_rendered`
2. fallback-preview не может считаться эквивалентом фирменного рендера
3. любые визуальные расхождения второго шага нужно сначала разбирать через generator/export pipeline, а не через CSS

### 2. React Konva / Konva

Использование:

- интерактивный overlay поверх backend-rendered slide preview
- выделение editable-зон
- drag/resize элементов editor layer

Техническая документация:

- React drag and drop:
  - `https://konvajs.org/docs/react/Drag_And_Drop.html`
- React Transformer:
  - `https://konvajs.org/docs/react/Transformer.html`

Обязательные правила:

1. Konva используется только как editor overlay, а не как движок фирменного preview
2. Konva-state должен маппиться обратно в `PresentationPlan`/edit payload
3. нельзя хранить бизнес-правду только в canvas coordinates без связи с backend model

### 3. AbortController

Использование:

- отмена устаревших generate/preview запросов
- защита от race conditions при быстрых правках

Техническая документация:

- `https://developer.mozilla.org/en-US/docs/Web/API/AbortController`

Обязательные правила:

1. каждый новый regenerate request должен уметь отменять предыдущий
2. устаревший preview response не должен перетирать более новый

### 4. ResizeObserver

Использование:

- синхронизация размеров editor overlay и preview frame

Техническая документация:

- `https://developer.mozilla.org/en-US/docs/Web/API/ResizeObserver`

Обязательные правила:

1. размер overlay должен вычисляться от фактического preview-frame
2. нельзя опираться только на `window resize`, если меняется размер контейнера

### 5. React concurrent UI APIs

Использование:

- неблокирующий UI для edit -> regenerate flow

Техническая документация:

- `useTransition`
  - `https://react.dev/reference/react/useTransition`
- `useActionState`
  - `https://react.dev/reference/react/useActionState`

Обязательные правила:

1. regenerate flow не должен подвешивать UI
2. pending-state должен быть явно виден пользователю
3. slow backend regenerate не должен ломать local editing flow

### 6. Playwright

Использование:

- smoke/visual contract для второго шага

Техническая документация:

- `https://playwright.dev/docs/screenshots`

Обязательные правила:

1. второй шаг нужно проверять на backend-rendered previews
2. visual baselines нельзя строить на локальном псевдо-рендере, если целевой режим уже image-based
3. проверки должны ловить regressions в `preview_fidelity`, regenerate flow и layout-stability

## Главный архитектурный запрет

Нельзя использовать новые frontend-инструменты как замену PowerPoint generator/export pipeline.

Разрешённая схема только такая:

`backend-rendered preview` + `frontend editor overlay` + `regenerate`

Запрещённая схема:

`frontend сам дорисовывает фирменный слайд и считает его итоговым`
