# Slide Review Render Contract

## Назначение

Этот документ фиксирует обязательный контракт для второго шага review после нажатия `Сгенерировать`.

Цель:

- второй шаг должен показывать тот же визуальный результат, который реально формирует generator
- фирменный стиль должен приходить из активного PowerPoint template, а не дорисовываться во frontend
- пользовательские изменения `вид слайда / таблица -> график / вариант графика` должны проходить через тот же backend generation pipeline

## Главный принцип

Второй шаг не является отдельным визуальным редактором.

Он является viewer/editor shell поверх backend-rendered артефакта:

`PresentationPlan + template_id -> PptxGenerator -> .pptx -> SlidePreviewService -> PNG previews -> review UI`

Если на втором шаге показывается локально собранный React-layout, это считается только временным fallback для диагностики и не может быть основным источником истины.

## Обязательный runtime flow

1. frontend строит `PresentationPlan`
2. frontend вызывает `/presentations/generate`
3. backend берёт активный `TemplateManifest`
4. backend через `PptxGenerator` создаёт реальный `.pptx`
5. backend через `SlidePreviewService` получает preview-картинки слайдов
6. frontend открывает второй шаг только на основании `slides_preview`
7. при изменении режима слайда или chart-variant frontend отправляет обновлённый plan на новый generate cycle
8. второй шаг обновляется из новых backend-rendered previews

## Обязательные инструменты

### Backend source of truth

- [template_registry.py](../src/a3presentation/services/template_registry.py)
- [template_analyzer.py](../src/a3presentation/services/template_analyzer.py)
- [pptx_generator.py](../src/a3presentation/services/pptx_generator.py)
- [slide_preview_service.py](../src/a3presentation/services/slide_preview_service.py)
- [deck_audit.py](../src/a3presentation/services/deck_audit.py)

### Frontend responsibility

- [App.tsx](../frontend/src/App.tsx)
- [api.ts](../frontend/src/api.ts)
- `Playwright` smoke/visual checks

Frontend отвечает только за:

- upload flow
- user controls
- launch generate cycle
- display backend-rendered previews

Frontend не должен:

- воспроизводить template typography своими руками как основной preview
- считать, что local HTML preview эквивалентен PowerPoint render
- открывать второй шаг на неподтверждённом локальном draft вместо backend-rendered preview

## Preview fidelity contract

API генерации должен явно сообщать качество preview:

- `powerpoint_rendered`
  - preview экспортирован из реального PowerPoint/COM
  - это целевой режим для второго шага
- `fallback_rendered`
  - preview собран обходным способом и не гарантирует полную template parity
  - этот режим допустим только как degradеd runtime path

Если backend не смог вернуть `powerpoint_rendered`, UI должен знать об этом явно, а не молча выдавать fallback за точный брендированный результат.

## Почему нужен PowerPoint export

Для второго шага нужна не approximate geometry, а фактический PowerPoint render активного шаблона.

Официальная техдокументация:

1. Microsoft Learn, `Presentation.Export`:
   `https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.export`
2. Microsoft Learn, `Presentation.ExportAsFixedFormat2`:
   `https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.exportasfixedformat2`
3. Playwright screenshots:
   `https://playwright.dev/docs/screenshots`

Из этого следует практическое правило:

- source of truth для второго шага должен опираться на реальный export из PowerPoint presentation object
- visual regression должен валидировать именно review-shell поверх backend-rendered preview images

## Обязательные правила реализации

1. review-step должен жить на `GeneratePresentationResponse`, а не на локальном `renderSlideCanvas`
2. thumbs слева и основной frame должны по умолчанию показывать только `slides_preview`
3. если preview для слайда не пришёл, UI должен показывать controlled loading/error state, а не silently fallback-styled pseudo-slide
4. любой edit на втором шаге должен инвалидировать старые previews и запускать новый generate cycle
5. backend contract должен различать `preview_source` / `preview_fidelity`
6. quality-check после таких изменений должен включать:
   - backend API tests
   - frontend smoke
   - frontend visual
   - свежую runtime-генерацию `.pptx`

## Ближайший implementation plan

1. добавить в `GeneratePresentationResponse` поле качества preview
2. научить `SlidePreviewService` возвращать не только файлы, но и режим получения preview
3. перевести второй шаг на image-only source of truth
4. убрать локальный React-render как основной preview path
5. добавить проверки на то, что review-screen открывается только с backend-rendered previews
6. отдельно валидировать degraded path, если PowerPoint export недоступен
