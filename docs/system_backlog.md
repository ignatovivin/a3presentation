# Системный backlog

## Назначение

Этот backlog фиксирует основные незавершённые системные задачи, которые остаются важными после текущих улучшений в `content_blocks`, planner, generator и deck-audit.

Его нужно читать вместе с:

- [analysis_rule.md](analysis_rule.md)
- [document_class_matrix.md](document_class_matrix.md)

Backlog упорядочен по архитектурной значимости, а не по одному проблемному документу.
Он также исходит из того, что текущий встроенный шаблон не является постоянной константой, а будущие корпоративные шаблоны должны поддерживаться без локальной подгонки.

## 1. Модель вместимости на уровне placeholder

Статус:

- сделано по базовому template-aware контуру

Текущее состояние:

- `LayoutCapacityProfile` уже существует в [layout_capacity.py](../src/a3presentation/services/layout_capacity.py)
- generator и deck-audit теперь уменьшают profile от фактической геометрии text placeholder, а не только от грубого `layout_key`
- `TemplateAnalyzer` уже извлекает геометрию placeholder, margins text frame и базовый XML-style catalog из реальных `.pptx` шаблонов:
  `theme color/font scheme`, `master text styles`, `layout placeholder text styles`, `prototype token text styles`
- загруженные шаблоны уже могут нести shape metadata через analyzer/manifests, и этот metadata-path подтверждён для uploaded layout и uploaded prototype templates
- свежие runtime-артефакты подтвердили template-aware path:
  - uploaded layout runtime даёт уменьшенный profile по `max_chars`
  - uploaded prototype runtime остаётся зелёным по quality gate и font bounds

Почему это важно:

- разные placeholder внутри одного layout могут требовать разных fill targets и font bounds
- грубые ограничения на уровне layout заставляют planner и generator опираться на слишком широкие эвристики
- будущие корпоративные шаблоны не должны зависеть от того, что геометрия `corp_light_v1` считается константой

Цель:

- policy вместимости как минимум на уровне `layout + placeholder kind`, а лучше на уровне `layout + placeholder idx`
- template-aware path, в котором analyzer metadata может переопределять или инициализировать generator/audit geometry для загружаемых шаблонов

Что ещё остаётся:

- перенести placeholder-aware fill targets глубже в continuation heuristics planner'а
- расширить этот же контракт на более плотные narrative continuation groups и layout-stress corpus

## 2. Стабильность continuation packing

Статус:

- частично сделано

Текущее состояние:

- continuation rebalance уже существует в [planner.py](../src/a3presentation/services/planner.py)
- dense text continuation path теперь использует отдельный `dense_text_full_width` profile вместо фиктивного fallback на обычный `text_full_width`
- planner реально выбирает `dense_text_full_width`, когда это уменьшает число continuation slides без нарушения quality bounds
- свежий runtime-артефакт с плотным narrative подтвердил два `dense_text_full_width` слайда без нарушений `deck_audit`

Почему это важно:

- это главный оставшийся источник quality-regressions в narrative и mixed документах

Цель:

- стабильная балансировка для text-only и mixed continuation groups по классам документов

Что ещё остаётся:

- mixed narrative + bullets continuation groups теперь тоже умеют идти через `dense_text_full_width`, если bullets только уточняют narrative и умеренно перегружают обычный text layout
- это подтверждено planner tests, regression cases и свежим runtime-артефактом mixed dense scenario
- pure narrative continuation path теперь тоже реально умеет выбирать `dense_text_full_width`, а не только mixed path;
  при этом compact/rebalance больше не схлопывает subtitle-bearing stress-case обратно в один слайд
- source-heavy narrative case с reference-tail уже добавлен в regression corpus и quality gate
- question/callout-heavy narrative case тоже добавлен в regression corpus и quality gate
- long-title/layout-stress и отдельный long-title + subtitle + dense continuation case уже добавлены в regression corpus и quality gate
- source-heavy case с long title + subtitle + reference-tail теперь тоже добавлен в regression corpus и quality gate
- добавить больше runtime-проверок на длинные narrative sections из реальных `docx`

## 3. Расширение инвариантов deck-audit

Статус:

- частично сделано

Текущее состояние:

- [deck_audit.py](../src/a3presentation/services/deck_audit.py) уже проверяет:
  - границы размеров шрифта
  - риск overflow
  - balance continuations
  - underfilled continuation
  - порядок контента
  - геометрию table/chart/image
  - placeholder-level underfill для body area
  - footer geometry через placeholder-aware width/left checks там, где footer привязан к явному placeholder idx
  - аномально большой `title/subtitle -> body` gap, если он выходит за разумный runtime-layout контракт
  - prototype/template-specific footer geometry через synthetic idx path, а не только через обычные placeholders
  - детерминированный auxiliary text placeholder для `list_with_icons`, если ожидаемый payload левой колонки теряется при рендере
  - детерминированный правый text placeholder `14` для `list_with_icons`, даже если runtime-path считает его body-slot, а не auxiliary
  - детерминированные contact placeholders `10/11/12/13` для `contacts`, если при рендере теряется имя, роль, телефон или email
  - детерминированные card placeholders `11/12/13` для `cards_3`, если при рендере теряется текст одной из карточек
  - детерминированный `subtitle` placeholder для обычных content-slides, если он должен жить отдельно от body и не является duplicate intro

Чего ещё не хватает:

- дальнейшего расширения fill targets на остальные layout-specific placeholder roles, где mapping между plan payload и shape реально подтверждён runtime-проверкой

## 4. Контроль semantic false positives

Статус:

- частично сделано

Текущее состояние:

- ложные facts и ложные contacts уже уменьшены в [semantic_normalizer.py](../src/a3presentation/services/semantic_normalizer.py)

Почему это важно:

- report documents не должны превращать appendix/fallback logic в мусорные слайды
- form/resume documents при этом должны сохранять полезное извлечение фактов

Цель:

- более сильное разделение между narrative prose и настоящим field-value content

## 5. Расширение корпуса

Статус:

- частично сделано

Текущее состояние:

- regression corpus уже покрывает report, mixed, form, resume, table-heavy, chart-heavy, image-heavy, markdown и fact-only
- quality gate теперь дополнительно включает long-title/layout-stress docx class
- quality gate теперь дополнительно включает long-title + subtitle + dense continuation docx class
- quality gate теперь дополнительно включает source-heavy long-title + subtitle reference-tail docx class
- quality gate теперь дополнительно включает appendix-heavy structured docx class

Чего ещё не хватает:

- большего числа плотных narrative reports
- большего числа appendix/source-heavy reports с разной длиной appendix payload, reference-tail и разной долей narrative before appendix
- большего числа question/callout-heavy documents
- большего числа layout-stress кейсов с разными типами длинных заголовков, subtitle, compact body и appendix-tail комбинаций

## 6. Расширение PPTX quality layer

Статус:

- не сделано полностью

Текущее состояние:

- quality contracts уже существуют
- покрытие хорошее, но ещё не эквивалентно более широкому PPTX quality gate

Цель:

- усилить golden checks вокруг:
  - плотных bullets
  - continuation pairs
  - appendix/question slides
  - mixed paragraph+bullets
  - long title + dense body
  - compact vs underfilled соседних слайдов

## 7. Parity между preview графиков и генерацией

Статус:

- частично сделано

Текущее состояние:

- цепочка `TableChartAnalyzer -> App chart controls -> ChartPreview -> PptxGenerator -> deck_audit` уже работает как единый функциональный pipeline
- analyzer отбрасывает небезопасные default-сценарии для ordinal/status таблиц и слишком неоднозначных mixed-unit таблиц
- generator рендерит `bar`, `column`, `line`, `stacked_bar`, `stacked_column`, `pie` и explicit `combo`
- deck audit валидирует rendered chart type, series count, combo structure, размеры title/subtitle и format value-axis
- frontend smoke уже покрывает поддерживаемую матрицу preview
- visual baselines добавлены для ключевых chart preview состояний: negative column, dense line, mixed-unit combo и combo fallback при hidden line series
- runtime quality gate теперь покрывает uploaded prototype template с `chart_image` binding:
  chart slide проходит через свежую генерацию, реальный chart shape, combo/secondary-axis audit, chart title/subtitle `28/18 pt` и template-aware chart/footer width reference
- safe mixed-unit market/share scenario теперь тоже закреплён как backend contract:
  analyzer должен отдавать `combo`, primary axis не может получать percent format от соседней `%`-series,
  а свежая runtime-генерация должна давать `combo + secondary axis` без `deck_audit` violations
- mixed-unit chart choice в UI не должен схлопываться до одинаковых labels:
  если существует более одной безопасной `combo` конфигурации, analyzer должен отдавать несколько
  `candidate_specs` с различимыми `variant_label`, а select должен позволять выбрать нужный вариант
- UX выбора chart variants должен быть двухступенчатым:
  сначала пользователь выбирает тип графика (`column` / `line` / `combo`),
  потом отдельным select выбирает безопасный вариант внутри выбранного типа
- variants layer не должен терять compare-сценарии:
  для multi-series charts пользователь должен видеть `Сравнение` вместе с `Единичный`,
  а для mixed-unit scenarios дополнительно `Комбинированный`

Почему это важно:

- preview всё ещё является HTML/SVG approximation PowerPoint-графиков, поэтому parity нужно навязывать явно
- chart controls всё ещё могут разъехаться с backend generator semantics, если render rules будут дублироваться
- transposed mode и hidden-series behavior должны оставаться семантически безопасными, а не только визуально удобными
- chart formatting должен опираться на documented `python-pptx` / Open XML contract, а не на локальные UI/backend эвристики

Цель:

- один явный chart render contract для generator и deck audit
- правила transpose в UI, ограниченные семантикой chart spec, а не только геометрией таблицы
- parity checks, которые валидируют один `ChartSpec` через preview, API payload, generator и audit
- более сильное visual baseline coverage для chart-heavy сценариев

Шаги выполнения:

1. вынести общую backend chart render semantics из generator/audit в отдельный модуль
2. ужесточить `transposed` режим и запретить небезопасные mixed-unit transpose сценарии
3. добавить parity-ориентированные тесты для hidden series, combo fallback, transpose behavior и стабильности payload
4. расширить visual и quality checks для dense labels, negative values и chart-heavy document classes

Сделано в текущем цикле:

- smoke закрепляет combo fallback при скрытии line series в preview
- visual snapshots зафиксированы для chart preview card-состояний с negative values, dense labels, mixed-unit combo и hidden-series fallback
- добавлен runtime browser/API smoke для chart-heavy flow:
  браузер мокает только extraction response, затем реально вызывает local backend для plan/generate/download
- runtime chart fixture приведён к форме реального extractor:
  таблица присутствует в ordered `blocks` как `kind="table"`, поэтому chart override заменяет table-slide, а не создаёт chart рядом с дублирующей таблицей
- свежая сервисная генерация runtime-equivalent chart deck прошла `deck_audit` без violations:
  план `title/bullets/chart`, rendered chart `column`, 2 series, chart width ratio `1.00`
- отдельно закрыт defect-класс из реального артефакта на слайде с market volume + share:
  single-axis chart больше не получает левую ось `0"%"` из-за одной процентной series,
  а analyzer для safe mixed `number/RUB + %` теперь выдаёт `combo` как основной candidate

Следующий шаг по этому разделу:

- расширить browser/API parity до явной проверки `/plans/from-text` payload/response:
  тот же `ChartSpec` должен быть виден в preview controls, в `chart_overrides` request, в plan response, в generated deck и в `deck_audit`

## 8. Второй шаг review как backend-rendered editable flow

Статус:

- начато

Текущее состояние:

- второй шаг уже переведён на image-first preview path
- backend уже возвращает `preview_fidelity`
- локальный React-render больше не должен считаться основным источником истины для review-screen
- архитектурный контракт зафиксирован в [slide-review-render-contract.md](slide-review-render-contract.md)
- правила новых инструментов зафиксированы в [new-tooling-rules.md](new-tooling-rules.md)
- implementation plan вынесен в [slide-review-implementation-plan.md](slide-review-implementation-plan.md)

Почему это важно:

- пользователь ожидает, что второй шаг показывает именно тот фирменный стиль, который реально попадёт в итоговый `.pptx`
- без этого review-screen остаётся визуальной аппроксимацией, а не надёжным этапом редактирования

Цель:

- превратить второй шаг в `backend-rendered preview + editable overlay + regenerate` flow

Что ещё остаётся:

- `preview_source` добавлен в backend/frontend contract
- причина fallback возвращается из preview service, включая `fallback_code`
- сделать `powerpoint_rendered` стабильным целевым runtime path
- локальная runtime-проверка подтвердила `powerpoint_rendered` через PowerPoint COM после перехода на фактическое число слайдов в `.pptx`
- contract tests для `powerpoint_rendered`, `powershell_missing` и `partial_export` preview service добавлены
- editor overlay внедрён как слой зон поверх backend-rendered preview
- поля ввода контента убраны из панели второго шага; content editing переведён в overlay inline editor поверх preview
- cancellation/revision model для regenerate flow внедрена через `AbortController` и request id
- smoke/visual checks под editable review-step расширены
- восстановление review-state больше не принимает старый `fallback_rendered` или пустой preview как валидный результат второго шага
- UI второго шага показывает информацию о последней генерации: имя `.pptx`, режим preview, источник, число preview-слайдов и ссылку скачивания
- первый инкремент editable block model добавлен во frontend: клик по body-зоне открывает меню действий блока прямо поверх слайда, а выбранное действие маппится в `PresentationPlan` и запускает regenerate-flow
- исправлен defect-класс image-only preview: backend-rendered PNG больше не наследует grid-rows от локального React slide canvas и не сжимается в маленькую полоску над overlay-зонами; smoke теперь проверяет размер картинки относительно canvas
- исправлен defect-класс preview reset: при переключении типа контента/редактировании старый backend-rendered preview больше не обнуляется, а остаётся на экране до прихода нового regenerate-response
- overlay-зоны второго шага больше не закрывают слайд как видимые блоки по умолчанию: они прозрачные и проявляются только при hover/focus/active
- начат переход на Gamma-like порядок второго шага: первая кнопка после загрузки документа строит editable deck model из `PresentationPlan` без немедленной генерации `.pptx`, а финальный PowerPoint/render запускается отдельной кнопкой на втором экране
- editable deck canvas теперь template-aware на первом уровне: frontend получает `TemplateManifest` через `/templates/{template_id}`, помечает canvas `data-template-id`/`data-layout-key` и показывает активный стиль/layout до финальной генерации `.pptx`
- `TemplateManifest` теперь несёт `design_tokens`, а editable deck canvas применяет их как CSS variables; это фиксирует источник стиля второго шага в manifest contract, а не во фронтовых константах
- первый вход во второй шаг теперь сразу запускает backend generation/preview cycle; review-shell больше не показывает local React draft как основной слайд, а ждёт image-based preview и держит overlay только поверх полученного backend-rendered кадра
- regenerate orchestration второго шага продвинута: изменения title/body/mode/chart variant теперь запускают debounce-based auto-regenerate, старый image preview остаётся на экране до прихода нового ответа, а request cancellation продолжает отбрасывать устаревшие ответы
- inline editing narrative-блоков на втором шаге усилен: title/subtitle/body теперь редактируются через contenteditable surface прямо внутри overlay-зоны на слайде, а не через обычные `input/textarea` как основной UX path
- second-step UX переведён дальше в overlay-first режим: верхние select-контролы `Вид/Тип графика` убраны из review-frame, а смена slide mode и chart variant теперь вызывается из overlay popup прямо на слайде
- block-toolbar второго шага унифицирован: `title`, `body` и `chart` теперь открывают один и тот же popup-механизм блока, а уже внутри него запускаются edit / mode switch / chart variant actions
- full block-toolbar model расширен и на `table`: табличный слайд теперь имеет собственную overlay-зону `table` и тот же popup-path для `table -> chart`, вместо старого использования body-зоны как proxy
- block-toolbar теперь несёт metadata header: popup явно показывает тип активного блока и его текущее состояние/режим, чтобы edit actions не были контекстно-слепыми
- второй шаг снова приведён к `editable deck model`: основной canvas и левые миниатюры теперь всегда рендерятся как живые template-aware slide blocks из `PresentationPlan`, а backend-generated PNG/PPTX оставлен только как export/validation слой и больше не подменяет собой рабочий экран редактирования
- следующий инкремент direct-editing добавлен в сам canvas: `cards` теперь редактируются поштучно прямо в карточной сетке внутри overlay-редактора, а для `table/chart` появились встроенные quickbar-действия в зоне блока без обязательного прохода через общий popup
- block-toolbar второго шага расширен до block-level операций: для narrative/cards/list добавлены `duplicate/reset`, а reset теперь возвращает конкретный слайд к исходному `sourceReviewPlan` вместо локальной ручной отмены
- narrative UX второго шага переведён на direct click-to-edit: `title/subtitle/body/cards` теперь редактируются кликом по самому текстовому блоку внутри live canvas, а overlay-зоны больше не являются основным входом для текстового редактирования; overlay оставлен только для `table/chart`
- editable canvas стал template-aware по геометрии: `title/subtitle/body/cards/footer` теперь позиционируются через placeholder metadata из `TemplateManifest`, а при неполном manifest используют layout-aware fallback geometry, синхронизированную с backend `layout_capacity` политиками
- следующий инкремент block-model добавлен и для data-blocks: `table/chart` теперь открывают inline data editor прямо поверх live canvas, редактируют заголовки колонок и строки и маппят изменения обратно в `PresentationPlan`, сохраняя единый regenerate flow
- следующий инкремент representation presets тоже продвинут: компактные data-blocks теперь могут безопасно переключаться не только `table <-> chart`, но и в `cards/list`, если табличная структура допускает это без смысловой деградации; regenerate path и API contract для `chart override` и `list/cards` representation закреплены отдельными smoke/backend tests
- template-style path второго шага усилен: `TemplateManifest.design_tokens` теперь несёт не только базовые brand colors/font family, но и role-level typography (`title/subtitle/body/footer` colors, font sizes, weights); `TemplateAnalyzer` читает их из theme/master XML `.pptx`, а editable canvas больше не держит собственные hardcoded `28/14/17/12`, а применяет template-driven CSS variables
- введён следующий уровень manifest contract: `TemplateManifest.component_styles` как reusable component-layer поверх плоских `design_tokens`; scope уже расширен с `cards/text` на `table/chart/image/cover/list_with_icons/contacts`, и generator теперь читает этот слой для card/numeric-card rendering, table cell styling, table render behavior, chart palette/title/geometry, image subtitle/geometry spacing, cover typography/geometry, two-column spacing и contacts font-threshold behavior
- XML base layer расширен с текста на все основные компоненты шаблона: manifest теперь может хранить `theme color/font scheme`, `master text styles`, `master background styles`, `layout background style`, `placeholder shape styles`, `prototype token shape styles`; analyzer уже тянет `fill/gradient/image fill`, `line`, `shadow`, `rotation`, `geometry preset` и готовит единый style catalog для `text/table/chart/image/background`
- advanced XML coverage тоже добавлен в base layer: analyzer и manifest уже умеют хранить `paragraph lvl2+`, `bullet type/font/hanging`, `line compound/cap/join`, `glow/softEdge/reflection`, `text inset/anchor`, `theme fill/line refs`, а также catalog для `table cell margins` и `chart placeholder offsets`; это уже не только metadata-path, а частично рабочий generator path для реального `.pptx`
- этап применения XML в generator расширен: layout `background_style`, placeholder `text_style`, paragraph catalog, `shape_style` (`fill/line/inset/anchor`, `line compound/cap/join`, `theme fill/line refs`) и `table cell margins` уже реально накладываются при генерации слайда; для chart placeholders генератор теперь сохраняет `manualLayout` offsets/size, legend manual offset и axis label offset через Open XML, что закреплено generator contract tests
- добавить smoke/api coverage для:
  - `fallback_rendered` review-state уже покрыт
  - hard refresh и восстановления review-step уже покрыт
  - race conditions и отмены устаревших preview-response уже покрыт
  - overlay-driven редактирование table/chart zones уже покрыто базовым smoke
- overlay-driven редактирование content zones покрыто базовым smoke
- довести второй шаг до полноценного `backend-rendered preview + overlay editing`, а не только `preview + side controls`

## 9. Generic editable slot contract для user templates

Статус:

- начато

Текущее состояние:

- `TemplateManifest` уже хранит первый generic editable-slot metadata layer:
  - `editable_role`
  - `editable_capabilities`
  - `slot_group`
  - `slot_group_order`
- `TemplateAnalyzer` уже заполняет этот metadata-path:
  - для layout placeholders через `kind -> binding/role/capability`
  - для prototype tokens через token-name grouping conventions
- analyzer backfill теперь переносит этот metadata layer и в templates с уже существующим `manifest.json`,
  а не только в freshly analyzed manifests

Почему это важно:

- review-step не должен зависеть от built-in layout names
- frontend должен получать generic editable semantics ещё до полной slot-driven UI сборки
- uploaded templates требуют единого manifest contract, иначе extraction metadata теряется между analyzer и UI/runtime

Цель:

- довести manifest до полноценной editable slot model для arbitrary user templates

Что ещё остаётся:

- добавить confidence / safe-editable / render-only distinctions
- добавить richer repeated-group inference вместо только token naming conventions
- протянуть этот contract в API/frontend review-step и plan mutation mapping
- убрать remaining hardcoded representation targets во frontend review-step;
  первый шаг уже сделан для text-to-cards chooser, который теперь выбирает card-capable target layout из manifest metadata, а не только `cards_3`
- переносить эту же логику с frontend-эвристик на backend-driven `representation_hints`,
  где analyzer уже умеет помечать хотя бы card-like layouts
- удерживать единый manifest-driven path и для built-in registry templates:
  review-step уже начал читать manifest не только из uploaded-template response, но и через обычный `/templates/{template_id}`
- убирать remaining layout-name heuristics из review-step:
  один из таких путей уже убран для data-slide filtering, который теперь смотрит в manifest metadata, а не в `layoutKey.includes(...)`
- расширять backend-driven `representation_hints` дальше:
  базовый слой уже покрывает хотя бы `cards`, `table`, `image`, `contacts`,
  но review-step пока использует это только частично
- после текущего цикла появился уже и manual-testable UI слой:
  следующий шаг теперь не “сделать хоть что-то видимое”, а расширять, какие transformations разрешаются из manifest contract
- layout inventory уже начал использоваться и в backend plan path:
  следующий шаг теперь не “увидеть макеты”, а делать content-to-layout matching умнее и глубже, а не только remap logical aliases

## Текущий порядок выполнения

Следующий implementation cycle должен идти в таком порядке:

1. сделать `deck_audit` template-aware для uploaded layout templates через analyzer/manifest placeholder metadata
2. распространить тот же template-aware audit path на prototype templates, где token-bound shapes заменяют обычные placeholders
3. вывести placeholder-level expectations по вместимости из template metadata, а не полагаться только на грубые `layout_key` defaults
4. расширить regression и quality contracts для uploaded templates по классам документов
5. продолжать hardening continuation-packing только после того, как template-aware quality layer станет достаточно сильным, чтобы ловить regressions
6. вынести общий chart render contract для generator и deck audit
7. ужесточить transposed chart mode и mixed-unit chart safety в UI/API flow
8. добавить parity checks между chart preview и generated deck до того, как расширять chart features дальше

Состояние после текущего цикла:

- пункты 1 и 2 подтверждены не только project-contract tests, но и quality gate
- пункт 3 закрыт по базовому template-aware контуру: generator и deck-audit теперь используют placeholder-aware capacity derivation, и это подтверждено свежей runtime-генерацией uploaded layout/prototype артефактов
- пункт 4 продвинут: `quality-contracts` теперь включает uploaded layout и uploaded prototype template scenarios
- пункт 2 продвинут дальше: dense continuation path теперь работает и для paragraph-dominant mixed narrative + bullets scenarios, подтверждён tests, quality gate и свежей runtime-генерацией
- следующий незакрытый архитектурный шаг теперь: расширять `deck_audit` на placeholder-level fill targets и соседние fill/gap инварианты

## Правило исполнения

Когда выбирается следующая задача:

1. предпочитать архитектурные улучшения, а не тюнинг под один документ
2. по возможности сначала добавлять или обновлять тесты
3. валидировать по классам документов, а не по одному regression file
4. автоматически переходить к следующему безопасному шагу, а не останавливаться ради лишнего подтверждения внутри уже согласованной задачи
