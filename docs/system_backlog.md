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
- `TemplateAnalyzer` уже извлекает геометрию placeholder и margins text frame из реальных `.pptx` шаблонов
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

Чего ещё не хватает:

 - разницы размера шрифта между соседними continuation slides
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
