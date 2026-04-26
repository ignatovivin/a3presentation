# Roadmap и CJM для user-template editable версии

## Контекст

Текущий проект уже хорошо продвинут в сторону:

- `docx -> plan -> pptx`
- template-aware generation
- review-step с editable canvas
- analyzer/manifest metadata

Но следующий продуктовый рубеж должен строиться не вокруг built-in layout keys вроде `cards_3`, `contacts`, `list_with_icons`,
а вокруг общего сценария:

`любой пользовательский .pptx шаблон -> извлечение slot/model/style contract -> editable review -> controlled generation`

Это означает, что built-in layout rules должны остаться как fallback/bootstrap слой,
но не как главная архитектурная единица продукта.

## Текущий constrained working set

Пока этот roadmap является активным вектором, рабочий контекст нужно держать узким.

### In scope сейчас

Файлы, которые можно считать основным working set под текущий flow
`любой пользовательский шаблон -> extraction contract -> editable review -> controlled generation`:

- `src/a3presentation/services/template_analyzer.py`
- `src/a3presentation/services/template_registry.py`
- `src/a3presentation/api/routes.py`
- `src/a3presentation/domain/template.py`
- `src/a3presentation/domain/api.py`
- `frontend/src/App.tsx`
- `frontend/src/api.ts`
- `frontend/src/types.ts`
- `frontend/src/index.css` только если это нужно для template-aware review UI
- `tests/test_api.py`
- `tests/test_project_contracts.py` только в части analyzer / manifest / registry / API contract
- `docs/user-template-roadmap-cjm.md`

### Parked / out of scope до отдельного возврата

Эти файлы сейчас не должны быть основной зоной работы, если нет прямого блокера для user-template flow:

- `src/a3presentation/services/planner.py`
- `src/a3presentation/services/pptx_generator.py`
- `src/a3presentation/services/deck_audit.py`
- `src/a3presentation/services/layout_capacity.py`
- `tests/test_planner.py`
- `tests/test_quality_contracts.py`
- `docs/quality-contracts.md`
- любые правки, завязанные на built-in layout tuning (`cards_3`, `contacts`, `list_with_icons`) как на главный продуктовый слой

### Правило работы

- сначала доводим generic template contract и editable review path
- built-in layouts считаем bootstrap/fallback слоем
- если правка не усиливает user-template extraction/review/generation flow, её лучше не тащить в текущий цикл

## Целевая продуктовая модель

Система должна уметь:

1. принять пользовательский `.pptx` шаблон
2. извлечь из него editable/template grammar
3. показать пользователю управляемую deck model до финальной генерации
4. позволить управлять content/representation/layout choices в рамках реально доступных template slots
5. сгенерировать финальный `.pptx` в стиле именно этого шаблона

## Новый roadmap

### Этап 1. Generic template extraction contract

Цель:
перестать мыслить шаблон как набор наших layout names и перейти к generic slot model.

Что нужно сделать:

- расширить `TemplateAnalyzer`, чтобы он извлекал не только placeholder geometry, но и slot semantics:
  - title
  - subtitle
  - body/text
  - secondary text
  - table
  - chart
  - image
  - footer/meta
  - multi-slot groups
  - repeated card-like regions
  - two-column regions
- расширить `TemplateManifest`, чтобы он хранил:
  - slot ids
  - slot roles
  - slot geometry
  - style tokens
  - fill/border/text rules
  - group/repeater metadata
  - editable capabilities
  - supported representation transforms
- ввести distinction между:
  - detected slot
  - editable slot
  - render-only slot

Definition of done:

- загруженный шаблон описывается generic manifest contract без знания наших built-in layout names
- manifest достаточно богат, чтобы frontend мог строить editable overlay по нему

Сделано в текущем цикле:

- `TemplateManifest` расширен первым generic editable-slot contract:
  - `editable_role`
  - `editable_capabilities`
  - `slot_group`
  - `slot_group_order`
- `TemplateAnalyzer` теперь заполняет этот metadata-path для:
  - layout placeholders через generic role/binding inference
  - prototype tokens через token-name grouping conventions
- existing `manifest.json` templates больше не теряют этот слой:
  analyzer backfill'ит не только geometry, но и editable slot metadata в уже сохранённые manifests
- `TemplateRegistry.normalize_manifest()` теперь синхронизирует editable metadata после binding-normalization,
  чтобы `table/contacts/footer` placeholders не расходились между binding и editable capabilities

### Этап 2. Generic editable deck model

Цель:
review-step должен строиться из manifest-driven slot model, а не из фронтовых допущений о layout’ах.

Что нужно сделать:

- перевести второй шаг на generic slot rendering:
  - canvas строится из slot tree/template grammar
  - overlay actions строятся из editable capabilities manifest
- добавить mapping:
  - `PresentationPlan block -> template slot`
  - `slot action -> plan mutation`
- отделить:
  - content editing
  - representation switching
  - layout slot targeting
- дать пользователю управлять:
  - текстом
  - таблицей
  - графиком
  - карточным/списочным представлением
  - видимостью отдельных блоков
  - выбором между несколькими безопасными target slots/layout variants

Definition of done:

- review-step работает на чужом шаблоне без ручного кодирования нового layout key

Сделано в текущем цикле:

- text-to-cards chooser во frontend больше не жёстко привязан к `cards_3`
- review-step теперь умеет выбирать card-capable target layout из analyzer-derived manifest metadata
  по geometry/slot heuristic, а не только по встроенному layout key
- `TemplateManifest` теперь также несёт `representation_hints` для layouts/prototype slides
- analyzer уже умеет помечать card-like layouts этим hint'ом, а backfill существующих manifests не должен терять этот слой
- frontend review logic теперь использует manifest metadata не только для uploaded template path,
  но и для обычного выбранного шаблона из registry через `fetchTemplate()`
- ещё один hardcoded review-step path убран:
  фильтр text-to-cards больше не определяет data layouts по `layoutKey.includes("table"|"chart")`,
  а смотрит в manifest slot/support metadata
- `representation_hints` расширены дальше:
  теперь manifest умеет явно помечать не только `cards`, но и `table`, `image`, `contacts`
- normalization path тоже синхронизирует эти hints после binding-normalization,
  чтобы `table/contacts` layouts не теряли representation semantics
- UI доведён до ручного testable state:
  пользователь уже видит active template analysis прямо в интерфейсе
  (`editable slots`, `representation hints`, `card target`, active template source)
- backend планирования сделал следующий шаг к реальному layout inventory usage:
  после `build_plan()` plan slides теперь могут получать `preferred_layout_key` из реально найденных layouts шаблона,
  а не оставаться только на logical aliases
- UI теперь дополнительно показывает и явный список обнаруженных макетов шаблона

### Этап 3. Representation engine поверх template slots

Цель:
не подгонять документ под жёстко прошитые slide types, а уметь выбирать representation внутри ограничений шаблона.

Что нужно сделать:

- ввести representation layer:
  - paragraph flow
  - bullet list
  - cards
  - KPI cards
  - table
  - chart
  - image + text
  - contacts/meta
- научить planner выбирать не только `slide kind`, но и candidate representations,
  которые совместимы с доступными slot groups шаблона
- научить manifest описывать:
  - какие representation types поддерживаются данным шаблоном
  - какие slot groups являются interchangeable

Definition of done:

- пользователь может до генерации менять representation там, где шаблон это реально поддерживает

### Этап 4. Generic template-aware audit

Цель:
quality layer должен валидировать не built-in layouts, а generic slot contract шаблона.

Что нужно сделать:

- перенести `deck_audit` от built-in layout assumptions к manifest-derived slot expectations
- валидировать:
  - slot fill
  - content loss
  - geometry drift
  - title/subtitle/body stack
  - table/chart/image bounds
  - repeated group integrity
- оставить built-in rules только как fallback, если manifest бедный

Definition of done:

- `deck_audit` может валидировать пользовательский шаблон по его же manifest metadata

### Этап 5. Full user-template runtime

Цель:
полный продуктовый путь для пользовательского шаблона должен быть воспроизводим end-to-end.

Что нужно сделать:

- upload custom template
- analyze template
- build editable review deck
- user edits content/representation
- regenerate preview
- final `.pptx`
- audit/export/download

Definition of done:

- сценарий работает без ручного добавления нового template-specific кода

## Приоритеты

### P0

- generic slot extraction
- manifest contract for editable capabilities
- manifest-driven review-step
- template-aware audit from manifest metadata

### P1

- representation switching across arbitrary templates
- repeated/card-group inference
- richer chart/table/image editable controls

### P2

- automatic suggestion quality
- smarter slot ranking
- template similarity / reusable pattern library

## Основные риски

- analyzer может неправильно понять slot semantics в нестандартных шаблонах
- не все template shapes реально editable одинаково безопасно
- repeated visual groups не всегда однозначно извлекаются из XML
- пользователь может ожидать абсолютную WYSIWYG-редактируемость, когда шаблон структурно бедный

## Как минимизировать риски

- ввести confidence levels для extracted slots
- отделять safe-editable и unsafe/readonly slots
- показывать пользователю explicit fallback/limitations
- поддерживать layered mode:
  - full editable
  - partial editable
  - render-only

## CJM

### Persona

Пользователь приносит свой корпоративный `.pptx` шаблон и хочет получить презентацию в собственном стиле,
не теряя контроль над структурой до финальной генерации.

### Шаг 1. Загрузка шаблона

Пользователь:
- загружает `.pptx`

Ожидание:
- система понимает, что это не “просто файл”, а будущая визуальная грамматика колоды

Система должна:
- проанализировать template
- извлечь slot map
- показать, что найдено:
  - title/body slots
  - chart/table/image slots
  - multi-column/card groups
  - editable vs readonly areas

Риски:
- шаблон анализируется как “непонятный”
- пользователь не понимает, что именно будет доступно для управления

### Шаг 2. Загрузка документа

Пользователь:
- загружает `docx`/text source

Ожидание:
- система строит содержательную структуру, а не просто накидывает текст на слайды

Система должна:
- извлечь структуру документа
- построить `PresentationPlan`
- подобрать candidate representations, совместимые с template slots

Риски:
- planner делает representation, которую данный шаблон не умеет качественно показать

### Шаг 3. Черновой review

Пользователь:
- открывает второй шаг

Ожидание:
- видит не абстрактный draft, а editable deck model, уже основанную на его шаблоне

Система должна:
- показать slot-driven canvas
- подсветить editable blocks
- показать, какие трансформации доступны для каждого блока

Риски:
- пользователь не понимает, почему один блок можно редактировать свободно, а другой нельзя

### Шаг 4. Управление содержимым

Пользователь:
- редактирует title/body/table/chart/cards
- переключает representation там, где это допустимо

Ожидание:
- изменения происходят в рамках реальных ограничений шаблона

Система должна:
- маппить edit action обратно в plan
- валидировать совместимость с template slots
- не давать unsafe transforms

Риски:
- потеря части контента
- визуальное переполнение
- пользователь выбирает representation, которое шаблон плохо поддерживает

### Шаг 5. Regenerate preview

Пользователь:
- ждёт обновлённый preview

Ожидание:
- preview соответствует финальному `pptx`

Система должна:
- собрать новый plan
- сгенерировать новый `.pptx`
- построить backend-rendered preview
- явно показывать fidelity и fallback state

Риски:
- preview и итоговый `.pptx` расходятся
- regenerate flow ломает пользовательский контекст

### Шаг 6. Финальная генерация

Пользователь:
- нажимает generate/download

Ожидание:
- получает презентацию в стиле собственного шаблона, уже проверенную на базовые quality rules

Система должна:
- сгенерировать `.pptx`
- прогнать template-aware audit
- отдать файл и preview metadata

Риски:
- silent content loss
- invalid PPTX
- визуальные дефекты, которые не были показаны заранее

## Product criteria для “рабочей версии”

- пользовательский шаблон можно загрузить без ручной подготовки
- analyzer извлекает достаточно metadata для editable review
- review-step работает от manifest-driven slot model
- representation switch зависит от возможностей шаблона
- generator рендерит в пользовательский шаблон без template-specific кода
- audit валидирует результат по manifest-derived contract
- пользователь проходит весь путь:
  `upload template -> upload document -> review/edit -> regenerate -> generate/download`

## Ближайший следующий implementation step

Самый прагматичный следующий шаг:

1. провести gap-analysis текущего `TemplateAnalyzer` и `TemplateManifest`
2. определить, каких generic slot metadata не хватает для manifest-driven review-step
3. реализовать первый инкремент `editable slot capabilities` в manifest/backend/frontend contract
