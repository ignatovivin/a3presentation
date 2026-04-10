# Frontend и visual contracts

## Текущее состояние

- frontend runtime contract включает `yarn verify`
- в проекте уже есть `playwright` smoke и visual tests
- UI и visual validation автоматизированы частично и всё ещё требуют более широкого сценарного покрытия
- backend deck-level quality checks описаны отдельно в [quality-contracts.md](quality-contracts.md)
- backend quality layer теперь также валидирует порядок mixed-content внутри body containers сгенерированной колоды
- `frontend smoke` уже подходит для CI
- `frontend visual` пока должен оставаться отдельным gate, пока не будет зафиксирована стабильная cross-platform snapshot policy
- chart preview smoke теперь покрывает supported chart layout matrix, включая geometry line marker/label и защиту от invalid coordinates
- controls для transpose должны оставаться согласованы с backend chart semantics и не должны появляться для unsafe mixed-unit chart specs
- visual baselines теперь зафиксированы не только для общего drawer, но и для chart preview card-состояний: negative column, dense line, mixed-unit combo и combo fallback после скрытия line series
- добавлен отдельный runtime chart smoke: браузер мокает только extraction layer, а `/plans/from-text`, `/presentations/generate` и download идут через реальный локальный backend; это закрывает разрыв между preview controls/request payload и свежей backend-генерацией
- fixture для runtime chart smoke должен содержать таблицу и в `tables`, и в ordered `blocks` как `kind="table"`; это соответствует реальному extractor и предотвращает дублирование chart-slide + table-slide при table->chart override
- безопасный mixed-unit сценарий `number/RUB + %` теперь должен приходить в UI уже как `combo` candidate;
  preview обязан показывать column+line с secondary semantics, а не single-axis percent-formatted column
- если mixed-unit таблица допускает несколько безопасных `combo` конфигураций, backend должен отдавать несколько
  `candidate_specs` с явными `variant_label`, а `chart type select` должен показывать эти различимые варианты,
  а не несколько одинаковых option labels
- в карточке chart preview текущий выбранный mixed-unit вариант должен быть явно показан в тексте UI,
  чтобы пользователь видел не просто `combo`, а конкретный способ построения
- пользовательский flow выбора графика должен быть двухступенчатым:
  сначала отдельный select по типу графика (`column` / `line` / `combo`),
  затем второй select по варианту внутри уже выбранного типа, если у этого типа есть несколько безопасных вариантов
- при этом option labels должны быть самодостаточными:
  без отдельных верхних label над select'ами, с явными формулировками `Единичный`, `Сравнение` и `Комбинированный`

## Обязательный frontend smoke flow

Минимальный автоматизированный UI flow должен проверять:

1. приложение загружается и получает templates с backend
2. пользователь может загрузить `.docx` документ
3. извлечённые text, tables и chart assessments появляются в UI state
4. пользователь может открыть structure drawer
5. пользователь может переключить chartable table между `table` и `chart`
6. пользователь может сгенерировать презентацию
7. пользователь получает downloadable result link
8. пользователь может закрывать success и error panels
9. chart controls остаются согласованы с backend candidates:
   первый select выбирает тип графика, второй select выбирает вариант внутри типа,
   а variants layer не теряет сценарии `Сравнение` для multi-series charts и `Комбинированный` для mixed-unit cases
10. hidden chart series сохраняются в `chart_overrides` request payload
11. drawer switch/select controls остаются доступными по role и видимыми для Playwright
12. transpose select появляется только для семантически безопасных chart specs и даёт ожидаемый transposed payload

## Обязательный набор visual regression

Минимальный visual regression layer должен покрывать:

1. cover slide
2. dense text slide
3. bullets slide
4. compact table slide
5. wide table slide
6. chart slide
7. image slide
8. положение footer на long-title slides

## Chart preview smoke matrix

Текущий автоматизированный chart preview smoke должен покрывать:

- `column`
- `bar`
- `line`
- `stacked_column`
- `stacked_bar`
- `pie`
- explicit `combo` specs для legacy/generator parity
- mixed-unit combo preview, где line series использует свою axis/value-format semantics
- dense line categories с большими compact values
- negative values без `NaN` marker coordinates
- hidden-series behavior в plan payload
- combo fallback preview при скрытии line series в mixed-unit сценарии

## Обязательные правила форматирования значений графиков

Frontend preview и chart controls должны оставаться согласованы с backend chart-formatting contract:

1. axis labels и data labels считаются разными слоями и не должны смешиваться
2. `%` допустим только для оси, где все series процентные
3. mixed `number/RUB + %` должен отображаться как `combo` с secondary semantics
4. если backend определил процентную series по semantic header вроде `Доля` или `Маржа`, UI не должен терять этот контракт

Внешняя документация для этого правила фиксируется в [analysis_rule.md](analysis_rule.md).

## Рекомендуемый tooling

- UI smoke: `playwright`
- visual snapshots: `playwright` screenshots или эквивалентный browser snapshot layer
- CI entry points:
  - `yarn verify`
  - `yarn test:smoke`
  - `yarn test:visual` после фиксации cross-platform snapshot policy

## Критерий готовности

Frontend и visual contract layer можно считать готовым, когда:

- smoke flow автоматизирован end-to-end против локального backend
- хотя бы один golden snapshot существует для каждого обязательного visual scenario
- regressions валят CI до ручного QA
- frontend gate остаётся согласован с backend quality-contract gate
