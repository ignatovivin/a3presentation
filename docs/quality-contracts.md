# Quality Contracts

`quality-contracts` это отдельный слой верификации для сгенерированных колод.

Он уже, чем полный backend test suite, и сосредоточен на end-to-end layout-контрактах:

- вместимость текста и bullets
- баланс continuation slides
- сохранение порядка mixed-content
- геометрия table layout
- геометрия chart layout
- семантические chart-контракты:
  - тип отрендеренного графика
  - число отрендеренных рядов
  - combo-структура bar/line
  - размеры шрифтов title/subtitle
  - number format value-axis для компактного отображения больших чисел
- геометрия image layout
- template-aware поведение для uploaded и analyzer-derived templates
- quality gate теперь включает отдельные зелёные сценарии для:
  - analyzer-derived uploaded layout template
  - uploaded prototype template
  - uploaded prototype template с `chart_image` binding
- для template-aware text paths quality gate теперь дополнительно подтверждает placeholder-aware вместимость:
  - audit-profile уменьшается от реальной геометрии placeholder, а не только от `layout_key`
  - generator подбирает body font bounds из фактических размеров shape
  - это подтверждается и test suite, и свежей runtime-генерацией `.pptx` в `storage/outputs`
- analyzer/manifest path теперь также несёт первый generic editable-slot metadata layer
  (`editable_role`, `editable_capabilities`, `slot_group`, `slot_group_order`) для user-template review flow;
  этот слой пока не является отдельным quality gate, но уже закреплён project-contract tests и должен сохраняться при backfill существующих manifests
- normalization layer тоже должен сохранять согласованность этого контракта:
  если `TemplateRegistry.normalize_manifest()` проставляет `binding` вроде `table` или `contact_*`,
  он не должен оставлять устаревшие `editable_role/capabilities`
- тот же принцип теперь относится и к `representation_hints`:
  analyzer-derived hints для card-like layouts/prototypes не должны теряться при загрузке существующего `manifest.json`
- это уже относится не только к `cards`, но и к `table` / `image` / `contacts` semantics,
  которые могут появляться или уточняться на normalize-этапе после binding sync
- представительные классы документов:
  - text-only
  - mixed text
  - report-like
  - source-heavy report с reference-tail
  - question/callout-heavy report
  - long-title/layout-stress report
  - long-title + subtitle + dense continuation report
  - source-heavy report с long title + subtitle stress
  - appendix-heavy structured docx
  - strategy-heavy
  - form-like
  - resume-like
  - table-heavy
  - chart-heavy
  - image-heavy
  - fact-only

Текущий runner:

- [run_quality_contracts.py](../scripts/run_quality_contracts.py)

## Почему используется отдельный runner

Проект уже использует `unittest` везде.

Отдельный runner здесь самый практичный вариант, потому что он:

- переиспользует текущий test stack
- не требует вводить ещё один task runner или marker system
- даёт стабильную команду для локальных проверок и CI gate
- удерживает quality gate как намеренно отобранный набор проверок, а не как смесь со всем backend suite

## Как запускать

Из корня репозитория:

```bash
.venv\Scripts\python.exe scripts/run_quality_contracts.py
```

## Когда запускать

Рекомендуемые случаи:

- после изменений в `planner`
- после изменений в `pptx_generator`
- после изменений в layout/template logic
- после изменений в template analyzer или template-aware geometry logic
- перед переносом изменений из `dev` в `test`
- перед release verification

## Граница ответственности

Этот слой должен отвечать на вопросы:

- отрендерилась ли колода с ожидаемыми layout-контрактами?
- остались ли planner и generator согласованы по slide capacity rules?
- сохранился ли порядок mixed paragraph/list content после рендера?
- остались ли стабильными представительные document classes?
- остаётся ли изменение корректным не только для одного built-in template или одного regression document?

Он не должен заменять:

- полный backend suite
- API contract tests
- frontend smoke и visual tests

Эти слои должны запускаться вместе с `quality-contracts`, а не вместо него.

## Контракты, связанные с графиками

Для axis labels, `numFmt`, compact values, `%`, `₽`, data labels и mixed-unit charts
обязательная внешняя документация фиксируется в [analysis_rule.md](analysis_rule.md).
Именно она должна определять chart-formatting contract, а не локальная эвристика одного компонента.

Текущее поведение chart pipeline:

- chartable tables могут превращаться в реальные chart slides через explicit chart overrides
- default chart candidates включают column, line, bar, stacked column/bar и pie там, где это уместно
- combo остаётся поддержанным в generator для explicit specs и legacy plans
- безопасный mixed-unit сценарий `number/RUB + %` теперь должен по умолчанию идти как `combo` с secondary value axis,
  а не как single-axis column/line
- shared backend chart render semantics теперь централизованы, поэтому generator и deck-audit используют один и тот же contract для visible series, combo fallback и axis formats
- mixed-unit combo charts теперь могут рендериться с secondary value axis, когда это требуется render contract'ом
- unsafe mixed-unit tables с слишком большим числом unit families считаются not chartable
- ordinal/status tables с индексоподобными значениями `1..N` считаются not chartable
- summary rows вроде `Итого` отфильтровываются до построения chart series
- column, bar, line, stacked, pie и explicit combo покрыты generator XML tests
- chart slides используют тот же layout-quality contract, что и остальная колода
- для chart value axis проверяются компактные форматы вроде `млн`, `млрд`, `%` и `₽`
- для chart title/subtitle проверяется общий контракт `28 pt` / `18 pt`
- template-aware audit не должен применять image-panel rules к text slides только потому, что prototype layout основан на `image_text`
- placeholder-aware capacity contract теперь общий для generator и deck-audit: smaller uploaded placeholders должны давать более строгий `max_chars` и не выходить за свои font bounds в реальном артефакте
- dense narrative continuation contract теперь тоже закреплён: planner обязан реально использовать `dense_text_full_width`, когда он уменьшает число continuation slides без нарушений quality bounds
- тот же dense contract теперь распространяется и на paragraph-dominant mixed narrative + bullets scenarios, если bullets только уточняют основной narrative, а не превращают слайд в list-first layout
- quality gate теперь отдельно подтверждает continuation-balance для narrative stress-case с длинным title, явным subtitle и многочастным body; subtitle должен оставаться только на первом слайде continuation-group
- quality gate теперь дополнительно подтверждает, что source-heavy reference-tail не попадает в main deck даже в stress-case с длинным title, явным subtitle и continuation
- `deck_audit` теперь проверяет body underfill как placeholder-level fill contract и использует явную footer geometry там, где footer привязан к реальному placeholder idx
- `deck_audit` теперь дополнительно ловит аномально большой разрыв между `title/subtitle` и `body`, но делает это как runtime-эвристику для реального layout-контракта, не штрафуя штатный стандартный gap встроенного и uploaded template path'ов
- для prototype templates footer geometry теперь тоже проходит через synthetic footer idx в `deck_audit`, поэтому проверки `narrow_footer` и `footer_left_misalignment` больше не ограничены только обычными placeholder-layout путями
- placeholder-aware fill contract расширен и на детерминированный auxiliary text placeholder в `list_with_icons`;
  если ожидаемый payload левой колонки потерялся при рендере, `deck_audit` теперь возвращает `underfilled_auxiliary_placeholder_fill`
- placeholder-aware fill contract расширен и на обе детерминированные текстовые колонки `list_with_icons`:
  `deck_audit` теперь дополнительно проверяет payload правой колонки по placeholder idx `14`
  через `underfilled_two_column_placeholder_fill`, даже если runtime-path резолвит этот slot как body, а не auxiliary
- placeholder-aware fill contract расширен и на `contacts` layout:
  `deck_audit` теперь проверяет loss имени, роли, телефона и email по конкретным placeholder idx `10/11/12/13`
  и возвращает `underfilled_contact_placeholder_fill`, если один из подтверждённых contact-slots потерял payload
- placeholder-aware fill contract расширен и на `cards_3`:
  `deck_audit` теперь проверяет loss карточного payload по placeholder idx `11/12/13`
  и возвращает `underfilled_card_placeholder_fill`, если один из card-slots потерял текст
- placeholder-aware fill contract расширен и на обычный `subtitle` stack для content-slides:
  если `subtitle` должен существовать как отдельный placeholder, а не быть намеренно очищен как duplicate body intro,
  `deck_audit` теперь возвращает `underfilled_subtitle_placeholder_fill`
- planner dense continuation contract усилен и для pure narrative path:
  `dense_text_full_width` теперь реально доходит до continuation render path,
  но compaction не имеет права схлопывать subtitle-bearing stress-case обратно в один слайд
- template-aware chart contract расширен на prototype templates с `chart_image` binding:
  generator обязан рендерить реальный chart shape, применять chart title/subtitle contract `28/18 pt`,
  а `deck_audit` обязан выбирать тот же chart prototype, что и generator, и сравнивать chart/footer width с template slot, а не с full-width встроенным layout

Смежные frontend-проверки:

- structure drawer smoke проверяет переключение между `table` и `chart`
- chart type select больше не показывает combo для default mixed-unit candidates
- hidden series сохраняются в `chart_overrides` payload
- transpose mode разрешён только для chart specs, которые остаются семантически безопасными после обмена местами series и categories
- preview smoke покрывает column, bar, line, stacked column, stacked bar, pie и explicit combo paths
- line preview smoke проверяет marker/line/label counts и защищает от `NaN` координат
- runtime chart smoke должен мокать только `/documents/extract-text`; далее `/plans/from-text`, `/presentations/generate`
  и download должны идти через реальный локальный backend
- runtime smoke должен использовать отдельный frontend port `4173` и локальный backend `8000` по умолчанию;
  для параллельных раннеров допускаются только явные override-переменные `PLAYWRIGHT_FRONTEND_PORT` / `PLAYWRIGHT_BACKEND_PORT`
- mocked extraction fixture в runtime smoke обязан повторять форму реального extractor: если таблица есть в `tables`,
  соответствующий ordered block должен идти как `DocumentBlock(kind="table", table=...)`; иначе planner может создать
  лишний table-slide рядом с chart-slide и smoke перестанет доказывать чистый table->chart override path
- chart axis format contract теперь дополнительно запрещает применять percent format к primary axis,
  если на ней есть непроцентные series; `%` допустим только для оси, где все series процентные

## Правило исполнения

Этот verification layer должен поддерживать общий рабочий стандарт проекта:

- исправление считается глобальным только если проходит document-class и template-aware проверки
- один проблемный файл является regression case, но не доказательством допустимости локальной подгонки
- после исправления одного проявления нужно проверять и соседние проявления в связанных компонентах, layout'ах и render path'ах
- верификация после изменений должна включать targeted regressions, representative document classes и более широкий backend suite, а не только один happy-path rerun
- для generation-related fixes обязательна свежая генерация `.pptx` после изменений и прямая проверка нового артефакта; анализ старой колоды допустим только для диагностики причины
- после нахождения следующего безопасного шага работа должна продолжаться автоматически без лишних stop-and-ask пауз
