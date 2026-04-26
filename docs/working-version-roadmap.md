# Roadmap до Fully Working Version

## Цель

Под fully working version здесь понимается не просто "генерация работает", а рабочая продуктовая версия, в которой:

- пользователь стабильно проходит путь `upload -> plan -> review -> generate -> download`
- backend, generator и review-step согласованы по одному контракту данных и preview
- representative document classes проходят quality gate без ручной подгонки под один кейс
- uploaded templates поддерживаются без ломки layout/capacity контрактов
- smoke/runtime checks ловят основные product regressions до деплоя

## Архитектурный принцип

- проект не должен опираться на один эталонный шаблон как на модель правильного поведения
- `corp_light_v1` или любой другой registry template допустим только как временный reference smoke anchor
- целевой контракт системы: извлекать usable inventory из произвольного `.pptx`, а не воспроизводить поведение одного template id
- analyzer/generator/review должны опираться на общие возможности PowerPoint:
  text/body/title placeholders, tables, charts, images, footer/aux slots, geometry, margins, component roles
- template-specific эвристики допустимы только как degradеd fallback и не должны определять основной runtime path
- regression и release gate должны постепенно переходить от template-id checks к invariant-based checks для произвольных презентаций

## Что уже есть

- устойчивый `docx -> plan -> pptx` pipeline
- template-aware generator/analyzer path
- первый generic editable-slot metadata contract в analyzer/manifest path для user-uploaded templates
- review-step начал использовать analyzer-derived manifest metadata для representation targeting вместо жёсткого `cards_3`
- analyzer/manifest path получил первый representation-level hint layer (`representation_hints`) для layout/prototype targeting
- frontend review path начал использовать manifest metadata и для registry-selected templates, а не только для uploaded template flow
- review-step начал убирать layout-name heuristics и в текстовом structure flow: data-slide detection теперь идёт через manifest lookup
- representation-level analyzer contract расширен дальше: `representation_hints` теперь покрывают не только `cards`, но и `table` / `image` / `contacts`
- появился ручной testable UI slice для нового template-driven flow:
  active template analysis виден в composer screen без чтения backend логов
- plan/build path начал использовать реальный layout inventory шаблона:
  `preferred_layout_key` может адаптироваться от logical alias к найденному layout key
- quality contracts и regression corpus по основным классам документов
- backend chart render contract и deck-audit для chart semantics
- второй шаг review уже переведён в editable/template-aware flow и backend-preview path
- placeholder-aware audit уже покрывает body area, `list_with_icons`, `contacts`

## Что ещё нужно закрыть

### P0. Стабильность ядра генерации

- Довести placeholder-level fill targets до остальных deterministic layout roles.
- Углубить continuation heuristics planner'а с учётом placeholder-aware capacity, а не только layout defaults.
- Расширить runtime/regression coverage на длинные narrative и layout-stress документы.
- Усилить PPTX quality layer для continuation pairs, dense bullets, appendix/question slides и соседних compact/underfilled состояний.

### P0. Product-ready review step

- Довести второй шаг до полностью надёжного `editable deck -> regenerate -> backend-rendered validation/export`.
- Закрыть остаточные smoke/api gaps вокруг review-state, regenerate race conditions и inline block editing.
- Сделать `powerpoint_rendered` целевым стабильным runtime path, а `fallback_rendered` строго деградационным сценарием.

### P0. End-to-end parity для chart flow

- Проверить один и тот же `ChartSpec` через preview controls, API payload, `plan response`, generator и `deck_audit`.
- Дожать transposed/mixed-unit safety так, чтобы UI не мог собрать semantically unsafe chart request.
- Расширить visual/runtime checks для dense labels, negative values и chart-heavy document classes.

### P1. Template robustness

- Перевести template-aware path с reference-template mindset на generic PowerPoint extraction contract.
- Продолжить перенос layout/component contracts из hardcoded defaults в manifest/analyzer metadata.
- Расширить новый generic editable-slot contract от role/capability metadata к confidence, repeated-group и safe-editability flags.
- Добрать placeholder/style/runtime coverage для uploaded templates по документным классам.
- Закрыть оставшиеся layout-specific rules, где generator и audit ещё опираются на грубый `layout_key`.
- Проверять template extraction по инвариантам PowerPoint-компонентов, а не по совпадению с конкретным registry template.

### P1. Semantic quality

- Дальше уменьшать false positives в semantic normalizer без поломки form/resume scenarios.
- Добавить больше source-heavy, appendix-heavy и question/callout-heavy regression fixtures.

### P1. Release hardening

- Свести обязательный pre-release gate в один воспроизводимый набор команд.
- Убедиться, что локальный dev/runtime/deploy path совпадает с production contract.
- Добрать эксплуатационные smoke checks вокруг template upload, output storage и download flow.

## Предлагаемый порядок выполнения

1. Закрыть remaining `deck_audit` fill/gap invariants для deterministic placeholders.
2. Перенести эти ограничения глубже в planner continuation heuristics.
3. Усилить regression corpus и `quality-contracts` на narrative/layout-stress cases.
4. Добить chart parity path `preview -> payload -> plan -> generated deck -> audit`.
5. Довести review-step до release-grade runtime с устойчивым `powerpoint_rendered`.
6. Провести release hardening: полный gate, smoke, deploy verification.

## Definition of Done для рабочей версии

- representative backend suite зелёный
- `quality-contracts` зелёный
- frontend `yarn verify` зелёный
- Playwright smoke/visual для app, chart preview и review-step зелёные
- свежая runtime-генерация `.pptx` проходит `deck_audit` на нескольких классах документов
- chart-heavy, narrative-heavy и uploaded-template scenarios проходят без ручных workaround
- пользовательский сценарий `upload -> edit/review -> generate -> download` воспроизводим локально и на сервере

## Ближайший practical plan

- вынести `generic uploaded template extraction` в отдельный основной контракт
- расширять corpus synthetic/real `.pptx` без привязки к одному reference template
- добавлять analyzer tests на общие PowerPoint-компоненты и degradеd handling плохих шаблонов
- расширять `deck_audit` на `cards` и остальные deterministic placeholder roles
- затем сразу перейти к planner/runtime hardening для dense continuation
- после этого закрывать chart parity payload checks
