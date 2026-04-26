# Project Instructions

## Response Economy

- Minimize token usage by default.
- Keep status updates short and only report decisions, results, blockers, and the next concrete action.
- Do not restate prior context unless it is necessary for the current step.
- Prefer references to files and completed work over long recaps.

## Chat Rollover

- Warn the user to start a new chat when context becomes large enough that long recaps or repeated file history would waste tokens.
- Give that warning before token pressure becomes a blocker.
- When suggesting a new chat, include a compact handoff summary with current plan, completed work, and the exact next step.

## graphify

В проекте используется граф знаний `graphify` в `graphify-out/`.

Правила навигации по контексту:
- Всегда сначала обращаться к графу знаний.
- Сырые файлы читать только если пользователь явно просит это сделать.
- Если существует `graphify-out/wiki/index.md`, использовать его как основной вход в контекст вместо прямого чтения файлов.
- Перед ответами по архитектуре и устройству кодовой базы сначала смотреть `graphify-out/GRAPH_REPORT.md`, чтобы увидеть ключевые узлы и структуру сообществ.
- После изменений в коде в рамках сессии запускать `graphify update .`, чтобы граф оставался актуальным без лишних токенов и без LLM-стоимости для code-only изменений.
