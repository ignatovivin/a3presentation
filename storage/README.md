# Storage Layout

`storage/` contains local runtime artifacts. Product builds do not ship bundled templates.

## Not versioned in git

- `generated/`
  - temporary or historical generated `.pptx` files
- `outputs/`
  - runtime output presentations created by the backend
- `templates/`
  - runtime user-uploaded templates when `TEMPLATES_DIR` points inside `storage`

These folders are local artifacts only. They can be cleaned, regenerated, or archived outside git.

## Rule of thumb

- do not commit runtime template sources
- keep test-only template files under `tests/fixtures/templates`
- do not commit generated presentation results
- do not use `storage/outputs` as a document archive
- use generated outputs only as local runtime or manual QA artifacts
