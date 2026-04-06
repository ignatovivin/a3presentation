# Storage Layout

`storage/` is split into versioned source data and local runtime artifacts.

## Versioned in git

- `templates/`
  - `manifest.json` files
  - source `template.pptx` files that define real presentation templates

These files are part of the product and should stay in the repository.

## Not versioned in git

- `generated/`
  - temporary or historical generated `.pptx` files
- `outputs/`
  - runtime output presentations created by the backend

These folders are local artifacts only. They can be cleaned, regenerated, or archived outside git.

## Rule of thumb

- commit template sources
- on the single-server Timeweb deploy, templates are read directly from bundled `storage/templates` inside the backend image
- this includes tracked source `template.pptx` files, not only `manifest.json`
- do not commit generated presentation results
- do not use `storage/outputs` as a document archive
- use generated outputs only as local runtime or manual QA artifacts
