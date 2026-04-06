#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

mkdir -p data/templates data/outputs

if command -v rsync >/dev/null 2>&1; then
  rsync -a --delete storage/templates/ data/templates/
else
  rm -rf data/templates/*
  cp -r storage/templates/. data/templates/
fi

docker compose -f docker-compose.server.yml up -d --build

