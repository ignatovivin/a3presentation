#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

mkdir -p data/outputs

export APP_COMMIT_SHA="$(git rev-parse HEAD)"
export APP_COMMIT_BRANCH="$(git branch --show-current)"

docker compose -f docker-compose.server.yml down
docker compose -f docker-compose.server.yml up -d --build
