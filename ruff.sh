#!/bin/bash
set -e

echo "🚀 Running Ruff (Linting & Formatting)..."
uv run ruff check .
uv run ruff format --check .
echo "✅ Ruff checks passed!"
