#!/bin/bash
set -e

echo "🚀 Running Mypy (Type Checking)..."
uv run mypy .
echo "✅ Mypy checks passed!"
