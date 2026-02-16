#!/bin/bash

# Remove custom LLM cache
if [ -d ".cache" ]; then
    rm -rf .cache
    echo "✅ Removed .cache"
fi

# Remove Python cache
find . -type d -name "__pycache__" -exec rm -rf {} +
echo "✅ Removed __pycache__ directories"

# Remove tool caches
if [ -d ".ruff_cache" ]; then
    rm -rf .ruff_cache
    echo "✅ Removed .ruff_cache"
fi

if [ -d ".mypy_cache" ]; then
    rm -rf .mypy_cache
    echo "✅ Removed .mypy_cache"
fi

echo "✨ Cache cleaning complete!"
