# Evaluation

The versioned corpus contains twelve small synthetic forms with independently authored expected document text and tables. It includes repeated blanks, checkboxes, Romanian labels, multiline text, hyperlink offsets, headers, semantic table columns, encoded arrays, null/zero cells, and expansion before a total row.

## Reproduce application regressions without an API key

```sh
uv sync --frozen
uv run python -m evaluation.run --mode fixture
```

Fixture mode supplies **authored response fixtures** at the model boundary and checks the saved DOCX against the expected outputs. These are not recordings from a live model. Each case runs with a cold isolated cache. A failed comparison exits nonzero and records the difference in `evaluation/output/results.json`.

The deterministic baseline is **12/12 exact document-content matches**. This measures the application's handling of these fixtures. It does **not** establish semantic model accuracy, document generality, or Word pagination fidelity. A regression test deliberately supplies a wrong field value and confirms the evaluator fails that case.

## Measure a real model explicitly

```sh
# Uses your local .env. This makes paid requests and sends the synthetic corpus to the provider.
uv run python -m evaluation.run --mode live --allow-api --output-dir evaluation/live-output
```

`--allow-api` is required. Live mode uses the same forms and expected outputs with the configured `OPENAI_MODEL`, bypasses warm-cache effects, and records document exact-match rate, model name, elapsed time, number of requests, and provider-reported input/output tokens. Estimated monetary cost is left null because it depends on current provider pricing and billing details. Supply token totals to your own billing calculation; no price is assumed.

No live-model accuracy result is published here: a live run has not been performed for this release. The small corpus is a regression/evaluation starting point, not a representative production dataset. Unexpected layouts and semantically ambiguous forms need additional labeled examples.

## Inspect rendered examples

The `Visual examples` workflow builds fixture outputs, installs LibreOffice and Poppler on a fresh Ubuntu runner, and converts four before/after forms to images using an isolated LibreOffice profile:

```sh
uv run python -m evaluation.render --libreoffice /usr/bin/libreoffice
```

The renderer verifies every selected example is one page and fails rather than silently truncating additional pages. CI uploads all eight PNGs and the evaluation report. The published [example gallery](../docs/examples/README.md) is generated from that workflow and visually reviewed; it demonstrates those specific examples with LibreOffice, not universal fidelity across Word versions.
