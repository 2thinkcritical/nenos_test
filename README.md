# DOCX Semantic Form Filler

[![CI](https://github.com/2thinkcritical/docx-form-filler/actions/workflows/ci.yml/badge.svg)](https://github.com/2thinkcritical/docx-form-filler/actions/workflows/ci.yml) ![Python 3.13](https://img.shields.io/badge/python-3.13-blue) [![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)

Fill existing Word forms from JSON when the document's labels and source-data keys do not match exactly. The tool detects visual blanks, asks an OpenAI model to propose field and table mappings, then edits the document with `python-docx`.

This is a command-line prototype built around Romanian procurement forms. Its architecture separates document analysis, semantic mapping, and document editing so each stage can be tested independently.

## Try the reproducible demo

No API key is needed to exercise thirteen synthetic forms end to end:

```sh
git clone https://github.com/2thinkcritical/docx-form-filler.git
cd docx-form-filler
uv sync --frozen
uv run python -m evaluation.run --mode fixture
```

This produces before/after DOCX files and a comparison report under `evaluation/output/`. The current baseline is **13/13 document-content matches with authored response fixtures**. This validates the pipeline on these examples; live-model semantic accuracy is not measured by fixture mode.

See the [rendered before/after gallery](docs/examples/README.md) and the [evaluation protocol](evaluation/README.md), including opt-in live runs with token and latency reporting.

## Quick start

Requires **Python 3.13+**, [uv](https://docs.astral.sh/uv/), and an OpenAI API key for uncached mappings.

```sh
git clone https://github.com/2thinkcritical/docx-form-filler.git
cd docx-form-filler
uv sync --frozen
cp .env.example .env
```

Set `OPENAI_API_KEY` in your local `.env`, then run the included example:

```sh
uv run docx-fill resources/sample_form.docx resources/sample_input.json --verbose
```

The example makes paid API requests and sends document context and source data to the configured model. Review the sample and use a key with appropriate usage limits. The default model is `gpt-5.1`; `OPENAI_MODEL` can select another model available to your account. Other models have not been validated by this repository's offline tests.

For command-line help:

```sh
uv run docx-fill --help
uv run docx-fill template.docx data.json --output output/filled.docx
```

## Inputs and outputs

The first argument is a `.docx` template; the second is a UTF-8 JSON file whose root must be an object.

```json
{
  "provider": "DEMO Example Company",
  "date": "2026-11-23",
  "services": [
    { "offering": "Consulting", "amount_ron": "500" },
    { "offering": "Development", "amount_ron": "1000" }
  ]
}
```

Supported visual markers include underscore runs (`______`), four or more dots (`......`), Unicode ellipses (`…`), and text checkboxes (`|_|`). The analyzer uses surrounding text and table headers as context. Named markers such as `{{name}}` are **not** a supported detection syntax.

Table data can be a JSON array or a string containing an encoded JSON array. Arrays of objects use semantic column mappings with a heuristic fallback; arrays of strings are written into an empty cell in each data row. Prefer native arrays in new inputs. Empty or unconventional tables, mixed row types, nested tables, and heavily merged layouts require manual review.

The filled copy is written beside the template unless `--output` is given. The report and cache are relative to the directory where you run the command:

| Output | Purpose |
| --- | --- |
| `resources/sample_form.filled.docx` | Filled copy beside the template; the original is not modified |
| `docx_full_log.txt` | Analysis, mapping statistics, unfilled fields, and value-level fill log |
| `.cache/mapping.json` | Field decisions and an input hash |
| `.cache/table_mapping.json` | Table and column decisions with cache keys |

Use `--output path/to/result.docx` to choose the destination. The CLI checks file identity, including hard links and case aliases, to protect the source template, and saves through an atomic replacement so a failed write leaves an existing result intact. JSON null/missing table values stay empty; numeric zero is preserved. Table expansion inserts new rows before existing totals and notes.

The default output filename depends on the template stem. Repeating a run with that stem overwrites the existing generated file. The log is replaced on each run, and the cache stores the latest field/table decisions rather than a historical audit trail.

## How it works

```text
DOCX ──► analyzer ──► placeholders + context + data-table descriptions
                              │
JSON ─────────────────────────┤
                              ▼
                      mapper + local cache
                              │
                              ▼
                      document filler ──► filled DOCX + report
```

| Module | Responsibility |
| --- | --- |
| `cli.py` | Argument parsing, logging setup, and user-facing errors |
| `pipeline.py` | Input validation, stage orchestration, and atomic output |
| `analyzer.py` | Locate blanks and tables; build context and stable field identities |
| `llm.py` | The only model boundary: client settings, prompts, and reply parsing |
| `mapper.py` | Ask for field/table/column decisions, validate them, and cache them |
| `cache.py` | JSON decision caches with atomic writes |
| `filler.py` | Replace text, handle text checkboxes, and populate table rows |
| `matching.py` | Offline header/key heuristics used when the model gives no table mapping |
| `reporting.py` | Console summary and value-level fill report |

Modules live in `src/docx_form_filler/`; system prompts are versioned as Markdown in `src/docx_form_filler/prompts/`.

Field IDs include position offsets to distinguish repeated blanks in the same run or cell. Field cache keys include document context, source values, and the model name. Table cache keys include table descriptions, array data, and the model; column keys include headers, source keys, and the model. Changing these inputs invalidates the relevant cache.

Model output is treated as untrusted structured data: unknown field IDs, invalid value types, and unknown source keys are discarded. Invalid or absent mappings leave fields unfilled. Shape validation cannot prove that a generated value is correct or derived faithfully from the source.

See [architecture and trade-offs](docs/architecture.md) for implementation details.

## Verify without API calls

The automated suite replaces the external model boundary with fixtures. It builds small DOCX documents in temporary directories and checks the filled content, table mappings, repeated blanks, checkbox selection, CLI validation, cache invalidation, and malformed model output. Tests do not need a real API key and do not write generated files into the checkout.

```sh
uv sync --frozen
uv run pytest --cov
uv run ruff check .
uv run ruff format --check .
uv run mypy
```

CI runs these checks on Python 3.13; mypy runs in strict mode. `pre-commit install` runs Ruff on each commit. Passing offline tests verify the application logic against the test fixtures; they do not establish live-model accuracy or Word rendering fidelity for arbitrary forms.

## Data handling

- The full source JSON and selected document context are sent to OpenAI through `langchain-openai`. Table requests also include table descriptions and source-column information. Use only documents and data you are authorized to send to that service.
- Logs, filled documents, and caches can contain source values. Cache filenames containing hashes do **not** mean cached values are encrypted. Keep these artifacts local and manage their retention explicitly.
- `.env`, caches, logs, generated documents, and Office lock files are excluded from Git. `.env.example` contains placeholders only.
- Deleting `.cache/` removes saved mapping decisions. Generated DOCX documents and the report stay in place. A later run may incur new API requests.

## Design limits

The mapper proposes values; a human must review the generated document before it is signed, submitted, or relied upon. The tool does not certify procurement or legal correctness.

Detection is heuristic and may miss blanks or misclassify ordinary punctuation. Formatting is preserved where editing existing runs and cloning rows allows it; there is no guarantee of identical pagination or complex layout preservation. Text checkboxes are supported, not native Word form controls. Nested table addressing, text boxes, unusual fields, and repeated multi-line edits within one paragraph are not covered by the current regression suite.

The cache improves repeatability for unchanged inputs but does not make new LLM calls deterministic. Model alias changes and prompt-code changes are not automatically detected by the cache; clear it after updating prompts or when re-evaluating decisions. API failures stop the run with a sanitized CLI error; corrupt DOCX files receive an actionable error. No completed document is saved when model mapping fails. Concurrent processes sharing the same working directory are not supported: logs and cache files use fixed names.

The versioned evaluation corpus and measured fixture baseline are published in `evaluation/`. No live-model accuracy, production reliability claim, or repository-wide license is declared.
