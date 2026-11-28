# Architecture and implementation

The application has a sequential analyze → map → fill → report pipeline. Its purpose is to adapt existing forms without requiring authors to rename every blank to a source-data key. The model is used for semantic decisions; document traversal and editing remain explicit Python code.

## Boundaries

`pipeline.py` loads a DOCX and a JSON object, invokes each stage, prints unresolved fields and unused source keys, writes a report, and saves a filled copy. `cli.py` turns failures into exit status 1 with a one-line message on stderr; `--help` works without input files. Missing input files, invalid JSON, and a non-object JSON root produce actionable CLI errors. Corrupt DOCX files and provider failures receive actionable errors; raw provider exceptions are not printed. `--output` controls the destination and cannot target the source template. `reporting.py` owns human-readable diagnostics separately from orchestration.

`analyzer.py` describes blanks with `PlaceholderInfo` and data tables with `TableInfo`. The descriptions include nearby text, paragraph/run positions, table coordinates, section context, and the original character offset. The shared `placeholder_id()` function gives the mapper and filler the same identity, including when two identical blanks occur in one run.

`mapper.py` sends batches of up to ten field descriptions together with the source JSON to the configured OpenAI model. Separate requests match tables to arrays and columns to object keys. All model traffic goes through `llm.py`, which owns client settings (timeout, retries) and prompt files and parses a reply into untrusted objects, accepting a fenced JSON response. Unknown IDs, columns, source keys, and invalid field value shapes are rejected; unresolved fields stay blank. These checks validate the output contract, not semantic truth.

`filler.py` updates existing runs when possible and uses XML operations for split runs, paragraph insertion, and row cloning. Replacements within a paragraph are applied from right to left so earlier character offsets remain valid. Table cells are addressed by paragraph before applying replacements. Checkboxes are textual `|_|` markers; selecting a later checkbox does not change an earlier unselected one.

## Table strategy

The orchestrator collects native JSON arrays and strings that decode to arrays. The model proposes a table-to-array mapping; object arrays receive a separate column mapping. If those mappings are unavailable, the filler can fall back to normalized header/key matching, including Romanian diacritics and word-prefix heuristics. This fallback is useful but can choose the wrong field when labels are ambiguous.

Rows beyond the template's allocated data rows are cloned from a pristine snapshot and inserted immediately after the data region, preserving trailing totals and notes. Missing/null object values stay empty and numeric zero is preserved. Object values are converted to text; arrays of strings use an empty cell. This is intended for conventional data tables. Nested/merged layouts and heterogeneous arrays are not a general-purpose spreadsheet or Word form engine.

## Caching and repeatability

| Cache | Inputs used in the key | Stored decisions |
| --- | --- | --- |
| Field mapping | Complete placeholder descriptions, source JSON, model name, cache version | Field values and reported source keys |
| Table mapping | Table descriptions, array contents, model name, cache version | Table index → source key |
| Column mapping | Column headers, available object keys, model name | Column index → object key |

Malformed or unreadable cache JSON is treated as a miss. Shape checks reject malformed field decisions and invalid table/column indices. Local cache files are plaintext and can contain filled values. Writes are atomic, but fixed filenames make the cache unsuitable for concurrent workers sharing a directory. Prompt-code changes and provider-side model-alias changes require manual cache clearing.

## Verification strategy

The test suite exercises real DOCX traversal and mutation. Only the external model is replaced. End-to-end tests use source keys that differ from table headers, ensuring that semantic table and column mapping is exercised instead of silently succeeding through a header-match fallback. Both native and encoded JSON arrays are covered.

Regression tests cover repeated blanks, selected checkboxes, malformed model response shapes, changes to document context/source-array schema/model selection, malformed cache JSON, CLI help, and JSON-root validation. An autouse fixture gives each test its own working directory, preventing deletion or reuse of a developer's cache and output documents.

Ruff checks imports, common errors, and formatting; Mypy checks the project's annotated Python interfaces. CI uses the locked dependency set on Python 3.13. The unit suite does not call a live model. The versioned evaluation corpus separately checks the versioned synthetic forms using authored response fixtures; an opt-in live mode records exact-document matches, requests, tokens, and elapsed time. A dedicated CI workflow renders four before/after examples in an isolated LibreOffice profile. These checks do not prove live-model accuracy or pagination for unseen documents.

## Trade-offs and next steps

- **Flexible detection versus explicit templates:** visual-blank heuristics allow legacy forms, but need review when punctuation resembles a blank. Explicit template markers would trade flexibility for stronger guarantees.
- **Generated values versus strict extraction:** semantic mapping can combine and transform source fields, but source-key attribution is model-reported. A stricter system would return validated extraction operations instead of free-form values.
- **Existing layout versus layout guarantees:** editing in place reduces unnecessary reconstruction but cannot promise that inserted text will fit or preserve pagination.
- **Local prototype versus service:** a service would need explicit output locations, isolated jobs, atomic cache writes, input limits, privacy controls, structured error handling, and model-quality evaluation.

These are boundaries of the implementation, not guarantees delegated to the model prompt.
