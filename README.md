# AI Agentic Form Filler

An intelligent tool for automatically filling DOCX templates with data from JSON files. This project uses Large Language Models (LLMs) to semantically map JSON data to document placeholders and tables, ensuring robust handling of complex forms without rigid template requirements.

## 🚀 Features

-   **Semantic Mapping**: Uses LLMs (`GPT-5.1`) to intelligently map JSON fields to DOCX placeholders, identifying the correct data even if field names don't match exactly.
-   **Table Automation**: Automatically detects data tables in the DOCX, matches them to JSON arrays, and maps columns to the correct object keys.
-   **Format Preservation**:
    -   Fills data in-place, preserving original styles, fonts, and layout.
    -   Handles newlines in JSON strings by creating proper new paragraphs in DOCX.
    -   Supports checkboxes ("CHECKED"/"UNCHECKED") and complex text structures.
-   **Caching System**: Implements a hashing mechanism to cache LLM decisions. Rerunning the tool on the same inputs is fast and deterministic.
-   **Detailed Logging**: Generates comprehensive logs (`docx_full_log.txt`) tracing every filled field back to its source JSON key.

## 🛠️ Setup & Installation

### Prerequisites
-   **Python 3.13+**
-   [uv](https://github.com/astral-sh/uv) (recommended) or `pip`
-   OpenAI API Key

### Installation

1.  **Clone the repository:**
    ```bash
    git clone <repository_url>
    cd ai-engeneer-test-task
    ```

2.  **Environment Setup:**
    Create a `.env` file in the root directory:
    ```bash
    OPENAI_API_KEY=sk-your-api-key-here
    ```

3.  **Install Dependencies:**
    Using `uv` (faster, recommended):
    ```bash
    uv sync
    ```
    Or using standard `pip`:
    ```bash
    pip install .
    ```

## 💻 Usage

Run the main script providing the path to your DOCX template and the input JSON data.

```bash
uv run python main.py <path_to_docx> <path_to_json>
```

**Example:**
```bash
uv run python main.py "resources/template.docx" "resources/input_data.json"
```

The script will:
1.  Analyze the DOCX to find placeholders (e.g., `{{Name}}`, `.......`).
2.  Consult the LLM/Cache to map these placeholders to the JSON data.
3.  Generate a filled document named `<original_name>.filled.docx` in the `resources/` folder.

### 🧹 Cache Management

To clear all caches (LLM mapping, Python, Ruff, Mypy) and force a full re-run:

```bash
chmod +x clean.sh
./clean.sh
```

## 🏗️ Architecture & Design Decisions

### 1. LLM-Driven Field Mapping (`field_mapper.py`)
Instead of relying on strict regex matching (e.g., waiting for `{{key}}` to match `json["key"]`), this tool extracts the *context* around every potential placeholder (underscores, dots, brackets) and asks an LLM: *"Given this text context and this JSON data, what value belongs here?"*.
-   **Justification**: This allows filling forms with inconsistent naming conventions or where the template uses visual cues (lines `_______`) instead of named keys.

### 2. Smart Caching
To avoid excessive API costs and latency, all LLM mapping decisions are cached in `.cache/mapping.json` and `.cache/table_mapping.json`.
-   **Mechanism**: A hash is computed from the {Placeholder Contexts + JSON Data}. If the inputs haven't changed, the cached mapping is used instantly.

### 3. Newline Handling
JSON strings containing `\n` are not just inserted as text (which often breaks DOCX formatting). They are split, and new DOCX paragraphs are created for each line, inheriting the style of the original placeholder paragraph.

### 4. Table Filling Strategy
Tables are handled in two steps:
1.  **Table-to-Array Matching**: The LLM decides which JSON array corresponds to which DOCX table based on headers and content.
2.  **Column-to-Key Matching**: For each column, the LLM maps the header text (e.g., "Unit Price (RON)") to the specific JSON key (e.g., `price_ron`).

## ⚠️ Known Limitations

1.  **LLM Dependency**: The tool requires an active internet connection and a valid OpenAI API key. It is not designed for offline use (unless the cache is already warm).
2.  **Performance**: Initial runs on large documents (50+ pages) may take time due to sequential batch processing of placeholders by the LLM.
3.  **Placeholder Detection**: The analyzer (`docx_analyzer.py`) uses heuristics to find "placeholders" (patterns like `....`, `____`, `{{...}}`). Extremely unconventional placeholder styles might be missed.
4.  **Complex Tables**: While standard tables are supported, nested tables or tables with heavy cell merging might experience layout shifts or mapping inaccuracies.

## 🧪 Running Tests
    
To ensure the reliability of the application, we have a comprehensive test suite covering unit logic and end-to-end integration.

### Run All Tests
Execute the helper script:
```bash
./run_tests.sh
```

Or run manually using `uv`:
```bash
uv run pytest
```

### Run Specific Tests
To run only the integration tests:
```bash
uv run pytest tests/integration/test_workflow.py
```

## 🧹 Code Quality

This project uses **Ruff** for linting/formatting and **Mypy** for static type checking.

### Running Linters (Ruff)
To check for linting errors and verify formatting:
```bash
./ruff.sh
```
Or manually:
```bash
uv run ruff check .
uv run ruff format --check .
```

### Running Type Checker (Mypy)
To verify type safety:
```bash
./mypy.sh
```
Or manually:
```bash
uv run mypy .
```