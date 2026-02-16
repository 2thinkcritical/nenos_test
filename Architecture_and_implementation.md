# Project Architecture and Implementation Strategy Analysis

Here is a detailed breakdown of the architectural approach and recommended implementation strategy.

## 1. Architectural Approach

The solution adopts an **Agentic/Semantic pipeline architecture**. Instead of relying on rigid template engines (like Jinja2 or Mustache) that require the DOCX template to have specific variable names matching the data (e.g., `{{customer_name}}`), this system uses **Large Language Models (LLMs)** to bridge the gap between the document's visual structure and the data.

### Core Philosophy
-   **Decoupling**: The document template and the JSON data structure are completely decoupled. They do not need to share naming conventions.
-   **Semantic Reasoning**: The system understands "context". If a line says "Signed by: ______", the system infers that a name is needed there, even if the placeholder is just underscores.
-   **In-Place Modification**: The system modifies the existing DOCX binary structure rather than generating a new document from scratch, ensuring 100% preservation of complex layouts, headers, and styles.

### High-Level Pipeline
The data flow follows a strictly sequential 4-stage pipeline:

1.  **Analysis (Scanner)**: The system scans the raw DOCX file to identify all "fillable" areas (placeholders) and structural elements (tables).
2.  **Semantic Mapping (Reasoner)**: It extracts the text surrounding each placeholder and sends it to an LLM along with the JSON data. The LLM decides *exactly* what value belongs in each spot.
3.  **Data Processing (Transformer)**: Complex data types (tables, checkboxes, multi-line strings) are processed to match the document's format.
4.  **Execution (Filler)**: The system injects the determined values into the specific XML nodes of the document, adjusting the internal structure (e.g., creating new paragraphs) as needed.

---

## 2. Implementation Strategy

The implementation is divided into three distinct modules, each solving a specific class of problems.

### A. Document Analysis (`docx_analyzer.py`)
**Strategy: Heuristic Pattern Recognition**

Instead of forcing users to add `{{brackets}}` to every field, the generalized implementation supports natural forms.

-   **Detection Logic**: Uses Regex to find visual cues commonly used in forms:
    -   Underscore lines (`______`)
    -   dotted lines (`......`)
    -   Unicode ellipses (`…`)
    -   Checkboxes (`|_|`)
-   **Context Extraction**: This is the most critical step. For every placeholder found, the analyzer extracts:
    -   **Immediate Context**: Text immediately before/after the blank in the same paragraph.
    -   **Surrounding Context**: Text from neighboring paragraphs to capture section headers (e.g., "SECTION 5: COST").
    -   **Location Metadata**: Whether it's in a header, footer, table cell, or body text.

### B. Semantic Mapping Engine (`field_mapper.py`)
**Strategy: Batch LLM Processing with Caching**

This module acts as the "brain".

-   **Prompt Engineering**: The prompt doesn't just ask "fill this". It provides the *context* (e.g., "The placeholder is next to 'Total Cost:'") and the *entire JSON*, evaluating the best match semantically.
-   **Batching**: To handle large documents efficiently, placeholders are grouped into batches (e.g., 10 at a time) to reduce HTTP overhead and stay within context windows.
-   **Determinism & Caching**:
    -   Since LLMs can be non-deterministic, a robust caching layer (SHA-256 hash of inputs) is implemented.
    -   If the code runs twice with the same inputs, it skips the LLM and uses the cached JSON mapping. This is crucial for development speed and cost control.
-   **Table Strategy**:
    -   **Step 1**: Match Table to Array. (Which JSON list corresponds to this table?)
    -   **Step 2**: Match Column to Key. (Which object key corresponds to the "Unit Price" column?)

### C. Document Filler (`docx_filler.py`)
**Strategy: Low-Level XML Manipulation**

Standard libraries like `python-docx` are good for reading but often struggle with complex *in-place* editing. The implementation uses a hybrid approach:

-   **Preserving Styles**: Instead of creating new runs, it modifies the `text` attribute of existing runs. This preserves bold/italic/font/size settings found in the template.
-   **The "Newline Problem"**: A common pitfall is inserting text with `\n` into a single Run, which DOCX ignores or renders poorly.
    -   **Solution**: The implementation detects newlines and performs a "Paragraph Split". It clones the paragraph's XML properties (`<w:pPr>`) and creates new physical paragraphs for each line of text, ensuring proper spacing and indentation is inherited.
-   **Row Cloning**: For dynamic tables, the system identifies the "data row" and clones its XML element (`<w:tr>`) for every item in the JSON array, ensuring the filled table looks exactly like the template's example row.

---

## 3. Key Components Summary

| Component | File | Responsibility |
| :--- | :--- | :--- |
| **Orchestrator** | `main.py` | Loads files, manages the pipeline, handles CLI arguments and logging. |
| **Analyzer** | `docx_analyzer.py` | Regex-based scanning. Returns `PlaceholderInfo` objects containing location and context. |
| **Mapper** | `field_mapper.py` | Interacts with OpenAI API. manages `mapping.json` cache. Handles fuzzy logic. |
| **Filler** | `docx_filler.py` | Manipulates `python-docx` objects and OXML elements. Writes the final `.docx` file. |


