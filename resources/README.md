# Example resources

- `sample_form.docx`: the existing Romanian procurement-form template used by the project.
- `sample_input.json`: illustrative input for that template. Identity-document values and representative names are explicitly marked as demo data; do not use this fixture for a real filing.

Some array values are encoded as JSON strings to demonstrate the original input format. New inputs can use native JSON arrays instead. The template and input are examples, not a compliance specification. Inspect both before sending their content to a model.

The CLI writes a separate `*.filled.docx` beside the template. Generated documents and Office lock files are ignored by Git.
