# folder-to-md

Batch converts documents (PDF, Word, PowerPoint, EPUB) to Markdown files.

## Supported Formats

| Extension       | Converter                   |
| --------------- | --------------------------- |
| `.pdf`          | `pymupdf4llm`               |
| `.docx`         | `markitdown`                |
| `.doc`          | Word COM API → `markitdown` |
| `.ppt`, `.pptx` | `markitdown`                |
| `.epub`         | `markitdown`                |

## Requirements

- Windows (required for `.doc` conversion via Word COM)
- Microsoft Word installed (for `.doc` files)
- Python 3.x

Install dependencies:

```bash
pip install -r requirements.txt
```

## Configuration

Default folders are defined in `config.py`:

```python
DEFAULT_INPUT_FOLDER = "sample"
DEFAULT_OUTPUT_FOLDER = "output"
```

Edit these values to change the defaults without using CLI arguments.

## Usage

```bash
python main.py [source_folder] [output_folder]
```

Both arguments are optional and fall back to the defaults in `config.py`.

**Examples:**

```bash
# Use defaults from config.py (sample/ → output/)
python main.py

# Convert documents in a specific folder, save to default output folder
python main.py /path/to/documents

# Convert documents and save Markdown files to a specific folder
python main.py /path/to/documents /path/to/output
```

Each converted file is saved as `<original_name>.md`. Conversion errors are reported per-file without stopping the batch.
