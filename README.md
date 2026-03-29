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

## Usage

```bash
python main.py <source_folder> [output_folder]
```

If `output_folder` is omitted, converted files are saved alongside the originals in `source_folder`.

**Examples:**

```bash
# Convert all documents in a folder, save Markdown files in the same folder
python main.py /path/to/documents

# Convert documents and save Markdown files to a separate folder
python main.py /path/to/documents /path/to/output
```

Each converted file is saved as `<original_name>.md`. Conversion errors are reported per-file without stopping the batch.
