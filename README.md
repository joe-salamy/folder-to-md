# folder-to-md

Batch convert documents (PDF, Word, PowerPoint, EPUB) to Markdown.

## Supported Formats

| Extension       | Converter                    | Platform     |
| --------------- | ---------------------------- | ------------ |
| `.pdf`          | `pymupdf4llm`                | Any          |
| `.docx`         | `markitdown`                 | Any          |
| `.ppt`, `.pptx` | `markitdown`                | Any          |
| `.epub`         | `markitdown`                 | Any          |
| `.doc`          | Word COM API → `markitdown`  | Windows only |

> Legacy `.doc` files require Windows with Microsoft Word installed. All other formats work on any platform.

## Setup

Requires Python 3.11+.

```bash
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install .
```

For legacy `.doc` support (Windows only):

```bash
pip install .[doc]
```

## Configuration

Copy the example config and edit it:

```bash
cp config.py.example config.py
```

Default input/output folders are defined in `config.py`. CLI arguments override them.

## Usage

```bash
python main.py [source_folder] [output_folder]
```

Both arguments are optional and fall back to `config.py` defaults.

**Options:**

- `-r`, `--recursive` — scan subdirectories for documents

**Examples:**

```bash
# Use defaults from config.py
python main.py

# Specific folders
python main.py /path/to/documents /path/to/output

# Include subdirectories
python main.py -r /path/to/documents
```

Each file is saved as `<original_name>.md`. Errors are reported per-file without stopping the batch.

## License

MIT
