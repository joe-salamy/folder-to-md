"""Batch convert documents (PDF, DOCX, DOC, PPT, PPTX, EPUB) to Markdown."""

from __future__ import annotations

import argparse
import contextlib
import sys
import tempfile
from pathlib import Path
from typing import Any

import pymupdf4llm  # type: ignore[import-untyped]
from markitdown import MarkItDown

from config import DEFAULT_INPUT_FOLDER, DEFAULT_OUTPUT_FOLDER

SUPPORTED_EXTENSIONS: set[str] = {".pdf", ".docx", ".doc", ".ppt", ".pptx", ".epub"}
WD_FORMAT_DOCX: int = 16


def convert_pdf(doc_path: Path) -> str:
    """Convert a PDF file to Markdown via pymupdf4llm.

    Args:
        doc_path: Path to the PDF file.

    Returns:
        Markdown text extracted from the PDF.
    """
    result: str = pymupdf4llm.to_markdown(str(doc_path))
    return result


def convert_doc(doc_path: Path, word_app: Any) -> str:
    """Convert a legacy .doc file to Markdown via Word COM automation.

    Opens the .doc in Word, saves as .docx to a temp file, then converts
    the .docx with MarkItDown.

    Args:
        doc_path: Path to the .doc file.
        word_app: A running Word COM application instance.

    Returns:
        Markdown text extracted from the document.
    """
    import os

    tmp_fd, tmp_path_str = tempfile.mkstemp(suffix=".docx")
    os.close(tmp_fd)
    tmp_path = Path(tmp_path_str)

    doc = None
    try:
        doc = word_app.Documents.Open(str(doc_path.resolve()))
        doc.SaveAs2(str(tmp_path), FileFormat=WD_FORMAT_DOCX)
        doc.Close()
        doc = None
        md_converter = MarkItDown()
        result = md_converter.convert(str(tmp_path))
        return str(result.text_content)
    finally:
        if doc is not None:
            with contextlib.suppress(Exception):
                doc.Close(SaveChanges=0)
        if tmp_path.exists():
            tmp_path.unlink()


def convert_generic(doc_path: Path) -> str:
    """Convert a DOCX, PPT, PPTX, or EPUB file to Markdown via MarkItDown.

    Args:
        doc_path: Path to the document file.

    Returns:
        Markdown text extracted from the document.
    """
    md_converter = MarkItDown()
    result = md_converter.convert(str(doc_path))
    return str(result.text_content)


def get_document_files(folder: Path, *, recursive: bool = False) -> list[Path]:
    """Collect all supported document files from a folder.

    Args:
        folder: Directory to scan for documents.
        recursive: If True, scan subdirectories as well.

    Returns:
        List of paths to supported document files.
    """
    if recursive:
        return [
            f
            for f in folder.rglob("*")
            if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS
        ]
    return [f for f in folder.iterdir() if f.suffix.lower() in SUPPORTED_EXTENSIONS]


def convert_file(doc_path: Path, word_app: Any | None) -> str:
    """Route a single file to the appropriate converter.

    Args:
        doc_path: Path to the document.
        word_app: Word COM instance (required for .doc files, None otherwise).

    Returns:
        Markdown text from the converted document.

    Raises:
        RuntimeError: If a .doc file is encountered but no Word instance exists.
        ValueError: If the file extension is not supported.
    """
    suffix = doc_path.suffix.lower()

    if suffix == ".pdf":
        return convert_pdf(doc_path)
    if suffix == ".doc":
        if word_app is None:
            msg = f"Word COM instance required for .doc files: {doc_path}"
            raise RuntimeError(msg)
        return convert_doc(doc_path, word_app)
    if suffix in SUPPORTED_EXTENSIONS:
        return convert_generic(doc_path)

    msg = f"Unsupported file type: {suffix}"
    raise ValueError(msg)


def main() -> None:
    """Parse CLI arguments and batch-convert documents to Markdown."""
    parser = argparse.ArgumentParser(
        description=(
            "Batch convert documents"
            " (PDF, DOCX, DOC, PPT, PPTX, EPUB) to Markdown."
        ),
    )
    parser.add_argument(
        "source",
        nargs="?",
        help=(
            "Folder containing documents to convert"
            f" (default: {DEFAULT_INPUT_FOLDER})"
        ),
    )
    parser.add_argument(
        "output",
        nargs="?",
        help=f"Folder to save Markdown files (default: {DEFAULT_OUTPUT_FOLDER})",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="Scan subdirectories for documents",
    )
    args = parser.parse_args()

    doc_folder = Path(args.source) if args.source else Path(DEFAULT_INPUT_FOLDER)
    md_folder = Path(args.output) if args.output else Path(DEFAULT_OUTPUT_FOLDER)

    if not doc_folder.exists():
        raise FileNotFoundError(f"Document folder not found: {doc_folder}")

    md_folder.mkdir(parents=True, exist_ok=True)

    doc_files = get_document_files(doc_folder, recursive=args.recursive)

    if not doc_files:
        print("No documents found in the folder.")
        return

    # Open Word once for all .doc files if any exist
    word: Any | None = None
    has_doc_files = any(f.suffix.lower() == ".doc" for f in doc_files)

    if has_doc_files:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

    failures = 0
    try:
        for doc_path in doc_files:
            output_path = md_folder / (doc_path.stem + ".md")

            try:
                print(f"Converting: {doc_path.name} ...")
                md_text = convert_file(doc_path, word)

                output_path.write_text(md_text, encoding="utf-8")
                print(f"Saved: {output_path}")
            except Exception as e:  # noqa: BLE001
                failures += 1
                print(f"Error converting {doc_path.name}: {e}")
    finally:
        if word is not None:
            word.Quit()
            import pythoncom  # noqa: F811

            pythoncom.CoUninitialize()

    if failures:
        print(f"Done with {failures} error(s).")
        sys.exit(1)
    else:
        print("Done.")


if __name__ == "__main__":
    main()
