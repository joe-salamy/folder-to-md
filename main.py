import argparse
import os
import tempfile
import pymupdf4llm
from markitdown import MarkItDown


def main():
    parser = argparse.ArgumentParser(
        description="Batch convert documents (PDF, DOCX, DOC, PPT, PPTX, EPUB) to Markdown."
    )
    parser.add_argument("source", help="Folder containing documents to convert")
    parser.add_argument(
        "output",
        nargs="?",
        help="Folder to save Markdown files (defaults to source folder)",
    )
    args = parser.parse_args()

    doc_folder = args.source
    md_folder = args.output if args.output else args.source

    if not os.path.exists(doc_folder):
        raise FileNotFoundError(f"Document folder not found: {doc_folder}")

    if not os.path.exists(md_folder):
        os.makedirs(md_folder)

    md_converter = MarkItDown()

    doc_files = [
        f
        for f in os.listdir(doc_folder)
        if f.lower().endswith((".pdf", ".docx", ".doc", ".ppt", ".pptx", ".epub"))
    ]

    if not doc_files:
        print("No documents found in the folder.")
        return

    # Open Word once for all .doc files if any exist
    word = None
    if any(f.lower().endswith(".doc") for f in doc_files):
        import win32com.client
        import pythoncom

        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

    try:
        for doc_file in doc_files:
            doc_path = os.path.join(doc_folder, doc_file)
            output_name = os.path.splitext(doc_file)[0] + ".md"
            output_path = os.path.join(md_folder, output_name)

            try:
                print(f"Converting: {doc_file} ...")

                if doc_file.lower().endswith(".pdf"):
                    md_text = pymupdf4llm.to_markdown(doc_path)
                elif doc_file.lower().endswith(".doc"):
                    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".docx")
                    os.close(tmp_fd)
                    try:
                        doc = word.Documents.Open(os.path.abspath(doc_path))
                        doc.SaveAs2(tmp_path, FileFormat=16)  # 16 = wdFormatXMLDocument
                        doc.Close()
                        result = md_converter.convert(tmp_path)
                        md_text = result.text_content
                    finally:
                        if os.path.exists(tmp_path):
                            os.unlink(tmp_path)
                else:
                    result = md_converter.convert(doc_path)
                    md_text = result.text_content

                with open(output_path, "w", encoding="utf-8") as md_file:
                    md_file.write(md_text)

                print(f"Saved: {output_path}")
            except Exception as e:
                print(f"Error converting {doc_file}: {e}")
    finally:
        if word is not None:
            word.Quit()
            pythoncom.CoUninitialize()

    print("Done.")


if __name__ == "__main__":
    main()
