# extract_docx_into_docx.py

Creates a copy of a `.docx` file where all embedded OLE objects (Excel, PDF, Word, email, CSV) are replaced with their extracted text content — at the same positions in the document.

The output is a valid `.docx` file that can be opened in Microsoft Word or any compatible editor, with the embedded icons replaced by readable text paragraphs.

## Installation

pip install -r requirements_extract_docx_into_docx.txt

## Usage

python3 extract_docx_into_docx.py <input.docx> [output.docx]

- If `output.docx` is not specified, the output is saved as `<input_name>_extracted.docx`

## Supported Embedded File Types

- `.docx` (Word documents)
- `.xlsx` / `.xlsm` (Excel spreadsheets)
- `.pdf` (PDF documents)
- `.eml` (email files)
- `.msg` (Outlook messages)
- `.csv` / `.txt` (plain text files)

## Output Format

The output is a `.docx` file where each embedded object is replaced by:

1. A bold separator line: **--- Embedded File: filename.xlsx ---**
2. The extracted text content as normal paragraphs
3. A bold closing line: **--- End: filename.xlsx ---**

The rest of the document (text, tables, formatting, styles) remains unchanged. Only the embedded objects are replaced.

## Notes

- The original file is never modified — a new copy is created
- Document structure (tables, sections, headers) is preserved
- The separator lines are bold with slightly smaller font to visually distinguish them from document content
