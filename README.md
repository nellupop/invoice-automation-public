# Python Invoice Processing Automation

This project processes a folder of PDF invoices, extracts structured fields with an LLM, falls back to OCR for scanned documents, and writes the results to Excel.

## Features

- Input: folder of PDF invoices in mixed formats
- Extraction:
  - vendor name
  - invoice number
  - date
  - line items
  - subtotal
  - tax
  - total
- LLM-based extraction for accuracy across invoice layouts
- OCR fallback for scanned PDFs
- Output to a structured Excel workbook using pandas/openpyxl
- Low-confidence flagging for manual review

## Files

- `main.py` — processing script
- `requirements.txt` — Python dependencies
- `sample_output.xlsx` — example output workbook

## Prerequisites

Install the Python packages in `requirements.txt`.

For OCR fallback, you also need system tools:

- Tesseract OCR
- Poppler (required by `pdf2image`)

### Example installation

macOS:

```bash
brew install tesseract poppler
```

Ubuntu / Debian:

```bash
sudo apt-get update
sudo apt-get install -y tesseract-ocr poppler-utils
```

## Environment variables

Set your OpenAI API key:

```bash
export OPENAI_API_KEY="your_api_key_here"
```

Optional:

```bash
export OPENAI_MODEL="gpt-4.1-mini"
```

## Usage

Place your invoice PDFs in a folder, then run:

```bash
python main.py /path/to/invoice_folder --output invoices.xlsx
```

If you want to change the model or confidence threshold:

```bash
python main.py /path/to/invoice_folder --output invoices.xlsx --model gpt-4.1-mini --confidence-threshold 0.75
```

## Output

The script creates an Excel workbook with two sheets:

- `Invoices` — one row per invoice with summary fields and confidence metadata
- `Line Items` — exploded line-item rows for analysis

## Confidence handling

Invoices are flagged when:

- the model returns a low confidence score
- required fields are missing
- OCR/text extraction fails
- an exception occurs during processing

## Notes

- The script uses PDF text extraction first.
- If the extracted text is too short, OCR is attempted.
- The LLM is instructed to return JSON only, which makes the workflow easier to validate and store.
