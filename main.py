import argparse
import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from pypdf import PdfReader

try:
    from pdf2image import convert_from_path
    import pytesseract
except Exception:
    convert_from_path = None
    pytesseract = None

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

TEXT_EXTRACTION_MIN_CHARS = 300
DEFAULT_CONFIDENCE_THRESHOLD = 0.75
DEFAULT_MODEL = "gpt-4.1-mini"

@dataclass
class LineItem:
    description: str = ""
    quantity: float | None = None
    unit_price: float | None = None
    amount: float | None = None

@dataclass
class InvoiceRecord:
    file_name: str
    vendor_name: str = ""
    invoice_number: str = ""
    invoice_date: str = ""
    subtotal: float | None = None
    tax: float | None = None
    total: float | None = None
    overall_confidence: float = 0.0
    low_confidence: bool = True
    extraction_method: str = ""
    status: str = "ok"
    error_message: str = ""
    line_items: list[LineItem] | None = None

def extract_text_from_pdf(pdf_path: Path) -> tuple[str, str]:
    text_parts: list[str] = []
    try:
        reader = PdfReader(str(pdf_path))
        for page in reader.pages:
            text = page.extract_text() or ""
            if text.strip():
                text_parts.append(text)
    except Exception:
        pass

    extracted_text = "\n".join(text_parts).strip()
    if len(extracted_text) >= TEXT_EXTRACTION_MIN_CHARS:
        return extracted_text, "pdf_text"

    if convert_from_path is None or pytesseract is None:
        return extracted_text, "pdf_text_partial"

    try:
        images = convert_from_path(str(pdf_path), dpi=250)
        ocr_parts: list[str] = []
        for image in images:
            ocr_text = pytesseract.image_to_string(image)
            if ocr_text.strip():
                ocr_parts.append(ocr_text)
        ocr_text = "\n".join(ocr_parts).strip()
        if len(ocr_text) > len(extracted_text):
            return ocr_text, "ocr"
    except Exception:
        pass

    return extracted_text, "pdf_text_partial"

def build_extraction_prompt(raw_text: str, file_name: str) -> list[dict[str, str]]:
    system = (
        "You are an invoice extraction engine. Extract fields from invoice text with high accuracy. "
        "Return only valid JSON matching the requested schema. If a field is missing, use null or an empty string. "
        "Do not invent values. If uncertain, lower confidence scores."
    )
    user = f"""
Extract the following from this invoice document ({file_name}):
- vendor_name
- invoice_number
- invoice_date (ISO 8601 preferred, else original string)
- line_items (array of objects with description, quantity, unit_price, amount)
- subtotal
- tax
- total
- confidence scores for each major field and an overall confidence from 0 to 1

Also return a low_confidence boolean. Set low_confidence to true if overall confidence is below 0.75 or if key fields are missing.

Invoice text:
---
{raw_text}
---

Return JSON with this shape:
{{
  "vendor_name": "",
  "invoice_number": "",
  "invoice_date": "",
  "line_items": [
    {{"description": "", "quantity": 1, "unit_price": 10.0, "amount": 10.0}}
  ],
  "subtotal": 0.0,
  "tax": 0.0,
  "total": 0.0,
  "field_confidence": {{
    "vendor_name": 0.0,
    "invoice_number": 0.0,
    "invoice_date": 0.0,
    "line_items": 0.0,
    "subtotal": 0.0,
    "tax": 0.0,
    "total": 0.0
  }},
  "overall_confidence": 0.0,
  "low_confidence": true
}}
"""
    return [{"role": "system", "content": system}, {"role": "user", "content": user}]

def llm_extract_invoice(raw_text: str, file_name: str, model: str = DEFAULT_MODEL) -> dict[str, Any]:
    if OpenAI is None:
        raise RuntimeError("openai package is not installed")
    client = OpenAI()
    messages = build_extraction_prompt(raw_text, file_name)
    response = client.chat.completions.create(
        model=model,
        messages=messages,
        temperature=0,
        response_format={"type": "json_object"},
    )
    content = response.choices[0].message.content or "{}"
    return json.loads(content)

def normalize_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    cleaned = re.sub(r"[^0-9.\-]", "", text.replace(",", ""))
    if cleaned in {"", ".", "-", "-."}:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None

def parse_line_items(items: Any) -> list[LineItem]:
    parsed: list[LineItem] = []
    if not isinstance(items, list):
        return parsed
    for item in items:
        if not isinstance(item, dict):
            continue
        parsed.append(LineItem(description=str(item.get("description", "") or ""), quantity=normalize_number(item.get("quantity")), unit_price=normalize_number(item.get("unit_price")), amount=normalize_number(item.get("amount"))))
    return parsed

def process_invoice(pdf_path: Path, model: str, confidence_threshold: float) -> InvoiceRecord:
    record = InvoiceRecord(file_name=pdf_path.name, line_items=[])
    try:
        raw_text, method = extract_text_from_pdf(pdf_path)
        record.extraction_method = method
        if not raw_text.strip():
            record.status = "no_text"
            record.error_message = "No extractable text found in PDF"
            return record
        extracted = llm_extract_invoice(raw_text, pdf_path.name, model=model)
        record.vendor_name = str(extracted.get("vendor_name", "") or "")
        record.invoice_number = str(extracted.get("invoice_number", "") or "")
        record.invoice_date = str(extracted.get("invoice_date", "") or "")
        record.subtotal = normalize_number(extracted.get("subtotal"))
        record.tax = normalize_number(extracted.get("tax"))
        record.total = normalize_number(extracted.get("total"))
        record.overall_confidence = float(extracted.get("overall_confidence") or 0.0)
        record.low_confidence = bool(extracted.get("low_confidence", record.overall_confidence < confidence_threshold))
        record.line_items = parse_line_items(extracted.get("line_items", []))
        if record.overall_confidence < confidence_threshold:
            record.low_confidence = True
            record.status = "low_confidence"
        return record
    except Exception as exc:
        record.status = "error"
        record.error_message = str(exc)
        return record

def write_excel(records: list[InvoiceRecord], output_path: Path) -> None:
    summary_rows: list[dict[str, Any]] = []
    line_item_rows: list[dict[str, Any]] = []
    for record in records:
        summary_rows.append({"file_name": record.file_name, "vendor_name": record.vendor_name, "invoice_number": record.invoice_number, "invoice_date": record.invoice_date, "subtotal": record.subtotal, "tax": record.tax, "total": record.total, "overall_confidence": record.overall_confidence, "low_confidence": record.low_confidence, "extraction_method": record.extraction_method, "status": record.status, "error_message": record.error_message})
        for idx, item in enumerate(record.line_items or [], start=1):
            line_item_rows.append({"file_name": record.file_name, "line_item_no": idx, "vendor_name": record.vendor_name, "invoice_number": record.invoice_number, "description": item.description, "quantity": item.quantity, "unit_price": item.unit_price, "amount": item.amount})
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(summary_rows).to_excel(writer, index=False, sheet_name="Invoices")
        pd.DataFrame(line_item_rows).to_excel(writer, index=False, sheet_name="Line Items")
        workbook = writer.book
        for sheet_name in ["Invoices", "Line Items"]:
            sheet = workbook[sheet_name]
            for cell in sheet[1]:
                cell.font = Font(bold=True)
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions

def main() -> int:
    parser = argparse.ArgumentParser(description="Process invoice PDFs into a structured Excel workbook.")
    parser.add_argument("input_folder", help="Folder containing PDF invoices")
    parser.add_argument("--output", default="invoice_output.xlsx", help="Output Excel file path")
    parser.add_argument("--model", default=os.getenv("OPENAI_MODEL", DEFAULT_MODEL), help="OpenAI model")
    parser.add_argument("--confidence-threshold", type=float, default=DEFAULT_CONFIDENCE_THRESHOLD, help="Confidence threshold below which invoices are flagged")
    args = parser.parse_args()

    input_folder = Path(args.input_folder)
    if not input_folder.exists() or not input_folder.is_dir():
        raise SystemExit(f"Input folder not found: {input_folder}")

    pdf_files = sorted(input_folder.glob("*.pdf"))
    if not pdf_files:
        raise SystemExit("No PDF files found in the input folder.")

    records = [process_invoice(pdf, args.model, args.confidence_threshold) for pdf in pdf_files]
    write_excel(records, Path(args.output))
    print(f"Processed {len(records)} invoice(s) -> {args.output}")
    flagged = sum(1 for r in records if r.low_confidence or r.status != "ok")
    print(f"Flagged {flagged} invoice(s) for review")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
