import io
import os
import platform
import shutil
import subprocess
import tempfile
from datetime import datetime
from calendar import monthrange

import fitz  # PyMuPDF
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from openpyxl import load_workbook

# ---------- Config ----------
EXCEL_TEMPLATE = os.getenv("EXCEL_TEMPLATE", "Invoice.xlsx")
SIGNATURE_IMAGE = os.getenv("SIGNATURE_IMAGE", "ManishaKhoriaSignature.png")
# Signature placement (points)
SIG_X, SIG_Y, SIG_W, SIG_H = 620, 370, 100, 100
ALLOWED_ORIGINS = os.getenv("ALLOWED_ORIGINS", "*").split(",")

app = FastAPI(title="Invoice Service")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[o.strip() for o in ALLOWED_ORIGINS],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def update_excel_inplace(path: str) -> None:
    wb = load_workbook(path)
    ws = wb.active

    today = datetime.today()
    current_month = today.month
    current_year = today.year
    year_two_digits = str(current_year)[-2:]

    if current_month >= 4:
        fy_month_number = current_month - 3
        fy_start_year = current_year
        fy_end_year = current_year + 1
    else:
        fy_month_number = current_month + 9
        fy_start_year = current_year - 1
        fy_end_year = current_year

    fy_start_short = str(fy_start_year)[-2:]
    fy_end_short = str(fy_end_year)[-2:]

    ws["A9"] = f"Invoice Number- {fy_month_number}/BIG/{fy_start_short}-{fy_end_short}"
    ws["J9"] = today.replace(day=1).strftime("%d/%m/%Y")
    month_shorthand = today.strftime("%b")
    ws["A22"] = f"Rent for the month of {month_shorthand},{year_two_digits}"

    last_day = monthrange(current_year, current_month)[1]
    first_day_str = f"01/{current_month:02d}/{current_year}"
    last_day_str = f"{last_day:02d}/{current_month:02d}/{current_year}"
    ws["A23"] = f"({first_day_str} - {last_day_str})"

    wb.save(path)

def convert_xlsx_to_pdf(input_xlsx: str, out_dir: str) -> str:
    # Ensure LibreOffice path (Dockerfile puts 'soffice' on PATH)
    if platform.system() == "Darwin":
        soffice = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    else:
        soffice = "soffice"

    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", input_xlsx, "--outdir", out_dir],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except subprocess.CalledProcessError as e:
        raise HTTPException(status_code=500, detail=f"LibreOffice conversion failed: {e.stderr.decode(errors='ignore')}")

    pdf_path = os.path.join(
        out_dir, os.path.splitext(os.path.basename(input_xlsx))[0] + ".pdf"
    )
    if not os.path.exists(pdf_path):
        raise HTTPException(status_code=500, detail="PDF not generated.")
    return pdf_path

def add_signature_bytes(pdf_path: str, signature_path: str) -> bytes:
    if not os.path.exists(signature_path):
        raise HTTPException(status_code=500, detail="Signature image not found on server.")

    doc = fitz.open(pdf_path)
    try:
        page = doc[0]
        rect = fitz.Rect(SIG_X, SIG_Y, SIG_X + SIG_W, SIG_Y + SIG_H)
        page.insert_image(rect, filename=signature_path)
        return doc.tobytes()  # in-memory bytes
    finally:
        doc.close()

@app.get("/download-latest-invoice")
def download_latest_invoice():
    # Validate template inputs exist
    if not os.path.exists(EXCEL_TEMPLATE):
        raise HTTPException(status_code=500, detail="Invoice.xlsx not found on server.")
    if not os.path.exists(SIGNATURE_IMAGE):
        raise HTTPException(status_code=500, detail="Signature image not found on server.")

    # Work in a temp directory to stay stateless
    with tempfile.TemporaryDirectory() as tmpdir:
        # Copy template Excel to temp working file
        working_xlsx = os.path.join(tmpdir, "Invoice.xlsx")
        shutil.copy(EXCEL_TEMPLATE, working_xlsx)

        # 1) Update Excel
        update_excel_inplace(working_xlsx)

        # 2) Convert to PDF in temp
        pdf_path = convert_xlsx_to_pdf(working_xlsx, tmpdir)

        # 3) Add image; get final bytes (no disk writes persisted)
        final_pdf_bytes = add_signature_bytes(pdf_path, SIGNATURE_IMAGE)

    # 4) Stream to client with download filename
    buf = io.BytesIO(final_pdf_bytes)
    headers = {"Content-Disposition": 'attachment; filename="Invoice_Signed.pdf"'}
    return StreamingResponse(buf, media_type="application/pdf", headers=headers)

@app.get("/ping")
def ping():
    return {"status": "Service is up"}
