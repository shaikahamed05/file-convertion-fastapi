# uvicorn api:app --reload --host 192.168.5.11 --port 8000

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, RedirectResponse
import tempfile
import os
import shutil

# Import your conversion modules
# # PDF modules
from modules.pdf.pdf_to_html import convert_pdf_to_html
from modules.pdf.pdf_to_word import convert_pdf_to_word
from modules.pdf.pdf_to_excel import convert_pdf_to_xlsx

# Word modules
from modules.word.word_to_pdf import convert_word_to_pdf
from modules.word.word_to_excel import convert_word_to_excel
from modules.word.word_to_html import convert_word_to_html

# Excel modules
from modules.excel.excel_to_pdf import excel_to_pdf_converter
from modules.excel.excel_to_html import excel_to_html_enhanced
from modules.excel.excel_to_word import excel_to_word_converter

# HTML modules
from modules.html.html_to_excel import convert_html_to_excel
from modules.html.html_to_word import convert_html_to_word
from modules.html.html_to_pdf import html_to_pdf_converter

app = FastAPI(title="File Conversion API", version="1.0.0")


@app.get("/", include_in_schema=False)
def root():
    return RedirectResponse(url="/docs")

#main convertions
@app.post("/convert/main_html-to-excel", tags=["Main"])
async def main_html_to_excel(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp_html:
        tmp_html.write(await file.read())
        tmp_html_path = tmp_html.name
    original_filename = os.path.splitext(file.filename)[0]
    output_path = tmp_html_path.replace(".html", ".xlsx")
    convert_html_to_excel(tmp_html_path, output_path,columns_to_merge=None)

    return FileResponse(path=output_path, filename=f"{original_filename}.xlsx", media_type="application/pdf")

@app.post("/convert/main_excel-to-pdf", tags=["Main"])
async def main_excel_to_pdf(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
        tmp_xlsx.write(await file.read())
        tmp_xlsx_path = tmp_xlsx.name
    original_filename = os.path.splitext(file.filename)[0]
    output_path = tmp_xlsx_path.replace(".xlsx", ".pdf")
    excel_to_pdf_converter(tmp_xlsx_path, output_path)

    return FileResponse(path=output_path, filename=f"{original_filename}.pdf", media_type="application/pdf")


# PDF
@app.post("/convert/pdf-to-word", tags=["PDF"])
async def pdf_to_word(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        tmp_pdf.write(await file.read())
        tmp_pdf_path = tmp_pdf.name
    original_filename = os.path.splitext(file.filename)[0]
    output_path = tmp_pdf_path.replace(".pdf", ".docx")
    convert_pdf_to_word(tmp_pdf_path, output_path)

    return FileResponse(path=output_path, filename=f"{original_filename}.docx", media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

@app.post("/convert/pdf-to-excel", tags=["PDF"])
async def pdf_to_excel(file: UploadFile = File(...)):
    # Save uploaded PDF to temp
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        tmp_pdf.write(await file.read())
        tmp_pdf_path = tmp_pdf.name
    original_filename = os.path.splitext(file.filename)[0]
    # Convert and get the actual output path
    output_path = convert_pdf_to_xlsx(tmp_pdf_path)

    # Ensure file exists
    if not os.path.exists(output_path):
        raise HTTPException(status_code=500, detail="Conversion failed: output file not found.")

    return FileResponse(
        path=output_path,
        filename=f"{original_filename}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.post("/convert/pdf-to-html", tags=["PDF"])
async def pdf_to_html(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        tmp_pdf.write(await file.read())
        tmp_pdf_path = tmp_pdf.name
    original_filename = os.path.splitext(file.filename)[0]
    output_path = tmp_pdf_path.replace(".pdf", ".html")
    convert_pdf_to_html(tmp_pdf_path, output_path)

    return FileResponse(path=output_path, filename=f"{original_filename}.html", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Excel
@app.post("/convert/excel-to-pdf", tags=["Excel"])
async def excel_to_pdf(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
        tmp_xlsx.write(await file.read())
        tmp_xlsx_path = tmp_xlsx.name
    original_filename = os.path.splitext(file.filename)[0]
    output_path = tmp_xlsx_path.replace(".xlsx", ".pdf")
    excel_to_pdf_converter(tmp_xlsx_path, output_path)

    return FileResponse(path=output_path, filename=f"{original_filename}.pdf", media_type="application/pdf")

@app.post("/convert/excel-to-word", tags=["Excel"])
async def excel_to_word(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
        tmp_xlsx.write(await file.read())
        tmp_xlsx_path = tmp_xlsx.name

    tmp_pdf_path = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
    tmp_docx_path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
    original_filename = os.path.splitext(file.filename)[0]
    # Step 1: Excel → PDF
    excel_to_word_converter(tmp_xlsx_path, tmp_pdf_path)

    if not os.path.exists(tmp_pdf_path):
        raise HTTPException(status_code=500, detail="Failed to export Excel to PDF.")

    # Step 2: PDF → Word
    convert_pdf_to_word(tmp_pdf_path, tmp_docx_path)

    if not os.path.exists(tmp_docx_path):
        raise HTTPException(status_code=500, detail="Failed to convert PDF to Word.")

    # Step 3: Return the DOCX file
    return FileResponse(
        path=tmp_docx_path,
        filename=f"{original_filename}.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.post("/convert/excel-to-html", tags=["Excel"])
async def excel_to_html(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
        tmp_xlsx.write(await file.read())
        tmp_xlsx_path = tmp_xlsx.name
    original_filename = os.path.splitext(file.filename)[0]
    output_path = tmp_xlsx_path.replace(".xlsx", ".html")
    excel_to_html_enhanced(tmp_xlsx_path, output_path)

    return FileResponse(path=output_path, filename=f"{original_filename}.html", media_type="application/pdf")


# HTML
@app.post("/convert/html-to-pdf", tags=["HTML"])
async def html_to_pdf(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp_html:
        tmp_html.write(await file.read())
        tmp_html_path = tmp_html.name
    original_filename = os.path.splitext(file.filename)[0]
    output_path = tmp_html_path.replace(".html", ".pdf")
    html_to_pdf_converter(tmp_html_path, output_path)

    return FileResponse(path=output_path, filename=f"{original_filename}.pdf", media_type="application/pdf")


@app.post("/convert/html-to-word", tags=["HTML"])
async def html_to_word(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp_html:
        tmp_html.write(await file.read())
        tmp_html_path = tmp_html.name
    original_filename = os.path.splitext(file.filename)[0]
    output_path = tmp_html_path.replace(".html", ".docx")
    convert_html_to_word(tmp_html_path, output_path)

    return FileResponse(path=output_path, filename=f"{original_filename}.docx", media_type="application/pdf")


@app.post("/convert/html-to-excel", tags=["HTML"])
async def html_to_excel(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp_html:
        tmp_html.write(await file.read())
        tmp_html_path = tmp_html.name
    original_filename = os.path.splitext(file.filename)[0]
    output_path = tmp_html_path.replace(".html", ".xlsx")
    convert_html_to_excel(tmp_html_path, output_path)

    return FileResponse(path=output_path, filename=f"{original_filename}.xlsx", media_type="application/pdf")


# Word
@app.post("/convert/word-to-pdf", tags=["Word"])
async def word_to_pdf(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
        tmp_docx.write(await file.read())
        tmp_docx_path = tmp_docx.name
    original_filename = os.path.splitext(file.filename)[0]
    output_path = tmp_docx_path.replace(".docx", ".pdf")
    convert_word_to_pdf(tmp_docx_path, output_path)

    return FileResponse(path=output_path, filename=f"{original_filename}.pdf", media_type="application/pdf")

@app.post("/convert/word-to-excel", tags=["Word"])
async def word_to_excel(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
        tmp_docx.write(await file.read())
        tmp_docx_path = tmp_docx.name
    original_filename = os.path.splitext(file.filename)[0]
    output_path = tmp_docx_path.replace(".docx", ".xlsx")
    convert_word_to_excel(tmp_docx_path, output_path)

    return FileResponse(path=output_path, filename=f"{original_filename}.xlsx", media_type="application/pdf")

@app.post("/convert/word-to-html", tags=["Word"])
async def word_to_html(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
        tmp_docx.write(await file.read())
        tmp_docx_path = tmp_docx.name
    original_filename = os.path.splitext(file.filename)[0]
    output_path = tmp_docx_path.replace(".docx", ".html")
    convert_word_to_html(tmp_docx_path, output_path)

    return FileResponse(path=output_path, filename=f"{original_filename}.html", media_type="application/pdf")



