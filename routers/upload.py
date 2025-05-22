from fastapi import APIRouter, Request, UploadFile, File
from fastapi.responses import HTMLResponse
from app.services.xlsx_parser import parse_xlsx, load_error_reference

router = APIRouter()

@router.get("/", response_class=HTMLResponse)
async def index(request: Request):
        return request.app.templates.TemplateResponse("upload.html", {"request": request})

@router.post("/upload", response_class=HTMLResponse)
async def upload(request: Request, file: UploadFile = File(...)):
        if not file.filename.endswith(".xlsx"):
            return request.app.templates.TemplateResponse("upload.html", {
                   "request": request,
                   "error": "❗ Поддерживаются только .xlsx файлы"
            })
        
        contents = await file.read()
        table_raw = parse_xlsx(contents)

        table_errors_dict = load_error_reference()

        return request.app.templates.TemplateResponse("upload.html", {
                "request": request,
                "filename": file.filename,
                "table_raw": table_raw,
                "table_errors_dict": table_errors_dict
        })