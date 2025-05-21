from fastapi import APIRouter, Request, UploadFile, File
from fastapi.responses import HTMLResponse
from app.services.xlsx_parser import parse_xlsx

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
        table_html = parse_xlsx(contents)

        return request.app.templates.TemplateResponse("upload.html", {
                "request": request,
                "table_html": table_html,
                "filename": file.filename
        })