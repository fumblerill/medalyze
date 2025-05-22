from fastapi import APIRouter, Request, UploadFile, File
from fastapi.responses import HTMLResponse
from services.xlsx_handler import generate_tables

router = APIRouter()


@router.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return request.app.templates.TemplateResponse("upload.html", {
        "request": request
    })

@router.post("/upload", response_class=HTMLResponse)
async def upload(request: Request, file: UploadFile = File(...)):
        if not file.filename.endswith(".xlsx"):
            return request.app.templates.TemplateResponse("upload.html", {
                   "request": request,
                   "error": "❗ Поддерживаются только .xlsx файлы"
            })
        
        contents = await file.read()
        tables_data = generate_tables(contents)

        return request.app.templates.TemplateResponse("upload.html", {
                "request": request,
                "filename": file.filename,
                "tables": tables_data["tables"],
                "pie_labels": tables_data["pie_labels"],
                "pie_values": tables_data["pie_values"]
        })