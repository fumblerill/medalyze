import pandas as pd
import io
from pathlib import Path

def parse_xlsx(file_bytes: bytes) -> str:
    df = pd.read_excel(io.BytesIO(file_bytes), skiprows=1)
    return df.to_html(
        table_id="mainTable",
        classes="table table-striped table-bordered",
        index=False,
        escape=False
    )

def load_error_reference() -> str | None:
    path = Path(__file__).resolve().parent.parent.parent / "error_reference.xlsx"
    print(f"[DEBUG] Путь к справочнику: {path}")  

    if not path.exists():
        print("[DEBUG] Справочник не найден")
        return None
    
    try:
        df = pd.read_excel(path)
        return df.to_html(classes="table table-sm table-striped", index=False, escape=False)
    except Exception as e:
        print(f"[ERROR] Не удалось загрузить справочник ошибок: {e}")
        return None