import pandas as pd
import io

def parse_xlsx(file_bytes: bytes) -> str:
    df = pd.read_excel(io.BytesIO(file_bytes), skiprows=1)
    return df.to_html(classes="table table-bordered table-striped", index=False, escape=False)