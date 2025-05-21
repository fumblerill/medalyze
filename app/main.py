from nicegui import ui
from io import BytesIO
import pandas as pd

DARK_BG = "#171b2b"
SIDEBAR_BG = "#232949"
SIDEBAR_BTN = "#5c7cff"
SIDEBAR_WIDTH = "240px"
CARD_BG = "#23263a"

ui.add_body_html(f"""
<style>
body, #root {{
    background: {DARK_BG};
    overflow-x: hidden !important;
}}
.q-uploader {{
    min-height: 0 !important;
    padding: 0 !important;
    margin: 0 !important;
    box-shadow: none !important;
}}
.q-uploader__header {{
    margin-bottom: 0 !important;
    padding-bottom: 0 !important;
}}
.q-uploader__list {{
    min-height: 0 !important;
    margin-top: 0 !important;
    margin-bottom: 0 !important;
}}
.q-table__middle, .q-table__container {{
    max-height: 210px !important;
    overflow-y: auto !important;
}}
</style>
""")

app_state = {'df': None}

def render_content():
    content_area.clear()
    df = app_state['df']

    if df is None:
        with content_area:
            ui.label('Загрузите Excel-файл, чтобы увидеть данные').style('color: #ccc; font-size: 1.05rem;')
        return

    df_display = df.head(5).copy()

    for col in df_display.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]']):
        df_display[col] = df_display[col].astype(str)

    # Жестко задаём ширину таблице
    table_min_width = f"{150 * len(df_display.columns)}px"

    with content_area:
        ui.label('Первые 5 строк файла:').style('color: #ccc; margin-bottom: 12px; font-size: 1.1rem;')

        # Ограничение ширины контентной зоны
        with ui.element('div').style('max-width: 1300px; width: 100%; margin: 0 auto;'):
            # Реальный скролл
            with ui.element('div').style('overflow-x: auto; width: 100%;'):
                ui.html('<div class="q-pa-none">')  # обёртка для контроля padding у q-table
                ui.table(
                    columns=[{'name': col, 'label': col, 'field': col} for col in df_display.columns],
                    rows=df_display.to_dict('records'),
                ).style(f'''
                    min-width: {table_min_width};
                    width: auto;
                    background: #23263a;
                    color: #fff;
                    border-radius: 8px;
                ''')
                ui.html('</div>')

with ui.row().style(f'width: 100%; min-height: 100vh; background: {DARK_BG}; margin: 0; overflow-x: hidden;'):
    # Sidebar
    with ui.column().style(
        f'''
        width: {SIDEBAR_WIDTH}; min-width: {SIDEBAR_WIDTH};
        min-height: 100vh; background: {SIDEBAR_BG}; 
        padding: 16px 0 0 0;
        position: sticky; left: 0; top: 0; z-index: 10;
        display: flex; flex-direction: column; align-items: center;
        '''
    ):
        ui.label('Remora.Analyze').style('font-size: 1.7rem; color: #89a3ff; font-weight: bold; margin-bottom: 10px; letter-spacing:1px;')
        
        def on_upload(e):
            file_bytes = e.content.read()
            try:
                df = pd.read_excel(BytesIO(file_bytes), skiprows=1)
                app_state['df'] = df
                ui.notify('Файл успешно загружен!', type='positive')
                render_content()
            except Exception as ex:
                ui.notify(f'Ошибка загрузки: {ex}', type='negative')

        # 👇 ВОТ ОН — uploader, которого у тебя не было
        ui.upload(on_upload=on_upload, auto_upload=True, label='Загрузить Excel (.xlsx)').style('width: 90%; margin-top: 8px;')

    # Main area
    with ui.column().style(
        f'''
        flex: 1; min-height: 100vh; background: {DARK_BG};
        margin: 0; padding: 24px;
        display: flex; flex-direction: column; justify-content: flex-start;
        '''
    ):
        content_area = ui.column().style(
            f'''
            flex: 1;
            width: 100%;
            background: {CARD_BG};
            border-radius: 14px;
            box-shadow: none;
            padding: 16px;
            min-height: 250px;
            display: flex;
            flex-direction: column;
            justify-content: flex-start;
            align-items: flex-start;
            '''
        )

render_content()
ui.run(title='Remora.Analyze', port=8080)