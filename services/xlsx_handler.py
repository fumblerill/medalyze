import pandas as pd
import io
from pathlib import Path

from settings import BASE_DIR, ERROR_REF_PATH

def load_error_reference(table_id: str = "errorsDictTable") -> str | None:
    if not ERROR_REF_PATH.exists():
        print("[DEBUG] Справочник не найден:", ERROR_REF_PATH)
        return None
    try:
        df = pd.read_excel(ERROR_REF_PATH)
        return df.to_html(
            table_id=table_id,
            classes="table table-sm table-striped",
            index=False,
            escape=False 
        )
    except Exception as e:
        print(f"[ERROR] Ошибка загрузки справочника: {e}")
        return None
    
def load_error_reference_dataframe() -> pd.DataFrame:
    if not ERROR_REF_PATH.exists():
        print("[DEBUG] Справочник не найден:", ERROR_REF_PATH)
        return pd.DataFrame()
    try:
        return pd.read_excel(ERROR_REF_PATH)
    except Exception as e:
        print(f"[ERROR] Ошибка загрузки справочника: {e}")
        return pd.DataFrame()
    
def render_df(df: pd.DataFrame, table_id: str) -> str:
    return df.to_html(
        table_id=table_id,
        classes="table table-sm table-striped table-bordered",
        index=False,
        escape=False,
        border=0
    ).replace("<table ", '<table style="width:100%" ')

# Функция обработки основного датафрейма по описанию ошибок из справочника
def process_main_file(df_main, df_errors):
    df_main['Код ошибки'] = df_main['Описание'].str.extract(r"Code:\s*(\w+)", expand=False)

    df_main['Код ошибки'] = df_main['Код ошибки'].fillna(
        df_main['Описание'].str.extract(r"Ошибка:\s*([\w\s]+?)(?:[:\[]|$)", expand=False)
    )

    df_main['Код ошибки'] = df_main['Код ошибки'].fillna(
        df_main['Описание'].str.extract(r"(CommunicationException|FaultException):", expand=False)
    )

    df_main['Код ошибки'] = df_main['Код ошибки'].str.strip()

    df_result = pd.merge(
        df_main,
        df_errors[['Код ответа', 'Описание']].rename(columns={'Описание': 'Описание ошибки'}),
        how='left',
        left_on='Код ошибки',
        right_on='Код ответа'
    )

    df_result['Описание'] = df_result.apply(
        lambda row: "Отсутствует в справочнике" if pd.notna(row['Код ошибки']) and pd.isna(row['Код ответа'])
        else row['Описание ошибки'] if pd.notna(row['Код ответа'])
        else row['Описание'],
        axis=1
    )

    df_result.drop(columns=['Код ответа', 'Описание ошибки'], inplace=True, errors='ignore')

    return df_result

# Функция вывода статистики по СЭМД в РЭМД
def create_summary_by_department(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    for col in df.select_dtypes(include=['category']).columns:
        df[col] = df[col].astype(str)

    df['Сотрудник, сформировавший СЭМД'] = df['Сотрудник, сформировавший СЭМД'].astype(str)

    df = df[
        df['Сотрудник, сформировавший СЭМД'].str.strip().notnull() &
        (df['Сотрудник, сформировавший СЭМД'] != 'nan') &
        (df['Сотрудник, сформировавший СЭМД'] != '')
    ]

    unique_pairs = df[['Сотрудник, сформировавший СЭМД', 'Отделение МО']].drop_duplicates()

    pivot = (
        df.groupby(['Отделение МО', 'Сотрудник, сформировавший СЭМД', 'Статус передачи СЭМД'])
        .size()
        .unstack(fill_value=0)
        .reset_index()
    )

    required_cols = [
        'Зарегистрирован в Нетрике',
        'Зарегистрирован в РЭМД',
        'Отказано в регистрации в Нетрике',
        'Отказано в регистрации в РЭМД'
    ]

    for col in required_cols:
        if col not in pivot.columns:
            pivot[col] = 0

    pivot['СЭМД успешно зарегистрированных в РЭМД'] = (
        pivot['Зарегистрирован в Нетрике'] + pivot['Зарегистрирован в РЭМД']
    )
    pivot['СЭМД отказано в регистрации в РЭМД'] = (
        pivot['Отказано в регистрации в Нетрике'] + pivot['Отказано в регистрации в РЭМД']
    )

    pivot.drop(columns=required_cols, inplace=True)

    pivot.insert(2, 'СЭМД успешно зарегистрированных в РЭМД',
                 pivot.pop('СЭМД успешно зарегистрированных в РЭМД'))

    result = pd.merge(unique_pairs, pivot,
                      on=['Отделение МО', 'Сотрудник, сформировавший СЭМД'],
                      how='left').fillna(0)

    cols = result.columns.tolist()
    cols.insert(0, cols.pop(cols.index('Отделение МО')))
    result = result[cols]

    for col in result.columns:
        if result[col].dtype == 'float':
            result[col] = result[col].astype(int)

    return result

# Функция сводной таблицы по сотрудникам, выполнившим норму в 500 СЭМД
def create_summary_by_employee_threshold(df: pd.DataFrame, threshold=500) -> pd.DataFrame:
    successful_records = df[df['Статус передачи СЭМД'].isin(['Зарегистрирован в РЭМД', 'Зарегистрирован в Нетрике'])]

    employee_summary = successful_records['Сотрудник, сформировавший СЭМД'].value_counts().reset_index()

    employee_summary.columns = ['Сотрудник, сформировавший СЭМД', 'Количество СЭМД']

    filtered_summary = employee_summary[employee_summary['Количество СЭМД'] >= threshold]

    return filtered_summary

# Функция сводной таблицы по видам СЭМД
def create_summary_by_type(df: pd.DataFrame) -> pd.DataFrame:
    pivot_table = df.groupby(['Вид СЭМД', 'Статус передачи СЭМД']).size().unstack(fill_value=0).reset_index()

    required_columns = [
        'Зарегистрирован в Нетрике',
        'Зарегистрирован в РЭМД',
        'Отказано в регистрации в Нетрике',
        'Отказано в регистрации в РЭМД'
    ]

    for col in required_columns:
        if col not in pivot_table.columns:
            pivot_table[col] = 0

    pivot_table['СЭМД успешно зарегистрированных в РЭМД'] = pivot_table[
        ['Зарегистрирован в Нетрике', 'Зарегистрирован в РЭМД']].sum(axis=1)

    pivot_table['СЭМД отказано в регистрации в РЭМД'] = pivot_table[
        ['Отказано в регистрации в Нетрике', 'Отказано в регистрации в РЭМД']].sum(axis=1)

    pivot_table.drop(columns=['Зарегистрирован в Нетрике', 'Зарегистрирован в РЭМД', 'Отказано в регистрации в Нетрике',
                              'Отказано в регистрации в РЭМД'], inplace=True, errors='ignore')

    pivot_table.insert(1, 'СЭМД успешно зарегистрированных в РЭМД',
                       pivot_table.pop('СЭМД успешно зарегистрированных в РЭМД'))

    return pivot_table

# Функция сводной таблицы по сотрудникам, с ошибочно выбранными сертификатами
def create_summary_by_value_mismatch_threshold(df, threshold='VALUE_MISMATCH_METADATA_AND_CERTIFICATE'):
    filtered_summary = df[df['Код ошибки'] == threshold]

    employee_summary = filtered_summary['Сотрудник, сформировавший СЭМД'].value_counts()

    employee_summary = employee_summary[employee_summary > 0].reset_index()

    employee_summary.columns = ['Сотрудник, сформировавший СЭМД', 'Частота ошибки']

    return employee_summary

# Функция диаграммы по статусам
def create_status_chart_data(df: pd.DataFrame) -> tuple[list[str], list[int]]:
    labels = [
        'СЭМД успешно зарегистрированных в РЭМД',
        'Не  отправлен',
        'СЭМД отказано в регистрации в РЭМД'
    ]

    values = []
    for label in labels:
        if label in df.columns:
            values.append(int(df[label].sum()))
        else:
            values.append(0)

    return labels, values

def generate_tables(xlsx_bytes: bytes) -> dict[str, str]:
    # 0. Загрузка справочника ошибок
    df_errors = load_error_reference_dataframe()
    
    # 1. Загрузка исходного .xlsx
    df = pd.read_excel(io.BytesIO(xlsx_bytes), skiprows=1)

    # 2. Обогащение ошибок
    df_processed = process_main_file(df, df_errors)

    # 3. Удаляем строки "Всего: ЧИСЛО" из raw
    df_raw = df_processed[~df_processed.apply(lambda row: row.astype(str).str.contains(r"^Всего:\s*\d+$", regex=True)).any(axis=1)]

    # ===== 4. Статистика по СЭМД в РЭМД =====
    statistics = create_summary_by_department(df)

    # ===== 5. Врачи, сделавшие > 500 =====
    doc_perf = create_summary_by_employee_threshold(df)

    # ===== 6. Статистика по видам СЭМД =====
    df_types = create_summary_by_type(df)

    # ===== 7. Ошибки при выборе сертификата =====
    df_cert_errors = create_summary_by_value_mismatch_threshold(df)

    pie_labels, pie_values = create_status_chart_data(statistics)

    return {
        "tables": {
            "raw": render_df(df_raw, "rawTable"),
            "statistics": render_df(statistics, "statisticsTable"),
            "doc_perf_500": render_df(doc_perf, "docPerfTable"),
            "types": render_df(df_types, "typesTable"),
            "cert_errors": render_df(df_cert_errors, "certErrorsTable"),
            "reference": load_error_reference(table_id="errorsDictTable"),
        },
        "pie_labels": pie_labels,
        "pie_values": pie_values
    }