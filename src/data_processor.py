import pandas as pd
import os


def load_and_merge_data(folder_path: str) -> pd.DataFrame:
    """Загружает все Excel-файлы из папки и объединяет их"""
    if not os.path.exists(folder_path):
        raise FileNotFoundError(f"Папка '{folder_path}' не найдена")

    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

    if not excel_files:
        raise ValueError(f"В папке '{folder_path}' не найдено Excel-файлов")

    dataframes = []
    for file in excel_files:
        filepath = os.path.join(folder_path, file)
        df = pd.read_excel(filepath)
        dataframes.append(df)
        print(f"✓ Прочитано: {file} — {len(df)} строк")

    df = pd.concat(dataframes, ignore_index=True)
    print(f"Объединено строк: {len(df)}")

    return df


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Очистка и предобработка данных"""
    df = df.drop_duplicates()
    df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce')

    # Удаляем строки с пропущенными ключевыми данными
    df = df.dropna(subset=['Дата', 'Количество', 'Цена_продажи'])

    # Удаляем некорректные значения
    df = df[(df['Количество'] > 0) & (df['Цена_продажи'] > 0)]

    return df


def calculate_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Расчёт всех необходимых метрик"""
    df = df.copy()
    df['Сумма_продажи'] = df['Количество'] * df['Цена_продажи']
    df['Маржа_на_единицу'] = df['Цена_продажи'] - df['Себестоимость']
    df['Прибыль'] = df['Количество'] * df['Маржа_на_единицу']
    df['Процент_маржи'] = (df['Маржа_на_единицу'] / df['Цена_продажи'] * 100).round(2)

    return df