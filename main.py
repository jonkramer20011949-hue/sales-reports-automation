import argparse
from datetime import datetime

from data_processor import load_and_merge_data, clean_data, calculate_metrics
from report_generator import generate_report
from config import DEFAULT_FOLDER


def main():
    parser = argparse.ArgumentParser(description='Автоматизация обработки отчётов продаж')
    parser.add_argument('--folder', '-f', default=DEFAULT_FOLDER,
                        help='Папка с исходными Excel-файлами')
    parser.add_argument('--output', '-o', default=None,
                        help='Имя выходного файла')

    args = parser.parse_args()

    # Определяем имя выходного файла
    if args.output is None:
        output_file = f'sales_report_{datetime.now().strftime("%Y-%m-%d")}.xlsx'
    else:
        output_file = args.output
        if not output_file.endswith('.xlsx'):
            output_file += '.xlsx'

    print("Запуск обработки отчётов продаж\n")
    print(f"Папка: {args.folder}")
    print(f"Выходной файл: {output_file}\n")

    try:
        # Загрузка и обработка данных
        df = load_and_merge_data(args.folder)
        df = clean_data(df)
        df = calculate_metrics(df)

        # Генерация отчёта
        generate_report(df, output_file)


    except Exception as e:
        print(f"\nОшибка: {e}")


if __name__ == "__main__":
    main()