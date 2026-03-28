from datetime import datetime

# Настройки проекта
DEFAULT_FOLDER = 'reports'
DEFAULT_OUTPUT_NAME = lambda: f'sales_report_{datetime.now().strftime("%Y-%m-%d")}.xlsx'

# Названия листов в Excel
SHEET_NAMES = {
    'main': 'Общие_данные',
    'managers': 'По_менеджерам',
    'categories': 'По_категориям',
    'top': 'Топ_10_товаров'
}

# Цвет шапки
HEADER_COLOR = "366092"