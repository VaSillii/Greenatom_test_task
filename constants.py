from setup import get_chromedriver_path

SITE_URL = 'https://www.moex.com'
EXCEL_FILE_NAME = 'report.xlsx'
NAME_COLUMN = {
    'date': 'Дата',
    'exchange_rate': 'Курс',
    'time': 'Время'
}
CHROMEDRIVER_PATH = get_chromedriver_path()
