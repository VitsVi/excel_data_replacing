import datetime as dt
import os
import unicodedata

TIME_TODAY = dt.date.today()
CURRENT_DIR = os.getcwd()
WORK_DIR = 'files'
RESULT_DIR = f'files_result_{TIME_TODAY.strftime("%d.%m.%Y")}'

FILENAMES = [
    'Ежедневный финансовый отчет',
    'Ежедневный_отчет_Склад_*',
    'Поставщики комплектующих'
]


DAILY_FINANCE_SHEET_NAME = '02.06.2025'
DAILY_FINANCE_FILENAME = 'Ежедневны* финансовы* отчет.xlsx'
#DAILY_FINANCE_FILENAME = unicodedata.normalize('NFC', DAILY_FINANCE_FILENAME)


STORAGE_COMPONENT_SHEET_NAME = '20.05.2025'
STORAGE_COMPONENT_FILENAME = 'Ежедневны*_отчет_Склад_*'

# выбор в электричество из расчетного счета
ELECTRICITY_COMPANY = 'ИП Савина О.В. Электроэнергия'