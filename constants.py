import datetime as dt
import os


TIME_TODAY = dt.date.today()
CURRENT_DIR = os.getcwd()
WORK_DIR = 'files'
RESULT_DIR = f'files_result_{TIME_TODAY.strftime("%d.%m.%Y")}'

FILENAMES = [
    'Ежедневный финансовый отчет',
    'Ежедневный_отчет_Склад_*',
    'Поставщики комплектующих'
]