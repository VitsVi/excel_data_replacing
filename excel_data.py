import xlwings as xw
import datetime as dt
from glob import glob
import logging
import os
import sys


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("логи_выполнения.log", encoding='utf-8', mode='w'),
        logging.StreamHandler()
    ]
)



TIME_TODAY = dt.date.today()
CURRENT_DIR = os.getcwd()
WORK_DIR = 'files'
RESULT_DIR = f'files_result_{TIME_TODAY.strftime("%m.%d.%Y")}'

FILENAMES = [
    'Ежедневный финансовый отчет',
    'Ежедневный_отчет_Склад_*',
    'Поставщики комплектующих'
]

# ООО ТД Л-Пак


def daily_finance_open(
    app: xw.App,
    sheet_name='02.06.2025',
    filename='Ежедневный финансовый отчет.xlsx'
):
    daily_finance_path = glob(f'{WORK_DIR}/{filename}')
    
    if not daily_finance_path:
        logging.error(f'Не найден файл: {filename}')
        sys.exit(1)

    book = app.books.open(daily_finance_path[0])
    filename = os.path.basename(daily_finance_path[0])
    try:
        try:
            sheet = book.sheets[sheet_name]
            logging.info(f'Лист найден: {sheet.name}')
        except:
            logging.error(f'Не найден лист: {sheet_name}')

        input_data = daily_finance_search_data(sheet)
    except:
        logging.error(f'Ошибка при работе с файлом: {filename}')
    finally:
        logging.info(f'Закрытие книги: {filename}')
        book.close()

    
    # book.close()
    # return None

def daily_finance_search_data(sheet: xw.Sheet):
    """ Ищет данные в заданном диапазоне книги Excel."""
    start_target_name = 'Расчётный счёт'
    finish_target_name = 'ИТОГО:'

    search_range = sheet.range('A1:E100')

    found_cell_start = None
    found_cell_finish = None
    for cell in search_range:
        if cell.value == start_target_name:
            found_cell_start = cell
        if cell.value == finish_target_name:
            found_cell_finish = cell
        if found_cell_start and found_cell_finish:
            break

    if found_cell_start:
        print(f"Значение '{start_target_name}' найдено в ячейке {found_cell_start.address}")
    else:
        print(f"Значение '{start_target_name}' не найдено")
        sys.exit(1)

    if found_cell_finish:
        print(f"Значение '{finish_target_name}' найдено в ячейке {found_cell_finish.address}")
    else:
        print(f"Значение '{finish_target_name}' не найдено")
        sys.exit(1)

    # filter_none_data    
    data_range = sheet.range(f'C{found_cell_start.row + 5}:D{found_cell_finish.row - 3}')

    detect_none_row = None
    for cell in data_range:
        if cell.value is None:
            detect_none_row = cell.row
            break

    if not detect_none_row:
        clear_data_range = data_range
    else:
        clear_data_range = sheet.range(f'C{found_cell_start.row + 5}:D{detect_none_row - 1}')

    print(f"Значение данных: {clear_data_range.value}")
    print(f'Диапазон данных: {clear_data_range.address}')
    logging.info(f'Данные из листа {sheet.name} получены')
    return clear_data_range



        
        

def main():
    try:
        app = xw.App(visible=False)
        daily_finance_open(app)
        

        # daily_finance_path = glob(f'{WORK_DIR}/Ежедневный финансовый отчет.xlsx')[0]
        # daily_finance_production_path = glob(f'{WORK_DIR}/Ежедневный_отчет_Склад_*')[0]
        ...
    finally:# Закрываем исходник
        app.quit()
    



if __name__ == '__main__':
    main()