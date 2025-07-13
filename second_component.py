from dopfuncs import autofill_formula_down, add_border
import sys
import shutil
import xlwings as xw
from glob import glob
import logging
import os
from constants import (
    RESULT_DIR, 
    WORK_DIR, 
    DAILY_FINANCE_SHEET_NAME,
    DAILY_FINANCE_FILENAME,

    STORAGE_COMPONENT_SHEET_NAME,
    STORAGE_COMPONENT_FILENAME,
)

log = logging.getLogger(__name__)

ELECTRICITY_FILENAME = 'Электричество*.xlsx'

def copy_electricity():
    try:
        os.mkdir(RESULT_DIR)
    except:
        log.info(f'Папка {RESULT_DIR} уже создана, повторное создание не нужно')
    
    orig_path = glob(f'{WORK_DIR}/{ELECTRICITY_FILENAME}')
    if not orig_path:
        log.error(f'Не найден файл: {ELECTRICITY_FILENAME}')
        sys.exit(1)
    #abs_path = Path(__file__).resolve().parent
    filename = os.path.basename(orig_path[0])
    result_path = f'{RESULT_DIR}/{filename}'
    print(result_path)

    try:
        shutil.copy2(orig_path[0], result_path)
        log.info(f'{filename} - Файл скопирован в {RESULT_DIR}')
    except Exception as e:
        print(e)
        log.error(f'Ошибка копирования файла {filename}')
        
        
def electicity_data(app):
    """Запись данных в Электричество"""
    
    daily_finance_path = glob(f'{WORK_DIR}/{DAILY_FINANCE_FILENAME}')
    storage_component_path = glob(f'{WORK_DIR}/{STORAGE_COMPONENT_FILENAME}')

    if not daily_finance_path:
        log.error(f'Не найден файл: {DAILY_FINANCE_FILENAME}')
        sys.exit(1)

    if not storage_component_path:
        log.error(f'Не найден файл: {DAILY_FINANCE_FILENAME}')
        sys.exit(1)
        
    daily_finance_book = app.books.open(daily_finance_path[0])
    log.info(f'Открытие книги: {DAILY_FINANCE_FILENAME}')
    storage_book = app.books.open(storage_component_path[0])
    log.info(f'Открытие книги: {STORAGE_COMPONENT_SHEET_NAME}')

    try:
        # Ежедневный отчет получение данных
        try:
            daily_sheet = daily_finance_book.sheets[DAILY_FINANCE_SHEET_NAME]
            log.info(f'Лист найден: {daily_sheet.name}')
        except:
            log.error(f'Не найден лист: {DAILY_FINANCE_SHEET_NAME}')

        daily_finance_input_data = daily_finance_data_search(daily_sheet)
        
        # Склад и производство получение данных
        try:
            storage_sheet = storage_book.sheets[STORAGE_COMPONENT_SHEET_NAME]
            log.info(f'Лист найден: {storage_sheet.name}')
        except:
            log.error(f'Не найден лист: {STORAGE_COMPONENT_SHEET_NAME}')
        try:
            storage_component_input_data = storage_data_search(storage_sheet)
        except Exception as e:
            print(e)
        #####
        # Запись полученных данных в поставщики комплектующих
        electicity_input(
            app, 
            daily_finance_input_data,
            storage_component_input_data
        )
    except Exception as e:
        print(e)
        log.error(
            f'Ошибка при работе с файлами: {DAILY_FINANCE_FILENAME}, '
            f'{STORAGE_COMPONENT_FILENAME}, {ELECTRICITY_FILENAME}'
        )
    finally:
        log.info(
            f'Закрытие книг: {DAILY_FINANCE_FILENAME}, '
            f'{STORAGE_COMPONENT_FILENAME}'
        )
        daily_finance_book.close()
        storage_book.close()
        
        
def daily_finance_data_search(daily_sheet):
    ...
    

def storage_data_search(sheet):
    """Поиск данных в файле Склад и производство"""

    start_target_name = 'Показания эл.счетчика кВТ на начало дня'

    search_range = sheet.range('A1:A500')

    found_cell_start = None
    for cell in search_range:
        if cell.value == start_target_name:
            found_cell_start = cell
            break

    if found_cell_start:
        print(f"Значение '{start_target_name}' найдено в ячейке {found_cell_start.address}")
    else:
        print(f"Значение '{start_target_name}' не найдено")
        sys.exit(1)

    
    data_range = sheet.range(
        # +2 из-за объекдиненной ячейки start
        f'A{found_cell_start.row + 2}:J{found_cell_start.row + 2}',
        )


    print(f"Значение данных: {data_range.value}")
    print(f'Диапазон данных: {data_range.address}')
    log.info(f'Данные из листа {sheet.name} получены')
    return data_range

def celecticity_input(
    app: xw.App,
    daily_data,
    storage_data,
):
    ...