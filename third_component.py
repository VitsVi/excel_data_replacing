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
from dopfuncs import copy_file_to_result



log = logging.getLogger(__name__)



SERVICES_TO_US_FILENAME = 'Услуги нам.xlsx'


# ООО ТД Л-Пак
def services_to_us_input(
    app: xw.App,
    daily_data,
):
    copy_file_to_result(SERVICES_TO_US_FILENAME)

    component_suppliers_path = glob(f'{RESULT_DIR}/{SERVICES_TO_US_FILENAME}')
    
    if not component_suppliers_path:
        log.error(f'Не найден файл: {SERVICES_TO_US_FILENAME}')
        sys.exit(1)
        
    book = app.books.open(component_suppliers_path[0])
    log.info(f'Открытие книги: {SERVICES_TO_US_FILENAME}')
    
    try:
        log.info('Поиск листа с поставщиками и вставка данных')

        # input DAILY_DATA
        log.info('Запись данных из Ежедневного отчета в Поставщики комплектующих')
        for row in daily_data.rows:
            list_found = False
            lower_ground = None
            for sheet_name in book.sheets:
                if sheet_name.name.strip() == row[0].value.strip():
                    list_found = True
                    log.info(f'Лист найден: {row[0].value}')
                    sheet = book.sheets[sheet_name.name]
                    looking_range = sheet.range('A1:A200')
                    formula_cell = None
                    for cell in reversed(looking_range):
                        
                        if sheet_name.range(f'E{cell.row}').formula and not formula_cell:
                            formula_cell = sheet_name.range(f'E{cell.row}')
                        if cell.value is not None:
                            lower_ground = sheet_name.range(f'A{cell.row + 1}')
                            break
                    if formula_cell.row - lower_ground.row > daily_data.rows.count:
                        try:
                           # lower_ground.value = daily_data.value
                            book.save()
                        except Exception as e:
                            print(e)
                        
                    else:
                        cell_input_data = sheet_name.range(f'G{lower_ground.row}')
                        cell_input_date = sheet_name.range(f'F{lower_ground.row}')
                        cell_input_data.value = row[1].value
                        cell_input_date.value = daily_data.sheet.name
                        book.save()
                        print('ELSE')
            if not list_found:
                log.error(f'НЕ найден лист: {row[0].value}')

        

    except:
        log.error(f'Ошибка при работе с файлом: {SERVICES_TO_US_FILENAME}')
        sys.exit(1)
    finally:
        log.info(f'Закрытие книги: {SERVICES_TO_US_FILENAME}')
        #book.save()
        book.close()
    ...

    
    # book.close()
    # return None

def daily_finance_data_search(sheet: xw.Sheet):
    """Поиск данных в файле Ежедневный отчет"""

    start_target_name = 'Расчётный счёт'
    finish_target_name = 'ИТОГО:'

    search_range = sheet.range('A1:E100')

    found_cell_start = None
    found_cell_finish = None
    for cell in search_range:
        if cell.value == start_target_name:
            found_cell_start = cell
        if cell.value == finish_target_name and found_cell_start is not None:
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
    log.info(f'Данные из листа {sheet.name} получены')
    return clear_data_range


def services_to_us_data(app):
    """Запись данных в Поставщики комплектующих"""
    
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
    log.info(f'Открытие книги: {STORAGE_COMPONENT_FILENAME}')

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
        #####
        # Запись полученных данных в поставщики комплектующих
        services_to_us_input(
            app, 
            daily_finance_input_data,
        )
    except Exception as e:
        print(e)
        log.error(
            f'Ошибка при работе с файлами: {DAILY_FINANCE_FILENAME}, '
            f'{STORAGE_COMPONENT_FILENAME}, {SERVICES_TO_US_FILENAME}'
        )
    finally:
        log.info(
            f'Закрытие книг: {DAILY_FINANCE_FILENAME}, '
            f'{STORAGE_COMPONENT_FILENAME}'
        )
        daily_finance_book.close()
        storage_book.close()
