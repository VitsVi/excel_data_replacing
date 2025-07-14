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
    
    ELECTRICITY_COMPANY
)
from dopfuncs import copy_file_to_result, autofill_formula_down, add_border

log = logging.getLogger(__name__)

ELECTRICITY_FILENAME = 'Электричество*.xlsx'

        
        
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
    log.info(f'Открытие книги: {STORAGE_COMPONENT_FILENAME}')

    try:
        # Ежедневный отчет получение данных
        try:
            daily_sheet = daily_finance_book.sheets[DAILY_FINANCE_SHEET_NAME]
            log.info(f'Лист найден: {daily_sheet.name}')
        except:
            log.error(f'Не найден лист: {DAILY_FINANCE_SHEET_NAME}')

        daily_finance_input_data = daily_finance_data_search(daily_sheet)
    

        # выход из функции электричества, если нет компании
        if not daily_finance_input_data:
            return
        
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
  
    data_range = sheet.range(f'C{found_cell_start.row + 5}:D{found_cell_finish.row - 3}')
    
    # поиск электроэнергии
    column_c_values = data_range.columns[0].value

    if not isinstance(column_c_values, list):
        column_c_values = [column_c_values]

    for idx, value in enumerate(column_c_values):
        if value == ELECTRICITY_COMPANY:
            found_row = data_range.rows[idx].row
            log.error(f'Компания "{ELECTRICITY_COMPANY}" найдена в строке: {found_row}')
            break
    else:
        log.error(f'Компания "{ELECTRICITY_COMPANY}" не найдена в Расчетном счете')
        return None
    
    data = sheet.range(f'C{found_row}:D{found_row}')

    print(f"Значение данных: {data.value}")
    log.info(f'Данные из листа {sheet.name} получены')
    return data
    

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
        f'C{found_cell_start.row + 2}:I{found_cell_start.row + 2}',
        )


    print(f"Значение данных: {data_range.value}")
    print(f'Диапазон данных: {data_range.address}')
    log.info(f'Данные из листа {sheet.name} получены')
    return data_range

def electicity_input(
    app: xw.App,
    daily_data,
    storage_data,
):
    copy_file_to_result(ELECTRICITY_FILENAME)
    component_suppliers_path = glob(f'{RESULT_DIR}/{ELECTRICITY_FILENAME}')
    
    if not component_suppliers_path:
        log.error(f'Не найден файл: {ELECTRICITY_FILENAME}')
        sys.exit(1)
        
    book = app.books.open(component_suppliers_path[0])
    sheet = book.sheets['Лист1']
    log.info(f'Открытие книги и листа: {ELECTRICITY_FILENAME}')
    
    
    try:
        looking_range = sheet.range('D1:D600')
        for cell in reversed(looking_range):
            if cell.value is not None:
                end_row = cell.row
                break
        # подготовка шаблона
        add_border(sheet.range(f'A{end_row}:O{end_row}'))
        autofill_formula_down(sheet, f'D{end_row}', 2)
        del_D_formula = sheet.range(f'D{end_row}')
        del_D_formula.formula = None
        autofill_formula_down(sheet, f'L{end_row}', 2)
        del_L_formula = sheet.range(f'L{end_row}')
        del_L_formula.formula = None
        autofill_formula_down(sheet, f'D{end_row-1}', 2)
        autofill_formula_down(sheet, f'I{end_row-1}', 2)
        #заполнение новых данных

        # input DAILY_DATA
        try:
            log.info('Запись данных из Ежедневного отчета')
            sheet.range(f'A{end_row}').value = daily_data.sheet.name   
            sheet.range(f'H{end_row}').value = daily_data[1].value
            sheet.range(f'N{end_row}').value = daily_data[1].value
        except:
            log.info('Ошибка записи из Ежедневного отчета')
        try:
            log.info('Запись данных из Склад и Производство')
            sheet.range(f'B{end_row}').value = storage_data[0].value
            sheet.range(f'C{end_row}').value = storage_data[1].value
            sheet.range(f'J{end_row}').value = storage_data[5].value
            sheet.range(f'K{end_row}').value = storage_data[6].value
        except:
            log.info("Ошибка записи из Склад и Производство")

    except:
        log.error(f'Ошибка при работе с файлом: {ELECTRICITY_FILENAME}')
        sys.exit(1)
    finally:
        log.info(f'Закрытие книги: {ELECTRICITY_FILENAME}')
        book.save()
        book.close()
