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


DAILY_BUYERS_FILENAME = 'Покупатели ежедневные.xlsx'


# ООО ТД Л-Пак
def daily_buyers_input(
    app: xw.App,
    daily_data_1,
    daily_data_2
):
    copy_file_to_result(DAILY_BUYERS_FILENAME)

    component_suppliers_path = glob(f'{RESULT_DIR}/{DAILY_BUYERS_FILENAME}')
    
    if not component_suppliers_path:
        log.error(f'Не найден файл: {DAILY_BUYERS_FILENAME}')
        sys.exit(1)
        
    book = app.books.open(component_suppliers_path[0])
    log.info(f'Открытие книги: {DAILY_BUYERS_FILENAME}')
    
    try:
        log.info(f'Запись данных из Ежедневного отчета в {DAILY_BUYERS_FILENAME}')
        log.info('Поиск листа и вставка данных')
        
        
        
        if daily_data_1:
            list_found = False
            for row in daily_data_1.rows:
                start_row = None
                company_name = row[0].value.strip()
                for sheet_name in book.sheets:
                    if company_name in sheet_name.name.strip():
                        list_found = True
                        log.info(f'Лист найден: {sheet_name.name.strip()}')
                        sheet = book.sheets[sheet_name.name]
                        looking_range = sheet.range('A1:A200')
                        for cell in reversed(looking_range):
                            if cell.value:
                                start_row = cell.row
                                break
                        if not start_row:
                            start_row = 2
                        else:
                            start_row += 1
                        print(f'Последняя заполненная строка в листе {sheet_name.name.strip()}: {start_row}')
                        try:
                            sheet_name.range(f'A{start_row}').value = daily_data_1.sheet.name
                            sheet_name.range(f'B{start_row}').value = daily_data_1[8].value
                            log.info(f'Данные записаны в лист: {sheet_name.name.strip()}')
                        except Exception as e:
                            log.error(f'Ошибка записи данных в лист: {sheet_name.name.strip()}')
            if not list_found:
                log.info('Не найден ни один подходящий лист для данных')

        if daily_data_2:
            try:
                sheet_2 = book.sheets['ЭЛЕМЕНТ-ТРЕЙД']
            except:
                log.error(f'Не найден лист ЭЛЕМЕНТ-ТРЕЙД в {DAILY_BUYERS_FILENAME}')
            looking_range = sheet_2.range('D2:D200')
            found_row = None
            for cell in reversed(looking_range):
                if str(cell.value) == cell.formula:
                    found_row = cell.row + 1
                    break
            sheet_2.range(f'C{found_row}').value = daily_data_2.sheet.name
            sheet_2.range(f'D{found_row}').value = daily_data_2[0].value

    except:
        log.error(f'Ошибка при работе с файлом: {DAILY_BUYERS_FILENAME}')
        sys.exit(1)
    finally:
        log.info(f'Закрытие книги: {DAILY_BUYERS_FILENAME}')
        book.save()
        book.close()


def daily_finance_data_search_first(sheet: xw.Sheet):
    """Поиск данных в файле Ежедневный отчет по безналу."""

    start_target_name = 'Реализация по безналичному расчёту'
    finish_target_name = 'ИТОГО:'

    search_range = sheet.range('A1:A200')

    found_cell_start = None
    found_cell_finish = None
    for cell in search_range:
        if cell.value is not None and start_target_name in cell.value:
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
    data_range = sheet.range(f'A{found_cell_start.row + 2}:F{found_cell_finish.row - 1}')
    detect_none_row = None
    for cell in data_range.rows:
        if cell[0].value is None:
            detect_none_row = cell[0].row
            # проверка если записей нет
            if detect_none_row == found_cell_start.row + 2:
                log.info(f'Записей в Ежедневном отчете из {start_target_name} - нет')
                return None
            break

    if not detect_none_row:
        clear_data_range = data_range
    else:
        clear_data_range = sheet.range(f'A{found_cell_start.row + 2}:I{detect_none_row - 1}')
    print(f"Значение данных: {clear_data_range.value}")
    print(f'Диапазон данных: {clear_data_range.address}')
    log.info(f'Данные из листа {sheet.name} получены')
    return clear_data_range


def daily_finance_data_search_second(sheet: xw.Sheet):
    """Поиск данных в расчетном счете по рефинансированию альфа."""

    target_name = 'АО "АЛЬФА-БАНК" факторинг'

    search_range = sheet.range('E1:E100')

    found_cell_start = None
    for cell in search_range:
        if cell.value == target_name:
            found_cell_start = cell
            break

    if found_cell_start:
        print(f"Значение '{target_name}' найдено в ячейке {found_cell_start.address}")
    else:
        print(f"Значение '{target_name}' не найдено")
        return None

    # filter_data    
    data_range = sheet.range(f'F{found_cell_start.row}')
    if not data_range.value:
        log.info('Данных по АО "АЛЬФА-БАНК" факторинг не найдено')
        return None

    print(f"Значение данных: {data_range.value}")
    print(f'Диапазон данных: {data_range.address}')
    log.info(f'Данные из листа {sheet.name} получены')
    return data_range


def daily_buyers_data(app):
    """Запись данных в Поставщики комплектующих"""
    
    daily_finance_path = glob(f'{WORK_DIR}/{DAILY_FINANCE_FILENAME}')
    storage_component_path = glob(f'{WORK_DIR}/{STORAGE_COMPONENT_FILENAME}')

    if not daily_finance_path:
        log.error(f'Не найден файл: {DAILY_FINANCE_FILENAME}')
        sys.exit(1)

    if not storage_component_path:
        log.error(f'Не найден файл: {STORAGE_COMPONENT_FILENAME}')
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

        daily_finance_input_data_first = daily_finance_data_search_first(daily_sheet)
        daily_finance_input_data_second = daily_finance_data_search_second(
            daily_sheet
        )
        if daily_finance_input_data_first is None and daily_finance_input_data_second is None:
            log.info(f'Данных для записи в {DAILY_BUYERS_FILENAME} не найдено')
            return
        
        # Склад и производство получение данных
        try:
            storage_sheet = storage_book.sheets[STORAGE_COMPONENT_SHEET_NAME]
            log.info(f'Лист найден: {storage_sheet.name}')
        except:
            log.error(f'Не найден лист: {STORAGE_COMPONENT_SHEET_NAME}')
        #####
        # Запись полученных данных в поставщики комплектующих
        daily_buyers_input(
            app, 
            daily_finance_input_data_first,
            daily_finance_input_data_second
        )
    except Exception as e:
        print(e)
        log.error(
            f'Ошибка при работе с файлами: {DAILY_FINANCE_FILENAME}, '
            f'{STORAGE_COMPONENT_FILENAME}, {DAILY_BUYERS_FILENAME}'
        )
    finally:
        log.info(
            f'Закрытие книг: {DAILY_FINANCE_FILENAME}, '
            f'{STORAGE_COMPONENT_FILENAME}'
        )
        daily_finance_book.close()
        storage_book.close()
