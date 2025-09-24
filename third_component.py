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
    DAILY_FINANCE_FILENAME,
    STORAGE_COMPONENT_FILENAME,
)
from dopfuncs import copy_file_to_result



log = logging.getLogger(__name__)



SERVICES_TO_US_FILENAME = 'Услуги нам.xlsx'


# ООО ТД Л-Пак
def services_to_us_input(
    app: xw.App,
    daily_data_1,
    daily_data_2
):
    copy_file_to_result(SERVICES_TO_US_FILENAME)

    component_suppliers_path = glob(f'{RESULT_DIR}/{SERVICES_TO_US_FILENAME}')
    
    if not component_suppliers_path:
        log.error(f'Не найден файл: {SERVICES_TO_US_FILENAME}')
        sys.exit(1)
        
    book = app.books.open(component_suppliers_path[0])
    log.info(f'Открытие книги: {SERVICES_TO_US_FILENAME}')
    
    try:
        log.info('Поиск листа и вставка данных')

        # input DAILY_DATA
        log.info('Запись данных из Ежедневного отчета в Услуги нам')
        if daily_data_1:
            log.info('Запись из ЕФО"Услуги оказывают нам" в файл Услуги нам')
            for row in daily_data_1.rows:
                start_row = None
                company_name = row[0].value.strip()
                for sheet_name in book.sheets:
                    if company_name in sheet_name.name.strip():
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
                            sheet_name.range(f'B{start_row}').value = daily_data_1[2].value
                            sheet_name.range(f'F{start_row}').value = daily_data_1[5].value
                            log.info(f'Данные записаны в лист: {sheet_name.name.strip()}')
                        except Exception as e:
                            log.error(f'Ошибка записи данных в лист: {sheet_name.name.strip()}')
        else:
            log.info('Записей из ЕФО"Услуги оказывают нам" не найдено')


        # daily_data_2 выбирает все и тут проверяет наличие нужных
        list_found = False
        for row in daily_data_2.rows:
            start_row = None
            company_name = row[0].value.strip().split('/')[0]

            for sheet_name in book.sheets:
                if company_name in sheet_name.name.strip():
                    list_found = True
                    log.info('Запись из ЕФО"Расчетный счет" в файл Услуги нам')
                    log.info(f'Лист найден: {sheet_name.name.strip()}')
                    sheet = book.sheets[sheet_name.name]
                    looking_range = sheet.range('G1:G200')
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
                        sheet_name.range(f'H{start_row}').value = row[1].value
                        sheet_name.range(f'G{start_row}').value = daily_data_2.sheet.name
                        log.info(f'Данные записаны в лист: {sheet_name.name.strip()}')
                        book.save()
                    except Exception as e:
                        log.error(f'Ошибка записи данных в лист: {sheet_name.name.strip()}')

        if not list_found:
            log.info('Записей из ЕФО"Расчетный счет" в Услуги нам не найдено')

    except:
        log.error(f'Ошибка при работе с файлом: {SERVICES_TO_US_FILENAME}')
        sys.exit(1)
    finally:
        log.info(f'Закрытие книги: {SERVICES_TO_US_FILENAME}')
        book.save()
        book.close()


def daily_finance_data_search_first(sheet: xw.Sheet):
    """Поиск данных в файле Ежедневный отчет по услугам нам."""

    start_target_name = 'Услуги оказывают нам'
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
                log.info('Записей в Ежедневном отчете из "Услуги оказывают нам" - нет')
                return None
            break

    if not detect_none_row:
        clear_data_range = data_range
    else:
        clear_data_range = sheet.range(f'A{found_cell_start.row + 2}:F{detect_none_row - 1}')
    print(f"Значение данных: {clear_data_range.value}")
    print(f'Диапазон данных: {clear_data_range.address}')
    log.info(f'Данные из листа {sheet.name} получены')
    return clear_data_range


def daily_finance_data_search_second(sheet: xw.Sheet):
    """Поиск данных в расчетном счете по агентам."""

    start_target_name = 'Расчетный счет'
    finish_target_name = 'ИТОГО'

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


def services_to_us_data(app, date_sheet):
    """Запись данных в Поставщики комплектующих"""
    log.info("РАБОТА С ФАЙЛОМ УСЛУГИ НАМ")
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
            daily_sheet = daily_finance_book.sheets[date_sheet]
            log.info(f'Лист найден: {daily_sheet.name}')
        except:
            log.error(f'Не найден лист: {date_sheet}')

        daily_finance_input_data_first = daily_finance_data_search_first(daily_sheet)
        daily_finance_input_data_second = daily_finance_data_search_second(
            daily_sheet
        )
        
        # Склад и производство получение данных
        try:
            storage_sheet = storage_book.sheets[date_sheet]
            log.info(f'Лист найден: {storage_sheet.name}')
        except:
            log.error(f'Не найден лист: {date_sheet}')
        #####
        # Запись полученных данных в поставщики комплектующих
        services_to_us_input(
            app, 
            daily_finance_input_data_first,
            daily_finance_input_data_second
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
