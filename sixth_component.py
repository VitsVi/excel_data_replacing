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
    STORAGE_COMPONENT_FILENAME,
)
from dopfuncs import copy_file_to_result



log = logging.getLogger(__name__)


FILENAME = 'Склад комплект.xlsx'


# ООО ТД Л-Пак
def file_input(
    app: xw.App,
    data_1,
):
    copy_file_to_result(FILENAME)

    component_suppliers_path = glob(f'{RESULT_DIR}/{FILENAME}')
    
    if not component_suppliers_path:
        log.error(f'Не найден файл: {FILENAME}')
        sys.exit(1)
        
    book = app.books.open(component_suppliers_path[0])
    log.info(f'Открытие книги: {FILENAME}')
    
    try:
        log.info(f'Запись данных из Склад и производство в {FILENAME}')
        log.info('Поиск листа и вставка данных')

        if data_1:
            try:
                sheet_1 = book.sheets['склад']
            except:
                log.error(f'Не найден лист склад в {FILENAME}')
            looking_range = sheet_1.range('A1:A50')
            first_row = None
            second_row = None
            for cell in reversed(looking_range):
                if 'на сумму' in str(cell.value):
                    first_row = cell.row
                if 'в единицах' in str(cell.value):
                    second_row = cell.row
                if first_row and second_row:
                    break
            if first_row and second_row:
                first_cell = sheet_1.range(f'D{first_row}')
                second_cell = sheet_1.range(f'D{second_row}')
                first_cell.formula = f'{first_cell.formula}+{data_1[6].value}'
                second_cell.formula = f'{second_cell.formula}+{data_1[1].value}'
                second_cell.formula = f'{second_cell.formula}-{data_1[2].value}'
            else:
                log.info('Ошибка записи. Не найдены строки "на сумму", "в единицах"')
                raise

    except:
        log.error(f'Ошибка при работе с файлом: {FILENAME}')
        sys.exit(1)
    finally:
        log.info(f'Закрытие книги: {FILENAME}')
        book.save()
        book.close()


def data_search_first(sheet: xw.Sheet):
    """Поиск данных в файле Ежедневный отчет по безналу."""

    start_target_name = 'Материалы'
    finish_target_name = 'ИТОГО'

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

    clear_data_range = sheet.range(f'C{found_cell_finish.row}:I{found_cell_finish.row}')
    print(f"Значение данных: {clear_data_range.value}")
    print(f'Диапазон данных: {clear_data_range.address}')
    log.info(f'Данные из листа {sheet.name} получены')
    return clear_data_range


def storage_complect_data(app, date_sheet):
    """Запись данных в Поставщики комплектующих"""
    log.info(f"РАБОТА С ФАЙЛОМ {FILENAME}")

    storage_component_path = glob(f'{WORK_DIR}/{STORAGE_COMPONENT_FILENAME}')

    if not storage_component_path:
        log.error(f'Не найден файл: {STORAGE_COMPONENT_FILENAME}')
        sys.exit(1)

    storage_book = app.books.open(storage_component_path[0])
    log.info(f'Открытие книги: {STORAGE_COMPONENT_FILENAME}')

    try:
        # Склад и производство получение данных
        try:
            sheet = storage_book.sheets[date_sheet]
            log.info(f'Лист найден: {sheet.name}')
        except:
            log.error(f'Не найден лист: {date_sheet}')

        input_data_first = data_search_first(sheet)

        if input_data_first is None:
            log.info(f'Данных для записи в {FILENAME} не найдено')
            return
        
        # Запись полученных данных в поставщики комплектующих
        file_input(
            app, 
            input_data_first,
        )
    except Exception as e:
        print(e)
        log.error(
            'Ошибка при работе с файлами: '
            f'{STORAGE_COMPONENT_FILENAME}, {FILENAME}'
        )
    finally:
        log.info(
            f'Закрытие книги: {STORAGE_COMPONENT_FILENAME}'
        )
        storage_book.close()
