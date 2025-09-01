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
)
from dopfuncs import copy_file_to_result



log = logging.getLogger(__name__)


FILENAME = 'Услуги мы и табель.xlsx'


# ООО ТД Л-Пак
def file_input(
    app: xw.App,
    data_1,
    data_2,
):
    copy_file_to_result(FILENAME)

    component_suppliers_path = glob(f'{RESULT_DIR}/{FILENAME}')
    
    if not component_suppliers_path:
        log.error(f'Не найден файл: {FILENAME}')
        sys.exit(1)
    ######################################################
            
    book = app.books.open(component_suppliers_path[0])
    log.info(f'Открытие книги: {FILENAME}')
    
    try:
        log.info(f'Запись данных из Ежедневного финансового отчета в {FILENAME}')
        log.info('Поиск листа и вставка данных')
        try:
            sheet_1 = book.sheets['Табель начисления ЗП']
        except:
            log.error(f'Не найден лист "Табель начисления ЗП" в {FILENAME}')

        first_date = True # первая вставка даты в таблицу листа

        ### поиск последней записи
        
        looking_range = sheet_1.range('B1:B700')
        back_row = None
        for cell in reversed(looking_range):
            if cell.value is not None:
                back_row = cell.row
                break
        ### вставка данных

        for row in data_1.value:   # row = список значений в строке
            if row[0] is not None:
                if 'зарплата' in row[0].lower():
                    autofill_formula_down(sheet_1, f'B{back_row}', 2)
                    autofill_formula_down(sheet_1, f'C{back_row}', 2)
                    autofill_formula_down(sheet_1, f'E{back_row - 1}', 2)
                    add_border(sheet_1.range(f'A{back_row}:F{back_row+1}'))
                    sheet_1.range(f'B{back_row}').formula = None
                    sheet_1.range(f'C{back_row}').formula = None
                    if first_date:
                        sheet_1.range(f'A{back_row}').value = data_1.sheet.name
                        first_date = False
                    sheet_1.range(f'C{back_row}').value = row[3]
                    sheet_1.range(f'F{back_row}').value = row[0]
                    back_row += 1

        # for data-2
        for row in data_2.value:   # row = список значений в строке
            if row[0] is not None:
                if 'зарплата' in row[0].lower():
                    autofill_formula_down(sheet_1, f'B{back_row}', 2)
                    autofill_formula_down(sheet_1, f'C{back_row}', 2)
                    autofill_formula_down(sheet_1, f'E{back_row - 1}', 2)
                    add_border(sheet_1.range(f'A{back_row}:F{back_row+1}'))
                    sheet_1.range(f'B{back_row}').formula = None
                    sheet_1.range(f'C{back_row}').formula = None
                    if first_date:
                        sheet_1.range(f'A{back_row}').value = data_1.sheet.name
                        first_date = False
                    sheet_1.range(f'C{back_row}').value = row[1]
                    sheet_1.range(f'F{back_row}').value = row[0]
                    back_row += 1
            if row[4] is not None:
                if 'зарплата' in row[4].lower():
                    autofill_formula_down(sheet_1, f'B{back_row}', 2)
                    autofill_formula_down(sheet_1, f'C{back_row}', 2)
                    autofill_formula_down(sheet_1, f'E{back_row - 1}', 2)
                    add_border(sheet_1.range(f'A{back_row}:F{back_row+1}'))
                    sheet_1.range(f'B{back_row}').formula = None
                    sheet_1.range(f'C{back_row}').formula = None
                    if first_date:
                        sheet_1.range(f'A{back_row}').value = data_1.sheet.name
                        first_date = False
                    sheet_1.range(f'C{back_row}').value = row[5]
                    sheet_1.range(f'F{back_row}').value = row[4]
                    back_row += 1

            
                    


        # if data_1:


        #     first_row = None
        #     second_row = None
        #     for cell in reversed(looking_range):
        #         if 'на сумму' in str(cell.value):
        #             first_row = cell.row
        #         if 'в единицах' in str(cell.value):
        #             second_row = cell.row
        #         if first_row and second_row:
        #             break
        #     if first_row and second_row:
        #         first_cell = sheet_1.range(f'D{first_row}')
        #         second_cell = sheet_1.range(f'D{second_row}')
        #         first_cell.formula = f'{first_cell.formula}+{data_1[6].value}'
        #         second_cell.formula = f'{second_cell.formula}+{data_1[1].value}'
        #         second_cell.formula = f'{second_cell.formula}-{data_1[2].value}'
        #     else:
        #         log.info('Ошибка записи. Не найдены строки "на сумму", "в единицах"')
        #         raise

    except:
        log.error(f'Ошибка при работе с файлом: {FILENAME}')
        sys.exit(1)
    finally:
        log.info(f'Закрытие книги: {FILENAME}')
        book.save()
        book.close()


def data_search_first(sheet: xw.Sheet):
    """Поиск данных в файле Ежедневный отчет по безналу."""

    start_target_name = 'Касса'
    finish_target_name = 'Остаток на конец дня'
    start_target_name_2 = 'Расчётный счёт'
    finish_target_name_2 = 'ИТОГО'

    clear_data_range_1 = looking_data(start_target_name, finish_target_name, sheet)
    clear_data_range_2 = looking_data(start_target_name_2, finish_target_name_2, sheet, alternative=False)
    
    return clear_data_range_1, clear_data_range_2
    



def looking_data(start_target_name, finish_target_name, sheet, alternative=True):
    search_range = sheet.range('A1:A200')
    found_cell_start = None
    found_cell_finish = None
    for cell in search_range:
        if cell.value is not None:
            if start_target_name in cell.value:
                found_cell_start = cell
            if finish_target_name in cell.value and found_cell_start is not None:
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

    if alternative:
        clear_data_range = sheet.range(f'A{found_cell_start.row + 3}:D{found_cell_finish.row - 2}')
    else:
        clear_data_range = sheet.range(f'C{found_cell_start.row + 5}:H{found_cell_finish.row - 3}')
    print(f"Значение данных: {clear_data_range.value}")
    print(f'Диапазон данных: {clear_data_range.address}')
    log.info(f'Данные из листа {sheet.name} получены')
    return clear_data_range


def we_services_and_report_card(app, date_sheet):
    """Запись данных в Поставщики комплектующих"""
    logging.info(f"РАБОТА С ФАЙЛОМ {FILENAME}")

    storage_component_path = glob(f'{WORK_DIR}/{DAILY_FINANCE_FILENAME}')

    if not storage_component_path:
        log.error(f'Не найден файл: {DAILY_FINANCE_FILENAME}')
        sys.exit(1)

    storage_book = app.books.open(storage_component_path[0])
    log.info(f'Открытие книги: {DAILY_FINANCE_FILENAME}')

    try:
        # Склад и производство получение данных
        try:
            sheet = storage_book.sheets[date_sheet]
            log.info(f'Лист найден: {sheet.name}')
        except:
            log.error(f'Не найден лист: {date_sheet}')

        input_data_first_1, input_data_first_2 = data_search_first(sheet)

        if input_data_first_1 is None:
            log.info(f'Данных для записи в {FILENAME} в не найдено')
            return
        if input_data_first_2 is None:
            log.info(f'Данных для записи в {FILENAME} в не найдено')
            return
        
        # Запись полученных данных в поставщики комплектующих
        file_input(
            app, 
            input_data_first_1,
            input_data_first_2,
        )
    except Exception as e:
        print(e)
        log.error(
            'Ошибка при работе с файлами: '
            f'{DAILY_FINANCE_FILENAME}, {FILENAME}'
        )
    finally:
        log.info(
            f'Закрытие книги: {DAILY_FINANCE_FILENAME}'
        )
        storage_book.close()
