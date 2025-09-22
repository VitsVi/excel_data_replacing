from dopfuncs import autofill_formula_down, add_border
import sys
import shutil
import xlwings as xw
from glob import glob
import logging
import traceback
import os
from constants import (
    RESULT_DIR, 
    WORK_DIR, 
    STORAGE_COMPONENT_FILENAME,
)
from dopfuncs import copy_file_to_result



log = logging.getLogger(__name__)


FILENAME = 'Сальдо.xlsx'


# ООО ТД Л-Пак
def file_input(
    app: xw.App,
    data_1,
):
    copy_file_to_result(FILENAME)

    saldo_path = glob(f'{RESULT_DIR}/{FILENAME}')
    
    if not saldo_path:
        log.error(f'Не найден файл: {FILENAME}')
        sys.exit(1)
        
    book = app.books.open(saldo_path[0])
    log.info(f'Открытие книги: {FILENAME}')
    
    try:
        log.info(f'Запись данных из Склад и производство в {FILENAME}')
        log.info('Поиск листа и вставка данных')

        try:
            sheet_1 = book.sheets[0]
        except:
            log.error(f'Не найден лист в {FILENAME}')
            raise
        looking_range = sheet_1.range('A1:A300')
        looking_names = {
            'Склад (комплектующие)': data_1[3],
            'Склад (сырье)': data_1[2],
            'Склад (полуфабрикаты)': data_1[1],
            'Склад (готовая продукция': data_1[0],
        }
        try:
            score = 0
            for cell in looking_range:
                for key in looking_names.keys():
                    if cell.value and key in cell.value:
                        sheet_1.range(f'B{cell.row}').value = looking_names[key].value
                        log.info(f'Найдена ячейка для записи, {cell.value} изменена на {looking_names[key].value}')
                        score += 1
                if score == 4:
                    log.info('Все ячейки успешно заполнены')
                    break
        except:
            traceback.print_exc()
            log.error('Ошибка при заполнении ячеек Сальдо')


    except:
        log.error(f'Ошибка при работе с файлом: {FILENAME}')
        sys.exit(1)
    finally:
        log.info(f'Закрытие книги: {FILENAME}')
        book.save()
        book.close()


def data_catching(sheet: xw.Sheet, start_target_name:str, finish_target_name:str, column:str):
    """Выбор конкретных ячеек из диапазона, п8 в файле тз"""

    # start_target_name = 'Склад готовой продукции'
    # finish_target_name = 'ИТОГО'
    try:
        search_range = sheet.range('A1:A400')

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
        #print(f"Значение данных: {clear_data_range.value}")
        print(f'Диапазон данных: {clear_data_range.address}')
        last_row = clear_data_range.last_cell.row
        result = sheet.range(f'{column}{last_row}')
        log.info(f'Последня ячейка для Сальдо: {column}{last_row} = {result.value}')
        if not result.value:
            log.error(f'В найденной ячейке нет значения: {column}{last_row} = {result.value}')
        return result
    except:
        traceback.print_exc()
        log.error(f'Не найдены данные для Сальдо по поиску таблицы: {start_target_name}, {finish_target_name} в ячейке {column}')


def data_search_first(sheet: xw.Sheet):
    """Поиск данных в файле Ежедневный отчет склад и производство"""
    results = []
    results.append(data_catching(sheet, start_target_name='Склад готовой продукции', finish_target_name='ИТОГО', column='N'))
    results.append(data_catching(sheet, start_target_name='СКЛАД Производства', finish_target_name='ИТОГО:', column='G'))
    results.append(data_catching(sheet, start_target_name='СКЛАД СЫРЬЕВЫХ КОМПОНЕНТОВ', finish_target_name='Остаток по складу в деньгах', column='I'))
    results.append(data_catching(sheet, start_target_name='Склад Материалов', finish_target_name='ИТОГО', column='I'))
    log.info(f'Данные из листа {sheet.name} получены')
    return results


def saldo_data(app, date_sheet):
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

        
        if not input_data_first:
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
