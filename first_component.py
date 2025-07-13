from dopfuncs import autofill_formula_down, add_border
import sys
import shutil
import xlwings as xw
from glob import glob
import logging
from constants import RESULT_DIR, WORK_DIR
import os


log = logging.getLogger(__name__)

DAILY_FINANCE_SHEET_NAME = '02.06.2025'
DAILY_FINANCE_FILENAME = 'Ежедневный финансовый отчет.xlsx'

STORAGE_COMPONENT_SHEET_NAME = '20.05.2025'
STORAGE_COMPONENT_FILENAME = 'Ежедневный_отчет_Склад_*'

COMPONENT_SUPPLIERS_FILENAME = 'Поставщики комплектующих.xlsx'


# ООО ТД Л-Пак
def component_suppliers_input(
    app: xw.App,
    daily_data,
    storage_data,
):
    component_suppliers_path = glob(f'{RESULT_DIR}/{COMPONENT_SUPPLIERS_FILENAME}')
    
    if not component_suppliers_path:
        log.error(f'Не найден файл: {COMPONENT_SUPPLIERS_FILENAME}')
        sys.exit(1)
        
    book = app.books.open(component_suppliers_path[0])
    log.info(f'Открытие книги: {COMPONENT_SUPPLIERS_FILENAME}')
    
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
                
        # input STORAGE_DATA
        log.info('Запись данных из Склад и производство в Поставщики комплектующих')
        used_storage_sheet = []
        for row in storage_data.rows:
            if row[1].value.strip() in used_storage_sheet:
                continue
            list_found = False
            lower_ground = None
            for sheet_name in book.sheets:
                if sheet_name.name.strip() == row[1].value.strip():
                    list_found = True
                    log.info(f'Лист найден: {row[1].value}')
                    sheet = book.sheets[sheet_name.name]
                    looking_range = sheet.range('A1:A200')
                    formula_cell_E = None
                    for cell in reversed(looking_range):
                        
                        if sheet_name.range(f'E{cell.row}').formula and not formula_cell_E:
                            try:
                                formula_cell_E = sheet_name.range(f'E{cell.row}')
                            except Exception as e:
                                print(e)
                        if cell.value is not None:
                            lower_ground = cell.row + 1
                            break

                        
                    for row_in in storage_data.rows:
                        if sheet_name.name.strip() == row_in[1].value.strip() and row_in[4].value:
                            try:
                                cell_input_date = sheet_name.range(f'A{lower_ground}')
                                cell_input_product = sheet_name.range(f'B{lower_ground}')
                                cell_input_cost = sheet_name.range(f'D{lower_ground}')
                                cell_input_count = sheet_name.range(f'C{lower_ground}')
                                cell_input_date.value = storage_data.sheet.name
                                cell_input_product.value = row_in[0].value
                                cell_input_cost.value = row_in[6].value
                                cell_input_count.value = row_in[4].value
                                
                                cell_input_formule_E = sheet.range(f'E{lower_ground}')
                                cell_input_formule_E.value = f'=C{lower_ground}*D{lower_ground}'

                                lower_ground += 1
                                if lower_ground == formula_cell_E.row:
                                    print('LOWER:', lower_ground)
                                    #delete formula
                                    cell_G = sheet_name.range(f'G{lower_ground}')
                                    cell_E = sheet_name.range(f'E{lower_ground}')
                                    cell_I = sheet_name.range(f'I{lower_ground}')
                                    cell_G.formula = None
                                    cell_E.formula = None
                                    cell_I.formula = None
                                    #replace formula
                                    formula_cell_G = sheet_name.range(f'G{lower_ground+5}')
                                    formula_cell_E = sheet_name.range(f'E{lower_ground+5}')
                                    formula_cell_I = sheet_name.range(f'I{lower_ground+5}')
                                    print(formula_cell_I)
                                    formula_cell_E.formula = f'=SUM(E2:E{lower_ground+4})'
                                    formula_cell_G.formula = f'=SUM(G2:G{lower_ground+4})'
                                    formula_cell_I.formula = f'=G{lower_ground+5}-E{lower_ground+5}'
                                    print(formula_cell_I.formula)
                                    main_formula_cell_I = sheet_name.range(f'I{lower_ground-1}')
                                    
                                    autofill_formula_down(sheet_name, main_formula_cell_I, 6)
                                    # formula_cell_I.value = ...
                                    # formula_cell_I = ...
                            except Exception as e:
                                print(e)
                    book.save()
                    print('ELSE')
                    if lower_ground:
                        add_border(sheet_name.range(f'A1:I{lower_ground+4}'))

            if not list_found:
                log.error(f'НЕ найден лист: {row[1].value}')
            # убираем дублирование проверок
            used_storage_sheet.append(row[1].value.strip())
        

    except:
        log.error(f'Ошибка при работе с файлом: {COMPONENT_SUPPLIERS_FILENAME}')
        sys.exit(1)
    finally:
        log.info(f'Закрытие книги: {COMPONENT_SUPPLIERS_FILENAME}')
        #book.save()
        book.close()
    ...

    
    # book.close()
    # return None

def daily_finance_search_data(sheet: xw.Sheet):
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


def storage_data_search(sheet):
    """Поиск данных в файле Склад и производство"""

    start_target_name = 'Материалы'
    finish_target_name = 'ИТОГО'

    search_range = sheet.range('A1:A500')

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
    
    data_range = sheet.range(
        f'A{found_cell_start.row + 1}:G{found_cell_finish.row - 1}',
        )


    print(f"Значение данных: {data_range.value}")
    print(f'Диапазон данных: {data_range.address}')
    log.info(f'Данные из листа {sheet.name} получены')
    return data_range

def daily_and_storage_data_search(sheet):
    """Поиск данных в файле Склад и производство"""

    start_target_name = 'Материалы'
    finish_target_name = 'ИТОГО'

    search_range = sheet.range('A1:A500')

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


def component_suppliers_data(app):
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
    log.info(f'Открытие книги: {STORAGE_COMPONENT_SHEET_NAME}')

    try:
        # Ежедневный отчет получение данных
        try:
            daily_sheet = daily_finance_book.sheets[DAILY_FINANCE_SHEET_NAME]
            log.info(f'Лист найден: {daily_sheet.name}')
        except:
            log.error(f'Не найден лист: {DAILY_FINANCE_SHEET_NAME}')

        daily_finance_input_data = daily_finance_search_data(daily_sheet)
        
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
        component_suppliers_input(
            app, 
            daily_finance_input_data,
            storage_component_input_data
        )
    except Exception as e:
        print(e)
        log.error(
            f'Ошибка при работе с файлами: {DAILY_FINANCE_FILENAME}, '
            f'{STORAGE_COMPONENT_FILENAME}, {COMPONENT_SUPPLIERS_FILENAME}'
        )
    finally:
        log.info(
            f'Закрытие книг: {DAILY_FINANCE_FILENAME}, '
            f'{STORAGE_COMPONENT_FILENAME}'
        )
        daily_finance_book.close()
        storage_book.close()



def copy_component_suppliers():
    try:
        os.mkdir(RESULT_DIR)
    except:
        log.info(f'Папка {RESULT_DIR} уже создана, повторное создание не нужно')
    
    orig_path = glob(f'{WORK_DIR}/{COMPONENT_SUPPLIERS_FILENAME}')
    #abs_path = Path(__file__).resolve().parent
    result_path = f'{RESULT_DIR}/{COMPONENT_SUPPLIERS_FILENAME}'
    print(result_path)
    
    if not orig_path:
        log.error(f'Не найден файл: {COMPONENT_SUPPLIERS_FILENAME}')
        sys.exit(1)
    try:
        shutil.copy2(orig_path[0], result_path)
        log.info(f'{COMPONENT_SUPPLIERS_FILENAME} - Файл скопирован в {RESULT_DIR}')
    except Exception as e:
        print(e)
        log.error(f'Ошибка копирования файла {COMPONENT_SUPPLIERS_FILENAME}')