from constants import RESULT_DIR, WORK_DIR
import logging
from glob import glob
import os
import shutil
import sys


log = logging.getLogger(__name__)

def autofill_formula_down(sheet, cell_address, num_rows):
    """
    Протягивает формулу вниз с автокоррекцией относительных ссылок (аналог drag в Excel).
    
    :param sheet: лист xlwings
    :param cell_address: строка вида "E25" — начальная ячейка с формулой
    :param num_rows: сколько строк вниз протянуть (включая начальную)
    """
    start_cell = sheet.range(cell_address)
    formula = start_cell.formula

    # Определяем диапазон, который надо "залить" формулой
    fill_range = sheet.range(cell_address).expand('right').resize(num_rows, 1)

    # Применяем формулу ко всему диапазону (как автозаполнение)
    fill_range.formula = formula
    
    
def add_border(range_obj, weight=2, linestyle=1):
    """
    Добавляет рамки диапазону через .api (COM) — без прямого импорта win32.
    :param range_obj: xlwings Range
    :param weight: 2 = xlMedium
    :param linestyle: 1 = xlContinuous
    """
    borders = range_obj.api.Borders
    for i in range(7, 13):  # xlEdgeLeft=7 до xlInsideHorizontal=12
        borders(i).LineStyle = linestyle
        borders(i).Weight = weight


def copy_file_to_result(filename):
    try:
        os.mkdir(RESULT_DIR)
    except:
        log.info(f'Папка {RESULT_DIR} уже создана, повторное создание не нужно')
    
    orig_path = glob(f'{WORK_DIR}/{filename}')
    #abs_path = Path(__file__).resolve().parent

    if not orig_path:
        log.error(f'Не найден файл: {filename}')
        sys.exit(1)
    try:
        result_path = f'{RESULT_DIR}/{os.path.basename(orig_path[0])}'
        shutil.copy2(orig_path[0], result_path)
        log.info(f'{filename} - Файл скопирован в {RESULT_DIR}')
    except Exception as e:
        print(e)
        log.error(f'Ошибка копирования файла {filename}')