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