import xlwings as xw
import datetime as dt
from glob import glob
import logging
import os

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("логи_выполнения.log", encoding='utf-8', mode='w'),
        logging.StreamHandler()
    ]
)


class File:
    def __init__(self, dataframe, filename):
        self.df = dataframe
        self.filename = filename


TIME_TODAY = dt.date.today()
CURRENT_DIR = os.getcwd()
WORK_DIR = 'files'
RESULT_DIR = f'files_result_{TIME_TODAY.strftime('%m.%d.%Y')}'

FILENAMES = [
    'Ежедневный финансовый отчет',
    'Ежедневный_отчет_Склад_*',
    'Поставщики комплектующих'
]


def daily_finance():
    daily_finance_path = glob(f'{WORK_DIR}/Ежедневный финансовый отчет.xlsx')
    daily_finance_production_path = glob(f'{WORK_DIR}/Ежедневный_отчет_Склад_*')
    
    
    
    if not daily_finance_path:
        logging.error('Не найден файл: Ежедневный финансовый отчет.xlsx')
    if not daily_finance_production_path:
        logging.error('Не найден файл: Ежедневный_отчет_Склад_...')
        

    daily_finance = File(
        pd.read_excel(daily_finance_path[0], sheet_name=None),
        os.path.basename(daily_finance_path[0])
    )
    daily_finance_production = File(
        pd.read_excel(daily_finance_production_path[0], sheet_name=None),
        os.path.basename(daily_finance_production_path[0])
    )
    
    
    return daily_finance, daily_finance_production


def result_add(*data_files):
    try:
        os.mkdir(RESULT_DIR)
    except FileExistsError:
        logging.info(f'Папка {RESULT_DIR} уже создана. Повторное создание не нужно')
        
    for file in data_files:
        with pd.ExcelWriter(f'{RESULT_DIR}/{file.filename}') as writer:
            for sheet_name, sheet_df in file.df.items():
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        

def main():
    app = xw.App(visible=False)
    daily_finance_path = glob(f'{WORK_DIR}/Ежедневный финансовый отчет.xlsx')[0]
    daily_finance_production_path = glob(f'{WORK_DIR}/Ежедневный_отчет_Склад_*')[0]
    wb1 = app.books.open(daily_finance_path) 
    wb2 = app.books.open(daily_finance_path)
    try:
        print([sheet.name for sheet in wb1.sheets])

        wl1 = wb1.sheets['29.05. ']
        search_range = wl1.range('A1:E100')
        print(wl1)
        target = 'Расчётный счёт'
        found_cell = None
        for cell in search_range:
            if cell.value == target:
                found_cell = cell
                break
        # save_path = os.path.join(RESULT_DIR, 'test.xlsx')
        # print(save_path)
        # Сохраняем копию с новым именем
        if found_cell:
            print(f"Значение '{target}' найдено в ячейке {found_cell.address}")
        else:
            print(f"Значение '{target}' не найдено")
        ...
        # wb.api.SaveCopyAs(save_path)     # Сохраняем копию под новым именем
        # wb.close()
    finally:# Закрываем исходник
        wb1.close()
        app.quit()
    



if __name__ == '__main__':
    main()