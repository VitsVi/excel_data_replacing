import xlwings as xw
import datetime as dt
from glob import glob
import logging
import os
from first_component import copy_component_suppliers, component_suppliers_data


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("логи_выполнения.log", encoding='utf-8', mode='w'),
        logging.StreamHandler()
    ]
)
###### FILES CONTS ######


#########################







def main():
    try:
        logging.info("Запуск приложения Excel")
        app = xw.App(visible=False)
        copy_component_suppliers()
        component_suppliers_data(app)

    finally:
        logging.info("Закрытие приложения Excel")
        app.quit()
    



if __name__ == '__main__':
    main()