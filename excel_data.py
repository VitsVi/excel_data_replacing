import xlwings as xw
import datetime as dt
from glob import glob
import logging
import os
from first_component import component_suppliers_data
from second_component import electicity_data
from third_component import services_to_us_data


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
        #logging.info("РАБОТА С ФАЙЛОМ ПОСТАВЩИКИ КОМЛПЕКТУЮЩИХ")
        #component_suppliers_data(app)
        logging.info("РАБОТА С ФАЙЛОМ ЭЛЕКТРИЧЕСТВО")
        electicity_data(app)
        #logging.info("РАБОТА С ФАЙЛОМ УСЛУГИ НАМ")
        #services_to_us_data(app)
        
    finally:
        logging.info("Закрытие приложения Excel")
        app.quit()
        input("Завершение работы. Нажмите любую клавишу для продолжения...")
    



if __name__ == '__main__':
    main()