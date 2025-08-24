import xlwings as xw
import datetime as dt
from datetime import datetime
from glob import glob
import logging
import os
from colorama import init, Fore, Style
from first_component import component_suppliers_data
from second_component import electicity_data
from third_component import services_to_us_data
from fourth_component import daily_buyers_data
from fifth_component import storage_material_data
from sixth_component import storage_complect_data
from seventh_component import we_services_and_report_card


init()

# Кастомный цветной форматтер
class ColorFormatter(logging.Formatter):
    def format(self, record):
        msg = super().format(record)
        if record.levelno == logging.ERROR:
            return Fore.RED + msg + Style.RESET_ALL
        elif record.levelno == logging.INFO:
            return Fore.GREEN + msg + Style.RESET_ALL
        elif record.levelno == logging.WARNING:
            return Fore.YELLOW + msg + Style.RESET_ALL
        else:
            return msg

# Формат логов
log_format = '%(asctime)s - %(levelname)s - %(message)s'

# Файл — обычный формат
file_handler = logging.FileHandler("логи_выполнения.log", encoding='utf-8', mode='w')
file_handler.setFormatter(logging.Formatter(log_format))

# Консоль — цветной формат
stream_handler = logging.StreamHandler()
stream_handler.setFormatter(ColorFormatter(log_format))

# Настройка логгера
logging.basicConfig(
    level=logging.INFO,
    handlers=[file_handler, stream_handler]
)
#########################

TEXT_OUT = {
    "all": 'Запуск всех стадий',
    '1': 'РАБОТА С ФАЙЛОМ ПОСТАВЩИКИ КОМПЛЕКТУЮЩИХ',
    '2': 'РАБОТА С ФАЙЛОМ ЭЛЕКТРИЧЕСТВО',
    '3': 'РАБОТА С ФАЙЛОМ УСЛУГИ НАМ',
    '4': 'РАБОТА С ФАЙЛОМ ЕЖЕДНЕВНЫЕ ПОКУПАТЕЛИ',
    '5': 'РАБОТА С ФАЙЛОМ СКЛАД СЫРЬЕ',
    '6': 'РАБОТА С ФАЙЛОМ СКЛАД КОМПЛЕКТ',
    '7': 'РАБОТА С ФАЙЛОМ УСЛУГИ МЫ И ТАБЕЛЬ',
    '0': 'ЗАПУСК ВЫБРАННЫХ ПРОГРАММ',
}
CHOOSEN_FILES = {
    '1': component_suppliers_data,
    '2': electicity_data,
    '3': services_to_us_data,
    '4': daily_buyers_data,
    '5': storage_material_data,
    '6': storage_complect_data,
    '7': we_services_and_report_card
}

def funcs_manager():
    
    console = True
    choosen_files = []
    choosen_files_id = []
    date = None
    error = None
    while console:
        if not date:
            if error:
                print(Fore.RED + error)
                error = None
            print(Fore.GREEN + 'Введите дату для работы(название листа) в формате ДД.ММ.ГГГГ')
            date = input('Ввод:').strip()
            try:
                datetime.strptime(date, "%d.%m.%Y")
                print("Формат корректный, дата:", date)
            except ValueError:
                print(Fore.RED + "Неверный формат даты или несуществующая дата."+ Style.RESET_ALL)
                input(Fore.GREEN + 'Нажмите Enter для продолжения...'+ Style.RESET_ALL)
                date = None
                os.system('cls')
        else:
            print(Fore.GREEN +'Выберите действие:')
            for key, text in zip(TEXT_OUT.keys(), TEXT_OUT.values()):
                print(Fore.GREEN + key + ' - ' + text + Style.RESET_ALL)
            if choosen_files:
                print()
                print(Fore.YELLOW + 'Выбранные файлы:')
                for text in choosen_files:
                    print(Fore.YELLOW + text + Style.RESET_ALL)
            inputter = input(Fore.GREEN + 'Ввод:').strip()
            if inputter == 'all':
                console = False
                return 'all'
            elif inputter in TEXT_OUT.keys():
                if TEXT_OUT[inputter] in choosen_files:
                    print(Fore.RED + 'Этот файл уже выбран')
                    input(Fore.GREEN + 'Нажмите Enter для продолжения...'+ Style.RESET_ALL)
                else:
                    if inputter == '0':
                        console = False
                        return choosen_files_id
                    choosen_files.append(TEXT_OUT[inputter])
                    choosen_files_id.append(inputter)
                    print(Fore.GREEN + f'Выбран файл: {TEXT_OUT[inputter]}')
                    input(Fore.GREEN + 'Нажмите Enter для продолжения...'+ Style.RESET_ALL)
            elif inputter:
                print(Fore.RED + 'Не найден вариант ответа!' + Style.RESET_ALL)
                input(Fore.GREEN + 'Нажмите Enter для продолжения...'+ Style.RESET_ALL)
            else:
                print(Fore.RED + 'Не надо оставлять ответ пустым!' + Style.RESET_ALL)
                input(Fore.GREEN + 'Нажмите Enter для продолжения...'+ Style.RESET_ALL)
            os.system('cls')
            



def funcs_runner(app: xw.App, result):
    """Запуск функций в зависимости от выбора пользователя"""
    if result == 'all':
        logging.info("Запуск всех стадий")
        for func in CHOOSEN_FILES.values():
            func(app)
    else:
        logging.info("Запуск выбранных функций")
        for res in result:
            CHOOSEN_FILES[res](app)



def main():
    try:
        #result = funcs_manager()
        logging.info("Запуск приложения Excel")
        app = xw.App(visible=False)
        # component_suppliers_data(app)
        # electicity_data(app)
        # services_to_us_data(app)
        # daily_buyers_data(app)
        # storage_material_data(app)
        # storage_complect_data(app)
        we_services_and_report_card(app)
    finally:
        logging.info("Закрытие приложения Excel")
        app.quit()
        print('Логи работы находятся в файле "логи_выполнения.log"')
        input("Завершение работы. Нажмите любую клавишу для продолжения...")
    



if __name__ == '__main__':
    main()