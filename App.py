import time
from datetime import datetime, timedelta

import gspread
from selenium import webdriver
from selenium.webdriver.common.by import By
from oauth2client.service_account import ServiceAccountCredentials

from config import get_config

url = get_config()['general'].get("url", "")  # Ссылка для входа на сайт
order_url = get_config()['general'].get("order_url", "")  # Ссылка на первую таблицу
order_else_url = get_config()['general'].get("order_else_url", "")  # Ссылка на вторую таблицу
# Тут указать Путь к Хром Драйверу !!! Хром драйвер по версии должен совпадать с самим хромом скачанным с сайта
# https://googlechromelabs.github.io/chrome-for-testing/
driver_path = "C:\\chromedriver\\chromedriver.exe"
options = webdriver.ChromeOptions() # при желании в ковычки можно вписать опции, на пример чтоб все делалось без открытия браузера
driver = webdriver.Chrome(options=options) # Открывается браузер
driver.get(url)  # Обращается к сайду


# ПОДКЛЮЧЕНИЕ ГУГЛ ТАБЛИЦЫ
credentials_file = 'D:\\PYCHARMPROJECT\\BotParcer\\secret.json'     # Секрет ключ для подключения таблицы
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
client = gspread.authorize(credentials)
sheet_id = '1b0MRV4EaIqt8Sntjt8iGaDoATF7cXsfysQyrDgqq77g'   # ИД Таблицы , берется прямо с ссылки на таблицу
sheet = client.open_by_key(sheet_id).sheet1
range_name = 'test_list' # указываем лист на котором будет работать

try:
    # Найти поле для ввода логина
    login_field = driver.find_element(By.XPATH, '/html/body/div/div/table/tbody/tr[2]/td/form/table/tbody/tr[1]/td[2]/div/input')  # Замените на актуальный селектор
    login_field.send_keys(get_config()['logins'].get("login", ""))  # Замените на ваш логин

    # Найти поле для ввода пароля
    password_field = driver.find_element(By.XPATH, '/html/body/div/div/table/tbody/tr[2]/td/form/table/tbody/tr[2]/td[2]/div/div/input')  # Замените на актуальный селектор
    password_field.send_keys(get_config()['logins'].get("password", ""))  # Замените на ваш пароль

    time.sleep(1)   # это все таймеры ожидания
    # Найти и нажать кнопку для отправки формы
    submit_button = driver.find_element(By.ID, 'enter-button')
    submit_button.click()

    # Ожидаем 5 секунд пока прогрузится страница и переходим на страницу заказов
    time.sleep(5)
    driver.get(order_url)   # подключаемся на сайт с таблицой

    # Находим, фильтруем таблицу, выбираем "Разширенный"
    time.sleep(1)
    choose_ordinary = driver.find_element(By.XPATH, '/html/body/div[1]/section/form/table/tbody/tr[3]/td[2]/div[1]')
    choose_ordinary.click()
    choose_expansive = driver.find_element(By.XPATH, '/html/body/div[1]/section/form/table/tbody/tr[3]/td[2]/div[1]/div/ul/li[2]')
    choose_expansive.click()

    # Устаналиваем сегодняшнюю дату, отнимаем один день, приводит формат даты в соотествие с сайтом
    today = datetime.now().date()
    yesterday_date = today - timedelta(days=1)
    formatted_yesterday_date = yesterday_date.strftime("%Y-%m-%d")

    time.sleep(2)
        # Находит поле для "вчерашней" даты, очищаем его и вставляем дату полученную выше, т.е. на один день раньше от "сегодня"
    choose_date1 = driver.find_element(By.XPATH, '/html/body/div[1]/section/form/table/tbody/tr[3]/td[2]/span/span[1]/input')
    choose_date1.clear()
    choose_date1.send_keys(formatted_yesterday_date)

    time.sleep(2)
    # "Застосувати" кнопку нажимаем
    order_submit_button = driver.find_element(By.XPATH, '/html/body/div[1]/section/form/table/tbody/tr[22]/td[2]/button')
    order_submit_button.click()

    time.sleep(5)

    # Устанавливаем фильтр по 500 штук на одной странице
    choose_show = driver.find_element(By.XPATH, '/html/body/div[1]/section/div[2]/div[3]/select/option[5]')
    choose_show.click()
    # Находим таблицу по ее ИД и устанавливаем значения счетчика на 0
    table = driver.find_element(By.ID, 'table-json-data')
    counter_table = 0
    if table:
        # Находим тэг <tbody> внутри таблицы по классу
        tbody = table.find_element(By.CLASS_NAME, 'data-container')
        # Если тело таблицу найдено начинаем копировать таблицу
        if tbody:
            rows = tbody.find_elements(By.TAG_NAME, 'tr')

            all_data = []
            for row in rows:
                cells = row.find_elements(By.XPATH, './/td | .//th')
                cell_texts = [cell.text.strip() for cell in cells]
                all_data.append(cell_texts)
                counter_table += 1  # Подключил счетчик, чтобы считало количество строк, так как в след таблице их больше
                time.sleep(1)  # В каждом шаге итератора мы делаем задержку, так как может быть ошибка

                # выравниваем таблицу по столбцу А
                sheet.update(range_name='A1', values=all_data)
            # Удаляем не нужные столбцы
    columns_to_delete = [2, 4, 5, 7, 8, 9, 12]  # Индексы колонок в Гугл таблицах считаються НЕ с "0", а с "1"
    for col_index in sorted(columns_to_delete, reverse=True):
        sheet.delete_columns(col_index)
    # else:
    #     driver.quit()

    time.sleep(5)
    driver.get(order_else_url)      #Заходим на следующую таблицу
    time.sleep(5)

    # жмем фильтр "Всі" на заказах
    choose_all_orders = driver.find_element(By.XPATH, '/html/body/div[1]/section/div[2]/ul/li[1]/a')
    choose_all_orders.click()

    time.sleep(5)
    order_table = driver.find_element(By.XPATH, '/html/body/div[1]/section/div[3]/form/table')
    if order_table:
        order_tbody = driver.find_element(By.XPATH, '/html/body/div[1]/section/div[3]/form/table/tbody')
        if order_tbody:
            # Если тело таблицу найдено начинаем копировать таблицу
            if order_tbody:
                rows = order_tbody.find_elements(By.TAG_NAME, 'tr')

                order_all_data = []
                counter_order = 0
                for row in rows:
                    if counter_order <= counter_table - 1:      # Копируем таблицу, пока количество строк не совпадет с предыдущей
                        cells = row.find_elements(By.XPATH, './/td | .//th')
                        cell_texts = [cell.text.strip() for cell in cells]
                        order_all_data.append(cell_texts)
                        counter_order += 1
                        # выравниваем по столбцу А
                        sheet.update(range_name='H1', values=order_all_data)
                        time.sleep(1)

        # Удаляем не нужные столбцы
    columns_to_delete = [9, 10, 11, 12, 13, 15, 16, 18,
                         19, 20, 21, 22, 23, 24, 25, 26, 27,
                         28, 29, 30, 31, 32, 33, 34, 35]  # Индексы колонок в Гугл таблицах считаються НЕ с "0", а с "1"
    for col_index in sorted(columns_to_delete, reverse=True):
        sheet.delete_columns(col_index)

finally:

    driver.quit()




