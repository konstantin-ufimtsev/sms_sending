from selenium.webdriver.common.by import By
from selenium import webdriver
import time
import logging
import openpyxl
import pprint


logging.basicConfig(
    format='%(asctime)s %(levelname)-8s [%(filename)s:%(lineno)d %(message)s',
    datefmt='[%d.%m.%Y-%H:%M%S]',
    filename='log.log',
    filemode='w',
    encoding='utf-8',
    level=logging.INFO
)


def get_data(filename = 'contacts.xlsx') -> list:
    result_data = []
    book = openpyxl.open(filename, read_only=True)
    sheet = book.active
    for row in sheet.iter_rows(2, sheet.max_row):
        row_dict = {}
        if len(str(row[1].value)) != 11:
            print(f"ВНИМАНИЕ: ошибка в записи телефона в строке {row[1], row[1].value}")
            logging.info(f"Ошибка в записи телефона в строке {row[1]}")
        else:
            pass
        dict_data = {"date": row[0].value,
                "phone": row[1].value,
                "name": row[2].value,
                "album_link": row[3].value,
                "promocode": row[4].value,
                "promo_date": row[5].value}
        result_data.append(dict_data)
    book.close()
    print(f"ВНИМАНИЕ! Отправляется активный сохраненный лист - считан лист {sheet.title}")
    print("Считаны строки:")
    for i in result_data:
        print(i, sep='\n')
    while True:
        answer = input("Отправляем СМС по этиому списку ?  (0 - НЕТ, 1 - ДА) :")
        try:
            answer = int(answer)
        except ValueError:
            print("Нужно ввести 0 или 1")
            continue
        if answer == 1:
            logging.info("Список передан на отправку")
            return result_data

        if answer == 0:
            logging.info("Завершение работы программы - пользоваталем отклонен список для рассылки")
            print("Завершение работы программы - пользоваталем отклонен список для рассылки")
            break
        elif print("Нужно ввести число 0 или 1"):
            continue


def send_sms(contact_list:list):

    for item in contact_list:
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--headless")
            url = "http://192.168.1.1/html/smsinbox.html"
            driver = webdriver.Chrome(options=options)
            driver.get(url)
            driver.find_element(By.ID, "span_message").click()
            time.sleep(2)
            driver.find_element(By.ID, "recipients_number").send_keys(item["phone"])
            
            #сообщение с промокодом
            #message = f"Здравствуйте {item['name']}, ваши фото с Арт-вечеринки SALVADOR.ONE {item['date']} по ссылке: {item['album_link']} Дарим вам или вашим друзьям промокод {item['promocode']} на посещение арт вечеринки в {item['promo_date']}"
            
            #сообщение без промокода
            message = f"Здравствуйте {item['name']}, ваши фото с Арт-вечеринки SALVADOR.ONE {item['date']} по ссылке: {item['album_link']} Присодиняйтесь к нам: https://t.me/salvador_kaliningrad и https://vk.com/salvador_kld" 
            
            driver.find_element(By.ID, "message_content").send_keys(message)
            time.sleep(1)
            driver.find_element(By.ID, "pop_send").click()
            print(f"Отправлено сообщение: {message}")
            logging.info(f"Отправлено сообщение: {message}")
            time.sleep(3)
        except Exception as ex:
            logging.info(f"Ошибка {ex}")
        logging.info(f"Отправлено {len(contact_list)} сообщений!")

if __name__ == '__main__':
    send_sms(get_data())
