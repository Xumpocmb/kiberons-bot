import time
import pandas as pd
import json
import os
import gspread
import threading
import logging
from tkinter import Tk, Label, Entry, Button, Checkbutton, IntVar, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException


logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

CREDENTIALS_FILE = 'credentials.json'
GOOGLE_CREDENTIALS_FILE = 'endless-gasket-436505-j1-7b21e7342be7.json'
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1XU9SzrkQGtd6atfctJc8fNK5nA7MYbvtbQaBmGMCZz0'


class GoogleSheet:
    def __init__(self, spreadsheet_url=SPREADSHEET_URL):
        try:
            self.account = gspread.service_account(filename=GOOGLE_CREDENTIALS_FILE)
            self.spreadsheet = self.account.open_by_url(spreadsheet_url)
            self.topics = {elem.title: elem.id for elem in self.spreadsheet.worksheets()}
            self.answers = self.spreadsheet.get_worksheet_by_id(self.topics.get("кибероны боровки"))
            logging.info("Успешное подключение к Google Sheets")
        except Exception as e:
            logging.error(f"Ошибка подключения к Google Sheets: {e}")
            raise e

    def load_data_from_google_sheet(self):
        """Загружает данные из Google Sheets."""
        try:
            worksheet = self.spreadsheet.worksheet("кибероны боровки")
            data = worksheet.get_all_records()
            df = pd.DataFrame(data)
            logging.info("Данные успешно загружены из Google Sheets")
            return df
        except Exception as e:
            logging.error(f"Ошибка загрузки данных из Google Sheets: {e}")
            raise e

    def save_data_to_google_sheet(self, df):
        """Сохраняет изменения обратно в Google Sheets."""
        try:
            worksheet = self.spreadsheet.worksheet("кибероны боровки")
            worksheet.clear()
            worksheet.update([df.columns.values.tolist()] + df.values.tolist())
            logging.info("Данные успешно сохранены в Google Sheets")
        except Exception as e:
            logging.error(f"Ошибка сохранения данных в Google Sheets: {e}")
            raise e


def load_credentials():
    """Загружает учетные данные из файла JSON."""
    try:
        if os.path.exists(CREDENTIALS_FILE):
            with open(CREDENTIALS_FILE, 'r') as file:
                data = json.load(file)
                login_entry.insert(0, data.get("login", ""))
                password_entry.insert(0, data.get("password", ""))
                remember_var.set(data.get("remember", 0))
            logging.info("Учетные данные успешно загружены из JSON")
    except Exception as e:
        logging.error(f"Ошибка загрузки учетных данных: {e}")


def save_credentials():
    """Сохраняет учетные данные в файл JSON."""
    try:
        if remember_var.get() == 1:
            credentials = {
                "login": login_entry.get(),
                "password": password_entry.get(),
                "remember": remember_var.get()
            }
            with open(CREDENTIALS_FILE, 'w') as file:
                json.dump(credentials, file)
            logging.info("Учетные данные успешно сохранены в JSON")
        else:
            if os.path.exists(CREDENTIALS_FILE):
                os.remove(CREDENTIALS_FILE)
                logging.info("Файл учетных данных удален")
    except Exception as e:
        logging.error(f"Ошибка сохранения учетных данных: {e}")


def init_driver():
    """Инициализирует и возвращает объект Selenium WebDriver."""
    try:
        service = Service('chromedriver-win64/chromedriver.exe')
        options = webdriver.ChromeOptions()
        logging.info("WebDriver успешно инициализирован")
        return webdriver.Chrome(service=service, options=options)
    except Exception as e:
        logging.error(f"Ошибка инициализации WebDriver: {e}")
        raise e


def login_to_site(driver, login, password):
    """Выполняет вход на сайт с указанными логином и паролем."""
    try:
        driver.get('https://kiber-one.club/')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'login')))
        driver.find_element(By.NAME, 'login').send_keys(login)
        driver.find_element(By.NAME, 'password').send_keys(password)
        driver.find_element(By.XPATH, '//*[@id="loginForm"]/table/tbody/tr[4]/td/input').click()
        WebDriverWait(driver, 10).until(EC.url_changes('https://kiber-one.club/'))
        logging.info("Успешный вход на сайт")
        return True
    except (NoSuchElementException, TimeoutException) as e:
        logging.error(f"Ошибка входа на сайт: {e}")
        messagebox.showerror("Ошибка входа", f"Не удалось войти на сайт: {e}")
        driver.quit()
        return False


def process_user(driver, row):
    """Обрабатывает каждого пользователя в таблице."""
    try:
        search_field = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div/div[2]/div[2]/div/div/div[2]/input')
        search_field.clear()
        search_field.send_keys(row['фио'])
        time.sleep(3)

        user_item = driver.find_element(By.XPATH, '//div[contains(@class, "user_item") and @style="display: table-row;"]')
        user_item.find_element(By.TAG_NAME, 'a').click()

        time.sleep(2)
        iter_count = int(row['кибероны']) // 5

        for _ in range(iter_count):
            button_change_kiberons = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div/div[2]/div[2]/div/div/div[1]/div[1]/span/span')
            button_change_kiberons.click()

            select1 = Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "fc_field_sign_id"))))
            select1.select_by_visible_text("Начисление")

            select2 = Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "fc_field_cause_id"))))
            select2.select_by_index(4)

            save_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "sendsave")))
            save_button.click()

            close_modal_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "uss_modal_close")))
            close_modal_element.click()

            time.sleep(2)

        logging.info(f"Кибероны успешно начислены для пользователя: {row['фио']}")

        row["кибероны"] = None
        driver.back()
        time.sleep(3)
        driver.refresh()

        return True
    except (NoSuchElementException, TimeoutException) as e:
        logging.error(f"Ошибка при обработке пользователя {row['фио']}: {e}")
        return False


def start_processing():
    """Основная логика обработки данных."""
    save_credentials()

    try:
        google_sheet = GoogleSheet()

        df = google_sheet.load_data_from_google_sheet()

        driver = init_driver()

        login = login_entry.get()
        password = password_entry.get()

        if not login_to_site(driver, login, password):
            return

        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div/div[1]/div/div[2]/div[1]/a').click()
        time.sleep(2)

        for index, row in df.iterrows():
            if pd.notna(row["фио"]):
                try:
                    kiberony_value = float(row["кибероны"])
                    if kiberony_value > 0:
                        logging.info(f"Начинается обработка для пользователя: {row['фио']}")
                        if process_user(driver, row):
                            df.at[index, "кибероны"] = None
                        else:
                            logging.warning(f"Не удалось обработать пользователя: {row['фио']}")
                    else:
                        logging.info(f"Пропуск пользователя {row['фио']} с нулевыми или отрицательными киберонами.")
                except ValueError:
                    logging.warning(f"Неверное значение киберонов для пользователя {row['фио']}: {row['кибероны']}")
            else:
                logging.info(f"Пропущена строка: ФИО отсутствует (ФИО: {row.get('фио', 'пусто')})")

        google_sheet.save_data_to_google_sheet(df)
        logging.info("Обработка завершена успешно")
        messagebox.showinfo("Завершено", "Обработка завершена успешно.")
    except Exception as e:
        logging.error(f"Ошибка во время обработки: {e}")
        messagebox.showerror("Ошибка", f"Произошла ошибка во время обработки: {e}")
    finally:
        driver.quit()




def start_processing_thread():
    """Запускает процесс обработки данных в отдельном потоке."""
    threading.Thread(target=start_processing).start()


if __name__ == "__main__":
    root = Tk()
    root.title("KIBER Club - Бот для начисления Киберонов")

    def center_window(window, width=400, height=300):
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        window.geometry(f'{width}x{height}+{x}+{y}')

    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=1)
    root.grid_rowconfigure(0, weight=1)
    root.grid_rowconfigure(5, weight=1)

    login_label = Label(root, text="Логин:")
    login_label.grid(row=1, column=0, sticky="e", padx=10)
    login_entry = Entry(root)
    login_entry.grid(row=1, column=1, padx=10)

    password_label = Label(root, text="Пароль:")
    password_label.grid(row=2, column=0, sticky="e", padx=10)
    password_entry = Entry(root, show="*")
    password_entry.grid(row=2, column=1, padx=10)

    remember_var = IntVar()
    remember_checkbutton = Checkbutton(root, text="Запомнить", variable=remember_var)
    remember_checkbutton.grid(row=3, column=0, columnspan=2, pady=10)

    start_button = Button(root, text="Начать", command=start_processing_thread)
    start_button.grid(row=4, column=0, columnspan=2, pady=20)

    load_credentials()

    root.update_idletasks()
    center_window(root, width=400, height=300)

    root.mainloop()
