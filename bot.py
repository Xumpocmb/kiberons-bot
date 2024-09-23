import time
import pandas as pd
import json
import os
import gspread
import threading
import logging
from tkinter import Tk, Label, Entry, Button, Checkbutton, IntVar, messagebox, filedialog

from openpyxl.xml.constants import WORKSHEET_TYPE
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException


logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

CREDENTIALS_FILE = 'credentials.json'
DEFAULT_GOOGLE_CREDENTIALS_FILE = 'google-credentials.json'
DEFAULT_SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1XU9SzrkQGtd6atfctJc8fNK5nA7MYbvtbQaBmGMCZz0'
DEFAULT_WORKSHEET_NAME = 'кибероны боровки'


class GoogleSheet:
    def __init__(self, spreadsheet_url: str = DEFAULT_SPREADSHEET_URL) -> None:
        """
        Initialize a GoogleSheet object.

        :param spreadsheet_url: The URL of the Google Sheets spreadsheet to connect to.
        :type spreadsheet_url: str
        :return: None
        :rtype: None
        """
        try:
            self.account = gspread.service_account(filename=google_credentials_file_entry.get())
            self.spreadsheet = self.account.open_by_url(spreadsheet_url)
            self.topics = {elem.title: elem.id for elem in self.spreadsheet.worksheets()}
            if worksheet_name_entry.get() not in self.topics:
                raise ValueError(f"Worksheet '{worksheet_name_entry.get()}' not found in spreadsheet")
            self.answers = self.spreadsheet.get_worksheet_by_id(self.topics[worksheet_name_entry.get()])
            logging.info("Успешное подключение к Google Sheets")
        except Exception as e:
            logging.error(f"Ошибка подключения к Google Sheets: {e}")
            raise e

    def load_data_from_google_sheet(self) -> pd.DataFrame:
        """
        Загружает данные из Google Sheets.

        :return: A Pandas DataFrame containing the data from the worksheet.
        :rtype: pd.DataFrame
        """
        try:
            if not self.spreadsheet:
                raise ValueError("Spreadsheet is not initialized")
            if not self.topics:
                raise ValueError("No topics found in the spreadsheet")
            worksheet = self.spreadsheet.worksheet(worksheet_name_entry.get())
            if not worksheet:
                raise ValueError(f"Worksheet '{worksheet_name_entry.get()}' not found in spreadsheet")
            data = worksheet.get_all_records()
            if not data:
                raise ValueError("No data found in the worksheet")
            df = pd.DataFrame(data)
            logging.info("Данные успешно загружены из Google Sheets")
            return df
        except Exception as e:
            logging.error(f"Ошибка загрузки данных из Google Sheets: {e}")
            raise e

    def save_data_to_google_sheet(self, df: pd.DataFrame) -> None:
        """Saves the DataFrame to the specified worksheet in the Google Sheets.

        Args:
            df (pd.DataFrame): The DataFrame to save.

        Returns:
            None
        """
        try:
            if not self.spreadsheet:
                raise ValueError("Spreadsheet is not initialized")
            if not self.topics:
                raise ValueError("No topics found in the spreadsheet")
            worksheet = self.spreadsheet.worksheet(worksheet_name_entry.get())
            if not worksheet:
                raise ValueError(f"Worksheet '{worksheet_name_entry.get()}' not found in spreadsheet")
            worksheet.clear()
            worksheet.update([df.columns.values.tolist()] + df.values.tolist())
            logging.info("Данные успешно сохранены в Google Sheets")
        except Exception as e:
            logging.error(f"Ошибка сохранения данных в Google Sheets: {e}")
            raise e


def choose_google_credentials_file() -> None:
    """Открывает диалог выбора файла для учетных данных Google.

    :return: None
    """
    file_path = filedialog.askopenfilename(title="Выберите файл учетных данных Google",
                                           filetypes=[("JSON files", "*.json")])
    if file_path:
        if google_credentials_file_entry is None:
            raise ValueError("google_credentials_file is None")
        google_credentials_file_entry.delete(0, 'end')
        google_credentials_file_entry.insert(0, file_path)


def load_credentials() -> None:
    """
    Loads credentials from a JSON file.

    Tries to load the credentials from the file specified by `CREDENTIALS_FILE`.
    If the file does not exist, does nothing.

    Args:
        None

    Returns:
        None
    """
    try:
        if os.path.exists(CREDENTIALS_FILE):
            with open(CREDENTIALS_FILE, 'r') as file:
                data: dict = json.load(file)
                if data is None:
                    raise ValueError("Loaded JSON is empty")
                login = data.get("login")
                if login is None:
                    raise ValueError("JSON does not contain 'login' key")
                login_entry.insert(0, login)
                password = data.get("password")
                if password is None:
                    raise ValueError("JSON does not contain 'password' key")
                password_entry.insert(0, password)
                spreadsheet_url = data.get("spreadsheet_url")
                if spreadsheet_url is None:
                    raise ValueError("JSON does not contain 'spreadsheet_url' key")
                spreadsheet_url_entry.insert(0, spreadsheet_url)
                worksheet_name = data.get("worksheet_name")
                if worksheet_name is None:
                    raise ValueError("JSON does not contain 'worksheet_name' key")
                worksheet_name_entry.insert(0, worksheet_name)
                google_credentials_file = data.get("google_credentials_file")
                if google_credentials_file is None:
                    raise ValueError("JSON does not contain 'google_credentials_file' key")
                google_credentials_file_entry.insert(0, google_credentials_file)
                remember = data.get("remember")
                if remember is None:
                    raise ValueError("JSON does not contain 'remember' key")
                remember_var.set(remember)
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
                "spreadsheet_url": spreadsheet_url_entry.get(),
                "worksheet_name": worksheet_name_entry.get(),
                "google_credentials_file": google_credentials_file_entry.get(),
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


def init_driver() -> webdriver.Chrome:
    """Инициализирует и возвращает объект Selenium WebDriver типа webdriver.Chrome.

    Returns:
        webdriver.Chrome: Инициализированный объект WebDriver.
    """
    try:
        if not os.path.exists('chromedriver-win64/chromedriver.exe'):
            logging.error("chromedriver.exe could not be found.")
            raise FileNotFoundError("chromedriver.exe could not be found.")
        service: Service = Service('chromedriver-win64/chromedriver.exe')
        if not service:
            logging.error("Service could not be created.")
            raise RuntimeError("Service could not be created.")
        options: webdriver.ChromeOptions = webdriver.ChromeOptions()
        driver: webdriver.Chrome = webdriver.Chrome(service=service, options=options)
        if not driver:
            logging.error("Driver could not be created.")
            raise RuntimeError("Driver could not be created.")
        logging.info("WebDriver успешно инициализирован")
        return driver
    except FileNotFoundError as e:
        logging.error(f"Ошибка инициализации WebDriver: {e}")
        raise e
    except RuntimeError as e:
        logging.error(f"Ошибка инициализации WebDriver: {e}")
        raise e


def login_to_site(driver: webdriver.Chrome, login: str, password: str) -> bool:
    """
    Выполняет вход на сайт с указанными логином и паролем.

    Args:
        driver (webdriver.Chrome): Инициализированный объект WebDriver.
        login (str): Логин для входа на сайт.
        password (str): Пароль для входа на сайт.

    Returns:
        bool: True, если вход выполнен успешно, False - в противном случае.
    """
    if not driver:
        logging.error("Driver is null")
        messagebox.showerror("Ошибка входа", "Driver is null")
        return False
    if not login or not password:
        logging.error("Login or password is null or empty")
        messagebox.showerror("Ошибка входа", "Login or password is null or empty")
        return False
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


def start_processing() -> None:
    """Основная логика обработки данных."""
    save_credentials()

    try:
        google_sheet = GoogleSheet()

        df = google_sheet.load_data_from_google_sheet()

        if df is None:
            raise ValueError("No data loaded from Google Sheet")

        driver = init_driver()

        login: str = login_entry.get()
        password: str = password_entry.get()

        if not login_to_site(driver, login, password):
            return

        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div/div[1]/div/div[2]/div[1]/a').click()
        time.sleep(2)

        for index, row in df.iterrows():
            if pd.notna(row["фио"]):
                try:
                    kiberones_value: float = float(row["кибероны"])
                    if kiberones_value > 0:
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

        if df is not None:
            google_sheet.save_data_to_google_sheet(df)
            logging.info("Обработка завершена успешно")
            messagebox.showinfo("Завершено", "Обработка завершена успешно.")
    except Exception as e:
        logging.error(f"Ошибка во время обработки: {e}")
        messagebox.showerror("Ошибка", f"Произошла ошибка во время обработки: {e}")
    finally:
        driver.quit()


def start_processing_thread() -> None:
    """Запускает процесс обработки данных в отдельном потоке.

    Args:
        None

    Returns:
        None
    """
    try:
        threading.Thread(target=start_processing).start()
    except Exception as e:
        logging.error("Ошибка при запуске потока обработки: %s", e)


if __name__ == "__main__":
    root = Tk()
    root.title("KIBER Club - Бот для начисления Киберонов")


    def center_window(window):
        window.update_idletasks()  # Обновить размеры
        window_width = window.winfo_width() + 20  # Добавить немного отступа
        window_height = window.winfo_height() + 20
        x = (window.winfo_screenwidth() // 2) - (window_width // 2)
        y = (window.winfo_screenheight() // 2) - (window_height // 2)
        window.geometry(f'{window_width}x{window_height}+{x}+{y}')

    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=1)
    root.grid_rowconfigure(0, weight=1)
    root.grid_rowconfigure(6, weight=1)

    login_label = Label(root, text="Логин:")
    login_label.grid(row=1, column=0, sticky="e", padx=10)
    login_entry = Entry(root, width=20)
    login_entry.grid(row=1, column=1, padx=10)

    password_label = Label(root, text="Пароль:")
    password_label.grid(row=2, column=0, sticky="e", padx=10)
    password_entry = Entry(root, show="*", width=20)
    password_entry.grid(row=2, column=1, padx=10)

    spreadsheet_url_label = Label(root, text="Ссылка на таблицу:")
    spreadsheet_url_label.grid(row=3, column=0, sticky="e", padx=10)
    spreadsheet_url_entry = Entry(root, width=80)
    spreadsheet_url_entry.grid(row=3, column=1, padx=10)

    worksheet_name_label = Label(root, text="Название листа:")
    worksheet_name_label.grid(row=4, column=0, sticky="e", padx=10)
    worksheet_name_entry = Entry(root, width=80)
    worksheet_name_entry.grid(row=4, column=1, padx=10)

    google_credentials_file_label = Label(root, text="Путь к файлу учетных данных:")
    google_credentials_file_label.grid(row=5, column=0, sticky="e", padx=10)
    google_credentials_file_entry = Entry(root, width=80)
    google_credentials_file_entry.grid(row=5, column=1, padx=10)

    choose_file_button = Button(root, text="Выбрать файл", command=choose_google_credentials_file)
    choose_file_button.grid(row=5, column=2, padx=10)

    remember_var = IntVar()
    remember_checkbutton = Checkbutton(root, text="Запомнить", variable=remember_var)
    remember_checkbutton.grid(row=6, column=0, columnspan=2, pady=10)

    start_button = Button(root, text="Начать", command=start_processing_thread)
    start_button.grid(row=7, column=0, columnspan=2, pady=20)

    load_credentials()

    center_window(root)

    root.mainloop()

