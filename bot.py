import json
import logging
import os
import threading
import time
from tkinter import Tk, Label, Entry, Button, Checkbutton, IntVar, messagebox, filedialog, StringVar

import gspread
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select

logging.basicConfig(
    level=logging.INFO,
    format='%(levelname)s - %(lineno)d - %(message)s'
)

CREDENTIALS_FILE = 'credentials.json'


class GoogleSheet:
    def __init__(self, google_credentials_file: str, spreadsheet_url: str, worksheet_name: str) -> None:
        """
        Initialize a GoogleSheet object.

        :param google_credentials_file: The path to the Google credentials JSON file.
        :type google_credentials_file: str
        :param spreadsheet_url: The URL of the Google Sheets spreadsheet to connect to.
        :type spreadsheet_url: str
        :param worksheet_name: The name of the worksheet to access.
        :type worksheet_name: str
        :return: None
        :rtype: None
        """
        try:
            self.account = gspread.service_account(filename=google_credentials_file)
            self.spreadsheet = self.account.open_by_url(spreadsheet_url)
            self.topics = {elem.title: elem.id for elem in self.spreadsheet.worksheets()}
            if worksheet_name not in self.topics:
                raise ValueError(f"Worksheet '{worksheet_name}' not found in spreadsheet")
            self.answers = self.spreadsheet.get_worksheet_by_id(self.topics[worksheet_name])
            logging.info("Успешное подключение к Google Sheets")
            update_status("Успешное подключение к Google Sheets")
        except Exception as e:
            logging.error(f"Ошибка подключения к Google Sheets: {e}")
            update_status(f"Ошибка подключения к Google Sheets: {e}")
            raise e

    def load_data_from_google_sheet(self) -> pd.DataFrame:
        """Загружает данные из Google Sheets."""
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
        """Сохраняет DataFrame в указанный лист Google Sheets."""
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
    """Открывает диалог выбора файла для учетных данных Google."""
    file_path = filedialog.askopenfilename(title="Выберите файл учетных данных Google",
                                           filetypes=[("JSON files", "*.json")])
    if file_path:
        google_credentials_file_entry.delete(0, 'end')
        google_credentials_file_entry.insert(0, file_path)


def load_credentials() -> None:
    """Loads credentials from a JSON file."""
    try:
        if os.path.exists(CREDENTIALS_FILE):
            with open(CREDENTIALS_FILE, 'r') as file:
                data: dict = json.load(file)
                login = data.get("login")
                login_entry.insert(0, login)
                password = data.get("password")
                password_entry.insert(0, password)
                spreadsheet_url = data.get("spreadsheet_url")
                spreadsheet_url_entry.insert(0, spreadsheet_url)
                worksheet_name = data.get("worksheet_name")
                worksheet_name_entry.insert(0, worksheet_name)
                google_credentials_file = data.get("google_credentials_file")
                google_credentials_file_entry.insert(0, google_credentials_file)
                remember = data.get("remember")
                remember_var.set(remember)
            logging.info("Учетные данные успешно загружены из JSON")
            update_status("Учетные данные успешно загружены из JSON")
    except Exception as e:
        logging.error(f"Ошибка загрузки учетных данных: {e}")
        update_status(f"Ошибка загрузки учетных данных: {e}")


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
            update_status("Учетные данные успешно сохранены в JSON")
        else:
            if os.path.exists(CREDENTIALS_FILE):
                os.remove(CREDENTIALS_FILE)
                logging.info("Файл учетных данных удален либо не найден")
                update_status("Файл учетных данных удален либо не найден")
    except Exception as e:
        logging.error(f"Ошибка сохранения учетных данных: {e}")
        update_status(f"Ошибка сохранения учетных данных: {e}")


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
        update_status("Успешный вход на сайт")
        return True
    except (NoSuchElementException, TimeoutException) as e:
        logging.error(f"Ошибка входа на сайт: {e}")
        update_status(f"Ошибка входа на сайт: {e}")
        messagebox.showerror("Ошибка входа", f"Не удалось войти на сайт: {e}")
        driver.quit()
        return False


def find_and_open_user(driver, row) -> bool:
    """Функция поиска и открытия профиля пользователя"""
    try:
        search_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH,
                                            '/html/body/div[1]/div/div/div/div/div[2]/div[2]/div/div/div[2]/input'))
        )
        search_field.clear()
        search_field.send_keys(row['фио'])

        user_item = WebDriverWait(driver, 1).until(
            EC.presence_of_element_located((By.XPATH,
                                            '//div[contains(@class, "user_item") and @style="display: table-row;"]'))
        )
        user_item.find_element(By.TAG_NAME, 'a').click()
        return True

    except (NoSuchElementException, TimeoutException):
        update_status("Не удалось найти пользователя или элементы поиска не загрузились")
        logging.error("Не удалось найти пользователя или элементы поиска не загрузились")
        return False


def activity_bonus(driver, row) -> bool:
    """Обрабатывает пользователя в таблице."""
    try:
        if find_and_open_user(driver, row):
            iter_count = row['активность'] // 5
            for _ in range(iter_count):
                apply_bonus(driver, 4)
            logging.info(f"Кибероны успешно начислены для пользователя: {row['фио']}")
            update_status(f"Кибероны успешно начислены для пользователя: {row['фио']}")
            driver.back()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*')))
            driver.refresh()
            return True
        else:
            logging.info(f"Не удалось найти пользователя: {row['фио']}")
            update_status(f"Не удалось найти пользователя: {row['фио']}")
            return False
    except (NoSuchElementException, TimeoutException) as e:
        logging.error(f"Ошибка при обработке пользователя {row['фио']}: {e}")
        return False


def other_bonus(driver, row, index) -> bool:
    """Обрабатывает остальные бонусы"""
    try:
        if find_and_open_user(driver, row):
            apply_bonus(driver, index)
            logging.info(f"Бонусные кибероны успешно начислены для пользователя: {row['фио']}")
            update_status(f"Бонусные кибероны успешно начислены для пользователя: {row['фио']}")
            driver.back()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*')))
            driver.refresh()
            return True
        else:
            logging.info(f"Не удалось найти пользователя: {row['фио']}")
            update_status(f"Не удалось найти пользователя: {row['фио']}")
            return False
    except (NoSuchElementException, TimeoutException) as e:
        logging.error(f"Ошибка при обработке бонуса {row['фио']}: {e}")
        return False


def process_penalty(driver, row) -> bool:
    """Запускает процесс обработки штрафов."""
    try:
        if find_and_open_user(driver, row):
            apply_penalty(driver, row)
            logging.info(f"Штраф успешно начислены для пользователя: {row['фио']}")
            update_status(f"Штраф успешно начислены для пользователя: {row['фио']}")
            driver.back()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*')))
            driver.refresh()
            return True
        else:
            logging.info(f"Не удалось найти пользователя: {row['фио']}")
            update_status(f"Не удалось найти пользователя: {row['фио']}")
            return False
    except (NoSuchElementException, TimeoutException) as e:
        logging.error(f"Ошибка при обработке пользователя {row['фио']}: {e}")
        return False


def apply_bonus(driver, index) -> bool:
    try:
        button_change_kiberons = driver.find_element(By.XPATH,
                                                     '/html/body/div[1]/div/div/div/div/div[2]/div[2]/div/div/div[1]/div[1]/span/span')
        button_change_kiberons.click()

        select1 = Select(
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "fc_field_sign_id"))))
        select1.select_by_visible_text("Начисление")

        select2 = Select(
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "fc_field_cause_id"))))
        select2.select_by_index(index)

        save_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "sendsave")))
        save_button.click()

        close_modal_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "uss_modal_close")))
        close_modal_element.click()

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*')))
        return True
    except (NoSuchElementException, TimeoutException) as e:
        logging.error(f"Ошибка при начислении бонуса: {e}")
        return False


def apply_penalty(driver, row) -> bool:
    """Запускает процесс обработки штрафов."""
    try:
        button_change_kiberons = driver.find_element(By.XPATH,
                                                     '/html/body/div[1]/div/div/div/div/div[2]/div[2]/div/div/div[1]/div[1]/span/span')
        button_change_kiberons.click()
        select1 = Select(
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "fc_field_sign_id"))))
        select1.select_by_visible_text("Списание")
        field_comment = driver.find_element(By.ID, "fc_field_comment_id")
        field_comment.clear()
        field_comment.send_keys("Замечания по поведению")
        field_amount = driver.find_element(By.ID, "fc_field_amount_id")
        field_amount.clear()
        field_amount.send_keys(int(row['штраф']))
        save_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "sendsave")))
        save_button.click()
        return True
    except (NoSuchElementException, TimeoutException) as e:
        logging.error(f"Ошибка при взыскании штрафа: {e}")
        return False


def start_processing() -> None:
    """Основная логика обработки данных."""
    save_credentials()
    driver: WebDriver | None = None
    try:
        update_status("Начинается обработка данных...")
        google_sheet = GoogleSheet(google_credentials_file, spreadsheet_url, worksheet_name)

        df = google_sheet.load_data_from_google_sheet()

        if df is None:
            raise ValueError("No data loaded from Google Sheet")

        driver: WebDriver = init_driver()

        login: str = login_entry.get()
        password: str = password_entry.get()

        if not login_to_site(driver, login, password):
            return

        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div/div[1]/div/div[2]/div[1]/a').click()
        time.sleep(3)

        for index, row in df.iterrows():
            if pd.notna(row["фио"]):
                if pd.notna(row["активность"]):
                    try:
                        kiberones_value: float = float(row["активность"])
                        if kiberones_value > 0:
                            logging.info(f"Начинается начисление киберонов для пользователя: {row['фио']}")
                            update_status(f"Начинается начисление киберонов для пользователя: {row['фио']}")
                            if activity_bonus(driver, row):
                                df.at[index, "активность"] = None
                                google_sheet.save_data_to_google_sheet(df)
                            else:
                                logging.warning(f"Не удалось обработать пользователя: {row['фио']}")
                                update_status(f"Не удалось обработать пользователя: {row['фио']}")
                        else:
                            logging.info(f"Пропуск пользователя {row['фио']} с нулевыми или отрицательными киберонами.")
                    except ValueError:
                        logging.warning(
                            f"Неверное значение киберонов для пользователя {row['фио']}: {row['активность']}")
                        update_status(f"Неверное значение киберонов для пользователя {row['фио']}: {row['активность']}")

                if pd.notna(row["штраф"]):
                    try:
                        penalty_value: float = float(row["штраф"])
                        if penalty_value > 0:
                            logging.info(f"Начинается начисление штрафа для пользователя: {row['фио']}")
                            update_status(f"Начинается начисление штрафа для пользователя: {row['фио']}")
                            if process_penalty(driver, row):
                                df.at[index, "штраф"] = None
                                google_sheet.save_data_to_google_sheet(df)
                            else:
                                logging.warning(f"Не удалось штраф обработать пользователя: {row['фио']}")
                                update_status(f"Не удалось штраф обработать пользователя: {row['фио']}")
                        else:
                            logging.info(f"Пропуск пользователя {row['фио']} с нулевыми или отрицательными штрафами.")
                    except ValueError:
                        logging.warning(f"Неверное значение штрафа для пользователя {row['фио']}: {row['штраф']}")
                else:
                    logging.info(f"Пропущена строка: штраф отсутствует (штраф: {row.get('штраф', 'пусто')})")
                    update_status(f"Пропущена строка: штраф отсутствует (штраф: {row.get('штраф', 'пусто')})")

                if pd.notna(row["дз"]):
                    try:
                        homework_value: float = float(row["дз"])
                        if homework_value > 0:
                            logging.info(f"Начинается начисление за ДЗ для пользователя: {row['фио']}")
                            update_status(f"Начинается начисление за ДЗ для пользователя: {row['фио']}")
                            if other_bonus(driver, row, 5):
                                df.at[index, "дз"] = None
                                google_sheet.save_data_to_google_sheet(df)
                            else:
                                logging.warning(f"Не удалось обработать начисление за ДЗ пользователя: {row['фио']}")
                                update_status(f"Не удалось обработать начисление за ДЗ пользователя: {row['фио']}")
                        else:
                            logging.info(f"Пропуск пользователя {row['фио']} с нулевыми или отрицательными ДЗ.")
                    except ValueError:
                        logging.warning(f"Неверное значение ДЗ для пользователя {row['фио']}: {row['дз']}")
                else:
                    logging.info(f"Пропущена строка: ДЗ отсутствует (ДЗ: {row.get('дз', 'пусто')})")
                    update_status(f"Пропущена строка: ДЗ отсутствует (ДЗ: {row.get('дз', 'пусто')})")

                if pd.notna(row["др"]):
                    try:
                        homework_value: float = float(row["др"])
                        if homework_value > 0:
                            logging.info(f"Начинается начисление за ДР для пользователя: {row['фио']}")
                            update_status(f"Начинается начисление за ДР для пользователя: {row['фио']}")
                            if other_bonus(driver, row, 6):
                                df.at[index, "др"] = None
                                google_sheet.save_data_to_google_sheet(df)
                            else:
                                logging.warning(f"Не удалось обработать начисление за ДР пользователя: {row['фио']}")
                                update_status(f"Не удалось обработать начисление за ДР пользователя: {row['фио']}")
                        else:
                            logging.info(f"Пропуск пользователя {row['фио']} с нулевыми или отрицательными ДЗ.")
                    except ValueError:
                        logging.warning(f"Неверное значение ДР для пользователя {row['фио']}: {row['др']}")
                else:
                    logging.info(f"Пропущена строка: ДР отсутствует (ДР: {row.get('др', 'пусто')})")
                    update_status(f"Пропущена строка: ДР отсутствует (ДР: {row.get('др', 'пусто')})")

                if pd.notna(row["бонус пропуск"]):
                    try:
                        no_skip_value: str = str(row["бонус пропуск"])
                        if no_skip_value == "да":
                            logging.info(
                                f"Начинается начисление бонуса за модуль без пропуска для пользователя: {row['фио']}")
                            update_status(
                                f"Начинается начисление бонуса за модуль без пропуска для пользователя: {row['фио']}")
                            if other_bonus(driver, row, 7):
                                df.at[index, "бонус пропуск"] = None
                                google_sheet.save_data_to_google_sheet(df)
                            else:
                                logging.warning(f"Не удалось обработать начисление бонуса за модуль без пропуска пользователя: {row['фио']}")
                                update_status(f"Не удалось обработать начисление бонуса за модуль без пропуска пользователя: {row['фио']}")
                        else:
                            logging.info(f"Пропуск пользователя {row['фио']} с нулевыми или отрицательными бонусами.")
                    except ValueError:
                        logging.warning(
                            f"Неверное значение бонуса для пользователя {row['фио']}: {row['бонус пропуск']}")
                else:
                    logging.info(f"Пропущена строка: бонус отсутствует (бонус: {row.get('бонус пропуск', 'пусто')})")
                    update_status(f"Пропущена строка: бонус отсутствует (бонус: {row.get('бонус пропуск', 'пусто')})")

                if pd.notna(row["бонус поведение"]):
                    try:
                        no_penalty_value: str = str(row["бонус поведение"])
                        if no_penalty_value == "да":
                            logging.info(
                                f"Начинается начисление бонуса за модуль без замечаний по поведению для пользователя: {row['фио']}")
                            update_status(
                                f"Начинается начисление бонуса за модуль без замечаний по поведению для пользователя: {row['фио']}")
                            if other_bonus(driver, row, 8):
                                df.at[index, "бонус поведение"] = None
                                google_sheet.save_data_to_google_sheet(df)
                            else:
                                logging.warning(f"Не удалось обработать начисление бонуса за модуль без замечаний по поведению пользователя: {row['фио']}")
                                update_status(f"Не удалось обработать начисление бонуса за модуль без замечаний по поведению пользователя: {row['фио']}")
                        else:
                            logging.info(f"Пропуск пользователя {row['фио']} с нулевыми или отрицательными бонусами.")
                    except ValueError:
                        logging.warning(
                            f"Неверное значение бонуса для пользователя {row['фио']}: {row['бонус поведение']}")
                else:
                    logging.info(f"Пропущена строка: бонус отсутствует (бонус: {row.get('бонус поведение', 'пусто')})")
                    update_status(f"Пропущена строка: бонус отсутствует (бонус: {row.get('бонус поведение', 'пусто')})")
            else:
                logging.info(f"Пропущена строка: ФИО отсутствует (ФИО: {row.get('фио', 'пусто')})")
                update_status(f"Пропущена строка: ФИО отсутствует (ФИО: {row.get('фио', 'пусто')})")

        if df is not None:
            google_sheet.save_data_to_google_sheet(df)
            logging.info("Обработка завершена успешно")
            update_status("Обработка завершена успешно")
            messagebox.showinfo("Завершено", "Обработка завершена успешно.")
    except Exception as e:
        logging.error(f"Ошибка во время обработки: {e}")
        update_status(f"Ошибка во время обработки: {e}")
        messagebox.showerror("Ошибка", f"Произошла ошибка во время обработки: {e}")
    finally:
        if driver is not None:
            driver.quit()


def start_processing_thread() -> None:
    """Запускает процесс обработки данных в отдельном потоке.

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

    status_message = StringVar()


    def update_status(message: str) -> None:
        status_message.set(message)


    def center_window(window: Tk) -> None:
        """
        Центрирует окно на экране.

        :param window: Окно, которое нужно центрировать
        :type window: Tk
        :return: None
        """
        if window is None:
            raise ValueError("Window is None")

        window.update_idletasks()  # Обновить размеры
        window_width = window.winfo_width() + 20  # Добавить немного отступа
        window_height = window.winfo_height() + 20
        x = (window.winfo_screenwidth() // 2) - (window_width // 2)
        y = (window.winfo_screenheight() // 2) - (window_height // 2)

        if window_width <= 0 or window_height <= 0:
            raise ValueError("Window size is invalid")

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

    status_label = Label(root, textvariable=status_message, relief="sunken", anchor="w")
    status_label.grid(row=8, column=0, columnspan=3, sticky="ew")

    load_credentials()

    google_credentials_file = google_credentials_file_entry.get()
    spreadsheet_url = spreadsheet_url_entry.get()
    worksheet_name = worksheet_name_entry.get()

    center_window(root)

    root.mainloop()
