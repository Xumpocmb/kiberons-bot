import time
import pandas as pd
import json
import os
import threading
from tkinter import Tk, Label, Entry, Button, Checkbutton, IntVar, filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

CREDENTIALS_FILE = 'credentials.json'


def load_credentials():
    """Загружает учетные данные из файла JSON."""
    if os.path.exists(CREDENTIALS_FILE):
        with open(CREDENTIALS_FILE, 'r') as file:
            data = json.load(file)
            login_entry.insert(0, data.get("login", ""))
            password_entry.insert(0, data.get("password", ""))
            remember_var.set(data.get("remember", 0))


def save_credentials():
    """Сохраняет учетные данные в файл JSON."""
    if remember_var.get() == 1:
        credentials = {
            "login": login_entry.get(),
            "password": password_entry.get(),
            "remember": remember_var.get()
        }
        with open(CREDENTIALS_FILE, 'w') as file:
            json.dump(credentials, file)
    else:
        if os.path.exists(CREDENTIALS_FILE):
            os.remove(CREDENTIALS_FILE)


def select_file():
    """Открывает диалог выбора файла и вставляет его путь в текстовое поле."""
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    file_entry.delete(0, 'end')
    file_entry.insert(0, file_path)


def start_processing_thread():
    """Запускает процесс обработки данных в отдельном потоке."""
    threading.Thread(target=start_processing).start()


def init_driver():
    """Инициализирует и возвращает объект Selenium WebDriver."""
    service = Service('chromedriver-win64/chromedriver.exe')
    options = webdriver.ChromeOptions()
    return webdriver.Chrome(service=service, options=options)


def login_to_site(driver, login, password):
    """Выполняет вход на сайт с указанными логином и паролем."""
    try:
        driver.get('https://kiber-one.club/')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'login')))
        driver.find_element(By.NAME, 'login').send_keys(login)
        driver.find_element(By.NAME, 'password').send_keys(password)
        driver.find_element(By.XPATH, '//*[@id="loginForm"]/table/tbody/tr[4]/td/input').click()
        WebDriverWait(driver, 10).until(EC.url_changes('https://kiber-one.club/'))
    except (NoSuchElementException, TimeoutException) as e:
        messagebox.showerror("Ошибка входа", f"Не удалось войти на сайт: {e}")
        driver.quit()
        return False
    return True


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

            time.sleep(1)

        row["кибероны"] = None
        driver.back()
        time.sleep(3)
        driver.refresh()

    except (NoSuchElementException, TimeoutException):
        print(f"Пользователь {row['фио']} не найден или произошла ошибка. Пропускаю.")
        return False

    return True


def start_processing():
    """Основная логика обработки данных."""
    save_credentials()
    file_path = file_entry.get()
    if not file_path:
        messagebox.showerror("Ошибка", "Выберите файл.")
        return

    login = login_entry.get()
    password = password_entry.get()

    df = pd.read_excel(file_path)
    driver = init_driver()

    if not login_to_site(driver, login, password):
        return

    try:
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div/div[1]/div/div[2]/div[1]/a').click()
        time.sleep(2)

        for index, row in df.iterrows():
            if pd.notna(row["фио"]) and pd.notna(row["кибероны"]):
                if process_user(driver, row):
                    df.at[index, "кибероны"] = None

        df.to_excel(file_path, index=False)
        messagebox.showinfo("Завершено", "Обработка завершена успешно.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка во время обработки: {e}")
    finally:
        driver.quit()


if __name__ == "__main__":
    root = Tk()
    root.title("KIBER Club - Бот для начисления Киберонов")

    file_label = Label(root, text="Выберите файл с данными:")
    file_label.grid(row=0, column=0)
    file_entry = Entry(root, width=40)
    file_entry.grid(row=0, column=1)
    file_button = Button(root, text="Выбрать файл", command=select_file)
    file_button.grid(row=0, column=2)

    login_label = Label(root, text="Логин:")
    login_label.grid(row=1, column=0)
    login_entry = Entry(root)
    login_entry.grid(row=1, column=1)

    password_label = Label(root, text="Пароль:")
    password_label.grid(row=2, column=0)
    password_entry = Entry(root, show="*")
    password_entry.grid(row=2, column=1)

    remember_var = IntVar()
    remember_checkbutton = Checkbutton(root, text="Запомнить", variable=remember_var)
    remember_checkbutton.grid(row=3, column=1)

    start_button = Button(root, text="Начать", command=start_processing_thread)
    start_button.grid(row=4, column=1, pady=20)

    load_credentials()

    root.mainloop()
