from organize import get_desktop_path, create_paths
from edit_files import prepare_file
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.firefox.options import Options


def get_data():
    df = pd.read_excel('files/users.xlsx')
    clients = df['CLIENT'].to_list()
    users = df['CUIT'].to_list()
    passwords = df['PASSWORD'].to_list()
    mode = df['MODE'].to_list()  
    return clients, users, passwords, mode

def set_driver(download_path=None):
    service = Service("C:\\Program Files (x86)\\Development\\geckodriver.exe")
    options = Options()
    options.set_preference("browser.download.folderList", 2)
    options.set_preference("browser.download.dir", download_path)
    options.set_preference("browser.helperApps.neverAsk.openFile", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    options.set_preference("pdfjs.disabled", True)
    driver = webdriver.Firefox(service=service, options=options)
    return driver

def wait(driver, method, route):
    if method == 'XPATH':
        WebDriverWait(driver, 20).until(lambda d: d.find_element(By.XPATH, route))
    elif method == 'CSS_SELECTOR':
        WebDriverWait(driver, 20).until(lambda d: d.find_element(By.CSS_SELECTOR, route))

def switch_tab(driver):
    driver.switch_to.window(driver.window_handles[-1])

def start_download():
    clients, users, passwords, mode = get_data()
    n = 0
    bills_path = (get_desktop_path() + '\\Facturas')
    for n in range(len(users)):
        download_path = bills_path + f'\\{clients[n]}'
        driver = set_driver(download_path)
        driver.get("https://auth.afip.gob.ar/contribuyente_/login.xhtml")

        print(f'Procesing {clients[n]}, please wait...')

        wait(driver, 'XPATH', '//*[@id="F1:username"]')
        cuit = driver.find_element(By.XPATH, '//*[@id="F1:username"]')
        cuit.clear()
        cuit.send_keys(users[n])

        wait(driver, 'XPATH', '//*[@id="F1:btnSiguiente"]')
        btn_next = driver.find_element(By.XPATH,'//*[@id="F1:btnSiguiente"]')
        btn_next.click()

        wait(driver, 'XPATH', '//*[@id="F1:password"]')
        password = driver.find_element(By.XPATH, '//*[@id="F1:password"]')
        password.clear()
        password.send_keys(passwords[0])

        wait(driver, 'XPATH', '//*[@id="F1:btnIngresar"]')
        btn_login = driver.find_element(By.XPATH,'//*[@id="F1:btnIngresar"]')
        btn_login.click()

        wait(driver, 'XPATH', '//*[text() = "Mis Comprobantes"]')
        mis_comprobantes = driver.find_element(By.XPATH, '//*[text() = "Mis Comprobantes"]')
        mis_comprobantes.click()

        time.sleep(2)

        switch_tab(driver)

        if mode[n].capitalize() == 'Emitidos':
            wait(driver, 'XPATH', '//*[text() = "Emitidos"]')
            bills = driver.find_element(By.XPATH, '//*[text() = "Emitidos"]')
        elif mode[n].capitalize() == 'Recibidos':
            wait(driver, 'XPATH', '//*[text() = "Recibidos"]')
            bills = driver.find_element(By.XPATH, '//*[text() = "Recibidos"]')
        bills.click()


        wait(driver, 'XPATH', '//*[@id="fechaEmision"]')
        select_period = driver.find_element(By.XPATH, '//*[@id="fechaEmision"]')
        select_period.click()

        wait(driver, 'XPATH', '//*[text() = "Mes Pasado"]')
        period = driver.find_element(By.XPATH, '//*[text() = "Mes Pasado"]')
        period.click()

        wait(driver, 'XPATH', '//*[@id="buscarComprobantes"]')
        search = driver.find_element(By.XPATH, '//*[@id="buscarComprobantes"]')
        search.click()

        time.sleep(5)

        wait(driver, 'XPATH', '/html/body/main/div/section/div[1]/div/div[2]/div[2]/div[2]/div[1]/div[1]/div/button[2]')
        export_excel = driver.find_element(By.XPATH, '/html/body/main/div/section/div[1]/div/div[2]/div[2]/div[2]/div[1]/div[1]/div/button[2]')
        export_excel.click()

        time.sleep(2)

        n =+ 1

        driver.quit()

        prepare_file(download_path)


if __name__ == '__main__':
    create_paths()
    start_download()