from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from selenium.webdriver.common.by import By
import pandas as pd
import os
from google.oauth2 import service_account
import gspread
from gspread_dataframe import set_with_dataframe
import sys

def obtener_datos_anuales(start_year, end_year):
    # Configuración inicial del WebDriver y opciones de Chrome
    perfil_directorio = "./gastosscrapper"
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument(f"--user-data-dir={perfil_directorio}")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

    # Crear el driver de Chrome
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    all_data_niveles = []
    all_data_sectores = []

    for year in range(start_year, end_year + 1):
        data_niveles = []
        data_sectores = []

        url = f"https://apps5.mineco.gob.pe/transparencia/Navegador/Navegar_7.aspx?y={year}"
        driver.get(url)
        sleep(4)

        select_year_query = '#ctl00_CPH1_DrpYear'
        select_year = driver.find_element(By.CSS_SELECTOR, select_year_query)
        value_year = select_year.get_attribute("value")

        print(f"El año seleccionado es: {value_year}")
        sleep(4)
        nivel_bg = '#ctl00_CPH1_BtnTipoGobierno'
        driver.find_element(By.CSS_SELECTOR, nivel_bg).click()
        sleep(2)

        div = 'ctl00_CPH1_UpdatePanel1'
        tableDom = driver.find_element(By.CSS_SELECTOR, f"#{div} table:nth-child(2)")
        rows = tableDom.find_elements(By.TAG_NAME, "tr")
        columns = ["ID", "NOMBRE NIVEL","PIA","PIM","CERTIFICACION","COMPROMISO ANUAL","EJE. ATENCION COMP ANUAL","EJE. DEVENGADO", "EJE. GIRADO","AVANCE"]

        for row in rows:
            row_temp = []
            cells = row.find_elements(By.TAG_NAME, "td")
            for cell in cells:
                if len(row_temp) == 0:
                    row_temp.append(cell.find_element(By.TAG_NAME, "input").get_attribute("value"))
                else:
                    row_temp.append(cell.text)
            row_temp.append(value_year)
            data_niveles.append(row_temp)

        all_data_niveles.extend(data_niveles)

        # Obtener datos de la primera tabla
        btn_sector_query = f'input[value="{data_niveles[0][0]}"]'
        print(f"Seleccionando el sector: {data_niveles[0]}")
        btn_sector = driver.find_element(By.CSS_SELECTOR, btn_sector_query)
        btn_sector.click()
        btn_sector_action_query = 'ctl00$CPH1$BtnSector'
        btn_sector_action = driver.find_element(By.NAME, btn_sector_action_query)
        btn_sector_action.click()
        sleep(2)

        div_container = '#PnlData'
        tableDom = driver.find_element(By.CSS_SELECTOR, f"{div_container} table:nth-child(3)")
        rows = tableDom.find_elements(By.TAG_NAME, "tr")
        columns = ["ID", "NOMBRE SECTOR","PIA","PIM","CERTIFICACION","COMPROMISO ANUAL","EJE. ATENCION COMP ANUAL","EJE. DEVENGADO", "EJE. GIRADO","AVANCE"]

        for row in rows:
            row_temp = []
            cells = row.find_elements(By.TAG_NAME, "td")
            for cell in cells:
                if len(row_temp) == 0:
                    row_temp.append(cell.find_element(By.TAG_NAME, "input").get_attribute("value"))
                else:
                    row_temp.append(cell.text)
            row_temp.append(value_year)
            data_sectores.append(row_temp)

        all_data_sectores.extend(data_sectores)

    driver.quit()

    # Crear DataFrames consolidados
    columns.append("AÑO")
    df_niveles = pd.DataFrame(all_data_niveles, columns=columns)
    df_sectores = pd.DataFrame(all_data_sectores, columns=columns)

    return df_niveles, df_sectores

def actualizar_google_sheets(df_niveles, df_sectores, spreadsheet_id):
    # Google Autenticación
    SERVICE_ACCOUNT_FILE = "/home/danmer/Desktop/python/pruebadriveapi-273722-2a88631c98cb.json"
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    gc = gspread.authorize(credentials)

    spreadsheet = gc.open_by_key(spreadsheet_id)

    try:
        worksheet_niveles = spreadsheet.add_worksheet(title='Niveles', rows="100", cols="20")
    except gspread.exceptions.APIError:
        worksheet_niveles = spreadsheet.worksheet('Niveles')
        worksheet_niveles.clear()
    set_with_dataframe(worksheet_niveles, df_niveles)

    try:
        worksheet_sectores = spreadsheet.add_worksheet(title='Sectores', rows="100", cols="20")
    except gspread.exceptions.APIError:
        worksheet_sectores = spreadsheet.worksheet('Sectores')
        worksheet_sectores.clear()
    set_with_dataframe(worksheet_sectores, df_sectores)

    print(f'Archivo de Google Sheets actualizado: {spreadsheet.url}')

# Parámetros de entrada
start_year = 2020
end_year = 2023
spreadsheet_id = "12kKSvnVUwch0tRUVzZ3ZmHh_pnFm-krigLn8BI_pEYA"

# Obtener y actualizar datos
df_niveles, df_sectores = obtener_datos_anuales(start_year, end_year)
actualizar_google_sheets(df_niveles, df_sectores, spreadsheet_id)
