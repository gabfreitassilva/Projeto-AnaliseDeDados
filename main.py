from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from time import sleep
from datetime import datetime


with open("credentials.txt", "r") as file:
    user = file.readline().strip()
    password = file.readline().strip()

service = Service(executable_path="geckodriver.exe")
driver = webdriver.Firefox(service=service)

driver.get("http://simg.metrofor.ce.gov.br")

# AUTOMATIZAÇÃO DE LOGIN
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "usuario"))
)

input_element = driver.find_element(By.ID, "usuario")
input_element.send_keys(user)

input_element = driver.find_element(By.ID, "senha")
input_element.send_keys(password + Keys.ENTER)
#############################################################################################################

# AUTOMATIZAÇÃO PARA BAIXAR BANCO DE DADOS SAF
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatórios"))
)
sleep(2)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatórios")
link.click()

WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatório de Formulários"))
)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatório de Formulários")
link.click()

WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "SAF"))
)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "SAF")
link.click()

sleep(2)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Data")
link.click()

input_element = driver.find_element(By.NAME, "dataPartirAbertura")
sleep(2)
input_element.send_keys(Keys.CONTROL + "a")
input_element.send_keys(Keys.DELETE)

current_month_number = datetime.now().strftime("%m")
current_year_number = datetime.now().year
date_input = "01" + str(current_month_number) + str(current_year_number)
input_element.send_keys(date_input)

sleep(2)
baixar_click = driver.find_element(By.CLASS_NAME, "btn.btn-primary.btn-lg")
baixar_click.click()
#############################################################################################################

# AUTOMATIZAÇÃO PARA BAIXAR BANCO DE DADOS OSM
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatórios"))
)
sleep(2)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatórios")
link.click()

WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatório de Formulários"))
)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatório de Formulários")
link.click()

WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "OSM"))
)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "OSM")
link.click()

sleep(2)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Data")
link.click()

input_element = driver.find_element(By.NAME, "dataPartirOsmAbertura")
sleep(2)
input_element.send_keys(Keys.CONTROL + "a")
input_element.send_keys(Keys.DELETE)

current_month_number = datetime.now().strftime("%m")
current_year_number = datetime.now().year
date_input = "01" + str(current_month_number) + str(current_year_number)
input_element.send_keys(date_input)

sleep(2)
baixar_click = driver.find_element(By.CLASS_NAME, "btn.btn-primary.btn-lg")
baixar_click.click()
#####################################################################################################3#######

# AUTOMATIZAÇÃO PARA BAIXAR BANCO DE DADOS SSP
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatórios"))
)
sleep(2)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatórios")
link.click()

WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatório de Formulários"))
)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatório de Formulários")
link.click()

WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "SSP"))
)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "SSP")
link.click()

sleep(2)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Status")
link.click()

input_element = driver.find_element(By.NAME, "dataPartirAbertura")
sleep(2)
input_element.send_keys(Keys.CONTROL + "a")
input_element.send_keys(Keys.DELETE)

current_month_number = datetime.now().strftime("%m")
current_year_number = datetime.now().year
date_input = "01" + str(current_month_number) + str(current_year_number)
input_element.send_keys(date_input)

sleep(2)
baixar_click = driver.find_element(By.CLASS_NAME, "btn.btn-primary.btn-lg")
baixar_click.click()
#####################################################################################################3#######

# AUTOMATIZAÇÃO PARA BAIXAR BANCO DE DADOS OSP
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatórios"))
)
sleep(2)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatórios")
link.click()

WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatório de Formulários"))
)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatório de Formulários")
link.click()

WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "OSP"))
)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "OSP")
link.click()

sleep(2)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Datas e Status")
link.click()

input_element = driver.find_element(By.NAME, "dataPartirOspAbertura")
sleep(2)
input_element.send_keys(Keys.CONTROL + "a")
input_element.send_keys(Keys.DELETE)

current_month_number = datetime.now().strftime("%m")
current_year_number = datetime.now().year
date_input = "01" + str(current_month_number) + str(current_year_number)
input_element.send_keys(date_input)

sleep(2)
baixar_click = driver.find_element(By.CLASS_NAME, "btn.btn-primary.btn-lg")
baixar_click.click()
#####################################################################################################3#######

sleep(10)
driver.quit()