from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from time import sleep
from datetime import datetime

# leitura das credenciais de acesso ao sistema
with open("credentials.txt", "r") as file:
    user = file.readline().strip()
    password = file.readline().strip()

# criação do driver para execução de qualquer automação que seja utilizando o firefox
service = Service(executable_path="./geckodriver.exe")
driver = webdriver.Firefox(service=service)

# modo de acesso e utilização do firefox de forma automatizada
driver.get("http://simg.metrofor.ce.gov.br")

# Automatização de login no sistema
WebDriverWait(driver, 40).until(
    EC.presence_of_element_located((By.ID, "usuario"))
)

input_element = driver.find_element(By.ID, "usuario")
input_element.send_keys(user)

input_element = driver.find_element(By.ID, "senha")
input_element.send_keys(password + Keys.ENTER)


# Automatização para obtenção do csv com os dados de Solicitações de Abertura de Falhas
WebDriverWait(driver, 40).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatórios"))
)
sleep(2)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatórios")
link.click()

WebDriverWait(driver, 40).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatório de Formulários"))
)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatório de Formulários")
link.click()

WebDriverWait(driver, 40).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "SAF"))
)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "SAF")
link.click()

sleep(2)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Data")
link.click()

input_element = driver.find_element(By.NAME, "dataPartirAbertura")
sleep(2)
# Comandos para limpar a caixa de seleção e digitar a data desejada
input_element.send_keys(Keys.CONTROL + "a")
input_element.send_keys(Keys.DELETE)

current_month_number = datetime.now().strftime("%m") # Identifica o mês atual de acordo com a máquina que o programa estiver rodando
current_year_number = datetime.now().year # Identifica o ano atual de acordo com a máquina que o programa estiver rodando
date_input = "01" + str(current_month_number) + str(current_year_number) # Formata a data obtida com o dia predefinido e o mês e ano obtidos anteriormente
# Ex: 01032025 (neste formato)
input_element.send_keys(date_input) # Comando para entrar com a data definida anteriormente

sleep(2)
baixar_click = driver.find_element(By.CLASS_NAME, "btn.btn-primary.btn-lg") # Classe do botão para fazer o download do csv com os dados
baixar_click.click()


# Automatização do processo de obtenção do csv com os dados de Ordens de Serviços de Manutenção
WebDriverWait(driver, 40).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatórios"))
)
sleep(2)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatórios")
link.click()

WebDriverWait(driver, 40).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatório de Formulários"))
)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatório de Formulários")
link.click()

WebDriverWait(driver, 40).until(
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

input_element.send_keys(date_input)

sleep(2)
baixar_click = driver.find_element(By.CLASS_NAME, "btn.btn-primary.btn-lg")
baixar_click.click()


# Automatização para obtenção do csv com os dados de Solicitações de Serviços Programados
WebDriverWait(driver, 40).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatórios"))
)
sleep(2)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatórios")
link.click()

WebDriverWait(driver, 40).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatório de Formulários"))
)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatório de Formulários")
link.click()

WebDriverWait(driver, 40).until(
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

input_element.send_keys(date_input)

sleep(2)
baixar_click = driver.find_element(By.CLASS_NAME, "btn.btn-primary.btn-lg")
baixar_click.click()

# Automatização para obtenção do csv com os dados de Ordens de Serviços Programados
WebDriverWait(driver, 40).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatórios"))
)
sleep(2)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatórios")
link.click()

WebDriverWait(driver, 40).until(
    EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Relatório de Formulários"))
)
link = driver.find_element(By.PARTIAL_LINK_TEXT, "Relatório de Formulários")
link.click()

WebDriverWait(driver, 40).until(
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

input_element.send_keys(date_input)

sleep(2)
baixar_click = driver.find_element(By.CLASS_NAME, "btn.btn-primary.btn-lg")
baixar_click.click()


sleep(10)
driver.quit() # Finaliza o navegador aberto em modo de automatização
