from webbrowser import Chrome

from selenium import webdriver as opcoesSelenium
from selenium.webdriver.common.by import By
import pyautogui as tempoEspera
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook

nome_arquivo_desafio = "C:\\Projetos\\Python\\pythonProject\\Desafio\\challenge.xlsx"
planilha_arquivo_desafio = load_workbook(nome_arquivo_desafio)


planilha_selecionada = planilha_arquivo_desafio["Sheet1"]

chrome_options = opcoesSelenium.ChromeOptions()
chrome_options.add_experimental_option("detach", True)
navegador = opcoesSelenium.Chrome(options=chrome_options)

#Abre o site do desafio
navegador.get("https://rpachallenge.com/")

#Aguarda a página carregar
tempoEspera.sleep(2)

navegador.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button').click()

for linha in range(2, len(planilha_selecionada['A']) + 1):

    tempoEspera.sleep(2)

    #Recebe o valor da celula corrente na coluna A
    firstName = planilha_selecionada['A%s' % linha].value

    lastName = planilha_selecionada['B%s' % linha].value

    companyName = planilha_selecionada['C%s' % linha].value

    labelRole = planilha_selecionada['D%s' % linha].value

    labelAddress = planilha_selecionada['E%s' % linha].value

    labelEmail = planilha_selecionada['F%s' % linha].value

    labelPhone = planilha_selecionada['G%s' % linha].value

    #Localiza o campo e envia o texto do Excel
    # xpath personalizado //*[@]
    navegador.find_element(By.XPATH, '//*[@ng-reflect-name="labelFirstName"]').send_keys(firstName)

    navegador.find_element(By.XPATH, '//*[@ng-reflect-name="labelLastName"]').send_keys(lastName)

    navegador.find_element(By.XPATH, '//*[@ng-reflect-name="labelCompanyName"]').send_keys(companyName)

    navegador.find_element(By.XPATH, '//*[@ng-reflect-name="labelRole"]').send_keys(labelRole)

    navegador.find_element(By.XPATH, '//*[@ng-reflect-name="labelAddress"]').send_keys(labelAddress)

    navegador.find_element(By.XPATH, '//*[@ng-reflect-name="labelEmail"]').send_keys(labelEmail)

    navegador.find_element(By.XPATH, '//*[@ng-reflect-name="labelPhone"]').send_keys(labelPhone)

    #Clicar no botão enviar
    navegador.find_element(By.XPATH, "/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input").click()
