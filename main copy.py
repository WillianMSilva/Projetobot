import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from  selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
import pandas as pd
from getpass import getpass

import openpyxl

sistemaURL = "https://scat.brasilcenter.com.br"

def cadastrar():
    navegador = webdriver.Chrome()
    navegador.maximize_window()
    navegador.get("https://scat.brasilcenter.com.br")
    time.sleep(5)
    navegador.find_element('xpath','//*[@id="txtLogin"]').send_keys(user)
    navegador.find_element('xpath','//*[@id="txtSenha"]').send_keys(senha)
    navegador.find_element('xpath','/html/body/div[2]/div[3]/div/div[3]/div/button').click()
    time.sleep(5)
    # menu
    navegador.find_element('xpath', '/html/body/div[2]/nav/ul[1]/li/a').click()
    time.sleep(1)
    # solicitações
    navegador.find_element('xpath', '/html/body/div[2]/aside[1]/div/div[4]/div/div/nav/ul/li/a').click()
    time.sleep(1)
    #cadastrar
    navegador.find_element('xpath', '/html/body/div[2]/aside[1]/div/div[4]/div/div/nav/ul/li/ul[1]/li[2]/a').click()
    #retirada
    navegador.find_element('xpath','/html/body/div[2]/div[2]/section[2]/div/div[1]/div[1]/div[1]/div/div/span/span[1]/span/span[2]').click()
    time.sleep(2)
    navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Retirada'+Keys.ENTER)
    time.sleep(2)
    #inserir matriculas
    navegador.find_element('xpath','//*[@id="txtMatriculaFuncionario"]').send_keys('954145')
    time.sleep(50)

def termo():
    navegador = webdriver.Chrome()
    navegador.maximize_window()
    navegador.get("https://scat.brasilcenter.com.br")
    time.sleep(5)
    navegador.find_element('xpath','//*[@id="txtLogin"]').send_keys(user)
    navegador.find_element('xpath','//*[@id="txtSenha"]').send_keys(senha)
    navegador.find_element('xpath','/html/body/div[2]/div[3]/div/div[3]/div/button').click()
    time.sleep(5)
    # menu
    navegador.find_element('xpath', '/html/body/div[2]/nav/ul[1]/li/a').click()
    time.sleep(2)
    navegador.find_element('xpath', '/html/body/div[2]/aside[1]/div/div[4]/div/div/nav/ul/li/a').click()
    time.sleep(2)
    #termo
    navegador.find_element('xpath','/html/body/div[2]/aside[1]/div/div[4]/div/div/nav/ul/li/ul[1]/li[3]/a').click()
    time.sleep(2)
    #navegador.find_element('xpath', '//*[@id="txtMatriculaFuncionario"]').send_keys(colar)
    tabela = pd.read_excel(planilha, engine='openpyxl', sheet_name=0, usecols=[1])
    for index, row in tabela.iterrows():
        matricula = navegador.find_element('xpath', '//*[@id="txtMatriculaFuncionario"]')
        matricula.send_keys(str(row['NU_MATR']))
        time.sleep(2)
        navegador.find_element('xpath', '/html/body/div[1]/div[2]/section[2]/div/div[1]/div[1]/div[1]/div[2]/div/span[2]/span[1]/span/span[2]').click()
        time.sleep(2)
        navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Retirada'+Keys.ENTER)
        time.sleep(2)
        navegador.find_element('xpath', '//*[@id="select2-sltrTermo-container"]').click()
        time.sleep(2)
        navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Retirada'+Keys.ENTER)
        time.sleep(2)
        navegador.find_element('xpath', '//*[@id="btnFiltrar"]').click()
        time.sleep(5)
        navegador.find_element('xpath', '//*[@id="txtMatriculaFuncionario"]').clear()

def inventario():
    navegador = webdriver.Chrome()
    navegador.maximize_window()
    navegador.get(sistemaURL)
    time.sleep(5)
    navegador.find_element(By.ID,"txtLogin").send_keys(user)
    navegador.find_element(By.ID,"txtSenha").send_keys(senha)
    navegador.find_element('xpath','/html/body/div[2]/div[3]/div/div[3]/div/button').click()
    time.sleep(5)
    # menu
    navegador.find_element('xpath', '/html/body/div[2]/nav/ul[1]/li/a').click()
    time.sleep(2)
    # inventario
    navegador.find_element('xpath',
                           '/html/body/div[2]/aside[1]/div/div[4]/div/div/nav/ul/li/ul[2]/li/ul[4]/li/ul[12]/li[1]/a').click()
    # equipamentos
    navegador.find_element('xpath',
                           '/html/body/div[2]/aside[1]/div/div[4]/div/div/nav/ul/li/ul[2]/li/ul[4]/li/ul[12]/li[1]/ul[1]/li/a').click()
    #ribeirão
    navegador.find_element('xpath', '/html/body/div[1]/div[2]/section[2]/div/div[1]/div[1]/div[1]/div[1]/div/span/span[1]/span/span[2]').click()
    navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Ribeirão Preto')
    navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
    # monitor
    navegador.find_element('xpath', '//*[@id="select2-sltrCategoria-container"]').click()
    navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Monitor')
    navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
    time.sleep(1)
    #filtrar
    navegador.find_element('xpath', '//*[@id="btnFiltrar"]').click()
    time.sleep(10)
    #pesquisar
    #navegador.find_element('xapth', '//*[@id="tblInventario_filter"]/label/input').send_keys(Keys.ENTER)
    tabela = pd.read_excel(planilha, engine='openpyxl', sheet_name=0, usecols=[5])
    for index, row in tabela.iterrows():
        #navegador.find_element('xpath', '//*[@id="tblInventario_filter"]/label/input').send_keys(str(row['Nº de Série']))
        navegador.find_element('xpath', '//*[@id="tblInventario_filter"]/label/input').send_keys('GYVL7XA003217')
        time.sleep(2)
        navegador.find_element('xpath', '/html/body/div[1]/div[2]/section[2]/div').click()
        navegador.execute_script("window.scrollBy(0, 500)")
        time.sleep(3)
        navegador.find_element('xpath', '//*[@id="tblInventario"]/tbody/tr/td[1]/a').click()
        time.sleep(5)
        #alterar status
        navegador.find_element('xpath', '//*[@id="select2-sltrStatusModal-container"]').click()
        time.sleep(1)
        navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('produção')
        navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
        time.sleep(2)
        #atualizar
        navegador.find_element('xpath', '//*[@id="modalInventarioEdicao"]/div/div/div[4]/button[4]').click()
        time.sleep(3)
        #navegador.execute_script("document.activeElement.dispatchEvent(new KeyboardEvent('keydown', {'keyCode': 13, 'which': 13}));")
        #navegador.switchTo().alert.accept(); 
        Alert(navegador).accept()
        time.sleep(6)
        #fechar
        navegador.find_element('xpath','//*[@id="modalInventarioEdicao"]/div/div/div[4]/button[5]').click()

#Programa Principal
user = input('Insira seu usuario com domínio:')
senha = getpass()
planilha = r'D:\PROJECTS\Projetobot-master\00.xlsx'

inventario()