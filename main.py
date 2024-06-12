import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from  selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
import pandas as pd
from getpass import getpass
from datetime import datetime
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select


import openpyxl


sistemaURL = "https://scat.brasilcenter.com.br"

def logarScat():
    navegador = webdriver.Chrome()
    navegador.maximize_window()
    navegador.get(sistemaURL)
    navegador.find_element(By.ID,"txtLogin").send_keys(user)
    navegador.find_element(By.ID,"txtSenha").send_keys(senha)
    navegador.find_element('xpath','/html/body/div[2]/div[3]/div/div[3]/div/button').click()

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
    # HEADER
    lines = ['GERAÇÃO DE TERMOS\n'] 
    arquivoSaida.writelines(lines)
    tabela = pd.read_excel(planilha, engine='openpyxl', sheet_name=0, usecols=[7])
    for index, row in tabela.iterrows():
        matricula = navegador.find_element('xpath', '//*[@id="txtMatriculaFuncionario"]')
        matricula.send_keys(str(row['MATRÍCULA']))
        time.sleep(2)
        navegador.find_element('xpath', '/html/body/div[1]/div[2]/section[2]/div/div[1]/div[1]/div[1]/div[2]/div/span[2]/span[1]/span/span[2]').click()
        time.sleep(2)
        navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Entrega'+Keys.ENTER)
        time.sleep(2)
        navegador.find_element('xpath', '//*[@id="select2-sltrTermo-container"]').click()
        time.sleep(2)
        navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Retirada'+Keys.ENTER)
        time.sleep(2)
        #ACIONA O BOTÃO FILTRAR
        navegador.find_element('xpath', '//*[@id="btnFiltrar"]').click()
        time.sleep(5)
        # TRY-CATCH PARA ERROR NO FLUXO DE EXECURÇÃO
        try:
            navegador.find_element('xpath', '//*[@id="gridCheckAll"]').click()
            navegador.find_element('xpath','//*[@id="btnGerarTermoMassa"]').click()
            navegador.find_element('xpath', '//*[@id="txtMatriculaFuncionario"]').clear()
            # IMPRIME A SAIDA NO ARQUIVO DE LOF
            
            lines = [str(row['MATRÍCULA'])+'    OK\n'] 
            arquivoSaida.writelines(lines)

        except Exception as Errou:
            print('Erro:'+ str(type(Errou)))
            # IMPRIME A SAIDA NO ARQUIVO DE LOg
            lines = [str(row['MATRÍCULA'])+'    ERROR\n'] 
            arquivoSaida.writelines(lines)
            navegador.find_element('xpath', '//*[@id="txtMatriculaFuncionario"]').clear()
    time.sleep(3)
    arquivoSaida.close()

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
    time.sleep(2)
    
    # monitor
    #navegador.find_element('xpath', '//*[@id="select2-sltrCategoria-container"]').click()
    #navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Monitor')
    #navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
    time.sleep(2)
    #pesquisar
    #navegador.find_element('xapth', '//*[@id="tblInventario_filter"]/label/input').send_keys(Keys.ENTER)
    tabela = pd.read_excel(planilha, engine='openpyxl', sheet_name=0, usecols=[0,1,2,3])
    for index, row in tabela.iterrows():
        try:
            #ribeirão
            navegador.find_element('xpath', '//*[@id="select2-ddlCasFiltro-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Ribeirão Preto')
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
            #Nº de serie
            navegador.find_element('xpath', '//*[@id="txtNumeroSerieFiltro"]').clear()
            navegador.find_element('xpath', '//*[@id="txtNumeroSerieFiltro"]').send_keys(str(row['Nº de Série']))
            #filtrar
            navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div/div/div/div[3]/div[1]/button[2]').click()
            #navegador.find_element('xpath', '//*[@id="gridInventario_filter"]/label/input').send_keys('GYVL7XA003217')
            time.sleep(2)
            navegador.execute_script("window.scrollBy(0, 500)")
            navegador.find_element('xpath', '//*[@id="gridInventario"]/tbody/tr[1]/td[1]/button[2]').click()
            time.sleep(3)
            navegador.find_element('xpath', '//*[@id="tblInventario"]/tbody/tr/td[1]/a').click()
            time.sleep(5)
            #EstoqueLocalização            
            navegador.find_element('xpath','//*[@id="select2-IdTipoLocalizacao-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(str(row['LOCALIZAÇÃO']))
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
            time.sleep(5)
            #alterar status
            navegador.find_element('xpath', '//*[@id="select2-IdTipoStatusEquipamento-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('segregado')
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)

            time.sleep(2)
            #alterar host
            navegador.find_element('xpath', '//*[@id="Hostname"]').clear()
            navegador.find_element('xpath', '//*[@id="Hostname"]').send_keys(str(row['Hostname']))
            time.sleep(2)
            #atualizar
            navegador.find_element('xpath', '//*[@id="modalInventario_dialog"]/div/div[3]/span[2]/button').click()
            time.sleep(3)
            #navegador.execute_script("document.activeElement.dispatchEvent(new KeyboardEvent('keydown', {'keyCode': 13, 'which': 13}));")
            navegador.switchTo().alert.accept(); 
            Alert(navegador).accept()
            time.sleep(6)
            Alert(navegador).accept()
            time.sleep(10)
    
  # IMPRIME A SAIDA NO ARQUIVO DE LOF
            
            lines = [str(row['Nº de Série'])+'    OK\n'] 
            arquivoSaida.writelines(lines)

        except Exception as Errou:
            #fechar
            #navegador.find_element('xpath', '//*[@id="modalInventario_dialog"]/div/div[1]/button').click()
            print('Erro:'+ str(type(Errou)))
            # IMPRIME A SAIDA NO ARQUIVO DE LOg
            lines = [str(row['Nº de Série'])+'    ERROR\n'] 
            arquivoSaida.writelines(lines)

def Fechar():
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
    # Acompanhar OS
    navegador.find_element('xpath', '//*[@id="lnkAcomapnharOs"]').click
    #Matricula
    navegador.find_element('xpath', '//*[@id="txtMatriculaFuncionario"]').click
    #header
    lines = ['GERAÇÃO DE TERMOS\n'] 
    arquivoSaida.writelines(lines)
    #inserir matricula
    tabela = pd.read_excel(planilha, engine='openpyxl', sheet_name=0, usecols=[7])
    for index, row in tabela.iterrows():
        try:
                    navegador.find_element('xpath','//*[@id="txtMatriculaFuncionario"]').send_keys(str(row['MATRÍCULA'])).clear()
    #navegador.find_element('xpath', '//*[@id="txtMatriculaFuncionario"]').send_keys('954145')
                    navegador.find_element('xpath', '//*[@id="collapseOne"]/div/div[1]/div[3]/div/div[2]/div/span/span[1]/span/ul/li/input').click
                    navegador.find_element('xpath', '//*[@id="collapseOne"]/div/div[1]/div[3]/div/div[2]/div/span/span[1]/span/ul/li/input').send_keys('Pendente')
                    navegador.find_element('xpath', '//*[@id="collapseOne"]/div/div[1]/div[3]/div/div[2]/div/span/span[1]/span/ul/li/input').send_keys(Keys.ENTER)
                    time.sleep(1)
                    navegador.find_element('xpath', '//*[@id="btnFiltrar"]').click()
                    time.sleep(2)
                    #selecionar tudo
                    navegador.find_element('xpath', '//*[@id="gridCheckAll"]').click()
                    time.sleep(1)
                    #alterar equipe para adm
                    #navegador.find_element('xpath', '//*[@id="grdOs"]/div[1]/div[2]/button').click()
                    #escolher equipe ADM 
                    #navegador.find_element('xpath', '//*[@id="select2-sltrEquipeAtendimentoModal-container"]').send_keys('ADM'+Keys.ENTER)
                    #Clicar Salvar
                #navegador.find_element('xpath', '//*[@id="mdlCodOSEQPM"]/div/div/b/div/button').click()
                    #Alert(navegador).accept()
                    #time.sleep(3)
                    #selecionar tudo
                    #navegador.find_element('xpath', '//*[@id="gridCheckAll"]').click()
                    time.sleep(1)
                    #clicar no concluir OS
                    navegador.find_element('xpath', '//*[@id="grdOs"]/div[1]/div[1]/button').click()
                    time.sleep(2)
                    navegador.find_element('xpath', '//*[@id="txtDataEntregaOsEmLote"]').click()
                    navegador.find_element('xpath', '//*[@id="txtDataEntregaOsEmLote"]').send_keys(datetime.now().strftime("%d/%m/%Y"))
                    time.sleep(1)
                    navegador.find_element('xpath', '//*[@id="mdlCodOSEQPM"]/div/div/b/div/button').click()
                    time.sleep(2)
                    #WebDriverWait(navegador, 120).until(EC.alert_is_present(),
                                   #'Timed out waiting for PA creation ' +
                                  # 'confirmation popup to appear.')

                    #alert = navegador.switch_to.alert
                    time.sleep(2)
                    Alert(navegador).accept()
                    time.sleep(2)
                    Alert(navegador).accept()           
                    
            # IMPRIME A SAIDA NO ARQUIVO DE LOF
            
                    lines = [str(row['MATRÍCULA'])+'    OK\n'] 
                    arquivoSaida.writelines(lines)

        except Exception as Errou:
            print('Erro:'+ str(type(Errou)))
            # IMPRIME A SAIDA NO ARQUIVO DE LOg
            lines = [str(row['MATRÍCULA'])+'    ERROR\n'] 
            arquivoSaida.writelines(lines)
            
def cadinventario():
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
    time.sleep(5)
        #adicionar
    navegador.find_element('xpath', '//*[@id="btnAdicionar"]').click()
    time.sleep(2)
    tabela = pd.read_excel(planilha, engine='openpyxl', sheet_name=0, usecols=[0,1,2,3,4])
    for index, row in tabela.iterrows():
        try: 
            #cas
            navegador.find_element('xpath', '//*[@id="frmInventario"]/div/div/div/div/div[2]/div[1]/div/span/span[1]/span/span[2]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Ribeirão Preto')
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
            time.sleep(3)
            #categoria
            navegador.find_element('xpath', '//*[@id="select2-ddlCategoriaFrmInventario-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys("computador")
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
            time.sleep(3)

            #fabricante
            navegador.find_element('xpath', '//*[@id="select2-IdFabricante-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Positivo')
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
            time.sleep(3)
            #EstoqueLocalização
            navegador.find_element('xpath', '//*[@id="select2-IdTipoLocalizacao-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Estoque - RPO')
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
            time.sleep(3)
            #modelo
            navegador.find_element('xpath', '//*[@id="select2-IdModelo-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Positivo C6400')
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
            time.sleep(3)
            #propriedade
            navegador.find_element('xpath', '//*[@id="select2-IdTipoPropriedade-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys("Alugado")
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)

            #hostname
            navegador.find_element('xpath', '//*[@id="Hostname"]').click()
            navegador.find_element('xpath', '//*[@id="Hostname"]').clear()
            navegador.find_element('xpath', '//*[@id="Hostname"]').send_keys('RPO-HO-AL-XXXX')

            #Nº Patrimonio
            navegador.find_element('xpath', '//*[@id="NumeroPatrimonio"]').click()
            navegador.find_element('xpath', '//*[@id="NumeroPatrimonio"]').clear()
            navegador.find_element('xpath', '//*[@id="NumeroPatrimonio"]').send_keys(str(row['Nº de Série']))
            #navegador.find_element('xpath', '//*[@id="NumeroPatrimonio"]').send_keys('xxxx')
            time.sleep(1)
            #Nº serial
            navegador.find_element('xpath', '//*[@id="NumeroSerie"]').click()
            navegador.find_element('xpath', '//*[@id="NumeroSerie"]').clear()
            navegador.find_element('xpath', '//*[@id="NumeroSerie"]').send_keys(str(row['Nº de Série']))
            #navegador.find_element('xpath', '//*[@id="NumeroSerie"]').send_keys('xxx')
            time.sleep(1)
            #Status Equipamento
            navegador.find_element('xpath', '//*[@id="select2-IdTipoStatusEquipamento-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys("Em Reserva")
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
            time.sleep(1)
            #Sistema Operacional
            navegador.find_element('xpath', '//*[@id="select2-IdTipoSistemaOperacional-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys("Windows 10")
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)    
            time.sleep(1)
            #Memória
            navegador.find_element('xpath', '//*[@id="Memoria"]').click()
            navegador.find_element('xpath', '//*[@id="Memoria"]').clear()
            navegador.find_element('xpath', '//*[@id="Memoria"]').send_keys("16GB")
            navegador.find_element('xpath', '//*[@id="Memoria"]').send_keys(Keys.ENTER)
            time.sleep(1)
            #btnsalvar
            navegador.execute_script("window.scrollBy(0, 500)")
            navegador.find_element('xpath','//*[@id="modalInventario_dialog"]/div/div[3]/span[2]/button').click()
            time.sleep(1)
            Alert(navegador).accept()
            time.sleep(7)
            Alert(navegador).accept()
            time.sleep(10)
  
  # IMPRIME A SAIDA NO ARQUIVO DE LOF
            
            lines = [str(row['Nº de Série'])+'    OK\n'] 
            arquivoSaida.writelines(lines)

        except Exception as Errou:
            #fechar
            #navegador.find_element('xpath', '//*[@id="modalInventario_dialog"]/div/div[1]/button').click()
            print('Erro:'+ str(type(Errou)))
            # IMPRIME A SAIDA NO ARQUIVO DE LOg
            lines = [str(row['Nº de Série'])+'    ERROR\n'] 
            arquivoSaida.writelines(lines)
    
    
    #pesquisar
    #navegador.find_element('xapth', '//*[@id="tblInventario_filter"]/label/input').send_keys(Keys.ENTER)

def cadinventariomonitor():
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
    time.sleep(5)
        #adicionar
    navegador.find_element('xpath', '//*[@id="btnAdicionar"]').click()
    time.sleep(2)
    tabela = pd.read_excel(planilha, engine='openpyxl', sheet_name=0, usecols=[0,1,2,3,4])
    for index, row in tabela.iterrows():
        try:
            #cas
            navegador.find_element('xpath', '//*[@id="frmInventario"]/div/div/div/div/div[2]/div[1]/div/span/span[1]/span/span[2]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Ribeirão Preto')
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
            time.sleep(3)
            #categoria
            navegador.find_element('xpath', '//*[@id="select2-ddlCategoriaFrmInventario-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys("monitor")
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
            time.sleep(3)

            #fabricante
            navegador.find_element('xpath', '//*[@id="select2-IdFabricante-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Positivo')
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
            time.sleep(3)
            #EstoqueLocalização
            navegador.find_element('xpath', '//*[@id="select2-IdTipoLocalizacao-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('Estoque - RPO')
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
            time.sleep(3)
            #modelo
            navegador.find_element('xpath', '//*[@id="select2-IdModelo-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys('22BN550Y')
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
            time.sleep(3)
            #propriedade
            navegador.find_element('xpath', '//*[@id="select2-IdTipoPropriedade-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys("Alugado")
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)

            #hostname
            navegador.find_element('xpath', '//*[@id="Hostname"]').click()
            navegador.find_element('xpath', '//*[@id="Hostname"]').clear()
            navegador.find_element('xpath', '//*[@id="Hostname"]').send_keys('RPO-HO-AL-XXXX')

            #Nº Patrimonio
            navegador.find_element('xpath', '//*[@id="NumeroPatrimonio"]').click()
            navegador.find_element('xpath', '//*[@id="NumeroPatrimonio"]').clear()
            navegador.find_element('xpath', '//*[@id="NumeroPatrimonio"]').send_keys(str(row['Nº de Série']))
            #navegador.find_element('xpath', '//*[@id="NumeroPatrimonio"]').send_keys('xxxx')
            time.sleep(1)
            #Nº serial
            navegador.find_element('xpath', '//*[@id="NumeroSerie"]').click()
            navegador.find_element('xpath', '//*[@id="NumeroSerie"]').clear()
            navegador.find_element('xpath', '//*[@id="NumeroSerie"]').send_keys(str(row['Nº de Série']))
            #navegador.find_element('xpath', '//*[@id="NumeroSerie"]').send_keys('xxx')
            time.sleep(1)
            #Status Equipamento
            navegador.find_element('xpath', '//*[@id="select2-IdTipoStatusEquipamento-container"]').click()
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys("Em Reserva")
            navegador.find_element('xpath', '/html/body/span/span/span[1]/input').send_keys(Keys.ENTER)
            time.sleep(1)
            #btnsalvar
            navegador.execute_script("window.scrollBy(0, 500)")
            navegador.find_element('xpath','//*[@id="modalInventario_dialog"]/div/div[3]/span[2]/button').click()
            time.sleep(1)
            Alert(navegador).accept()
            time.sleep(7)
            Alert(navegador).accept()
            time.sleep(10)
  
  # IMPRIME A SAIDA NO ARQUIVO DE LOF
            
            lines = [str(row['Nº de Série'])+'    OK\n'] 
            arquivoSaida.writelines(lines)

        except Exception as Errou:
            #fechar
            #navegador.find_element('xpath', '//*[@id="modalInventario_dialog"]/div/div[1]/button').click()
            print('Erro:'+ str(type(Errou)))
            # IMPRIME A SAIDA NO ARQUIVO DE LOg
            lines = [str(row['Nº de Série'])+'    ERROR\n'] 
            arquivoSaida.writelines(lines)

def prodMaquina():
    #LOGIN NO SISTEMA
    navegador = webdriver.Chrome()
    navegador.maximize_window()
    navegador.get(sistemaURL)
    navegador.find_element(By.ID,"txtLogin").send_keys(user)
    navegador.find_element(By.ID,"txtSenha").send_keys(senha)
    navegador.find_element('xpath','/html/body/div[2]/div[3]/div/div[3]/div/button').click()
    time.sleep(5)
    # MENU -> SOLICITAÇÕES -> ACOMPANHAR OS
    navegador.find_element('xpath', '/html/body/div[2]/nav/ul[1]/li/a').click()
    navegador.find_element('xpath', '/html/body/div[2]/aside[1]/div/div[4]/div/div/nav/ul/li/a').click()
    navegador.find_element('xpath', '//*[@id="lnkAcomapnharOs"]').click()
    time.sleep(5)
    #PESQUISANDO MATRICULAS
    navegador.find_element('xpath','//*[@id="collapseOne"]/div/div[1]/div[1]/div/div[2]/div/span/span[1]/span/ul/li').click()
    navegador.find_element('xpath','//*[@id="collapseOne"]/div/div[1]/div[1]/div/div[2]/div/span/span[1]/span/ul/li[2]/input').send_keys("RIBEIRÃO PRETO")
    navegador.find_element('xpath','//*[@id="collapseOne"]/div/div[1]/div[1]/div/div[2]/div/span/span[1]/span/ul/li[2]/input').send_keys(Keys.ENTER)
    navegador.find_element('xpath','//*[@id="collapseOne"]/div/div[1]/div[3]/div/div[2]/div/span/span[1]/span/ul/li[3]/input').send_keys("NOVO")
    navegador.find_element('xpath','//*[@id="collapseOne"]/div/div[1]/div[3]/div/div[2]/div/span/span[1]/span/ul/li[3]/input').send_keys(Keys.ENTER)
    navegador.find_element('xpath','//*[@id="collapseOne"]/div/div[1]/div[3]/div/div[2]/div/span/span[1]/span/ul/li[3]/input').send_keys("PENDENTE")
    navegador.find_element('xpath','//*[@id="collapseOne"]/div/div[1]/div[3]/div/div[2]/div/span/span[1]/span/ul/li[3]/input').send_keys(Keys.ENTER)
    navegador.find_element('xpath','//*[@id="collapseOne"]/div/div[1]/div[4]/div/div[2]/div/span/span[1]/span/ul/li/input').send_keys("VALIDACAO DE ENDERECO")
    navegador.find_element('xpath','//*[@id="collapseOne"]/div/div[1]/div[4]/div/div[2]/div/span/span[1]/span/ul/li/input').send_keys(Keys.ENTER)
    navegador.find_element('xpath','//*[@id="collapseOne"]/div/div[1]/div[4]/div/div[2]/div/span/span[1]/span/ul/li/input').send_keys("SUPORTE TI")
    navegador.find_element('xpath','//*[@id="collapseOne"]/div/div[1]/div[4]/div/div[2]/div/span/span[1]/span/ul/li/input').send_keys(Keys.ENTER)


#Programa Principal
user = input('Insira seu usuario com domínio:')
senha = getpass()
planilha = input('Insira o caminho da planilha: ')
arquivoSaida = open('OutputLog.txt', 'w')

cadinventariomonitor()  