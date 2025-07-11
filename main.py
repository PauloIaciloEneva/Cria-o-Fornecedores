import sys
import time
import subprocess
#import login
import win32com.client
import pythoncom
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import pandas as pd
import os
import pyodbc
import numpy as np
import datetime
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import numpy as np 


#Abrindo navegador no portal de cadastro
driver = webdriver.Chrome()
driver.get('https://portaldecadastro.eneva.com.br/client/')
driver.maximize_window()



#Logando no portal
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mat-input-0"]'))).send_keys('paulo.iacilo@eneva.com.br')
driver.find_element(By.XPATH,'//*[@id="mat-input-1"]').send_keys('Vascodagama2027')
driver.find_element(By.XPATH,'/html/body/app-root/app-login/div/div/mat-card/div[1]/div[2]/form/div[2]/button').click()
time.sleep(5)
try:
    driver.find_element(By.XPATH,'//*[@id="mat-input-0"]').send_keys('paulo.iacilo@eneva.com.br')
    driver.find_element(By.XPATH,'//*[@id="mat-input-1"]').send_keys('Vascodagama2027')
    driver.find_element(By.XPATH,'/html/body/app-root/app-login/div/div/mat-card/div[1]/div[2]/form/div[2]/button').click()
except:
    pass





#Indo para todos os chamados de fornecedores
driver.get('https://portaldecadastro.eneva.com.br/client/ticket/fornecedores')
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/app-root/app-collaborator/div/div/app-ticket/div/div[1]/div/div[2]/span'))).click()



#Abrindo menu de filtros
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mat-expansion-panel-header-0"]/span[2]'))).click()


time.sleep(1)
#Filtrando chamados de Expansão
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mat-select-3"]/div/div[2]'))).click()

time.sleep(1)

for i in ['Fornecedor - Criação SAP NACIONAL', 'Fornecedor - Criação SAP ESTRANGEIRO', 'Fornecedor - Criação Colaborador Eneva']:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f"//span[contains(text(), '{i}')]"))).click()

webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()


#Filtrando chamados em análise
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mat-select-1"]/div/div[2]/div'))).click()
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Em análise')]"))).click()
webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()


WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mat-select-6"]/div/div[2]'))).click()
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), '50')]"))).click()
webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()



dataCHAMADOS = pd.DataFrame()


#for coluna in ['TIPO_CHAMADO', 'TIPO_FORNECEDOR', 'ORGANIZACAO_COMPRAS', 'EMPRESA', 'MATRICULA', 'NOME', 'CPF/CNPJ', 'ENDEREÇO', 'NUMERO', 'BAIRRO', 'MUNICIPIO', 'ESTADO', 'CEP', 'TELEFONE', 'CELULAR', 'E-MAIL', 'BANCO', 'CONTA CORRENTE', 'AGENCIA']:
#    dataCHAMADOS[coluna] = np.nan



def criar_df(tipo, linha):
    tipo_chamado = definir_chamado(tipo)
    dados = {}

    if tipo_chamado == ['COLABORADOR']:
        dados = {
            'SETOR_COLABORADOR': linha[1],
            'ORGANIZACAO_COMPRAS': linha[2],
            'EMPRESA': linha[3],
            'MATRICULA': linha[4],
            'NOME_COMPLETO': linha[5],
            'CPF/CNPJ': linha[6],
            'MUNICIPIO': linha[10],
            'CEP': linha[11]
        }

    elif tipo_chamado == ['SAP NACIONAL']:
        endereco = linha[8].split(",")
        numero = endereco[1] if len(endereco) > 1 else ""
        numero = ''.join(filter(str.isdigit, numero))

        dados = {
            'TIPO FORNECEDOR': linha[1],
            'ORGANIZACAO_COMPRAS': linha[2],
            'EMPRESA': linha[3],
            'MATRICULA': linha[4],
            'NOME': linha[5],
            'CPF/CNPJ': linha[7],
            'ENDEREÇO': endereco[0],
            'NUMERO': numero,
            'BAIRRO': linha[9],
            'MUNICIPIO': linha[10],
            'ESTADO': linha[11],
            'CEP': linha[12],
            'TELEFONE': linha[17],
            'CELULAR': linha[18],
            'E-MAIL': linha[19],
            'BANCO': linha[20],
            'CONTA_CORRENTE': linha[21],
            'AGENCIA': linha[22]
        }

    return pd.DataFrame([dados])

        
        


def definir_chamado(tipo):
    if tipo == 'Formulário: CRIAÇÃO FORNECEDOR COLABORADOR': return ["COLABORADOR"]
    elif tipo == 'Formulário: Fornecedor - Criação SAP NACIONAL': return ["SAP NACIONAL"]
    else: return ["SAP ESTRANGEIRO"]


numero_chamado=[]
for i in driver.find_elements(By.CSS_SELECTOR, 'tr.mat-row.cdk-row.ng-star-inserted'):
 numero_chamado.append(i.text.split(' ')[0])

for chamado in numero_chamado:
    
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f"//td[contains(text(), '{chamado}')]"))).click()
    chamado_atual=driver.current_url.split('/')[-1]
    
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), ' Editar o Formulário Online')]"))).click()

    time.sleep(2)
    tipo_chamado = driver.find_element(By.XPATH, "/html/body/app-root/app-collaborator/div/div/app-online-form/div/div[2]/div[1]/p").text
    
    #qtd_expansoes=len(driver.find_elements(By.XPATH, "//tr[starts-with(@class, 'mat-row.cdk-row.ng-star-inserted')]"))
    
    linhas = driver.find_elements(By.XPATH, "//table/tbody/tr")
    
    tabela_dados = []
    
    for linha in linhas:
        dados_linha = []
        celulas = linha.find_elements(By.XPATH, ".//td | .//th")

        for celula in celulas:
            try:
                input_element = celula.find_element(By.XPATH, ".//input")
                valor = input_element.get_attribute("value")
            except:
                valor = celula.text.strip()

            dados_linha.append(valor)
        tabela_dados.append(dados_linha)

    
    
    for linha in tabela_dados:
        if definir_chamado(tipo_chamado) == ["COLABORADOR"]:
            dataCOLABORADOR = criar_df(tipo_chamado, linha)
            dataCHAMADOS_COLABORADOR = pd.concat([dataCHAMADOS, dataCOLABORADOR], ignore_index=True)
        elif definir_chamado(tipo_chamado) == ['SAP NACIONAL']:
            dataSAP_NACIONAL = criar_df(tipo_chamado, linha)
            dataCHAMADOS_SAP_NACIONAL = pd.concat([dataCHAMADOS, dataSAP_NACIONAL], ignore_index=True)
        else:
            dataSAP_INTERNACIONAL = criar_df(tipo_chamado, linha)
            dataCHAMADOS_SAP_INTERNACIONAL = pd.concat([dataCHAMADOS, dataSAP_INTERNACIONAL], ignore_index=True)
        
        
        
    driver.back()
    driver.back()
















class SapGui(object):
    def __init__(self):
        self.path = "C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe"
        subprocess.Popen(self.path)
        time.sleep(5)
        
        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = self.SapGuiAuto.GetScriptingEngine
        #nome da conexão SAP 
        self.connection = application.OpenConnection("SAP S/4Hana - Homologação", True)
        time.sleep(3)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize()
        #self.session.findById("wnd[0]/usr/tblSAPMSYSTTC_IUSRACL/btnIUSRACL-BNAME[1,0]").setFocus()
        #self.session.findById("wnd[0]/usr/tblSAPMSYSTTC_IUSRACL/btnIUSRACL-BNAME[1,0]").press()

SAP_instancia = SapGui()





def acessarXK01():
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "xk01"
    session.findById("wnd[0]").sendVKey(0)
    
    if definir_chamado(tipo_chamado) == ['COLABORADOR']:
        session.findById("wnd[0]").sendVKey(0)
        
    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/btnSCREEN_3100_BUTTON_CLOSE").press()
        
    
acessarXK01()



def informacoes_endereco(primeiro_nome, segundo_nome):
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    
    if definir_chamado(tipo_chamado) == ['COLABORADOR']:
        try:
            session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/subSCREEN_1000_HEADER_AREA:SAPLBUPA_DIALOG_JOEL:1500/cmbBUS_JOEL_MAIN-CREATION_GROUP").key = "ZCOB"
        except:
            session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/subSCREEN_1000_HEADER_AREA:SAPLBUPA_DIALOG_JOEL:1500/cmbBUS_JOEL_MAIN-CREATION_GROUP").key = "ZCOB"
    
    #primeiras informações
    
    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7014/subA02P03:SAPLBUD0:1301/txtBUT000-NAME_FIRST").text = primeiro_nome
    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7014/subA02P04:SAPLBUD0:1302/txtBUT000-NAME_LAST").text = segundo_nome
    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7014/subA02P07:SAPLBUD0:1360/ctxtBUS000FLDS-LANGUCORR").text = "PT"
    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7014/subA03P01:SAPLBUD0:1110/txtBUS000FLDS-BU_SORT1_TXT").text = primeiro_nome + ' ' + segundo_nome
    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7014/subA03P01:SAPLBUD0:1110/txtBUS000FLDS-BU_SORT1_TXT").caretPosition = 3
    
    
    ###############
    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7014/subA02P04:SAPLBUD0:1302/btnPUSH_BUPK").press()
    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7014/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/txtADDR2_DATA-STREET").text = "dsa"
    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7014/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/txtADDR2_DATA-HOUSE_NUM1").text = "dsa"
    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7014/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/txtADDR2_DATA-POST_CODE1").text = "dsa"
    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7014/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/txtADDR2_DATA-CITY1").text = "dsa"
    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7014/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/ctxtADDR2_DATA-COUNTRY").text = "dsa"
    session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7014/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/ctxtADDR2_DATA-REGION").text = "dsa"
    
    
    
    

for index in range(0,len(dataCHAMADOS_COLABORADOR)):        
    informacoes_endereco(primeiro_nome=dataCHAMADOS_COLABORADOR['NOME_COMPLETO'].str.split(" ")[index][0], segundo_nome= dataCHAMADOS_COLABORADOR['NOME_COMPLETO'].str.split(" ")[index][1])











# referencia, parceiro = acessarTexto()


# def realizarmodificacao(parceiro, referencia):
#     sapguiauto = win32com.client.GetObject("SAPGUI")
#     application = sapguiauto.GetScriptingEngine
#     connection = application.Children(0)
#     session = connection.Children(0)
    
#     session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1102/subSCREEN_1100_SUB_HEADER_AREA:SAPLFS_BP_ECC_DIALOGUE:0001/btnPUSH_FSBP_CC_DETAIL").press()
#     session.findById("wnd[1]/usr/subBDT_SUBSCREEN_PP:SAPLBUSS:0029/ssubGENSUB:SAPLBUSS:7021/subA02P01:SAPLFS_BP_ECC_DIALOGUE:0002/btnPUSH_FSBP_CC_COPYREF").press()
#     session.findById("wnd[2]/usr/ctxtGS_SUPP_CC_TARGET-BUKRS").text = "MP93" ### MODIFICAR
#     session.findById("wnd[2]/usr/ctxtGS_BUT000-PARTNER").text = parceiro
#     session.findById("wnd[2]/usr/ctxtGS_BUT000-PARTNER").setFocus
#     session.findById("wnd[2]/usr/ctxtGS_BUT000-PARTNER").caretPosition = 3
#     session.findById("wnd[2]").sendVKey(0)
#     session.findById("wnd[2]/usr/ctxtGS_SUPP_CC_REF-BUKRS").text = referencia
#     session.findById("wnd[2]/usr/ctxtGS_SUPP_CC_REF-BUKRS").setFocus
#     session.findById("wnd[2]/usr/ctxtGS_SUPP_CC_REF-BUKRS").caretPosition = 4
#     session.findById("wnd[2]/tbar[0]/btn[5]").press()

# realizarmodificacao(parceiro=parceiro, referencia=referencia)