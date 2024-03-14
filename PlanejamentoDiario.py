import time
from datetime import date, timedelta
from tkinter.simpledialog import askstring
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException
from lancamentoWeb import executar_lancamento_web
from navegacao import executar_navegação
from geralDoc import executar_geral_documento
from contratoEmpresa import executar_contrato_empresa
from historicoWeb import executar_historico_web
 
def selecionar_opcao_com_tentativa(driver, element_id, valor, tentativas=3):
    for _ in range(tentativas):
        try:
            elemento_modulo = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, element_id))
            )
            dropdown = Select(elemento_modulo)
            dropdown.select_by_value(valor)
            return
        except StaleElementReferenceException:
            continue


# Inicialize o ChromeDriver
driver = webdriver.Chrome()

# Aguarde um pouco
driver.implicitly_wait(10)

try:
    #Data de hoje
    hoje = date.today()
    #Definindo e formatando data fim de vigencia que é o dia atual
    datafim = hoje.strftime("%d%m%Y")
    
    #Decrementando 119 dias da data atual para ter a data inicio da vigencia
    datainicio = hoje - timedelta(days=119)
    #Definindo data fim e formatando
    datainicio = datainicio.strftime("%d%m%Y")
      
    #Chamado condigo do menu de login e navegação
    executar_navegação(driver)

    #Chamando codigo de extração do relatorio de NEXA BRASIL- Documentos Lançados via WEB
    executar_lancamento_web(driver, datainicio, datafim)
    time.sleep(10)
     
    #Chamando codigo de extração de relatorio NEXA -Histórico de Lançamento de Documento via WEB
    executar_historico_web(driver, datainicio, datafim)
    time.sleep(10) 
        
    #Chamando codigo de extração do relatorio NEXA BRASIL - Contratos por Empresa
    executar_contrato_empresa(driver)
    time.sleep(10)  
    
    #Chamando codigo de extação do relatorio NEXA BRASIL - Relatório Geral de Documentos
    executar_geral_documento(driver)
    
    time.sleep(200)
    resposta = ""
    while resposta not in ["1", "Sim", "sim", "SIM"]:  
        time.sleep(120)  #120' seconds
        resposta = askstring("Informações", "Os relatórios foram baixados?\n1 - Sim\n2 - Não")
    

finally:
    driver.implicitly_wait(10)
    # Feche o navegador
    driver.quit()
