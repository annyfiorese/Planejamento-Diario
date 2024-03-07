import time
from datetime import date
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
    
    #Definindo data fim e formatando
    datafim = hoje.strftime("%d%m%Y")
    
    #Decrementando 119 dias da data atual para ter a data inicio
    
    print(hoje)
    #Chamado condigo do menu de login e navegação
    executar_navegação(driver)
    
    #Solicitando informação de data inicio de inclusão e final 
    datainicio = askstring("Informações", "Digite a data de início da incluisão(DDMMYYYY):")
    datafim = askstring("Informações", "Digite a data de final de inclusão (DDMMYYYY):")
    
    #Chamando codigo de extração do relatorio de NEXA BRASIL- Documentos Lançados via WEB
    executar_lancamento_web(driver, datainicio, datafim)
    time.sleep(20)
     
    #Chamando codigo de extração de relatorio NEXA -Histórico de Lançamento de Documento via WEB
    executar_historico_web(driver, datainicio, datafim)
    time.sleep(10) 
        
    #Chamando codigo de extração do relatorio NEXA BRASIL - Contratos por Empresa
    executar_contrato_empresa(driver)
    time.sleep(10)  
    
    #Chamando codigo de extação do relatorio NEXA BRASIL - Relatório Geral de Documentos
    executar_geral_documento(driver)
    
    time.sleep(40)
    resposta = ""
    while resposta not in ["1", "Sim", "sim", "SIM"]:  
        time.sleep(40)  # Delay for 40 seconds
        resposta = askstring("Informações", "Os relatórios foram baixados?\n1 - Sim\n2 - Não")
        
    #validation = askstring("Informações", "Automação irá finalizar!! se tudo estiver correto digite 1")

finally:
    driver.implicitly_wait(10)
    # Feche o navegador
    driver.quit()
