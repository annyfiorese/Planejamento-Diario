import time
from datetime import date, timedelta
from tkinter.simpledialog import askstring
from selenium import webdriver
import tkinter
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from lancamentoWeb import executar_lancamento_web
from navegacao import executar_navegação
from geralDoc import executar_geral_documento
from contratoEmpresa import executar_contrato_empresa
from historicoWeb import executar_historico_web

def selecionar_opcao_com_tentativa(driver, element_id, valor, tentativas=5, espera=20):
    for _ in range(tentativas):
        try:
            elemento_modulo = WebDriverWait(driver, espera).until(
                EC.presence_of_element_located((By.ID, element_id))
            )
            dropdown = Select(elemento_modulo)
            dropdown.select_by_value(valor)
            return
        except (StaleElementReferenceException, TimeoutException):
            time.sleep(1)
            continue
    print(f"Erro ao selecionar a opção {valor} no elemento {element_id}")

# Inicialize o ChromeDriver
driver = webdriver.Chrome()

# Aguarde um pouco
driver.implicitly_wait(10)

try:
    
    # Abra o site
    driver.get("https://nexa-brasil.rainbowtec.com.br")

    # Aguarde um pouco
    driver.implicitly_wait(10)  # Espera 10 segundos

    # Pegando as informações de login
    root = tkinter.Tk()
    root.withdraw()  # Oculta a janela principal
    usuario = askstring("Informações", "Digite seu usuário:")
    senha = askstring("Informações", "Digite sua senha:", show='*')
    
    # Chamado código do menu de login e navegação
    
    executar_navegação(driver,usuario,senha)
    time.sleep(5)
    
    # Data de hoje
    hoje = date.today()
    
    # Definindo e formatando data fim de vigência que é o dia atual
    datafim = hoje.strftime("%d%m%Y")
    
    # Decrementando 119 dias da data atual para ter a data início da vigência
    datainicio = hoje - timedelta(days=119)
    datainicio = datainicio.strftime("%d%m%Y")
    
    # Chamando código de extração do relatório de NEXA BRASIL - Documentos Lançados via WEB
    executar_lancamento_web(driver, datainicio, datafim)
    time.sleep(4)
    
    # Chamando código de extração de relatório NEXA - Histórico de Lançamento de Documento via WEB
    executar_historico_web(driver, datainicio, datafim)
    time.sleep(10)
    
    # Chamando código de extração do relatório NEXA BRASIL - Contratos por Empresa
    executar_contrato_empresa(driver)
    time.sleep(5)
    
    # Chamando código de extração do relatório NEXA BRASIL - Relatório Geral de Documentos
    executar_geral_documento(driver)
    
    time.sleep(200)
    resposta = ""
    while resposta not in ["1", "Sim", "sim", "SIM"]:
        fim = 1
        time.sleep(120)  # 120 segundos
        resposta = askstring("Informações", "Os relatórios foram baixados?\n1 - Sim\n2 - Não")
    
finally:

    driver.implicitly_wait(10)
    # Feche o navegador
    driver.quit()