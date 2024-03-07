import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains

def selecionar_opcao_com_tentativa(driver, element_id, valor, tentativas=3):
    for _ in range(tentativas):
        try:
            elemento_modulo = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, element_id))
            )
            dropdown = Select(elemento_modulo)
            dropdown.select_by_value(valor)
            return
        except StaleElementReferenceException:
            continue

def executar_historico_web(driver, datainicio, datafim):
    try:
        mensagem = "historico web"
        # Encontrar campo de reletorio para selecionar NEXA -Histórico de Lançamento de Documento via WEB
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_ddlRelatorio_ddlCombo", "428")
        
        # Encontrar campo de Grupo terceiro e selecionar Grupo geral  
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_003_ddlCombo", "0001")
        
        # Encontrar campo de status e selecionar todos 
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_007_ddlCombo", "1")
        
        # Encontrar campo de prestadores ativos e seleci0nar Apresentar ativos e inativos
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_012_ddlCombo", "0")
        
        # Encontrar campo de tipo de doc e selecionar todos 
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_013_ddlCombo", "2")

        # Localize o  campo data inicio de inclusão para setar o data inicio
        campo_datainicio_nome = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.NAME, "ctl00$ctl00$MainContentPlaceHolder$cphMainColumn$tbxFiltro_005$mskEdit"))
            )

        # Aplicar um clique triplo no campo data inicio de inclusão
        action = ActionChains(driver)
        action.click_and_hold(campo_datainicio_nome).click_and_hold(campo_datainicio_nome).click_and_hold(campo_datainicio_nome).release().perform()

        # Inserir data no campo data inicio de inclusão
        campo_datainicio_nome.send_keys(datainicio)
        
        # Localize o  campo data inicio de inclusão para setar o data dim
        campo_datafim_nome = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.NAME, "ctl00$ctl00$MainContentPlaceHolder$cphMainColumn$tbxFiltro_006$mskEdit"))
            )

        # Aplicar um clique triplo no campo data fim de inclusão
        action = ActionChains(driver)
        action.click_and_hold(campo_datafim_nome).click_and_hold(campo_datafim_nome).click_and_hold(campo_datafim_nome).release().perform()

        # Inserir data no campo data dim de inclusão
        campo_datafim_nome.send_keys(datafim)
        
        # Encontrar campo de Grupo terceiro e selecionar Grupo geral  
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_003_ddlCombo", "0001")
        
        driver.implicitly_wait(10)
        # Adicione uma pausa de 3 segundos
        time.sleep(4)
        
         # Encontrar campo de gerar documento
        mensagem = "gerar doc - historico web"
        elemento_visualizador = WebDriverWait(driver, 250).until(
            EC.visibility_of_element_located((By.ID, "MainContentPlaceHolder_cphMainColumn_btnShowOptions"))
        )
        elemento_visualizador.click()
        
        # Encontrar campo de formato excel
        mensagem = "formato excel - historico web"
        elemento_visualizador = WebDriverWait(driver, 400).until(
            EC.visibility_of_element_located((By.ID, "MainContentPlaceHolder_cphMainColumn_btnRelXls"))
        )
        elemento_visualizador.click()
        
        # Encontrar o x para fechar o pop-up de download
        mensagem = "pop- up- historico web"
        elemento_visualizador = WebDriverWait(driver, 150).until(
            EC.visibility_of_element_located((By.ID, "MainContentPlaceHolder_cphMainColumn_btnCloseOptions"))
        )
        elemento_visualizador.click()
        
        driver.implicitly_wait(10)
    
        # Adicione uma pausa de 3 segundos
        time.sleep(4)
        
    except:
       print ("Erro " + mensagem) 
       