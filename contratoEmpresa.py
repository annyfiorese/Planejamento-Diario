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


def executar_contrato_empresa(driver):
    try:
        mensagem = "contrato empresa"
        # Encontrar campo de reletorio para selecionar NEXA BRASIL - Contratos por Empresa
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_ddlRelatorio_ddlCombo", "511")
        driver.implicitly_wait(10)
        
        # Encontrar campo de Grupo terceiro e selecionar Grupo geral  
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_003_ddlCombo", "0001")
        driver.implicitly_wait(10)
        
        # Encontrando campo contratos vencidos há 
        campo_contratos_vencidos = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located(
                (By.ID, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_006_tbxEdit"))
        )
        # Insira o texto no contratos vencidos 
        campo_contratos_vencidos.send_keys("10000")
        
        # Encontrando campo contratos vencendo em  
        campo_contratos_vencendo = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located(
                (By.ID, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_007_tbxEdit"))
        )
        # Insira o texto no contratos vencendo em  
        campo_contratos_vencendo.send_keys("10000")
        
        driver.implicitly_wait(30)
        
        # Localize o  campo data inicio de inclusão para setar o data inicio
        campo_datainicio_nome = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.NAME, "ctl00$ctl00$MainContentPlaceHolder$cphMainColumn$tbxFiltro_008$mskEdit"))
            )

        # Aplicar um clique triplo no campo data inicio de inclusão
        action = ActionChains(driver)
        action.click_and_hold(campo_datainicio_nome).click_and_hold(campo_datainicio_nome).click_and_hold(campo_datainicio_nome).release().perform()

        # Inserir data no campo data inicio de inclusão
        campo_datainicio_nome.send_keys("01042017")
        
        driver.implicitly_wait(10)
        # Adicione uma pausa de 3 segundos
        time.sleep(4)
        
        # Encontrar campo de Grupo terceiro e selecionar Grupo geral  
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_003_ddlCombo", "0001")
        
        # Adicione uma pausa de 3 segundos
        time.sleep(4)
        
        mensagem = "gerar doc - contrato empresa"
         # Encontrar campo de gerar documento
        elemento_visualizador = WebDriverWait(driver, 250).until(
            EC.visibility_of_element_located((By.ID, "MainContentPlaceHolder_cphMainColumn_btnShowOptions"))
        )
        elemento_visualizador.click()
        
        mensagem = "formato excel - contrato empresa"
        # Encontrar campo de formato excel
        elemento_visualizador = WebDriverWait(driver, 250).until(
            EC.visibility_of_element_located((By.ID, "MainContentPlaceHolder_cphMainColumn_btnRelXls"))
        )
        elemento_visualizador.click()
        
        mensagem = "pop-up - contrato empresa"
        # Encontrar o x para fechar o pop-up de download
        elemento_visualizador = WebDriverWait(driver, 150).until(
           
            EC.visibility_of_element_located((By.ID, "MainContentPlaceHolder_cphMainColumn_btnCloseOptions"))
        )
        elemento_visualizador.click()
        
        driver.implicitly_wait(10)
        
    except:
       print ("Erro " + mensagem) 
       