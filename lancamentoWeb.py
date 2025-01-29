import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

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

def executar_lancamento_web(driver, datainicio, datafim):
    try:
        mensagem = "lançamento web"
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_ddlRelatorio_ddlCombo", "425")
        time.sleep(1)
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_003_ddlCombo", "0001")
        time.sleep(1)
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_007_ddlCombo", "1")
        time.sleep(2)
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_012_ddlCombo", "0")
        
        driver.implicitly_wait(10)
        time.sleep(4)

        campo_datainicio_nome = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.NAME, "ctl00$ctl00$MainContentPlaceHolder$cphMainColumn$tbxFiltro_005$mskEdit"))
        )
        action = ActionChains(driver)
        action.click_and_hold(campo_datainicio_nome).click_and_hold(campo_datainicio_nome).click_and_hold(campo_datainicio_nome).release().perform()
        campo_datainicio_nome.send_keys(datainicio)
        
        campo_datafim_nome = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.NAME, "ctl00$ctl00$MainContentPlaceHolder$cphMainColumn$tbxFiltro_006$mskEdit"))
        )
        action.click_and_hold(campo_datafim_nome).click_and_hold(campo_datafim_nome).click_and_hold(campo_datafim_nome).release().perform()
        campo_datafim_nome.send_keys(datafim)
        
        mensagem = "gerar doc - lançamento web"
        elemento_visualizador = WebDriverWait(driver, 250).until(
            EC.visibility_of_element_located((By.ID, "MainContentPlaceHolder_cphMainColumn_btnShowOptions"))
        )
        elemento_visualizador.click()
        
        mensagem = "formato excel- lançamento web"
        elemento_visualizador = WebDriverWait(driver, 250).until(
            EC.visibility_of_element_located((By.ID, "MainContentPlaceHolder_cphMainColumn_btnRelXls"))
        )
        elemento_visualizador.click()
        
        mensagem = "pop-up - lançamento web"
        elemento_visualizador = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "MainContentPlaceHolder_cphMainColumn_btnCloseOptions"))
        )
        elemento_visualizador.click()
        
        driver.implicitly_wait(10)
    
    except Exception as e:
        print(f"Erro {mensagem}: {e}")
