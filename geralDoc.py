import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException

def encontrar_elemento_e_clicar(driver, id, timeout=10):
    try:
        elemento_visualizador = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((By.ID, id))
        )
        elemento_visualizador.click()
    except TimeoutException:
        print(f"Elemento com ID '{id}' não encontrado após {timeout} segundos.")
        # Adicione aqui o tratamento adequado, como registrar ou levantar uma exceção personalizada.

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

def esperar_download_concluir(driver, nome_do_arquivo, timeout=30):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, f"//a[@href='{nome_do_arquivo}']"))
        )
        print("Download concluído com sucesso!")
    except TimeoutException:
        print("Tempo limite atingido. Download pode não ter sido concluído.")

def executar_geral_documento(driver):
    try:
        mensagem = "geral doc"
        # Encontrar campo de reletorio para selecionar NEXA BRASIL - Relatório Geral de Documentos
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_ddlRelatorio_ddlCombo", "430")
        
        # Encontrar campo de Grupo terceiro e selecionar Grupo geral  
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_003_ddlCombo", "0001")
        
        # Encontrar campo de exibir subcontratada e selecionar Sim 
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_008_ddlCombo", "1")
        
        # Encontrar campo de bloqueia acesso e selecionar ambos 
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_009_ddlCombo", "A")
        
        # Encontrar campo de tipo de documento e selecionar ambos 
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_015_ddlCombo", "0")
        
        # Encontrar campo de Grupo terceiro e selecionar Grupo geral  
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_003_ddlCombo", "0001")
        
        # Adicione uma pausa de 3 segundos
        time.sleep(4)
        
         # Encontrando campo de dias a vencer 
        campo_dias_a_vencer = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located(
                (By.ID, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_005_tbxEdit"))
        )
        # Insira o texto no dias a vencer 
        campo_dias_a_vencer.send_keys("10")
        
        # Encontrar campo de tipo de documento e selecionar ambos 
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_015_ddlCombo", "0")
        
        driver.implicitly_wait(10)
        # Adicione uma pausa de 3 segundos
        time.sleep(4)
        
        # Encontrar campo de exibir subcontratada e selecionar Sim 
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_008_ddlCombo", "1")
        
        # Encontrar campo de bloqueia acesso e selecionar ambos 
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_009_ddlCombo", "A")
        
        driver.implicitly_wait(10)
        # Adicione uma pausa de 3 segundos
        time.sleep(4)
        
        # Encontrar campo de gerar documento
        mensagem = "gerar doc - geral doc"
        encontrar_elemento_e_clicar(driver, "MainContentPlaceHolder_cphMainColumn_btnShowOptions", 300)

        time.sleep(300)
        
        # Encontrar campo de formato excel
        procurandoelement_xpath = '//*[@id="MainContentPlaceHolder_cphMainColumn_btnRelXls"]'

        tempo_limite = 500
        
        tempo_inicial = time.time()
        
        while time.time() - tempo_inicial < tempo_limite:
            try:
                elemento = driver.find_element(By.XPATH, procurandoelement_xpath)
                break
            except NoSuchElementException:   
                time.sleep(2)
      
      # Encontrar campo de Pop-Up
        elementPopUp_xpath = '//*[@id="MainContentPlaceHolder_cphMainColumn_btnCloseOptions"]'

        tempo_limite = 300
        
        tempo_inicial = time.time()
        
        while time.time() - tempo_inicial < tempo_limite:
            try:
                elemento = driver.find_element(By.XPATH, elementPopUp_xpath)
                break
            except NoSuchElementException:   
                time.sleep(2) 
                
        driver.implicitly_wait(10)
    
        # Adicione uma pausa de 3 segundos
        time.sleep(7)
             
        
    except Exception as e:
        print(f"Erro: {mensagem}. Detalhes: {str(e)}")
        # Adicione tratamento adequado para o erro, como registrar ou levantar uma exceção personalizada
