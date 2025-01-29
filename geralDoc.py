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
            elemento_modulo = WebDriverWait(driver, 20).until(
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
        
        # Encontrar campo de tipo de documento e selecionar ambos 
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_011_ddlCombo", "0")
        
        # Encontrar campo de Grupo terceiro e selecionar Grupo geral  
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_003_ddlCombo", "0001")
        
        # Encontrar campo de exibir subcontratada e selecionar Sim 
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_008_ddlCombo", "1")
        
        # Encontrar campo de bloqueia acesso e selecionar ambos 
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_009_ddlCombo", "A")
        
        # Encontrar campo de Grupo terceiro e selecionar Grupo geral  
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_003_ddlCombo", "0001")
        
        # Adicione uma pausa de 3 segundos
        time.sleep(1)
        # Encontrar campo de tipo de documento e selecionar ambos 
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_015_ddlCombo", "0")
        # Adicione uma pausa de 3 segundos
        time.sleep(2)
        
          # Encontrando campo de dias a vencer 
        campo_dias_a_vencer = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located(
                (By.ID, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_005_tbxEdit"))
        )
        # Insira o texto no dias a vencer 
        campo_dias_a_vencer.send_keys("10")
        # Adicione uma pausa de 3 segundos
        time.sleep(5)
        
        # Encontrar campo de exibir subcontratada e selecionar Sim 
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_008_ddlCombo", "1")
        
        # Encontrar campo de bloqueia acesso e selecionar ambos 
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_009_ddlCombo", "A")
        
        # Encontrar campo de tipo de documento e selecionar ambos 
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphMainColumn_tbxFiltro_015_ddlCombo", "0")
        
        
        
        # Encontrar campo de gerar documento
        mensagem = "gerar doc - geral doc"
        encontrar_elemento_e_clicar(driver, "MainContentPlaceHolder_cphMainColumn_btnShowOptions", 60)

        time.sleep(50)
        
        mensagem = "formato excel - geral doc"
        # Encontrar campo de formato excel
        elemento_visualizador = WebDriverWait(driver, 250).until(
            EC.visibility_of_element_located((By.ID, "MainContentPlaceHolder_cphMainColumn_btnRelXls"))
        )
        elemento_visualizador.click()
        
        mensagem = "pop-up - geral doc"
        # Encontrar o x para fechar o pop-up de download
        elemento_visualizador = WebDriverWait(driver, 150).until(
           
            EC.visibility_of_element_located((By.ID, "MainContentPlaceHolder_cphMainColumn_btnCloseOptions"))
        )
        elemento_visualizador.click()
        
        driver.implicitly_wait(10)
        
        
        # Encontrar o x para fechar o pop-up de download
        
        
        encontrar_elemento_e_clicar(driver, "MainContentPlaceHolder_cphMainColumn_btnCloseOptions", 150)
    
        
        driver.implicitly_wait(10)
        
    except Exception as e:
        print(f"Erro: {mensagem}. Detalhes: {str(e)}")
        