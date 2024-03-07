import tkinter
from tkinter.simpledialog import askstring
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException

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



def executar_navegação(driver):

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
        
        # Encontrando campo de usuário
        campo_usuario = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located(
                (By.ID, "MainContentPlaceHolder_tbxUsuario_tbxEdit"))
        )
        # Insira o texto no campo de usuário
        campo_usuario.send_keys(usuario)

        # Encontrando campo de senha
        campo_senha = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located(
                (By.ID, "MainContentPlaceHolder_tbxSenha_tbxEdit"))
        )
        # Insira o texto no campo de senha
        campo_senha.send_keys(senha)

        # Buscando o botão de login
        botao_login = WebDriverWait(driver, 60).until(
            EC.visibility_of_element_located(
                (By.ID, "MainContentPlaceHolder_btnLogin"))
        )
        # Clicando para fazer o login
        botao_login.click()
        
        # Localize o elemento pelo ID "liLeftMenuModulo7" e clique nele
        elemento_modulo = WebDriverWait(driver, 60).until(
            EC.visibility_of_element_located((By.ID, "liLeftMenuModulo7"))
        )
        elemento_modulo.click()

        # Após clicar em "liLeftMenuModulo7", localize e clique no link "Visualizador de Relatórios"
        elemento_visualizador = WebDriverWait(driver, 60).until(
            EC.visibility_of_element_located(
                (By.LINK_TEXT, "Visualizador de Relatórios"))
        )
        elemento_visualizador.click()
        
        
        # Localize o campo de seleção de relatório terceiro
        selecionar_opcao_com_tentativa(driver, "MainContentPlaceHolder_cphLeftColumn_ddlModulo_ddlCombo", "03")

        
    finally:
        print("ERRO PAGINA DE NAVEGAÇÃO")