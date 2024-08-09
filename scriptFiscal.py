from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
import getpass
import os
import csv
import sqlite3
import shutil
import time
import sys
import calendar
import pandas as pd
import mysql.connector
from mysql.connector import Error
from datetime import datetime, date
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.service import Service as ChromeService
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from time import sleep
from datetime import datetime, timedelta
from selenium.webdriver.support.ui import Select
from pynput.keyboard import Key, Controller
import re
from shareplum import Site
from shareplum.site import Version
import pyautogui
import chardet


def fiscal():
    
    navegador = webdriver.Chrome()

    url = "https://www.google.com.br/"

    url_unidocs = "https://unidocs.ambev.com.br/#/login"

    url_entrar_share_point ='https://login.microsoftonline.com/1824dd00-dd6f-44fe-baec-00e5460bc813/oauth2/authorize?client%5Fid=00000003%2D0000%2D0ff1%2Dce00%2D000000000000&response%5Fmode=form%5Fpost&response%5Ftype=code%20id%5Ftoken&resource=00000003%2D0000%2D0ff1%2Dce00%2D000000000000&scope=openid&nonce=01E1987B3D6E89CB8F3164D93B36F17CFFAAAD765015DC98%2D201826037737AC37433077D90158F26518608387E7E95F9B9E741AF05A495A7F&redirect%5Furi=https%3A%2F%2Fgrupohorizonte%2Esharepoint%2Ecom%2F%5Fforms%2Fdefault%2Easpx&state=ND1UcnVlJjg9MA&claims=%7B%22id%5Ftoken%22%3A%7B%22xms%5Fcc%22%3A%7B%22values%22%3A%5B%22CP1%22%5D%7D%7D%7D&wsucxt=1&cobrandid=11bd8083%2D87e0%2D41b5%2Dbb78%2D0bc43c8a8e8a&client%2Drequest%2Did=0c5415a1%2D507a%2D5000%2D0b90%2Df4a2014fccc4'

    url_CTE = "https://grupohorizonte.sharepoint.com/sites/spinteligencia/Documentos%20Compartilhados/Forms/AllItems.aspx?id=%2Fsites%2Fspinteligencia%2FDocumentos%20Compartilhados%2F01%5FControladoria%2F22%5FFiscal%2FCTRC&viewid=8a2d1e42%2D2f5e%2D43f5%2D8cfd%2D16d4ab81e2d7"

    url_NFE = "https://grupohorizonte.sharepoint.com/sites/spinteligencia/Documentos%20Compartilhados/Forms/AllItems.aspx?id=%2Fsites%2Fspinteligencia%2FDocumentos%20Compartilhados%2F01%5FControladoria%2F22%5FFiscal%2FNFS&viewid=8a2d1e42%2D2f5e%2D43f5%2D8cfd%2D16d4ab81e2d7"



    def config_options(usuario):
        download_directory = "C:\\Arquivos_Fiscais\\"
        
        options = Options() 
        options.add_argument('--disable-gpu')
        options.add_argument("--start-maximized")
        options.add_argument("--force-device-scale-factor=0.8")
        options.add_argument("--disable-javascript")
        
        prefs = {
            "download.default_directory": download_directory,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
            
        }
        options.add_experimental_option("prefs", prefs)
        return options


    input_email_unidocs = "/html/body/div[2]/div/div[2]/div/form/div[1]/div[2]/input"
    input_senha_unidocs = "/html/body/div[2]/div/div[2]/div/form/div[2]/div[2]/input"
    button_entrar = "/html/body/div[2]/div/div[2]/div/form/button"
    button_documentos = "/html/body/div[4]/nav/ul/li[6]/a"
    button_CTe = "/html/body/div[4]/nav/ul/li[6]/ul/li[1]/a"
    button_data_inicial = "//input[contains(@class,'form-control no-spin')]" 
    button_data_final = "(//input[contains(@class,'form-control no-spin')])[2]"
    button_exportar_CSV = "/html/body/div[5]/div[2]/div/div[3]/ng-include/form/div[3]/div/button[1]"
    button_NFe = "/html/body/div[4]/nav/ul/li[6]/ul/li[2]/a" 
    button_MDFe = "/html/body/div[4]/nav/ul/li[6]/ul/li[3]/a"
    button_NFe_csv_exportar = "/html/body/div[5]/div[2]/div/div[3]/ng-include/div[1]/div/button[1]"
    button_MDFe_csv_exportar = "/html/body/div[5]/div[2]/div/div[3]/ng-include/form/div[2]/button[1]"
    xpath_entrar_share_point = "/html/body/form/div[4]/div[2]/div/div/div[2]/div[2]/div[2]/div[1]/div[2]/a"
    url_email_para_sharepoint = '/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div[2]/div/div/div[2]/div/div/div[2]'
    xpath_input_email_share_point = "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[2]/div[2]/div/input[1]"
    xpath_avancar_shahere_point = "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[4]/div/div/div/div/input"
    xpath_input_senha_share_point = "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[3]/div/div[2]/input"
    xpath_entrar_email_share_point = "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[5]/div/div/div/div/input"
    xpath_meus_sites = "(//a[@class='sp-appBar-link'])[2]"
    xpath_SP_INTELIGENCIA = "//a[@title='SP-Inteligência']"
    xpath_documentos_SP_INT = "(//span[@class='ms-Nav-linkText linkText_4f56e93d'])[3]"
    xpath_sim = "/html/body/div/form/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[3]/div[2]/div/div/div[2]/input"







    "-------------------------------------------------------------------------------------------------"

    def abrir_navegador():
        global url,navegador
        navegador.get(url)
        
    def change_tab_with_interactions(driver, url):
        # Salvar a janela atual
        current_window = driver.current_window_handle
        
        # Abrir uma nova aba e mudar para ela
        driver.execute_script("window.open('" + url + "', '_blank');")
        
        # Mudar o foco para a nova aba
        for window_handle in driver.window_handles:
            if window_handle != current_window:
                driver.switch_to.window(window_handle)
                break

    def click_js(elemento):
        navegador.execute_script("arguments[0].click();", elemento)         


    def wait(time):
        global navegador
        return WebDriverWait(navegador, time)


    def click(elemento):
        try:
            wait(50).until(EC.visibility_of_element_located((By.XPATH, elemento))).click()
        except Exception as e:
            print(f"Erro ao clicar no elemento: {e}")

            
    def escreva_digitando(elemento, valor):
        campo = wait(50).until(
            EC.visibility_of_element_located((By.XPATH, elemento))
        )
        
        # Limpar o campo
        campo.clear()

        # Escrever o valor letra por letra 
        for letra in valor:
            campo.send_keys(letra) 
            time.sleep(0.1)


    def entrarIframe(xpath):
        wait(50).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, xpath))
        )

        
    def sairIframe():
        navegador.switch_to.default_content()

    def hoveraction(driver, elemento):
        # Use o ActionChains para passar o mouse sobre o elemento
        ActionChains(driver).move_to_element(elemento).perform()    
        

    def change_tab_with_interactions(driver, url):
        # Salvar a janela atual
        current_window = driver.current_window_handle
        
        # Abrir uma nova aba e mudar para ela
        driver.execute_script("window.open('" + url + "', '_blank');")
        
        # Mudar o foco para a nova aba
        for window_handle in driver.window_handles:
            if window_handle != current_window:
                driver.switch_to.window(window_handle)
                break



    def hover(elemento):
        wait(50).until(
            EC.presence_of_element_located((By.XPATH, elemento))
        )
        elemento_hover = navegador.find_element(By.XPATH, elemento)
        ActionChains(navegador).move_to_element(elemento_hover).perform() 

    def fechar_primeira_janela(driver):
        ## Certificar-se de que existem pelo menos duas janelas abertas
        if len(driver.window_handles) >= 1:
            ## Trocar para a primeira janela
            driver.switch_to.window(driver.window_handles[0])
            ## Fechar a primeira janela
            driver.close()
            ## Se houver mais de uma janela, trocar para a próxima janela disponível 
            if len(driver.window_handles) > 0:
                driver.switch_to.window(driver.window_handles[0])
        else:
            print("Não há janelas abertas para fechar.") 


    def scroll_down(driver):
        try:
            # Rola a página para baixo usando as teclas de seta 
            body = driver.find_element(By.TAG_NAME, 'body')
            body.send_keys(Keys.ARROW_DOWN)
            time.sleep(1)  # Aguarda um segundo para a rolagem ser executada
            # print("Rolou a página para baixo.")
        except Exception as e:
            print(f"Erro ao rolar a página para baixo: {e}")
            

    def verificar_arquivos_no_diretorio(nomes_arquivos, diretorio="C:\\Arquivos_Controladoria\\"):
        # Lista para armazenar os arquivos que estão presentes no diretório
        arquivos_presentes = []

        # Iterar sobre a lista de nomes de arquivos
        for nome_arquivo in nomes_arquivos:
            # Construir o caminho completo do arquivo
            caminho_arquivo = os.path.join(diretorio, nome_arquivo)
            
            # Verificar se o arquivo existe
            if os.path.isfile(caminho_arquivo):
                arquivos_presentes.append(nome_arquivo)
            # else:
            # print(f"O arquivo {nome_arquivo} não foi encontrado no diretório {diretorio}")

        return arquivos_presentes

    def get_data_primeiro_dia_mes_atual():
                # Obtém a data atual
                data_atual = datetime.now()
                
                # Formata a data para o primeiro dia do mês atual
                primeiro_dia_mes_atual = data_atual.replace(day=1)
                
                # Formata a data no formato desejado (dd/mm/aaaa)
                primeiro_dia_mes_formatado = primeiro_dia_mes_atual.strftime("%d/%m/%Y")
                
                return primeiro_dia_mes_formatado          
                'SELECT DISTINTIC FROM ESTHALL WHERE CODFIL = 87 AND CODCGA =23'
    def get_data_ultimo_dia_mes_atual():
        # Obtém a data atual
        data_atual = datetime.now()
        ultimo_dia_mes = calendar.monthrange(data_atual.year, data_atual.month)[1]
        # Formata a data para o primeiro dia do mês atual
        ultimo_dia_mes_atual = data_atual.replace(day=ultimo_dia_mes)
        
        # Formata a data no formato desejado (aaaa/mm/dd)
        ultimo_dia_mes_formatado = ultimo_dia_mes_atual.strftime("%d/%m/%Y")
        
        return ultimo_dia_mes_formatado            

    def get_data_dia_atual():
                return date.today().strftime('%d/%m/%Y')
    

    def clear_field(driver, xpath):
                try:
                    element = driver.find_element(By.XPATH, xpath)
                    element.clear()
                    print("Campo limpo com sucesso.")
                except Exception as e:
                    print(f"Erro ao limpar o campo: {e}")    


    def escreve_primeiro_dia_mes_atual():
        # Obtém a data atual
        data_atual = datetime.now()
        
        # Formata a data para o primeiro dia do mês atual
        primeiro_dia_mes_atual = data_atual.replace(day=1)
        
        # Formata a data no formato desejado (dd/mm/aaaa)
        primeiro_dia_mes_formatado = primeiro_dia_mes_atual.strftime("%d/%m/%Y")
        
        # Usa o pyautogui para digitar a data onde o cursor de texto está posicionado
        pyautogui.typewrite(primeiro_dia_mes_formatado)


    def escreve_ultimo_dia_mes_atual():
        # Obtém a data atual
        data_atual = datetime.now()
        
        # Calcula o último dia do mês atual
        ultimo_dia_mes = calendar.monthrange(data_atual.year, data_atual.month)[1]
        
        # Cria uma data com o último dia do mês atual
        ultimo_dia_mes_atual = data_atual.replace(day=ultimo_dia_mes)
        
        # Formata a data no formato desejado (dd/mm/yyyy)
        ultimo_dia_mes_formatado = ultimo_dia_mes_atual.strftime("%d/%m/%Y")
        
        # Usa o pyautogui para digitar a data onde o cursor de texto está posicionado 
        pyautogui.typewrite(ultimo_dia_mes_formatado)

    ## 
    def excluir_arquivos_em_pasta(caminho_pasta):
        try:
            # Listar todos os arquivos na pasta
            arquivos = os.listdir(caminho_pasta)

            # Iterar sobre os arquivos e excluí-los
            for arquivo in arquivos:
                caminho_completo = os.path.join(caminho_pasta, arquivo)
                
                # Verificar se é um arquivo (não é um diretório)
                if os.path.isfile(caminho_completo):
                    os.remove(caminho_completo)
                    print(f'Arquivo excluído: {caminho_completo}')
                    
            print('Todos os arquivos foram excluídos.')
        
        except Exception as e:
            print(f'Erro ao excluir arquivos: {e}')



    def get_data_dia_atual():
        return date.today().strftime('%d/%m/%Y')


    def renomear_cte(caminho_arquivo):
        endereco = os.path.dirname(caminho_arquivo)
        data_atual = datetime.now().strftime("%Y-%m-%d")
        nome_novo_arquivo = os.path.join(endereco, f"{data_atual}.csv")
        
        # Renomeia o arquivo
        print(f"Renomeando o arquivo de {caminho_arquivo} para {nome_novo_arquivo}.")
        os.rename(caminho_arquivo, nome_novo_arquivo)
        print(f"Arquivo renomeado para {nome_novo_arquivo}")

        

    def renomear_nfe():
        endereco = "C:\\Arquivos_Fiscais\\NFes\\"
        arquivo_antigo = os.path.join(endereco, "nome_do_arquivo.xlsx")  # Substitua "nome_do_arquivo.xlsx" pelo nome atual do arquivo
        
        # Obtém a data atual no formato desejado (por exemplo, AAAA-MM-DD)
        data_atual = datetime.now().strftime("%Y-%m-%d")
        nome_novo_arquivo = os.path.join(endereco, f"{data_atual}.csv")
        
        # Renomeia o arquivo
        os.rename(arquivo_antigo, nome_novo_arquivo)
        print(f"Arquivo renomeado para {nome_novo_arquivo}")

        # Espera 15 segundos
        time.sleep(15)




    def renomear_mdfe():
        endereco = "C:\\Arquivos_Fiscais\\MDFes\\"
        arquivo_antigo = os.path.join(endereco, "nome_do_arquivo.xlsx")  # Substitua "nome_do_arquivo.xlsx" pelo nome atual do arquivo
        
        # Obtém a data atual no formato desejado (por exemplo, AAAA-MM-DD)
        data_atual = datetime.now().strftime("%Y-%m-%d")
        nome_novo_arquivo = os.path.join(endereco, f"{data_atual}.csv")
        
        # Renomeia o arquivo
        os.rename(arquivo_antigo, nome_novo_arquivo)
        print(f"Arquivo renomeado para {nome_novo_arquivo}")

        # Espera 15 segundos
        time.sleep(15)



    def mover_arquivo(tipo):
        origem = "C:\\Arquivos_Fiscais\\"
        if tipo == "cte":
            endereco_destino = "C:\\Arquivos_Fiscais\\CTes\\"
        if tipo == "mdfe":
            endereco_destino = "C:\\Arquivos_Fiscais\\MDFes\\"
        if tipo == "nfe":
            endereco_destino = "C:\\Arquivos_Fiscais\\NFes\\"
            
        
        
        # Verifica se a pasta de destino existe, caso contrário, cria a pasta
        if not os.path.exists(endereco_destino):
            os.makedirs(endereco_destino)
            print(f"Pasta {endereco_destino} criada.")
        
        # Move o arquivo
        #print(f"Movendo o arquivo de {caminho_arquivo_origem} para {caminho_arquivo_destino}.")
        #shutil.move(caminho_arquivo_origem, caminho_arquivo_destino)
        #print(f"Arquivo {nome_arquivo} movido para a pasta CTes.")
        #return caminho_arquivo_destino
        for arquivo in os.listdir(origem):
        # Verifica se o arquivo termina com '.csv'
            if arquivo.endswith('.csv'):
                # Caminho completo do arquivo
                match = re.search(r'\d{4}-\d{2}-\d{2}', arquivo)
                if match:
                # Extrair a data do nome do arquivo
                    data = match.group(0)
            
                    # Definir o novo nome do arquivo
                    novo_nome = f"{data}.csv"
                
                    #novo_nome = arquivo[:-13] + '.csv'
                    caminho_arquivo = os.path.join(origem, arquivo)
                    novo_caminho_arquivo = os.path.join(endereco_destino, novo_nome)
                
                    # Move o arquivo para a pasta de destino
                    shutil.move(caminho_arquivo, novo_caminho_arquivo)
                    print(f"Arquivo {caminho_arquivo} movido para a pasta {endereco_destino}.")



    def mover_arquivo_nfe(nome_arquivo):
        endereco_origem = "C:\\Arquivos_Fiscais\\"
        endereco_destino = "C:\\Arquivos_Fiscais\\NFes\\"
        
        if not os.path.exists(endereco_destino):
            os.makedirs(endereco_destino)
        shutil.move(endereco_origem + nome_arquivo, endereco_destino + nome_arquivo)
        print(f"Arquivo {nome_arquivo} movido para a pasta NFes.")



    def mover_arquivo_mdfe(nome_arquivo):
        endereco_origem = "C:\\Arquivos_Fiscais\\"
        endereco_destino = "C:\\Arquivos_Fiscais\\MDFes\\"
        
        if not os.path.exists(endereco_destino):
            os.makedirs(endereco_destino)
        shutil.move(endereco_origem + nome_arquivo, endereco_destino + nome_arquivo)
        print(f"Arquivo {nome_arquivo} movido para a pasta MDFe.")


    def renomear_arquivo_para_data_csv(nome_arquivo_atual):
        endereco_origem = "C:\\Arquivos_Fiscais"
        caminho_arquivo_atual = os.path.join(endereco_origem, nome_arquivo_atual)
        novo_caminho_arquivo = os.path.join(endereco_origem, "data.csv")
        
        # Verifica se o arquivo de origem existe
        if not os.path.exists(caminho_arquivo_atual):
            print(f"Erro: O arquivo {caminho_arquivo_atual} não foi encontrado.")
            return False
        
        # Renomeia o arquivo para data.csv
        os.rename(caminho_arquivo_atual, novo_caminho_arquivo)
        print(f"Arquivo {nome_arquivo_atual} renomeado para data.csv.")
        return True


    def buscar_arquivo_por_data(diretorio):
        # Obtém a data atual no formato desejado (por exemplo, AAAA-MM-DD)
        data_atual = datetime.now().strftime("%Y-%m-%d")
        nome_arquivo = f"{data_atual}.csv"
        caminho_arquivo = os.path.join(diretorio, nome_arquivo)
        
        # Verifica se o arquivo existe
        if os.path.exists(caminho_arquivo):
            return caminho_arquivo
        else:
            print(f"Erro: O arquivo {caminho_arquivo} não foi encontrado.")
            return None


    def digitar_caminho_arquivo(caminho):
        pyautogui.typewrite(caminho)
        time.sleep(4)
        pyautogui.press('enter')
        print(f"O arquivo está na pasta do share point")
        print(f"Arquivo movido {caminho}")




    def excluir_arquivos_em_pasta(caminho_pasta):
        try:
            # Listar todos os arquivos na pasta
            arquivos = os.listdir(caminho_pasta)

            # Iterar sobre os arquivos e excluí-los
            for arquivo in arquivos:
                caminho_completo = os.path.join(caminho_pasta, arquivo)
                
                # Verificar se é um arquivo (não é um diretório)
                if os.path.isfile(caminho_completo):
                    os.remove(caminho_completo)
                    print(f'Arquivo excluído: {caminho_completo}')
                    
            print('Todos os arquivos foram excluídos.')
        except Exception as e:
            print(f'Ocorreu um erro: {e}')
        finally:
            print(f'Execução da função excluir_arquivos_em_pasta concluída.')



    pasta_cte = "C:\\Arquivos_Fiscais\\CTes\\"
    excluir_arquivos_em_pasta(pasta_cte)

    pasta_nfe = "C:\\Arquivos_Fiscais\\NFes\\"
    excluir_arquivos_em_pasta(pasta_nfe)



    usuario = getpass.getuser()
    options = config_options(usuario)
    navegador = webdriver.Chrome(options=options) ##DRIVER DA FUNÇÃO = NAVEGADOR 
    abrir_navegador()
    time.sleep(5)
    change_tab_with_interactions(navegador,url_unidocs)
    time.sleep(25)
    fechar_primeira_janela(navegador)
    time.sleep(10)
    click(input_email_unidocs)
    time.sleep(3)
    escreva_digitando(input_email_unidocs,"flavia.holanda@grupohorizonte.com.br")
    time.sleep(2)
    click(input_senha_unidocs)
    time.sleep(2)
    escreva_digitando(input_senha_unidocs,"AmbeV@1234*")
    time.sleep(2)
    click(button_entrar)
    time.sleep(6)
    click(button_documentos)
    time.sleep(3)
    click(button_CTe)
    time.sleep(3)
    click(button_data_inicial)
    time.sleep(2)
    escreve_primeiro_dia_mes_atual()
    time.sleep(15)
    click(button_data_final)
    time.sleep(3)
    escreve_ultimo_dia_mes_atual()
    time.sleep(15)
    click(button_exportar_CSV)
    time.sleep(35)
    print("CT'es Baixados de " + get_data_dia_atual())

    mover_arquivo("cte")
    time.sleep(15)


    ##CONSULTA NFS-E

    click(button_NFe)
    time.sleep(8)
    click(button_data_inicial)
    time.sleep(5) 
    escreve_primeiro_dia_mes_atual()
    time.sleep(15)
    click(button_data_final)
    time.sleep(3)
    escreve_ultimo_dia_mes_atual()
    time.sleep(15)
    click(button_NFe_csv_exportar)
    time.sleep(35)
    print("NF'es Baixados de " + get_data_dia_atual())

    mover_arquivo("nfe")
    time.sleep(15)


    # ###Up SHARE POINT

    change_tab_with_interactions(navegador,url_entrar_share_point)
    time.sleep(10)
    fechar_primeira_janela(navegador)
    time.sleep(10)
    click(xpath_input_email_share_point) 
    time.sleep(10)
    escreva_digitando(xpath_input_email_share_point,"joao.tavares@grupohorizonte.com.br")
    time.sleep(10)
    click(xpath_avancar_shahere_point)
    time.sleep(10)
    click(xpath_input_senha_share_point) 
    time,sleep(4)
    escreva_digitando(xpath_input_senha_share_point,"Shinji@0305")
    time.sleep(15)
    click(xpath_entrar_email_share_point)
    wait(10)

    change_tab_with_interactions(navegador,url_CTE)
    time.sleep(20)
    fechar_primeira_janela(navegador)
    click("(//i[@data-icon-name='ChevronDown'])[2]")
    time.sleep(10)
    click("//span[text()='Arquivos']")
    time.sleep(5) 

    diretorio_cte = "C:\\Arquivos_Fiscais\\CTes\\"
    for arquivo in os.listdir(diretorio_cte):
        if os.path.isfile(os.path.join(diretorio_cte,arquivo)):
            digitar_caminho_arquivo(diretorio_cte+arquivo)
    time.sleep(30)


    change_tab_with_interactions(navegador,url_NFE)
    time.sleep(20)
    fechar_primeira_janela(navegador)
    click("(//i[@data-icon-name='ChevronDown'])[2]")
    time.sleep(10)
    click("//span[text()='Arquivos']")
    time.sleep(5) 

    diretorio_nfe = "C:\\Arquivos_Fiscais\\NFes\\"
    for arquivo in os.listdir(diretorio_nfe):
        if os.path.isfile(os.path.join(diretorio_nfe, arquivo)):
            digitar_caminho_arquivo(os.path.join(diretorio_nfe, arquivo))
            time.sleep(10)


    time.sleep(200)


