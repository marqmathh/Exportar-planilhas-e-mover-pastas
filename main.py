from selenium import webdriver
from turtle import position
import pyautogui
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time
import os
import shutil
import win32com.client

def exportar():
    time.sleep(30)
    try:
        
        export_button = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="ZDBExportButton"]'))
        )
        export_button.click()

        embed_export_button = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="EmbedExportXLSMenuItem"]'))
        )
        time.sleep(1)
        embed_export_button.click()
        print("Botão 'EmbedExportXLSMenuItem' clicado com sucesso!")
        
        final_button = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[32]/div/div[2]/div/footer/div/button[1]'))
        )
        time.sleep(1)
        final_button.click()

        time.sleep(1)

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

def mudar_Link(texto,i):
    navegador.execute_script("window.open('');")
    navegador.switch_to.window(navegador.window_handles[i])
    navegador.get(texto)

def exporta_Nomus():
    navegador.get('https://reports.nomus.com.br/open-view/751489003498986462') # PC - Follow up
    time.sleep(10)
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003495455916',1) # SC COMPLETA
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003499869977',2) # Cotação
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003499311067',3) # NF Tipo de movimentação
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003497642590',4) # Fornecedores - TB-06
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003385051213',5) # Ordens Lead
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003418257074',6) # Necessidade
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003348744953',7) # ID-04e
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003473916593',8) # Tempo de produção
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003396077915',9) # Entradas
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003343390887',10) # Pedido de compra
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003375648746',11) # Pv's
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003343401232',12) # SC's
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003343402198',13) # Entradas compras
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003383877015',14) # Faturamento periodo
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003384381506',15) # CRM
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003381705320',16) # Vendas
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003396836362',17) # Emissão de PV
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003384833393',18) # Status do PV
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003501209517',19) # Familia
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003299922759',20) # CONTROLE DA PRODUÇÃO
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003399788047',21) # ENTRADAS (LIVRO)
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003415814176',22) # Nfs emitidas
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003399436517',23) # Saída 
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003406446494',24) # Dw Compras
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003342815143',25) # Modelo NF's 
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003415906765',26) # Grupos
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003432511966',27) # movimentação - pvs
    exportar()

    time.sleep(60)
    navegador.quit()

def move_Arquivos():
    for arquivo in os.listdir(pasta_destino):
        caminho_arquivo = os.path.join(pasta_destino, arquivo)
        if os.path.isfile(caminho_arquivo):
            os.remove(caminho_arquivo)

    for arquivo in os.listdir(pasta_downloads):
        caminho_arquivo = os.path.join(pasta_downloads, arquivo)
        if os.path.isfile(caminho_arquivo):
            shutil.move(caminho_arquivo, pasta_destino)

    print('Arquivos deletados e novos arquivos movidos com sucesso!')    

def titulo():
    pyautogui.alert("""O código vai começar.

Favor não útilizar o PC até a finalização.""")

def finalizacao():
    pyautogui.alert("""O código foi finalizado.

Não esqueça de pressionar o botão de atualizar nas suas planilhas.

Obrigado por aguardar. """)

def atualizar_Planilhas(local_Planilha):
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = True

    try:
        wb = excel.Workbooks.Open(local_Planilha)
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        wb.Close(SaveChanges=True)
        
    except Exception as e:
        print(f"Erro ao abrir ou atualizar a planilha: {e}")

    finally:
        excel.Quit()

def atualizar_Excel():
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\9 - PPCP\\Lead time\\Banco de dados\\Banco de dados - Lead.xlsx')
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\9 - PPCP\\Lead time\\Lead time - Familia.xlsx')
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\9 - PPCP\\.PCP\\Fechamento de estoque\\Monitoramento de estoque.xlsx')
    # atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\9 - PPCP\\.PCP\\Controle\\Controle.xlsx')
    # atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\5-PROCESSO PROJETOS\\REGISTROS\\Controle - Ass.xlsx')
    # atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\3-GESTÃO DA QUALIDADE\\REGISTROS\\Gráficos-IDs-PCP-2024 - IDs 04a 04b 13b.xlsx')
    # atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\3-GESTÃO DA QUALIDADE\\REGISTROS\\Graficos-IDs-Compras-2024 - IDs 09 13a 13c.xlsx')

titulo()

pasta_downloads = 'c:\\Users\\Tecnico4\\Downloads'
pasta_destino = 'Z:\\ISO 9000 - SGQ\\12-SISTEMA\\0. Banco de dados - Planilhas'
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

exporta_Nomus()

move_Arquivos()

atualizar_Excel()

finalizacao()
