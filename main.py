from selenium import webdriver
import pyautogui
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time
import win32com.client

pyautogui.PAUSE = 1

def caminho_Download():
    pyautogui.press('winleft')
    pyautogui.write('Downloads')
    pyautogui.press('enter')

def caminho_publico():
    pyautogui.press('winleft')
    pyautogui.write('Z:\PUBLICO')
    pyautogui.press('enter')
    pyautogui.press('pgup')
    pyautogui.press('enter')

def caminho_Iso():
    pyautogui.press('winleft')
    pyautogui.write('Z:\PUBLICO')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('down')
    pyautogui.press('enter')

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

def exportar_Diff():
    time.sleep(30)
    try:
        
        export_button = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="ZDBExportButton"]'))
        )
        export_button.click()

        embed_export_button = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="SumAndPvtEmbedExportXLSMenuItem"]'))
        )
        time.sleep(1)
        embed_export_button.click()
        print("Botão 'EmbedExportXLSMenuItem' clicado com sucesso!")
        
        final_button = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[34]/div/div[2]/div/footer/div/button[1]'))
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
    exportar_Diff()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003399788047',21) # ENTRADAS (LIVRO)
    exportar_Diff()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003415814176',22) # Nfs emitidas
    exportar_Diff()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003399436517',23) # Saída 
    exportar_Diff()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003406446494',24) # Dw Compras
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003342815143',25) # Modelo NF's 
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003415906765',26) # Grupos
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003432511966',27) # movimentação - pvs
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003502954549',28) # Descrições
    exportar()
    mudar_Link('https://reports.nomus.com.br/open-view/751489003502944448',29) # Estoque - tempo real
    exportar_Diff()
    time.sleep(60)
    navegador.quit()

def salva_Arquivos():
    caminho_Download()
    time.sleep(0.5)
    pyautogui.hotkey('ctrl','a')
    pyautogui.hotkey('ctrl','x')
    pyautogui.hotkey('ctrl','w')
    caminho_publico()
    time.sleep(0.5)
    pyautogui.hotkey('ctrl','v')
    time.sleep(0.5)
    pyautogui.press('esc')
    pyautogui.hotkey('ctrl','w')
    time.sleep(0.5)

def move_Arquivos():
    caminho_Iso()
    time.sleep(0.5)
    pyautogui.press('pgdn')
    pyautogui.press('enter')
    pyautogui.press('pgup')
    pyautogui.press('enter')
    pyautogui.hotkey('ctrl','a')
    pyautogui.press('delete')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(5)
    caminho_Download()
    time.sleep(0.5)
    pyautogui.hotkey('ctrl','a')
    pyautogui.hotkey('ctrl','x')
    pyautogui.hotkey('ctrl','w')
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('ctrl','w')
    time.sleep(0.5)

def voltar_Arquivos():
    caminho_publico()
    time.sleep(0.5)
    pyautogui.hotkey('ctrl','a')
    pyautogui.hotkey('ctrl','x')
    pyautogui.hotkey('ctrl','w')
    caminho_Download()
    time.sleep(0.5)
    pyautogui.hotkey('ctrl','v')
    time.sleep(0.5)
    pyautogui.press('esc')
    pyautogui.hotkey('ctrl','w')
    time.sleep(0.5)

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

def atualizar_Planilhas_Diff(local_Planilha):
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = True

    try:
        wb = excel.Workbooks.Open(local_Planilha)
        time.sleep(10)
        pyautogui.hotkey('ctrl','alt')
        pyautogui.press('enter')
        time.sleep(10)
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
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\6-PROCESSO SUPRIMENTOS\\REGISTROS\\Controle de inventários.xlsx')
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\6-PROCESSO SUPRIMENTOS\\REGISTROS\\TB-06_AvalProvedoresExternos-Rev05.xlsx')
    atualizar_Planilhas_Diff('Z:\\ISO 9000 - SGQ\\9 - PPCP\\.PCP\\Controle\\Controle.xlsm')
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\5-PROCESSO PROJETOS\\REGISTROS\\Controle - AT.xlsm')
    atualizar_Planilhas('Z:\\PUBLICO\Araujo\\Planilhas-Indicadores\\SCS vs CARTEIRAS.xlsx')
    atualizar_Planilhas('Z:\\ISO 9000 - SGQ\\3-GESTÃO DA QUALIDADE\\REGISTROS\\Gráficos-IDs-PCP-2024 - IDs 04a 04b 13b.xlsx')
    atualizar_Planilhas('Z:\\PUBLICO\\Araujo\\Planilhas-Indicadores\\Gráfico-ID-Produção-2024 - ID-02-01-10-2024.xlsx')
    atualizar_Planilhas('Z:\\PUBLICO\\Araujo\\Planilhas-Indicadores\\Graficos-IDs-Compras-2024 - IDs 09 13a 13c.xlsx')
    atualizar_Planilhas('Z:\\PUBLICO\\Araujo\\Planilhas-Indicadores\\Graficos-IDs-Engenharia-2024 - IDs10a 10b 11 12 e 14.xlsx')

titulo()

salva_Arquivos()

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

exporta_Nomus()

move_Arquivos()

voltar_Arquivos()

atualizar_Excel()

finalizacao()

# os.system("shutdown /s /t 10")
