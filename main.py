from turtle import position
import pyautogui
import time

def limpar_Downloads():
    ''' Entra em Downloads e excluí todos os arquivos '''
    pyautogui.press('winleft')
    pyautogui.write('Download')
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.hotkey('winleft','up')
    pyautogui.hotkey('ctrl','a')
    pyautogui.press('delete')
    pyautogui.hotkey('ctrl','w')

def click_Mouse_Padrao_exportar():
    ''' Aperta nos botões para exportar um relatório do Nomus '''
    time.sleep(25)
    pyautogui.press('esc')
    pyautogui.moveTo(1643,109)
    pyautogui.click()
    pyautogui.moveTo(1664,166)
    pyautogui.click()
    pyautogui.moveTo(1458,1012)
    pyautogui.click()

def click_Mouse_Padrao_exportar_Pausas():
    ''' Aperta nos botões para exportar um relatório do Nomus com pausas de 5 segundos entre cada click '''
    time.sleep(25)
    pyautogui.press('esc')
    pyautogui.moveTo(1643,109)
    pyautogui.click()
    time.sleep(5)
    pyautogui.moveTo(1664,166)
    pyautogui.click()
    time.sleep(5)
    pyautogui.moveTo(1458,1012)
    pyautogui.click()
    time.sleep(5)

def click_Mouse_Diferente_exportar():
    ''' Aperta nos botões para exportar um relatório do Nomus com a visualização diferente '''
    pyautogui.hotkey('ctrl','shift','tab')
    time.sleep(15)
    pyautogui.press('esc')
    pyautogui.moveTo(1761,111)
    pyautogui.click()
    pyautogui.moveTo(1696,164)
    pyautogui.click()
    pyautogui.moveTo(1458,1012)
    pyautogui.click()
    pyautogui.hotkey('ctrl','t')
    
def entrar_Link(texto):
    ''' Entra no link '''
    pyautogui.write(texto)
    pyautogui.press('enter')
    pyautogui.hotkey('ctrl','t')

def abrir_Navegador(texto):
    ''' Abre o navegador '''
    time.sleep(5)
    pyautogui.press('winleft')
    pyautogui.write(texto)
    pyautogui.press('enter')
    pyautogui.hotkey('winleft','up')
    pyautogui.moveTo(839,62)
    pyautogui.click()

def exportar_relatorio_padrao():
    ''' Baixa os arquivos '''
    pyautogui.hotkey('ctrl','tab')
    click_Mouse_Padrao_exportar()

def exportar_relatorio_padrao_Pausas():
    ''' Baixa os arquivos com pausas de 5 segundos por click'''
    pyautogui.hotkey('ctrl','tab')
    click_Mouse_Padrao_exportar_Pausas()
    
def exportar_Relatorio_Diferente():
    ''' Baixa os arquivos com os relatórios com visualizações diferentes'''
    click_Mouse_Padrao_exportar()
    pyautogui.hotkey('ctrl','shift','tab')

def exportar_Fechar_Navegador_Padrao():
    '''Conjunto de funções para exportar todos os arquivos e fechar o navegador ao terminar'''
    time.sleep(1)
    pyautogui.hotkey('ctrl','w')
    exportar_relatorio_padrao()
    exportar_relatorio_padrao()
    exportar_relatorio_padrao()
    exportar_relatorio_padrao()
    exportar_relatorio_padrao()
    exportar_relatorio_padrao()
    exportar_relatorio_padrao()
    exportar_relatorio_padrao()
    exportar_relatorio_padrao()
    exportar_relatorio_padrao()
    time.sleep(60)
    pyautogui.hotkey('ctrl','shift','w')

def exportar_Fechar_Navegador_Padrao_Pausas():
    '''Conjunto de funções para exportar todos os arquivos e fechar o navegador ao terminar com pausas de 5 segundos por click'''
    time.sleep(1)
    pyautogui.hotkey('ctrl','w')
    exportar_relatorio_padrao_Pausas()
    exportar_relatorio_padrao_Pausas()
    exportar_relatorio_padrao_Pausas()
    exportar_relatorio_padrao_Pausas()
    exportar_relatorio_padrao_Pausas()
    exportar_relatorio_padrao_Pausas()
    exportar_relatorio_padrao_Pausas()
    exportar_relatorio_padrao_Pausas()
    exportar_relatorio_padrao_Pausas()
    exportar_relatorio_padrao_Pausas()
    time.sleep(60)
    pyautogui.hotkey('ctrl','shift','w')

def exportar_Fechar_Navegador_Diferente():
    '''Conjunto de funções para exportar todos os arquivos com relatorios de visualizações diferentes e fechar o navegador ao terminar'''
    pyautogui.hotkey('ctrl','w')
    exportar_Relatorio_Diferente()
    exportar_Relatorio_Diferente()
    exportar_Relatorio_Diferente()
    time.sleep(60)
    pyautogui.hotkey('ctrl','shift','w')

def atualizar_Novos_Arquivos():
    ''' Recorta os arquivos em Download e cola no servidor '''
    pyautogui.press('winleft')
    pyautogui.write('Download')
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.hotkey('winleft','up')
    pyautogui.hotkey('ctrl','a')
    pyautogui.hotkey('ctrl','x')
    pyautogui.hotkey('ctrl','w')
    pyautogui.press('winleft')
    pyautogui.write('Z:\ISO 9000 - SGQ')
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('pgdn')
    pyautogui.press('enter')
    pyautogui.press('pgup')
    pyautogui.press('enter')    
    pyautogui.hotkey('winleft','up')
    pyautogui.hotkey('ctrl','v')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

pyautogui.alert("""O código vai começar.

Favor deixar o Capsloock desativado.""")

pyautogui.PAUSE = 0.5

limpar_Downloads()

abrir_Navegador('chrome')

entrar_Link('https://reports.nomus.com.br/open-view/751489003498986462') # PC - Follow up
entrar_Link('https://reports.nomus.com.br/open-view/751489003495455916') # SC COMPLETA
entrar_Link('https://reports.nomus.com.br/open-view/751489003499869977') # Cotação
entrar_Link('https://reports.nomus.com.br/open-view/751489003499311067') # NF Tipo de movimentação
entrar_Link('https://reports.nomus.com.br/open-view/751489003497642590') # Fornecedores - TB-06
entrar_Link('https://reports.nomus.com.br/open-view/751489003385051213') # Ordens Lead
entrar_Link('https://reports.nomus.com.br/open-view/751489003418257074') # Necessidade
entrar_Link('https://reports.nomus.com.br/open-view/751489003348744953') # ID-04e
entrar_Link('https://reports.nomus.com.br/open-view/751489003473916593') # Tempo de produção
entrar_Link('https://reports.nomus.com.br/open-view/751489003396077915') # Entradas
exportar_Fechar_Navegador_Padrao()

time.sleep(3)

abrir_Navegador('brave')

entrar_Link('https://reports.nomus.com.br/open-view/751489003343390887') # Pedido de compra
entrar_Link('https://reports.nomus.com.br/open-view/751489003375648746') # Pv's
entrar_Link('https://reports.nomus.com.br/open-view/751489003343401232') # SC's
entrar_Link('https://reports.nomus.com.br/open-view/751489003343402198') # Entradas compras
entrar_Link('https://reports.nomus.com.br/open-view/751489003383877015') # Faturamento periodo
entrar_Link('https://reports.nomus.com.br/open-view/751489003384381506') # CRM
entrar_Link('https://reports.nomus.com.br/open-view/751489003381705320') # Vendas
entrar_Link('https://reports.nomus.com.br/open-view/751489003396836362') # Emissão de PV
entrar_Link('https://reports.nomus.com.br/open-view/751489003384833393') # Status do PV
entrar_Link('https://reports.nomus.com.br/open-view/751489003501209517') # Familia
exportar_Fechar_Navegador_Padrao_Pausas()

time.sleep(3)

abrir_Navegador('firefox')

entrar_Link('https://reports.nomus.com.br/open-view/751489003299922759') # CONTROLE DA PRODUÇÃO
click_Mouse_Diferente_exportar()
entrar_Link('https://reports.nomus.com.br/open-view/751489003399788047') # ENTRADAS (LIVRO)
click_Mouse_Diferente_exportar()
entrar_Link('https://reports.nomus.com.br/open-view/751489003415814176') # Nfs emitidas
click_Mouse_Diferente_exportar()
entrar_Link('https://reports.nomus.com.br/open-view/751489003399436517') # Saída 
click_Mouse_Diferente_exportar()
pyautogui.hotkey('ctrl','tab')
pyautogui.hotkey('ctrl','tab')
pyautogui.hotkey('ctrl','w')
pyautogui.hotkey('ctrl','tab')
pyautogui.hotkey('ctrl','w')
pyautogui.hotkey('ctrl','tab')
pyautogui.hotkey('ctrl','w')
pyautogui.hotkey('ctrl','tab')
pyautogui.hotkey('ctrl','w')
entrar_Link('https://reports.nomus.com.br/open-view/751489003406446494') # Dw Compras
entrar_Link('https://reports.nomus.com.br/open-view/751489003342815143') # Modelo NF's 
entrar_Link('https://reports.nomus.com.br/open-view/751489003415906765') # Grupos
exportar_Fechar_Navegador_Diferente()

time.sleep(3)

atualizar_Novos_Arquivos()

pyautogui.alert("""Código finalizado.

Obrigado por aguardar.""")