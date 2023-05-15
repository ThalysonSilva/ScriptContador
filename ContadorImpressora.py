import time
import schedule
import pyautogui as py
import pyperclip as pyper
import datetime
from datetime import datetime, timedelta, date
import pandas as pd

date_atual = date.today()
dateToday = '{}_{}_{}'.format(date_atual.day, date_atual.month, date_atual.year)

def PegarContadorGeral():
    #Entrar no Site da impressora
    py.PAUSE = 2
    time.sleep(2)
    py.hotkey('win','d')
    py.press('win')     
    time.sleep(5)
    py.write('chrome')
    time.sleep(5)
    py.press('enter')
    time.sleep(5)
    py.write('http://10.10.84.28/web/guest/en/websys/webArch/mainFrame.cgi')
    py.press('enter')
    time.sleep(10)

    #Ir para pagina de login, e preencher o login e senha
    quantidade = 1
    for quantidade in range(4):
        py.press('tab')
    time.sleep(2)    
    py.press('enter')  
    time.sleep(5)
    py.write('admin')
    py.press('tab')
    py.write('1234567')
    py.press('tab')
    py.press('enter') 
    time.sleep(10)
    py.press('enter')

    #Entrar na opção de contador
    quantTab = 1
    for quantTab in range(19):
        py.press('tab')  
    py.press('enter')   
    py.press('tab')
    py.press('tab')
    py.press('tab')
    py.press('enter')
    time.sleep(2)   

    #Salvar arquivo no tic com a data atual
    py.PAUSE=2
    py.hotkey('ctrl', 'p')
    time.sleep(2)
    py.press('enter')
    time.sleep(2)
    py.write(dateToday)
    py.hotkey('func', 'f4')
    py.hotkey('ctrl', 'a')
    py.press('del')
    time.sleep(1)
    pyper.copy(f"\\\\10.10.84.250\\"+"tic\Contador geral RICOH sede") #em pyton pra sair duas barras invertidas, tem que colocar 4 na String
    py.hotkey('ctrl', 'v')
    py.press('enter')
    quantTab = 1
    for quantTab in range(4):
        py.hotkey('ctrl', 'shift', 'tab')   
        time.sleep(2)
    py.press('enter')
    py.hotkey('alt', 'func', 'f4')

def PegarContadorPorUsuario():
    #Entrar no Site da impressora
    py.PAUSE = 2
    time.sleep(2)
    py.hotkey('win','d')
    py.press('win')     
    time.sleep(5)
    py.write('chrome')
    time.sleep(5)
    py.press('enter')
    time.sleep(5)
    py.write('http://10.10.84.28/web/guest/en/websys/webArch/mainFrame.cgi')
    py.press('enter')
    time.sleep(10)

    #Ir para pagina de login, e preencher o login e senha
    quantidade = 1
    for quantidade in range(4):
        py.press('tab')
    time.sleep(2)    
    py.press('enter')  
    time.sleep(5)
    py.write('admin')
    py.press('tab')
    py.write('1234567')
    py.press('tab')
    py.press('enter') 
    time.sleep(10)
    py.press('enter')

    #Entrar na opção de contador por usuário 
    time.sleep(2)
    quantTab = 1
    for quantTab in range(19):
        py.press('tab')  
    py.press('enter')   
    for quantTab in range(4):
        py.press('tab') 
    py.press('enter')
    time.sleep(5)  

    #Fazer o download do arquivo e renomear
    time.sleep(2)
    for quantTab in range(4):
        py.press('tab')
    time.sleep(1)
    py.press('enter')
    time.sleep(10)
    for quantTab in range(8):
        py.press('tab')
    time.sleep(1)
    py.press('enter')
    py.press('up')
    py.press('enter')
    time.sleep(7)

    #Configurar para salvar no tic
    time.sleep(2)
    py.PAUSE=2
    py.hotkey('ctrl','x')
    py.hotkey('func', 'f4')
    py.hotkey('ctrl', 'a')
    py.press('del')
    time.sleep(1)
    py.write(f"\\\\10.10.84.250\\"+"tic\Contador por usuario Ricoh Sede") #em pyton pra sair duas barras invertidas, tem que colocar 4 na String
    py.press('enter')
    time.sleep(3)
    py.hotkey('ctrl','v')
    py.hotkey('ctrl','v')
    py.hotkey('ctrl','v')
    py.hotkey('func', 'f2')
    py.write(dateToday)
    py.press('enter')
    time.sleep(2)
    py.hotkey('alt', 'fun', 'f4')
    py.hotkey('alt', 'fun', 'f4')
    time.sleep(4)

def ConvertCsvInExcel():
    caminho = (f'//10.10.84.250/tic/Contador por usuario Ricoh Sede/{dateToday}.csv')
    csv = pd.read_csv(caminho)
    novaCsv = csv.loc[:, 'User':'Color (Total Prints)'] #Codigo para fazer fazer filtrar de uma coluna tal pra outra
    excel = novaCsv.to_excel(f'//10.10.84.250/tic/Contador por usuario Ricoh Sede/{dateToday}.xlsx') #Aqui ele converte e já salva o arquivo no caminho citado

def BloquearMaquina():
    py.hotkey('win','l')
   
schedule.every().day.at("12:00").do(PegarContadorGeral)
schedule.every().day.at("12:05").do(PegarContadorPorUsuario)
schedule.every().day.at("12:10").do(ConvertCsvInExcel)
schedule.every().day.at("12:11").do(BloquearMaquina)

while 1:
    schedule.run_pending()
    time.sleep(1)