from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
import openpyxl

fundos=['rbff11', 'bcff11', 'hsml11', 'visc11', 'bbpo11', 'knri11','vino11', 'rztr11','vghf11', 'mxrf11', 'kncr11', 'bcri11','hglg11','ggrc11', 'mcci11']

url = 'https://statusinvest.com.br/fundos-imobiliarios/'
nome_do_arquivo = openpyxl.load_workbook('Investimento.xlsx')
dados = nome_do_arquivo['planilha_automatica']
urls_fiis = []

#Config webdrive
crome_op = Options()
crome_op.headless = True
google = webdriver.Chrome(options=crome_op)

for fiis in fundos:
    urls_fiis= (url+fiis)
    google.get(urls_fiis)
    dy= google.find_element(
        By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
    v_vp= google.find_element(
        By.XPATH, '/html/body/main/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
    valor_atual = google.find_element(
        By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[1]/div/div[1]/strong').text
    varia = google.find_element(
        By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[5]/div/div[1]/strong').text
    valor_pa = google.find_element(
        By.XPATH,'//*[@id="main-2"]/div[2]/div[5]/div/div[1]/div/div[1]/strong').text
    
    dy = dy.replace(',', '.')
    v_vp = v_vp.replace(',', '.')
    valor_atual = valor_atual.replace(',', '.')
    varia = varia.replace(',', '.')
    valor_pa = valor_pa.replace(',', '.')

    google.quit
    print('\033[32m', 'Valores de '+urls_fiis[-6:]+'- OK')

    for rows in dados.iter_rows(min_row=2, max_row=16):
        
        if rows[0].value == urls_fiis[-6:]:  # Nome do Fundo
            rows[1].value = v_vp  # Valor sobre valor patrimonial
            rows[2].value = dy  # Porcentagem do dividendo anual
            rows[3].value = valor_atual 
            rows[4].value = varia
            rows[5].value = valor_pa #Valor patrimonial por cota
        
    nome_do_arquivo.save('Investimento.xlsx')
print('Finalizado...')
