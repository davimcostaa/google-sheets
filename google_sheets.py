import gspread
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import os
import glob
from time import sleep
import sys
sys.path.insert(1, "C:\\Users\\davi.costa\\Desktop")
from login import credenciais

CODE = ''

credencial = {
  
}
 
gc = gspread.service_account_from_dict(credencial)

sh = gc.open_by_key(CODE)

ws = sh.worksheet('Projetos 2023')

nome_usuario = credenciais.get('NOME_USUARIO')
senha = credenciais.get('SENHA')

servico = Service(ChromeDriverManager().install())
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

navegador = webdriver.Chrome(service=servico, options=chrome_options)

navegador.get('')

navegador.maximize_window()

navegador.find_element("xpath", '//*[@id="details-button"]').click()

navegador.find_element("xpath", '//*[@id="proceed-link"]').click()

navegador.find_element("xpath", '//*[@id="user"]').send_keys(nome_usuario)

navegador.find_element("xpath", '//*[@id="password"]').send_keys(senha)

navegador.find_element("xpath", '//*[@id="loginform"]/div[2]/div/p/button').click()

navegador.find_element("xpath", '//*[@id="panel-2"]').click()

navegador.find_element("xpath", '//*[@id="Survey_searched_value"]').send_keys('Projetos de Etnodesenvolvimento 2023')

navegador.find_element("xpath", '//*[@id="yw0"]/input[2]').click()

navegador.find_element("xpath", '//*[@id="survey-grid"]/table/tbody/tr').click()

navegador.find_element("xpath", '//*[@id="surveybarid"]/div/div[1]/div[3]/button').click()

navegador.find_element("xpath", '//*[@id="surveybarid"]/div/div[1]/div[3]/ul/li[1]/a').click()

navegador.implicitly_wait(10)

navegador.find_element("xpath", '//*[@id="browsermenubarid"]/div/div/div[1]/button').click()

navegador.find_element("xpath", '//*[@id="browsermenubarid"]/div/div/div[1]/ul/li[1]/a').click()

navegador.implicitly_wait(3)

navegador.find_element("xpath", '//*[@id="panel-1"]/div[2]/div/div[1]/div/div[2]').click()

navegador.find_element("xpath", '//*[@id="export-button"]').click()

sleep(10)

lista_download = glob.glob("C:/Users/davi.costa/Downloads/*")
arquivo_baixado = max(lista_download, key=os.path.getmtime)

print(arquivo_baixado)

resultado = pd.read_excel(arquivo_baixado) 
resultado.dropna(subset=["Data de envio"], inplace=True)

print(len(resultado))

resultado = resultado.loc[:, ~resultado.columns.duplicated()]
resultado.drop(columns = resultado.iloc[:,2:7], inplace=True) # apaga colunas de um determinado intervalo pelo método iloc
resultado = resultado.iloc[:, :-3]

print(resultado)

resultado.drop(columns = ["Telefone", "Email", "URL de referência","CTL 1 [Outros]", "CTL 7", "CTL 8", "CTL 9", "CTL 10", "CTL 11", "CTL 12", "Outros públicos beneficiados indiretamente"], inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 1º subelemento de despesa selecionado "])#transforma em número e preenche colunas vazias com zero, podemos também criar uma função para fazer isso com as outras colunas
resultado["Valor solicitado para o 1º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 1º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 1º subelemento de despesa selecionado "].fillna(0, inplace=True)
print(sum(resultado["Valor solicitado para o 1º subelemento de despesa selecionado "]))

pd.to_numeric(resultado["Valor solicitado para o 2º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 2º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 2º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 2º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 3º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 3º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 3º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 3º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 4º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 4º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 4º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 4º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 5º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 5º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 5º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 5º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 6º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 6º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 6º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 6º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 7º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 7º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 7º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 7º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 8º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 8º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 8º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 8º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 9º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 9º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 9º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 9º subelemento de despesa selecionado "].fillna(0, inplace=True)

pd.to_numeric(resultado["Valor solicitado para o 10º subelemento de despesa selecionado "])
resultado["Valor solicitado para o 10º subelemento de despesa selecionado "] = resultado["Valor solicitado para o 10º subelemento de despesa selecionado "].astype(float)
resultado["Valor solicitado para o 10º subelemento de despesa selecionado "].fillna(0, inplace=True)

resultado = resultado.fillna('')

os.remove(arquivo_baixado)
navegador.close()

tabelaAntiga = ws.get_all_values()

df = pd.DataFrame(tabelaAntiga)

df.columns = df.iloc[0]
df = df[1:]

indiceUltimaLinha = len(df) - 1

ultima_linha = df.iloc[indiceUltimaLinha]
maiorID = ultima_linha['ID da resposta']

for index, row in resultado.iterrows():
    if int(row['ID da resposta']) > int(maiorID):
      row = row.tolist()
      ws.append_row(row)

print('the end')

