#pip install pandas
#pip install numpy
#pip install openpyxl
#pip install selenium

#Chrome - chromedriver - Baixar e colocar no PATH do python
#Firefox - GeckoDriver

from selenium import webdriver # Navegador
from selenium.webdriver.common.by import By #Localizar elementos (os items do site)
from selenium.webdriver.common.keys import Keys # Permite clicar com teclas no teclado
import pandas as pd

navegador = webdriver.Chrome()

#Passo 1: Entrar no navegador
navegador.get("https://www.google.com/")

#Passo 2: Pesquisar a cotação do dólar
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação dólar")
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

#Passo 3: Pegar a cotação do dólar
cotacaoDolar = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacaoDolar)

#Passo 4: Pegar a cotação do euro
navegador.get("https://www.google.com/")

navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação euro")
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

cotacaoEuro = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacaoEuro)

#Passo 5:  Pegar a cotação do ouro
navegador.get("https://www.melhorcambio.com/ouro-hoje")

cotacaoOuro = navegador.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute('value')
cotacaoOuro = cotacaoOuro.replace(",", ".")
print(cotacaoOuro)

#Passo 6: Atualizar a minha base de dados com novas cotações
tabela = pd.read_excel("Produtos.xlsx")
#print(tabela)

# Atualizar a cotação de acordo com a moeda correspondente

#tabela.loc[linha,coluna]
#[tabela["Moeda"] == "Dólar", "Cotação"] - Condição
tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = float(cotacaoDolar)
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacaoEuro)
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = float(cotacaoOuro)


# Atualizar preço de compra = preço original * cotação
tabela["Preço de Compra"] = tabela["Preço Original"] * tabela["Cotação"]

# Atualizar preço de venda = preço de compra * margem
tabela["Preço de Venda"] = tabela["Preço de Compra"] * tabela["Margem"]
print(tabela)

# Exportar arquivo excel atualizado
tabela.to_excel("Produtos Novos.xlsx", index=False)
navegador.quit()