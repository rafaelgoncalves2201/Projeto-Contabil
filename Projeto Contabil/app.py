from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from docx import Document
import os

# Entrar no site https://contabil-devaprender.netlify.app/
driver = webdriver.Chrome()
driver.get('https://contabil-devaprender.netlify.app/')

# Preencher email
email = driver.find_element(By.XPATH,"/html/body/div/div/form/div[1]/input")
sleep(0.5)
email.click()
email.send_keys('teste@gmail.com')

# Pausa para clicar no proximo campo
sleep(0.5)

# Preencher senha
senha = driver.find_element(By.XPATH, "/html/body/div[1]/div/form/div[2]/input")
sleep(0.5)
senha.click()
senha.send_keys('teste')

# Pausa para clicar no proximo campo
sleep(0.5)

# Clicar em entrar
botao_entrar = driver.find_element(By.XPATH, '/html/body/div[1]/div/form/button')
botao_entrar.click()
sleep(5)

# Acessa "Cadastrar Balanço Patrimonial"
acessa_balanco = driver.find_element(By.XPATH, '/html/body/div/div/div/div[1]/div/div/a')
sleep(1)
acessa_balanco.click()

def inserir_valor_de_documento_word(caminho_arquivo):
    # Extrair os dados do arquivo word
    caminho_arquivo = r'C:\Users\rgcar\OneDrive\Área de Trabalho\Projeto Python\relatorios\Relatorio_Contabilidade_Delícias_Foodie.docx'
    arquivo_word = Document(caminho_arquivo)

    # Variaveis para os items desejados
    ativo_circulante = ''
    caixas_equivalentes = ''
    contas_receber = ''
    estoques = ''
    ativo_nao_circulante = ''
    imobilizado = ''
    intangivel = ''
    total_ativo = ''

    # Acessa todas as tabelas
    for tabela in arquivo_word.tables:
        for linha in tabela.rows:
            if 'Ativo Circulante' in linha.cells[0].text.strip():
                ativo_circulante = linha.cells[1].text.strip()
            elif 'Caixa e Equivalentes' in linha.cells[0].text.strip():
                caixas_equivalentes = linha.cells[1].text.strip()
            elif 'Contas a Receber' in linha.cells[0].text.strip():
                contas_receber = linha.cells[1].text.strip()
            elif 'Estoques' in linha.cells[0].text.strip():
                estoques = linha.cells[1].text.strip()
            elif 'Ativo Não Circulante' in linha.cells[0].text.strip():
                ativo_nao_circulante = linha.cells[1].text.strip()
            elif 'Imobilizado' in linha.cells[0].text.strip():
                imobilizado = linha.cells[1].text.strip()
            elif 'Intangível' in linha.cells[0].text.strip():
                intangivel = linha.cells[1].text.strip()
            elif 'Total do Ativo' in linha.cells[0].text.strip():
                total_ativo = linha.cells[1].text.strip()

    # Preencher os campos com os valores extraídos do arquivo word

    # Preencher o campo Ativo Circulante
    campo_ativo_circulante = driver.find_element(By.XPATH, '//*[@id="ativo_circulante"]')
    sleep(0.5)
    campo_ativo_circulante.click()
    campo_ativo_circulante.send_keys(ativo_circulante)

    # Preencher o campo Caixas Equivalentes
    campo_caixas_equivalente = driver.find_element(By.XPATH, '//*[@id="caixa_equivalentes"]')
    sleep(0.5)
    campo_caixas_equivalente.click()
    campo_caixas_equivalente.send_keys(caixas_equivalentes)

    # Preencher o campo Contas a Receber 
    campo_contas_receber = driver.find_element(By.XPATH, '//*[@id="contas_receber"]')
    sleep(0.5)
    campo_contas_receber.click()
    campo_contas_receber.send_keys(contas_receber)

    # Preencher o campo Estoques
    campo_estoques = driver.find_element(By.XPATH, '//*[@id="estoques"]')
    sleep(0.5)
    campo_estoques.click()
    campo_estoques.send_keys(estoques)

    # Preencher o campo Ativo Não Circulante
    campo_ativo_nao_circulante = driver.find_element(By.XPATH, '//*[@id="ativo_nao_circulante"]')
    sleep(0.5)
    campo_ativo_nao_circulante.click()
    campo_ativo_nao_circulante.send_keys(ativo_circulante)

    # Preencher o campo Imobilizado
    campo_imobilizado = driver.find_element(By.XPATH, '//*[@id="imobilizado"]')
    sleep(0.5)
    campo_imobilizado.click()
    campo_imobilizado.send_keys(imobilizado)

    # Preencher o campo Intangível
    campo_intangivel = driver.find_element(By.XPATH, '//*[@id="intangivel"]')
    sleep(0.5)
    campo_intangivel.click()
    campo_intangivel.send_keys(intangivel)

    # Preencher o campo Total Ativo
    campo_total_ativo = driver.find_element(By.XPATH, '//*[@id="total_ativo"]')
    sleep(0.5)
    campo_total_ativo.click()
    campo_total_ativo.send_keys(total_ativo)

    sleep(0.9)

    # Clicar no botão cadastrar
    botao_cadastrar = driver.find_element(By.XPATH, '//*[@id="balanco-form"]/div[2]/button')
    sleep(0.5)
    botao_cadastrar.click()

#  Repite os passos 4 e 5 até chegar ao ultimo arquivo word
pasta_relatorios = r'C:\Users\rgcar\OneDrive\Área de Trabalho\Projeto Python\relatorios'
for nome_arquivo in os.listdir(pasta_relatorios):
    if nome_arquivo.endswith('.docx'):
        caminho_arquivo_word = os.path.join(pasta_relatorios,nome_arquivo)
        inserir_valor_de_documento_word(caminho_arquivo_word)
