from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import openpyxl

#navegar ate o site
driver = webdriver.Chrome()
driver.get('https://contabilidade-devaprender.netlify.app/')
driver.maximize_window()
sleep(2)

#digitar email
email = driver.find_element(By.XPATH,"//input[@id='email']")
sleep(2)
email.send_keys('admin@contabilidade.com')

#digitar senha
senha = driver.find_element(By.XPATH, "//input[@id='senha']")
sleep(2)
senha.send_keys('contabilidade123456')

#entrar no sistema
btn_entrar = driver.find_element(By.XPATH, "//button[@id='Entrar']")
sleep(2)
btn_entrar.click()
sleep(5)

#extrair as informações da planilha 

empresas = openpyxl.load_workbook("./empresas.xlsx")
pagina_empresas = empresas['dados empresas']

for linha in pagina_empresas.iter_rows(min_row=2, values_only=True):
    nome_empresa, email, telefone, endereco, cnpj, area_atuacao, quantidade_funcionarios, data_fundacao = linha
    print(linha)

    #preencher campos
    driver.find_element(By.ID,'nomeEmpresa').send_keys(nome_empresa)
    sleep(2)
    driver.find_element(By.ID,'emailEmpresa').send_keys(email)
    sleep(2)
    driver.find_element(By.ID,'telefoneEmpresa').send_keys(telefone)
    sleep(2)
    driver.find_element(By.ID,'enderecoEmpresa').send_keys(endereco)
    sleep(2)
    driver.find_element(By.ID,'cnpj').send_keys(cnpj)
    sleep(2)
    driver.find_element(By.ID,'areaAtuacao').send_keys(area_atuacao)
    sleep(2)
    driver.find_element(By.ID,'numeroFuncionarios').send_keys(quantidade_funcionarios)
    sleep(2)
    driver.find_element(By.ID,'dataFundacao').send_keys(data_fundacao)
    sleep(2)

    btn_cadastrar = driver.find_element(By.ID,'Cadastrar')
    btn_cadastrar.click()
    sleep(3)

