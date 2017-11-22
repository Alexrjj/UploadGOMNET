from selenium.common.exceptions import NoSuchElementException
import os
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import openpyxl

#  Acessa os dados de login fora do script, salvo numa planilha existente, para proteger as informações de credenciais
dados = openpyxl.load_workbook('C:\\gomnet.xlsx')
login = dados['Plan1']
url = 'http://gomnet.ampla.com/'
url2 = 'http://gomnet.ampla.com/Upload.aspx?numsob='
username = login['A1'].value
password = login['A2'].value

driver = webdriver.Chrome()
if __name__ == '__main__':
    driver.get(url)
    # Faz login no sistema
    uname = driver.find_element_by_name('txtBoxLogin')
    uname.send_keys(username)
    passw = driver.find_element_by_name('txtBoxSenha')
    passw.send_keys(password)
    submit_button = driver.find_element_by_id('ImageButton_Login').click()

    # Modifica os campos necessários e envia o anexo de cada sob contido nos arquivos txt.
    with open('sobs.txt') as sobs:
        for sob in sobs:
            sob = sob.strip()
            driver.get(url2 + sob.partition("_")[0])
            # Preenche o campo "Descrição" com "PONTO DE SERVIÇO"
            atividade = driver.find_element_by_id('txtBoxDescricao')
            atividade.send_keys('PONTO DE SERVIÇO')
            # Identifica o menu " Categoria de Documento" e seleciona a opção "EXECUCAO"
            categoria = Select(driver.find_element_by_id('drpCategoria'))
            categoria.select_by_visible_text('EXECUCAO')
            # Identifica o menu " Tipo de Documento" e seleciona a opção "OUTROS"
            documento = Select(driver.find_element_by_id('DropDownList1'))
            documento.select_by_visible_text('OUTROS')
            driver.find_element_by_id('fileUPArquivo').send_keys(os.getcwd() + "\\" + sob + ".PDF")
            driver.find_element_by_id('Button_Anexar').click()
            try:
                # Verifica se a sob está no status desejado
                status = driver.find_element_by_xpath('//*[@id="txtBoxMessage"][contains(text(),'
                                                      '"Arquivo salvo com sucesso.")]')
                if status.is_displayed():
                    print(sob + " anexado com sucesso.\n")
                    driver.save_screenshot(sob + ".png")
            except NoSuchElementException:
                log = open("log.txt", "a")
                log.write(sob + " não foi anexado. \n")
                log.close()
                continue
