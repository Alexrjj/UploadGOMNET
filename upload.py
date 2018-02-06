from selenium.common.exceptions import NoSuchElementException
import os
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import openpyxl  #  Acessar os dados de login

dirListing = os.listdir("./")
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
    for item in dirListing:
        if ".PDF" in item:
            if item.startswith(('SG_REF', 'SG_QUAL', 'SG_RNT')):
                driver.get(url2 + '_'.join(item.split('_', 3)[1:3]))
            elif item.startswith('SG_PQ'):
                driver.get(url2 + '_'.join(item.split('_', 4)[1:4]))
            else:
                driver.get(url2 + item.split('_', 2)[1])

            try:  # Verifica se a sob foi digitada incorretamente.
                erro = driver.find_element_by_xpath('*//tr/td[contains(text(),'
                                                    '"Não existem dados para serem exibidos.")]')
                if erro.is_displayed():
                    print("Sob " + item.partition("_")[0] + " não encontrada. Favor verificar.")
            except NoSuchElementException:
                try:  # Verifica se o arquivo já foi anexado.
                    anexo = driver.find_element_by_xpath(
                        "*//a[contains(text(), '" + item + "')]")
                    if anexo.is_displayed():
                        print("Arquivo " + item + " já foi anexado.")
                except NoSuchElementException:
                    # Preenche o campo "Descrição" com "PONTO DE SERVIÇO"
                    atividade = driver.find_element_by_id('txtBoxDescricao')
                    atividade.send_keys('PONTO DE SERVIÇO')
                    # Identifica o menu " Categoria de Documento" e seleciona a opção "EXECUCAO"
                    categoria = Select(driver.find_element_by_id('drpCategoria'))
                    categoria.select_by_visible_text('EXECUCAO')
                    # Identifica o menu " Tipo de Documento" e seleciona a opção "OUTROS"
                    documento = Select(driver.find_element_by_id('DropDownList1'))
                    documento.select_by_visible_text('OUTROS')
                    driver.find_element_by_id('fileUPArquivo').send_keys(os.getcwd() + "\\" + item)
                    driver.find_element_by_id('Button_Anexar').click()
                    try:
                        # Verifica se o arquivo foi anexado com êxito
                        status = driver.find_element_by_xpath("*//a[contains(text(), '" + item + "')]")
                        if status.is_displayed():
                            print("Arquivo " + item + " anexado com sucesso.")
                            driver.save_screenshot(item.partition(".")[0] + ".png")
                    except NoSuchElementException:
                        log = open("log.txt", "a")
                        log.write(item + " não foi anexado.\n")
                        log.close()
                        continue
    print("Fim da execução.")
