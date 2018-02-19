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

Headless_Mode = input('Headless Mode? (y/n) \n')
if Headless_Mode == 'y':
    chromeOptions = webdriver.ChromeOptions()
    prefs = {"download.default_directory" : os.getcwd(),
             "download.prompt_for_download": False}
    chromeOptions.add_experimental_option("prefs",prefs)
    chromeOptions.add_argument('--headless')
    chromeOptions.add_argument('--window-size= 1600x900')
    driver = webdriver.Chrome(chrome_options=chromeOptions)
elif Headless_Mode == 'n':
    driver = webdriver.Chrome()
else:
    print('Invalid argument.')
    input()

if __name__ == '__main__':
    driver.get(url)
    # Faz login no sistema
    uname = driver.find_element_by_name('txtBoxLogin')
    uname.send_keys(username)
    passw = driver.find_element_by_name('txtBoxSenha')
    passw.send_keys(password)
    submit_button = driver.find_element_by_id('ImageButton_Login').click()

    # Modifica os campos necessários e envia o anexo de cada sob.
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
                        # Verificação da categoria dos arquivos de acordo com palavras-chave no nome do mesmo
                        if 'FORM_FISC' in item:  # Verifica se é Formulário de Fiscalização de Obra
                            categoria = Select(driver.find_element_by_id('drpCategoria'))
                            categoria.select_by_visible_text('ENCERRAMENTO')
                            documento = Select(driver.find_element_by_id('DropDownList1'))
                            documento.select_by_visible_text('FORMULARIO DE FISCALIZACAO DE OBRA')
                        elif 'AS_BUILT' in item:  # Verifica se é AS BUILT
                            categoria = Select(driver.find_element_by_id('drpCategoria'))
                            categoria.select_by_visible_text('ENCERRAMENTO')
                            documento = Select(driver.find_element_by_id('DropDownList1'))
                            documento.select_by_visible_text('AS BUILT')
                        elif 'APOIO_TRANSITO' in item:  # Verifica se é Apoio de Trânsito
                            categoria = Select(driver.find_element_by_id('drpCategoria'))
                            categoria.select_by_visible_text('PROJETO')
                            documento = Select(driver.find_element_by_id('DropDownList1'))
                            documento.select_by_visible_text('CARTAS/OFICIOS')
                        elif '_SGD_' in item:  # Verifica se é SGD
                            categoria = Select(driver.find_element_by_id('drpCategoria'))
                            categoria.select_by_visible_text('EXECUCAO')
                            documento = Select(driver.find_element_by_id('DropDownList1'))
                            documento.select_by_visible_text('SGD/TET')
                        elif 'CLIENTE_VITAL' in item:  # Verifica se é carta de cliente vital
                            categoria = Select(driver.find_element_by_id('drpCategoria'))
                            categoria.select_by_visible_text('PROJETO')
                            documento = Select(driver.find_element_by_id('DropDownList1'))
                            documento.select_by_visible_text('CARTAS/OFICIOS')
                        elif '_LOCACAO_' in item or '_APR_' in item or '_DESENHO_' in item: # Verifica se é planejamento
                            categoria = Select(driver.find_element_by_id('drpCategoria'))
                            categoria.select_by_visible_text('EXECUCAO')
                            documento = Select(driver.find_element_by_id('DropDownList1'))
                            documento.select_by_visible_text('LOCACAO')
                        else: # Caso não seja nenhuma das anteriores, reconhece como Ponto de Serviço
                            categoria = Select(driver.find_element_by_id('drpCategoria'))
                            categoria.select_by_visible_text('EXECUCAO')
                            documento = Select(driver.find_element_by_id('DropDownList1'))
                            documento.select_by_visible_text('PONTO DE SERVICO')

                        # Seleciona o arquivo  a ser upado e clica no botão "Adicionar Documento"
                        driver.find_element_by_id('fileUPArquivo').send_keys(os.getcwd() + "\\" + item)
                        # driver.find_element_by_id('Button_Anexar').click()
                        try:
                            # Verifica se o arquivo foi anexado com êxito
                            status = driver.find_element_by_xpath("*//a[contains(text(), '" + item + "')]")
                            if status.is_displayed():
                                print("Arquivo " + item + " anexado com êxito.")
                                if item.startswith(('SG_REF', 'SG_QUAL', 'SG_RNT')):
                                    driver.save_screenshot('_'.join(item.split('_', 1)[0]) + ".png")
                                elif item.startswith('SG_PQ'):
                                    driver.save_screenshot('_'.join(item.split('_', 1)[0]) + ".png")
                                else:
                                    driver.save_screenshot(item.split('.', 1)[0] + ".png")
                        except NoSuchElementException:
                            log = open("log.txt", "a")
                            log.write(item + " não foi anexado.\n")
                            log.close()
                            continue
    print("Fim da execução.")
