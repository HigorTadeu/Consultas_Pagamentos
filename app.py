import openpyxl
from selenium import webdriver #Possibilitar abrir o navegador
from selenium.webdriver.common.by import By # Possibilita navegar pela página
from time import sleep

def load_xlsx_file(file):
    #Load file XLSX
    return openpyxl.load_workbook(file)

def load_xlsx_sheet(file, sheet):
    """
    Load an XLSX file and return the specified worksheet.

    Args:
        file (str): The path to the XLSX file.
        sheet (str): The name of the sheet to be loaded.

    Returns:
        openpyxl.worksheet.worksheet.Worksheet: The loaded worksheet.
    """
    
    #Load the sheet that contains the information
    sheet_loaded = file[sheet]
    return sheet_loaded

def load_browser_firefox():
    return webdriver.Firefox() #Chama o naegador Firefox

def load_website(driver,site):
    driver.get(site) #Load website

def input_field_search(driver,input_xpath,value_search):
    search_field = driver.find_element(By.XPATH,input_xpath) #Procura na tela carregada pelo elemento através da técnica XPATH
    sleep(1)
    search_field.clear() #Apaga as informações existentes no campo
    search_field.send_keys(value_search)
    sleep(1)

def click_buton_search(driver,input_xpath):
    buton_search = driver.find_element(By.XPATH,input_xpath)
    sleep(1)
    buton_search.click()
    sleep(4)

def recover_data(driver,input_xpath):
    data = driver.find_element(By.XPATH,input_xpath)
    return data.text

def main():
    file_clients = load_xlsx_file('dados_clientes.xlsx')
    sheet_clients = load_xlsx_sheet(file_clients,'Sheet1')

    driver = load_browser_firefox()
    load_website(driver,'https://consultcpf-devaprender.netlify.app/')
    sleep(5)

    file_closure = load_xlsx_file('planilha fechamento.xlsx')
    sheet_closure = load_xlsx_sheet(file_closure,'Sheet1')

    for linha in sheet_clients.iter_rows(min_row=2, values_only=True):
        nome, valor, cpf, vencimento = linha 
    #2 - Entrar no site e pesquisar e usar o cpf 
        input_field_search(driver,"//input[@id='cpfInput']",cpf)

    # 3 - Verificar se está "em dia" ou "atrasado"
        click_buton_search(driver,"//button[@class='btn btn-custom btn-lg btn-block mt-3']")
        
    # 4 - Verificar se está "em dia"
        status = recover_data(driver,"//span[@id='statusLabel']")
        if status == 'em dia':
            data_pagamento = recover_data(driver,"//p[@id='paymentDate']")
            data_pagamento_limpo = data_pagamento.split()[3]
            metodo_pagamento = recover_data(driver,"//p[@id='paymentMethod']")
            metodo_pagamento_limpo = metodo_pagamento.split()[3]

            sheet_closure.append([nome, valor, cpf, vencimento, 'em dia', data_pagamento_limpo, metodo_pagamento_limpo])

            file_closure.save('planilha fechamento.xlsx')
        else:
            sheet_closure.append([nome, valor, cpf, vencimento, 'pendente'])
            file_closure.save('planilha fechamento.xlsx')
        

if __name__ == '__main__':
    main()