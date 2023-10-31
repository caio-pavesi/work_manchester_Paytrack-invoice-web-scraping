import datetime as dt
import warnings as wr
import pandas as pd
import time as tm
import re
import os

from pathlib import Path
from modules.pdf import Pdf
from modules.log import Log
from modules.outlook import Outlook
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

wr.filterwarnings('ignore')

def get_invoices(account: str, inbox: str, sender: str, search: str, folder: Path):
    Outlook.output_folder = folder
    
    yesterday = dt.datetime.now() - dt.timedelta(1)
    yesterday = dt.date(yesterday.year, yesterday.month, yesterday.day)
    
    emails = Outlook(account).search_emails(look_folder = inbox, look_sender = sender, look_subject = None, date_interval = [yesterday, yesterday])
    attachments = Outlook(account).download_attachments(emails, search)
    
    return attachments

def read_invoices(invoice_file):
    bill_data = []    # Compile all data from invoices
    
    ## Applies on each invoice
    for i in range(len(invoice_file)):
        ## Reads the content of the invoice
        content = Pdf( Path.cwd() / "PRD" / "data" / "600000" / "notas" / invoice_file[i] ).read_text()
        
        bill_maturity = []    # Expiracy date of bill (equal for all)
        invoice_num = []      # Number of the invoice
        invoice_sum = []      # Total of the invoice
        bill_num = []         # Total of the bill
        # ! Invoice and bill can have different sum, A bill can have more than one invoice
        
        ## filter to apply on each line
        filter = re.compile(r'\w+[A-Z]\s[0123456789]{5}|Vencimento: dia [0-9/]+ no valor de R[$] [0-9,]+|Nota de Débito nº \d+')
        
        ## Applies on each line
        for text in content.split('\n'):
            
            ## Applies filter
            if filter.search(text):
                
                ## Returns invoice_num
                if invoice_num == []:
                    invoice_num = re.findall(r'nº \d+$', text)
                    if invoice_num != []:
                        invoice_num = str(invoice_num[0]).split(" ")[1]    # returns only the number
                        invoice_num = int(invoice_num)                     # formats number from str to int
            
                ## Returns invoice_sum
                if invoice_sum == []:
                    invoice_sum = re.findall(r'R[$] [0-9,.]+', text)
                    if invoice_sum != []:
                        invoice_sum = str(invoice_sum[0])
                        invoice_sum = float(invoice_sum.split(" ")[1].replace(".", "").replace(",", "."))    # formats number from str to float

                ## Returns bill_num
                if bill_num == []:
                    bill_num = re.findall(r'\w{6}\s\d+', text)
                    if bill_num != []:
                        bill_num = str(bill_num[0]).split(" ")[1]    # returns only the number as str
                        bill_num = int(bill_num)                     # formats number from str to int
                        
                ## TODO: Add account maturity from bill
                if bill_maturity == []:
                    bill_maturity = re.findall(r'\d{2}/\d{2}/\d{4}', text)
                    if bill_maturity != []:
                        bill_maturity = bill_maturity[0]    # returns only the date as str
        
        ## None values to zero
        if bill_maturity == []:
            bill_maturity = 0
        if invoice_num == []:
            invoice_num = 0
        if invoice_sum == []:
            invoice_sum = 0
        if bill_num == []:
            bill_num = 0

        ## Stores values
        invoice_data = dict(
            Fatura = bill_num,
            Nota = invoice_num,
            Valor = invoice_sum,
            vencimento = bill_maturity
        )
        bill_data.append(invoice_data)
                    
    return bill_data

def donwload_extract(bill: pd.DataFrame, folder: Path):
    
    bill_file = []
    bill = bill['Fatura'].unique()
    dowloads_Folder = Path(str(Path.cwd()).split("\\OneDrive")[0] + "\\Downloads")

    ## Open edge
    web = webdriver.Edge()
    web.maximize_window()
    web.get('https://login.paytrack.com.br/')
    tm.sleep(1)
    
    ## login
    authLogin = "email@example.com" 
    XauthLogin = web.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/div[2]/form/div[2]/div/div/span/input')
    XauthLogin.send_keys(authLogin)
    authLogin_Button = web.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/div[2]/form/div[4]/button')
    web.execute_script("arguments[0].click()", authLogin_Button)
    tm.sleep(2)
    
    ## password
    authPassword = "some_password"
    XauthPassword = web.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/div[2]/form/div[2]/div/div/span/span/input')
    XauthPassword.send_keys(authPassword)
    authPassword_Button = web.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/div[2]/form/div[4]/div/button')
    web.execute_script("arguments[0].click()", authPassword_Button)
    tm.sleep(1)
    
    ## 
    web.get('https://app.paytrack.com.br/#/agencia-faturamento')
    tm.sleep(5)
    
    for i in range(len(bill)):
        Fatura = bill[i]
        if Fatura != 0:
            ## Pesquisando a fatura
            Xsearch = web.find_element(By.CLASS_NAME, 'ant-input')
            Xsearch.send_keys(Keys.CONTROL + "a")
            Xsearch.send_keys(Keys.DELETE)
            Xsearch.send_keys(int(Fatura))
            search_Button = web.find_element(By.XPATH, '/html/body/div[3]/div/div/ng-include/div[2]/div[1]/div[1]/div/div/div/agencia-faturamento/div/div/div[2]/div/div[1]/div/div[1]/div/div[3]/button')
            web.execute_script("arguments[0].click()", search_Button)
            tm.sleep(1)
            
            ## Achando o menu de download da fatura
            X_Hover = web.find_element(By.XPATH, '/html/body/div[3]/div/div/ng-include/div[2]/div[1]/div[1]/div/div/div/agencia-faturamento/div/div/div[2]/div/div[1]/div/div[2]/div/div/div/div/div/div/div[2]/table/tbody/tr[2]/td[6]/div/span[1]/span')
            ActionChains(web).move_to_element(X_Hover).perform()
            tm.sleep(1)
            
            ## Achando o botão de download da fatura
            X_Hover = web.find_element(By.XPATH, '/html/body/div[5]/div/div/ul/li[3]')
            ActionChains(web).move_to_element(X_Hover).perform()
            tm.sleep(1)
            
            ## Fazendo o download da fatura
            Xdownload_Button = web.find_element(By.XPATH, '/html/body/div[6]/div/div/ul/li[2]/p')
            tm.sleep(1)
            web.execute_script("arguments[0].click()", Xdownload_Button)
            tm.sleep(3)
            
            ## Renomeando o arquivo e movendo o arquivo
            fatura_FileName = str(Fatura) + ".xlsx"
            os.replace(dowloads_Folder / "itens-cobranca.xlsx", folder / fatura_FileName)
            
            bill_file.append(fatura_FileName)
            tm.sleep(1)
        
        
    return bill_file

def compile_extract(bill: pd.DataFrame, folder: Path):
    
    ## Final product will look like this
    extract = pd.DataFrame(
        columns = [
            'REL', 'ID', 'ND', 'FAT', 'VencimentoFatura', 'StatusFatura', 'StatusRelatorio',
            'DataConciliacao', 'StatusSAP', 'Unidade de Negócio', 'CNPJ Unidade de Negócio',
            'Relatório', 'Descrição', 'Motivo', 'Viajante', 'Serviço', 'Data de Emissão',
            'Destino', 'Fornecedor', 'Data Início da Viagem', 'Data Fim da Viagem',
            'Localizador', 'Código Centro de Custo', 'Descrição do Centro de Custo',
            'Rateio de Centro de Custo', 'Valor', 'Código Projeto', 'Nome Projeto'
        ]
    )
    
    ## Reutnrs only unique Fatura values
    bills = bill['Fatura'].unique()
    
    for i in range(len(bills)):

        ## Current Fatura number
        bill_num = bills[i]

        ## Ignores zero values in Fatura
        if bill_num != 0:
            ## Read the extract
            data = pd.read_excel(folder / str(str(bill_num) + ".xlsx"))
    
            ## Hotel services
            bill_hotel = data.loc[data['Serviço'] == str("Hotel")]
            bill_hotel_sum = round(bill_hotel['Valor'].sum(), 2)
    
            ## Other services
            bill_else = data.loc[data['Serviço']!=str("Hotel")]
            bill_else_sum = round(bill_else['Valor'].sum(), 2)
            
            ## Maturity date
            bill_maturity = bill.loc[bill['Fatura'] == int(bill_num)]['vencimento'].iloc[0]
    
            # Hotel services data
            ## Nota number
            invoice_hotel_num = bill.loc[(bill['Fatura'] == int(bill_num)) & (bill['Valor'] == bill_hotel_sum)]['Nota']
            if invoice_hotel_num.empty == False:
                invoice_hotel_num.values[0]
                ## Add extra columns
                bill_hotel.insert(0, "REL", None,True)
                bill_hotel.insert(1, "ID", None,True)
                bill_hotel.insert(2, "ND", int(invoice_hotel_num),True)
                bill_hotel.insert(3, "FAT", int(bill_num), True)
                bill_hotel.insert(4, "VencimentoFatura", bill_maturity, True)
                bill_hotel.insert(5, "StatusFatura", None, True)
                bill_hotel.insert(6, "DataConciliacao", None, True)
                bill_hotel.insert(7, "StatusRelatorio", None, True)
                bill_hotel.insert(8, "StatusSAP", None, True)
    
            # Other services data
            ## Nota number
            invoice_else_num = bill.loc[(bill['Fatura'] == int(bill_num)) & (bill['Valor'] == bill_else_sum)]['Nota']
            if invoice_else_num.empty == False:
                invoice_else_num.values[0]
                ## Add extra columns
                bill_else.insert(0, "REL", None,True)
                bill_else.insert(1, "ID", None,True)
                bill_else.insert(2, "ND", int(invoice_else_num),True)
                bill_else.insert(3, "FAT", int(bill_num), True)
                bill_else.insert(4, "VencimentoFatura", bill_maturity, True)
                bill_else.insert(5, "StatusFatura", None, True)
                bill_else.insert(6, "DataConciliacao", None, True)
                bill_else.insert(7, "StatusRelatorio", None, True)
                bill_else.insert(8, "StatusSAP", None, True)
            
            ## Concat both services
            bill_compiled = pd.concat([bill_hotel, bill_else]).sort_index()
            extract = pd.concat([extract, bill_compiled]).reset_index(drop = True)

    ## Removes NaN values
    extract = extract.where(pd.notnull(extract), None)
    return extract

def save_extract(content: pd.DataFrame, folder: Path):
    extract = folder / "extrato.xlsx"
    with pd.ExcelWriter(path = extract, datetime_format= "DD-MM-YYYY", date_format= "DD-MM-YYYY", mode = 'w') as excel:
        content.to_excel(excel, sheet_name = "Relatórios", index = False, engine = "openpyxl")
    
    return extract

def send_extract(account: str, extract: Path):
    to = "caio.pavesi@manchesterinvest.com.br"
    sb = f"Extrato compilado faturas paytrack {dt.date.today()}"
    bd = """<body>
        <div>
            <p>
                Bom dia,
                <br/><br/>
                Anexo segue faturas do paytrack compiladas.
                <br/><br/>
                Atenciosamente,
                <br/>
                Automação 600000, Caio Pavesi
            </p>
        </div>
    </body>"""
    
    Outlook(account).send_email(to, sb, [extract], bd)
    
    return sb

def cinco_esse(folders: list):
    deleted = []
    for i in range(len(folders)):
        for file in os.listdir(folders[i]):
            try:
                os.remove(folders[i] / file)
                deleted.append(file)
            except:
                pass
    
    return deleted

def main():
    
    Log.output_file = Path.cwd().parent.resolve() / "log.log"
    Log().log(f"600000 [PRD] - Processo iniciado - {dt.datetime.now()}\n")
    
    outlook_account = 'company@account.com'
    outlook_sender = 'partner@email.com.br'
    outlook_folder = "Folder"
    invoice_folder = Path.cwd() / "PRD" / "data" / "600000" / "notas"
    bill_folder = Path.cwd() / "PRD" / "data" / "600000" / "faturas"
    file = "nota_de_debito"

    # download the invoices
    invoice_file = get_invoices(outlook_account, outlook_folder, outlook_sender, search = file, folder = invoice_folder)
    Log().log(f"600000 [PRD] - Notas de débito obtidas - {dt.datetime.now()}\n")
    
    # read the invoices
    bill_data = pd.DataFrame(read_invoices(invoice_file))
    Log().log(f"600000 [PRD] - Notas de débito lidas - {dt.datetime.now()}\n")

    # download bill extract
    bill_file = donwload_extract(bill_data, bill_folder)
    Log().log(f"600000 [PRD] - Faturas baixadas - {dt.datetime.now()}\n")
    
    # compile extract
    bill_comp = compile_extract(bill_data, bill_folder)
    Log().log(f"600000 [PRD] - Fatura compilada - {dt.datetime.now()}\n")
    
    # save extract
    extract = save_extract(bill_comp, bill_folder)
    Log().log(f"600000 [PRD] - Fatura compilada salva - {dt.datetime.now()}\n")
    
    # send extract
    report = send_extract(outlook_account, extract)
    Log().log(f"600000 [PRD] - Fatura compilada enviada - {dt.datetime.now()}\n")
    
    # organize folder
    deleted = cinco_esse([bill_folder, invoice_folder])
    Log().log(f"600000 [PRD] - Arquivos temporários excluídos - {dt.datetime.now()}\n")
    
    Log().log(f"600000 [PRD] - Processo encerrado com sucesso - {dt.datetime.now()}\n")
    
    return 0

main()