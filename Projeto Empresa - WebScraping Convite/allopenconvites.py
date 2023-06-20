import time

#Importando o debugger do python para criar breakpoints e verificar se está funcionando o código criado
import pdb

#Definir um determinado período de tempo para que um elemento apareça antes de prosseguir é com Web Driver Wait. 
from selenium.webdriver.support.wait import WebDriverWait

#EC significa condições esperadas, que são condições que devem ser cumpridas para que uma determinada ação possa ser tomada
from selenium.webdriver.support import expected_conditions as EC

import os

##Importando o webdriver da biblioteca Selenium para acessar a Internet e carregar páginas
from selenium import webdriver

##Importando o Options para poder manipular o webdriver e suas propriedades, 
#como executar várias operações, desativar extensões, desativar pop-ups, etc
from selenium.webdriver.chrome.options import Options

#Importando as chaves especiais para utilizar teclas no teclado ao usar Selenium e 
#apagar textos pré-preenchido em campos de entradas
from selenium.webdriver.common.keys import Keys

#Importando o By para localizar elementos em uma página web
from selenium.webdriver.common.by import By

#Importando ActionChains para utilizar o mouse e teclas do teclado
from selenium.webdriver.common.action_chains import ActionChains

#Pandas para criar, manipular e visualizar tabelas
import pandas as pd

#Para formatar data e horário
from datetime import datetime

#OpenXL para escrever arquivos .xls (Excel) com o Pandas por meio do módulo "Xmlt" para arquivos .xls
import openpyxl 

#Time para utilizar datas quando salvar os arquivos .xls
import time

import re

#Pegando a data e horário de hoje, no momento que criou o arquivo .xls 
TodayDate = time.strftime("%d-%m-%Y %H-%M-%S")
DateSheet = time.strftime("%d-%m-%Y")

#Criando um nome padronizado para os arquivos .xls que terá os detalhes de cada OC
excelfilename = "Convite Detalhado Completo - " + TodayDate + ".xls"

username = os.getlogin()  # Obtém o nome do usuário da máquina
path_with_filename = os.path.join("C:\\", "WebScraping Licitações - Convite", "Tabela OCs - Convite", excelfilename)

#home_dir = os.path.expanduser("~")
#path_with_filename = os.path.join(home_dir, "OneDrive - Ponto Mix Comercial e Serviços", "Miscelânea", "Programas", "Ferramentas - BEC", "Webscraping - Resultados", "Tabela OCs - Convite", excelfilename)

#Criando um nome padronizado para os arquivos .xls que terá as informações da tabela das OCs
excelfilenameallocs = "Tabela Convites Completo - " + TodayDate + ".xls" 

path_with_filenameallocs = os.path.join("C:\\", "WebScraping Licitações - Convite", "Detalhes Produtos - Convite", excelfilenameallocs)

#Criando um nome padronizado para as folhas do .xls
excelsheet = "Convite Completo - " + DateSheet

#Importando as variáveis de ambiente para utilizar com segurança o login e senha do usuário
from config import database_infos

get_login = database_infos['login']
get_pass = database_infos['password']
get_username_pc = database_infos['username_pc']

def bec_allconvites():
    
    browser_driver = webdriver.Chrome()

    #Fazendo solicitação para abrir e navegar na página da BEC
    browser_driver.get("https://www.bec.sp.gov.br/BECSP/Home/Home.aspx")

    #Inicializando o WebDriverWait
    waitWDW = WebDriverWait(browser_driver, 10)

    #Maximizando a Tela do Browser
    browser_driver.maximize_window()

    #Confirmando que é o site correto aquele que está aberto
    assert "BEC" in browser_driver.title

    #Procurando a tag certa do botão "Negociações Eletrônicas" 
    btn_ne = browser_driver.find_element(By.LINK_TEXT, "Negociações Eletrônicas")
 
    ##Fazendo com que clique no botão
    btn_ne.send_keys(Keys.RETURN)

    #Procurando as tags certas com XPATH e preenchendo os campos "CNJP/CPF" e "Senha"
    login = browser_driver.find_element(By.XPATH, "//input[@id='TextLogin']") #Se parar de funcionar, utilize a class="TextLogin" ou o id="TextLogin"
    login.send_keys(get_login)

    password = browser_driver.find_element(By.XPATH, "//*[@id='TextSenha']") #Se parar de funcionar, utilize a class="TextSenha" ou o id="TextSenha"
    password.send_keys(get_pass)

    #Marcando a caixa de declaração
    statement_box = browser_driver.find_element(By.XPATH, "//*[@id='chkAceite']") #Se parar de funcionar, utilize a class="chkAceite" ou o id="chkAceite"
    statement_box.click()

    #Clicando no botão de entrar
    btn_enter = browser_driver.find_element(By.ID, "Btn_Confirmar") #Se parar de funcionar, utilize o id="Btn_Confirmar"
    btn_enter.click()

    current_url = browser_driver.current_url
    
    if current_url == "https://www.bec.sp.gov.br/fornecedor_ui/TermoResponsabilidade.aspx?Dzqeio6gALuoR%2flQf2tFB6zBkp9ETq5P44%2bgrURdFf66JmFgqUpWHFjTKO2RLNZR":
        waitWDW = WebDriverWait(browser_driver, 10)
        reconfirm_checkbox = browser_driver.find_element(By.ID, "//*[@id='ctl00_c_area_conteudo_chkDeclaracao']")
        reconfirm_checkbox.click()
        ok_button = browser_driver.find_element(By.ID, "//*[@id='ctl00_c_area_conteudo_Button1']")
        ok_button.click()
        #Passando o mouse por cima da lista "Participar"
        join_menu_list = waitWDW.until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Participar']")))
        actions = ActionChains(browser_driver)
        actions.move_to_element(join_menu_list).pause(2).perform()
        
        #Escolhendo o item da lista certa, que é o Convite e clicando nele
        pe_item_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//li[@id='2100121']//a[contains(text(),'Convite Eletrônico')]")))
        pe_item_list.click()

        pe_btn_search = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_bt_Pesquisa']")))  #Se parar de funcionar, utilizar o id="ctl00_conteudo_Pesquisa", css_selector="#pesquisa" ou text_link="Pesquisar"
        pe_btn_search.click()
        time.sleep(2) 
    
    else:
        #Passando o mouse por cima da lista "Participar"
        join_menu_list = waitWDW.until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Participar']")))
        actions = ActionChains(browser_driver)
        actions.move_to_element(join_menu_list).pause(2).perform()
        
        
        #Escolhendo o item da lista certa, que é o Convite e clicando nele
        pe_item_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//li[@id='2100121']//a[contains(text(),'Convite Eletrônico')]")))
        pe_item_list.click()
        
        pe_btn_search = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_bt_Pesquisa']")))  #Se parar de funcionar, utilizar o id="ctl00_conteudo_Pesquisa", css_selector="#pesquisa" ou text_link="Pesquisar"
        pe_btn_search.click()
        time.sleep(2)
    
    ###Tabela para armazenar informações das OCs###
    
    result_all_ocs_table = []    
        
    ###Lista para armazenar os resultados da coleta de dados do Convite###
    result_all_ocs = []
    
    #Lista para armazenar os resultados da coleta de dados da descrição, quantidade, uf, telefone e e-mails do Convite
    details_infos_invitation = []
    
    #Lista no qual unirá os valores (Informações básicas de cada OC) da primeira tabela com os detalhes de cada OC (segunda tabela)
    merged_list = []
    
    
    global iterator, n, i, j
    iterator = 1
    
    #A quantidade de páginas presentes na aba Convite (pegando o comprimento -> total)
    total_pages = len(browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr[32]/td/table/tbody/tr/td/a")) + 1

    n = 0
    i = 1
    j = 1
    
    
    while True:   
        

                for page in range(total_pages):

                                    #A quantidade de linhas da tabela presente em cada página
                                    rows = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr")[1:]

                                    #for i, row in enumerate(rows[iterator:], start=iterator):
                                    for i, row in enumerate(rows[iterator:], start=iterator):
                                        
                                        
                                            #Procurando todos os elementos da tabela dos Convites
                                            #Pegando os valores das colunas Oferta de Compra, Previsão de Abertura Fim e Unidade Comprador
                                            invitation_oc = browser_driver.find_elements(By.XPATH, f"//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr[{i+1}]/td[4]") [0]
                                            invitation_uncompradora_orgao = browser_driver.find_elements(By.XPATH, f"//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr[{i+1}]/td[5]/table/tbody/tr[2]") [0]
                                            invitation_uncompradora_municipio = browser_driver.find_elements(By.XPATH, f"//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr[{i+1}]/td[5]/table/tbody/tr[3]") [0]
                                            invitation_date_proposal_end = browser_driver.find_elements(By.XPATH, f"//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr[{i+1}]/td[7]/table/tbody/tr/td[2]")[0].text
                                            
                                            

                                            #date_time_list = invitation_date_proposal_end.split(" ")
                                            
                                            #date_value = date_time_list[0]
                                            #time_value = date_time_list[1]
                                            
                                            #Convertendo nos tipos de dados corretos
                                            #date_obj = datetime.strptime(date_value, '%d/%m/%Y').date()
                                            #time_obj = datetime.strptime(time_value, '%H:%M:%S').time()
                                            
                                            result_all_ocs_table.append({
                                                        "Unidade Compradora": invitation_uncompradora_orgao.text,
                                                        "Município": invitation_uncompradora_municipio.text,
                                                        "OC":invitation_oc.text,
                                                        "Data Final":invitation_date_proposal_end,
                                                        "Horário Final":invitation_date_proposal_end})

                                            result_all_ocs.append({"Unidade Compradora": invitation_uncompradora_orgao.text,
                                                    "Município": invitation_uncompradora_municipio.text,
                                                    "OC":invitation_oc.text,
                                                    "Data Final":invitation_date_proposal_end,
                                                    "Horário Final":invitation_date_proposal_end})
                                    
                                            #link = waitWDW.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/form/div[3]/div/div/div/div/div[2]/div[4]/div[2]/div/table/tbody/tr[{i+1}]/td[4]/a[2]")))
                                            link = waitWDW.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/form/div[3]/div/div/div/div/div[2]/div[4]/div[2]/div/table/tbody/tr[{i+1}]/td[4]/a[2]")))
                                            link.click()

                                            #Botão do Convite para pegar as informações essenciais (descrição dos produtos)
                                            invitation_button =  waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='topMenu']/li[3]/a")))
                                            ActionChains(browser_driver).key_down(Keys.CONTROL).click(invitation_button).perform()
                                            browser_driver.switch_to.window(browser_driver.window_handles[-1])
                                            time.sleep(2)

                                            #Pegando o número da OC para colocar junto com os detalhes dos itens
                                            oc_number_invitation = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_DetalhesOfertaCompra1_txtOC']")
                                            #Tabela com os detalhes dos itens que estão sendo solicitidas pela OC no Convite
                                            details_table_oc_invitation = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdv_item']")
                                            #Linhas da tabela
                                            rows_details_oc_invitation = details_table_oc_invitation.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdv_item']/tbody/tr")[1:]

                                            #Se a tabela possuir mais do que 1 item (1 linha com detalhes da descrição), o seguinte código será executado:
                                            if len(rows_details_oc_invitation) > 1:
                                            
                                                            item_values = []
                                                            code_values = []
                                                            description_values = []
                                                            quantity_values = []
                                                            uf_values = []


                                                            for row in rows_details_oc_invitation:
                                                            
                                                                        item_value = row.find_element(By.XPATH, "./td[4]").text
                                                                        code_value = row.find_element(By.XPATH, "./td[5]").text
                                                                        description_value = row.find_element(By.XPATH, "./td[6]").text
                                                                        quantity_value = row.find_element(By.XPATH, "./td[7]").text
                                                                        quantity_value = quantity_value.replace(".","")
                                                                        uf_value = row.find_element(By.XPATH, "./td[8]").text
                                                                        #oc_number_invitation = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_DetalhesOfertaCompra1_txtOC']").text
                                                                        #oc_number_invitation = oc_number_invitation[:22].strip()

                                                                        item_values = int(item_value)
                                                                        code_values = int(code_value)  
                                                                        description_values = description_value
                                                                        quantity_values = int(float(quantity_value))
                                                                        uf_values = uf_value
                                                                        #oc_number_invitation = oc_number_invitation

                                                                        details_infos_invitation.append({
                                                                            #"OC": oc_number_invitation,
                                                                            "Item": item_values,
                                                                            "SIAF.": code_values,
                                                                            "Desc.": description_values,
                                                                            "Qtd.": quantity_values,
                                                                            "UN": uf_values})
                                                            
                                                            for item in result_all_ocs:
                                                                    for i, detail in enumerate(details_infos_invitation):
                                                                            if i == 0:
                                                                                    merged_item = item.copy()
                                                                                    merged_item.update(detail)
                                                                                    merged_list.append(merged_item)
                                                                            else:
                                                                                    merged_item = item.copy()
                                                                                    merged_item.update(detail)
                                                                                    merged_list.append(merged_item)
                                                            
                                                            #Fechando a aba e voltando para a aba principal
                                                            browser_driver.close()
                                                            browser_driver.switch_to.window((browser_driver.window_handles[0]))
                                                            browser_driver.back()
                                                            iterator+=1
                                                            result_all_ocs = []
                                                            details_infos_invitation = []


                                            #O seguinte código será executado se apenas tiver um item na tabela do pedido da OC        
                                            else:
                                            
                                                            for row in rows_details_oc_invitation:
                                                            
                                                                        item_value = row.find_element(By.XPATH, "./td[4]").text
                                                                        code_value = row.find_element(By.XPATH, "./td[5]").text
                                                                        description_value = row.find_element(By.XPATH, "./td[6]").text
                                                                        quantity_value = row.find_element(By.XPATH, "./td[7]").text
                                                                        quantity_value = quantity_value.replace(".","")
                                                                        uf_value = row.find_element(By.XPATH, "./td[8]").text
                                                                        #oc_number_invitation = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_DetalhesOfertaCompra1_txtOC']").text
                                                                        #oc_number_invitation = oc_number_invitation[:22].strip()

                                                                        details_infos_invitation.append({
                                                                            #"OC": oc_number_invitation,
                                                                            "Item": int(item_value),
                                                                            "SIAF.": int(code_value),
                                                                            "Desc.": description_value,
                                                                            "Qtd.": int(float(quantity_value)),
                                                                            "UN": uf_value})
                                                                        
                                                            for item in result_all_ocs:
                                                                    for i, detail in enumerate(details_infos_invitation):
                                                                            if i == 0:
                                                                                    merged_item = item.copy()
                                                                                    merged_item.update(detail)
                                                                                    merged_list.append(merged_item)
                                                                            else:
                                                                                    merged_item = item.copy()
                                                                                    merged_item.update(detail)
                                                                                    merged_list.append(merged_item)
                                                            
                                                            #Fechando a aba e voltando para a aba principal
                                                            browser_driver.close()
                                                            browser_driver.switch_to.window((browser_driver.window_handles[0]))
                                                            browser_driver.back()
                                                            iterator+=1
                                                            result_all_ocs = []
                                                            details_infos_invitation = []

                                            time.sleep(3.5)
                                            #Pause de 20 segundos depois de fazer Scraping de 15 páginas, para depois continuar e diminuir as chances de dar algum erro             
                                            #if (i+1) % 15 == 0:
                                            #    time.sleep(20)

                                    if j < total_pages:
                                    
                                            next_button = browser_driver.find_element(By.XPATH, f"//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr[32]/td/table/tbody/tr/td[{j + 1}]/a")   
                                            next_button.click()

                                            j+=1
                                            iterator = 1
                                            time.sleep(10)
                                    else:
                                        break
                n +=1
        
                time.sleep(2)
            


                #Fechando o Browser depois de terminar
                browser_driver.close()
                browser_driver.quit()
    
                break
                                                ######################################################################################
                                                ######Criando uma tabela para visualizar os valores coletados da lista do Convite#####
                                                ######################################################################################
            
    #Utilizando pandas para criar e visualizar uma tabela formatada com os valores coletados da lista do Convite
    
    #Tabela de Tabela Convites Completa 
    df_table_allocs = pd.DataFrame(result_all_ocs_table)
    
    #Valores coletados dentro de cada OC (detalhes)
    df_oc_details_invitation = pd.DataFrame(merged_list) 
    

    #Criando um Pandas Excel Writer para usar o Openpyxl como engine e salvar os detalhes das OCs selecionadas.
    writer = pd.ExcelWriter(path_with_filename, engine='openpyxl')
    
    #Criando um Pandas Excel Writer para salvar os dados atuais da tabela de OCs
    writer2 = pd.ExcelWriter(path_with_filenameallocs, engine='openpyxl')
    
    #Criando um arquivo .xls para utilizar os dados dos detalhes de OCs no Excel
    df_final_data = df_oc_details_invitation.to_excel(writer, sheet_name=DateSheet, header=True, index=False) #axis 1 é para colocar depois da coluna da primeira tabela, enquanto 0 é para colocar depois da última linha
    
    #Criando arquivo .xls para ver os dados gerais da tabela de OCs
    df_table_allocs.to_excel(writer2, sheet_name=DateSheet, header=True, index=False)
    
    print(df_oc_details_invitation.dtypes)
    
    #Fechando o Pandas Excel Writer e fazendo o output do arquivo .xls
    writer.close()
    writer2.close() 

    print(df_final_data)
    print('DataFrame is written to Excel File successfully!!!')
    
    
#bec_allconvites()