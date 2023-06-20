import time
import pdb
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
from datetime import datetime
import openpyxl 
import time
import re


TodayDate = time.strftime("%d-%m-%Y %H-%M-%S")
DateSheet = time.strftime("%d-%m-%Y")

# Obtenha o caminho absoluto para a pasta raiz do projeto
root_directory = os.path.dirname(os.path.abspath(__file__))

destination_directory = "Tabela OCs - Convite"

excelfilename = "Convite Detalhado Reduzido - " + TodayDate + ".xls"

username = os.getlogin()  # Obtém o nome do usuário da máquina
path_with_filename = os.path.join("C:\\", "WebScraping Licitações - Convite", "Tabela OCs - Convite", excelfilename)

excelfilenameallocs = "Tabela Convites Reduzido - " + TodayDate + ".xls" 

path_with_filenameallocs = os.path.join("C:\\", "WebScraping Licitações - Convite", "Detalhes Produtos - Convite", excelfilenameallocs)

excelsheet = "Convite Reduzido - " + DateSheet

from config import database_infos

get_login = database_infos['login']
get_pass = database_infos['password']
get_username_pc = database_infos['username_pc']

def bec_filterconvites(field_value):
    
    name_category = field_value
   
    browser_driver = webdriver.Chrome()

    browser_driver.get("https://www.bec.sp.gov.br/BECSP/Home/Home.aspx")
   
    waitWDW = WebDriverWait(browser_driver, 10)
    
    browser_driver.maximize_window()

    assert "BEC" in browser_driver.title

    btn_ne = browser_driver.find_element(By.LINK_TEXT, "Negociações Eletrônicas")
 
    btn_ne.send_keys(Keys.RETURN)

    login = browser_driver.find_element(By.XPATH, "//input[@id='TextLogin']")
    login.send_keys(get_login)

    password = browser_driver.find_element(By.XPATH, "//*[@id='TextSenha']")
    password.send_keys(get_pass)

    statement_box = browser_driver.find_element(By.XPATH, "//*[@id='chkAceite']") 
    statement_box.click()

    btn_enter = browser_driver.find_element(By.ID, "Btn_Confirmar") 
    btn_enter.click()

    current_url = browser_driver.current_url
    
    if current_url == "https://www.bec.sp.gov.br/fornecedor_ui/TermoResponsabilidade.aspx?Dzqeio6gALuoR%2flQf2tFB6zBkp9ETq5P44%2bgrURdFf66JmFgqUpWHFjTKO2RLNZR":
        waitWDW = WebDriverWait(browser_driver, 10)
        reconfirm_checkbox = browser_driver.find_element(By.ID, "//*[@id='ctl00_c_area_conteudo_chkDeclaracao']")
        reconfirm_checkbox.click()
        ok_button = browser_driver.find_element(By.ID, "//*[@id='ctl00_c_area_conteudo_Button1']")
        ok_button.click()

        join_menu_list = waitWDW.until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Participar']")))
        actions = ActionChains(browser_driver)
        actions.move_to_element(join_menu_list).pause(2).perform()
        
        pe_item_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//li[@id='2100121']//a[contains(text(),'Convite Eletrônico')]")))
        pe_item_list.click()

        pe_btn_search = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_bt_Pesquisa']")))
        pe_btn_search.click()
        time.sleep(2) 
    
    else:
        join_menu_list = waitWDW.until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Participar']")))
        actions = ActionChains(browser_driver)
        actions.move_to_element(join_menu_list).pause(2).perform()
        
        pe_item_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//li[@id='2100121']//a[contains(text(),'Convite Eletrônico')]")))
        pe_item_list.click()
        
        pe_btn_search = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_bt_Pesquisa']")))  
        pe_btn_search.click()
        time.sleep(2)
        
        
        
    #Campo para inserir o nome do produto que será pesquisado
    # //*[@id="ctl00_c_area_conteudo_wuc_filtroPesquisaOc1_c_txt_ItemMaterial_desc"]
    
    input_category = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_c_area_conteudo_wuc_filtroPesquisaOc1_c_txt_ItemMaterial_desc']") 
    input_category.send_keys(name_category)
    
    button_advanced_search = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_c_area_conteudo_bt_Pesquisa']"))) 
    button_advanced_search.click()
    
    result_all_ocs_table = []

    result_all_ocs = []
    
    details_infos_invitation = []
    
    merged_list = []

    
    global iterator, n, i, j
    iterator = 1

    n = 0
    i = 1
    j = 1
    
    
    
    #Verificando se existe ou não paginação na página para fazer o Scraping dos campos e verificar os detalhes de cada Convite
        
        
    try:
        
        pagination_element = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr[32]/td/table/tbody/tr/td/a")
        total_pages = len(pagination_element) + 1
        
        if total_pages > 1:
                        
                            while True:
                            
                                for page in range(total_pages):
                                
                                            rows = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr")[1:]
                
                
                                            for i, row in enumerate(rows[iterator:], start=iterator):
                                                
                                                    invitation_oc = browser_driver.find_elements(By.XPATH, f"//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr[{i+1}]/td[4]") [0]
                                                    invitation_uncompradora_orgao = browser_driver.find_elements(By.XPATH, f"//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr[{i+1}]/td[5]/table/tbody/tr[2]") [0]
                                                    invitation_uncompradora_municipio = browser_driver.find_elements(By.XPATH, f"//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr[{i+1}]/td[5]/table/tbody/tr[3]") [0]
                                                    #join_orgao_municipío = f"{invitation_uncompradora_orgao.text}-{invitation_uncompradora_municipio.text}"
                                                    invitation_date_proposal_end = browser_driver.find_elements(By.XPATH, f"//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr[{i+1}]/td[7]/table/tbody/tr/td[2]")[0].text
                                                    
                                                    #Separando a data do horário 
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

                                                    result_all_ocs.append({
                                                    "Unidade Compradora": invitation_uncompradora_orgao.text,
                                                    "Município": invitation_uncompradora_municipio.text,
                                                    "OC":invitation_oc.text,
                                                    "Data Final":invitation_date_proposal_end,
                                                    "Horário Final":invitation_date_proposal_end})

                                            
                                                    link = waitWDW.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/form/div[3]/div/div/div/div/div[2]/div[4]/div[2]/div/table/tbody/tr[{i+1}]/td[4]/a[2]")))
                                                    link.click()
                
                                                    invitation_button =  waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='topMenu']/li[3]/a")))
                                                    ActionChains(browser_driver).key_down(Keys.CONTROL).click(invitation_button).perform()
                                                    browser_driver.switch_to.window(browser_driver.window_handles[-1])
                                                    time.sleep(2)
                
                                                    oc_number_invitation = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_DetalhesOfertaCompra1_txtOC']")
                                                    details_table_oc_invitation = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdv_item']")
                                                    rows_details_oc_invitation = details_table_oc_invitation.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdv_item']/tbody/tr")[1:]
                
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
                                                                                    "Filtro": name_category,
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

                
                                                                    browser_driver.close()
                                                                    browser_driver.switch_to.window((browser_driver.window_handles[0]))
                                                                    browser_driver.back()
                                                                    iterator+=1
                                                                    result_all_ocs = []
                                                                    details_infos_invitation = []
                    
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
                                                                                    "Filtro": name_category,
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

                                                                            browser_driver.close()
                                                                            browser_driver.switch_to.window((browser_driver.window_handles[0]))
                                                                            browser_driver.back()
                                                                            iterator+=1
                                                                            result_all_ocs = []
                                                                            details_infos_invitation = []
                                                    
                                                    time.sleep(4)                
                                                    #Pause de 20 segundos depois de fazer Scraping de 15 páginas, para depois continuar e diminuir as chances de dar algum erro             
                                                    #if (i+1) % 15 == 0:
                                                    #    time.sleep(20)
                                                                    
                                            if j < total_pages:
                                
                                                    next_button = browser_driver.find_element(By.XPATH, f"//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr[32]/td/table/tbody/tr/td[{j + 1}]/a")   
                                                    next_button.click()

                                                    j+=1
                                                    iterator = 1
                                            else:
                                                break
                                n +=1
                        
                                time.sleep(2)
                                browser_driver.close()
                                browser_driver.quit()

        else:
            
            
                    for page in range(total_pages):
            
                            rows = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr")[1:]
                
                            for i, row in enumerate(rows, start=1):
                        
                        
                                            invitation_oc = browser_driver.find_elements(By.XPATH, f"//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr[{i+1}]/td[4]") [0]
                                            invitation_uncompradora_orgao = browser_driver.find_elements(By.XPATH, f"//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr[{i+1}]/td[5]/table/tbody/tr[2]") [0]
                                            invitation_uncompradora_municipio = browser_driver.find_elements(By.XPATH, f"//*[@id='ctl00_c_area_conteudo_grdvOC']/tbody/tr[{i+1}]/td[5]/table/tbody/tr[3]") [0]
                                            #join_orgao_municipío = f"{invitation_uncompradora_orgao.text}-{invitation_uncompradora_municipio.text}"
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
                    
                                            link = waitWDW.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/form/div[3]/div/div/div/div/div[2]/div[4]/div[2]/div/table/tbody/tr[{i+1}]/td[4]/a[2]")))
                                            link.click()

                                            invitation_button =  waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='topMenu']/li[3]/a")))
                                            ActionChains(browser_driver).key_down(Keys.CONTROL).click(invitation_button).perform()
                                            browser_driver.switch_to.window(browser_driver.window_handles[-1])
                                            time.sleep(2)

                                            oc_number_invitation = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_DetalhesOfertaCompra1_txtOC']")
                                            details_table_oc_invitation = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdv_item']")
                                            rows_details_oc_invitation = details_table_oc_invitation.find_elements(By.XPATH, "//*[@id='ctl00_c_area_conteudo_grdv_item']/tbody/tr")[1:]

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
                                                                            "Filtro": name_category,
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

                                                            browser_driver.close()
                                                            browser_driver.switch_to.window((browser_driver.window_handles[0]))
                                                            browser_driver.back()
                                                            iterator+=1
                                                            result_all_ocs = []
                                                            details_infos_invitation = []

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
                                                                            "Filtro": name_category,
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

                                                            browser_driver.close()
                                                            browser_driver.switch_to.window((browser_driver.window_handles[0]))
                                                            browser_driver.back()
                                                            iterator+=1
                                                            result_all_ocs = []
                                                            details_infos_invitation = []
                                            
                                            time.sleep(4)                
                                            #Pause de 20 segundos depois de fazer Scraping de 15 páginas, para depois continuar e diminuir as chances de dar algum erro             
                                            #if (i+1) % 15 == 0:
                                            #         time.sleep(20)

    

                    time.sleep(2)
                    browser_driver.close()
                    browser_driver.quit()

            
            
    except:
        print("EXCEPTION OCCURRED")
        
    
                                                ######################################################################################
                                                ######Criando uma tabela para visualizar os valores coletados da lista do Convite#####
                                                ######################################################################################
            

    df_table_allocs = pd.DataFrame(result_all_ocs_table)
    df_oc_details_invitation = pd.DataFrame(merged_list)
    writer = pd.ExcelWriter(path_with_filename, engine='openpyxl')
    writer2 = pd.ExcelWriter(path_with_filenameallocs, engine='openpyxl')
    df_final_data = df_oc_details_invitation.to_excel(writer, sheet_name=DateSheet, header=True, index=False)
    df_table_allocs.to_excel(writer2, sheet_name=DateSheet, header=True, index=False)
    print(df_oc_details_invitation.dtypes)
    print(df_table_allocs.dtypes)
    writer.close()
    writer2.close() 
    print(df_final_data)
    print('DataFrame is written to Excel File successfully!!!')
    