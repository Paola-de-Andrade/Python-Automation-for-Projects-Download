#Download automático propostas de projeto CPP 

#pip install openpyxl==2.6.3
#pip install selenium
#pip install time
#pip install datetime
#pip install tkinter
#pip install pandas
#pip install pywin32
#pip install os
#pip install requests
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium import webdriver
import time
import openpyxl
from openpyxl.styles import colors
from openpyxl.styles import Font, Color 
import tkinter as tk
from tkinter import messagebox
import win32com.client
import os
import requests

id_inicial = 1690
ultimo_id = 1786
link_login = "http://sefic.cpfl.com.br/SEFICad"
k=0

#iteração para buscar fazer login no SEFIC mesmo se cachê estiver cheio e ficar dando erro
while k <100:
    try:
        driver = webdriver.Chrome()
        driver.get(link_login)
        time.sleep(5)
        teste_home = driver.find_element_by_xpath("/html/body/div/nav/div/div/div[3]/div[2]/ul/li[1]/a")
        break
    except:
        driver.close()
        time.sleep(20)
        k=k+1
        continue

k=0

#loop para entrar em cada proposta de projeto que necessita de download
for id_projeto in range(id_inicial, ultimo_id+1):
    link_projeto = "http://sefic.cpfl.com.br/SEFIC/ModuloProjeto/ParticipanteProcessoChamadaPublica/Visualizar/"+str(id_projeto)+"?origem=interno"
    driver.get(link_projeto)
    UC = driver.find_element_by_xpath("/html/body/div/div/div/div/div[2]/div[1]/div[1]/fieldset[1]/div[2]/div/div").text
    nome_cliente = driver.find_element_by_xpath("/html/body/div/div/div/div/div[2]/div[1]/div[1]/fieldset[1]/div[3]/div/div").text
    #print(UC)
    #print(nome_cliente)
    processo = driver.find_element_by_xpath("/html/body/div/div/div/div/div[2]/div[1]/div[1]/fieldset[1]/div[1]/div/div").text
    if processo.find("PAULISTA") > 0:
        distribuidora = "\CPFL Paulista"
    else:
        if processo.find("PIRA") > 0:
                distribuidora = "\CPFL Piratininga"
        else:
            if processo.find("CRUZ") > 0:
                    distribuidora = "\CPFL Santa Cruz"
            else:
                if processo.find("RGE") > 0:
                        distribuidora = "\RGE Sul"
                else:
                    ROOT = tk.Tk()

                    ROOT.withdraw()

                    messagebox.showinfo("Erro distribuidora", "Não consegui achar de qual distribuidora é o processo:"+ processo)
    #print(distribuidora)
    parent_dir = r"C:\Users\2003305\Documents\Download Projetos SEFIC CPP21" + distribuidora 
    #print(parent_dir)
    
    folder_name = str(UC)+ "_" + nome_cliente[:20]
    #print (folder_name)
    path = os.path.join(parent_dir, folder_name)
    # Create the directory
    os.mkdir(path)
    
    #loop para baixar todos os arquivos de cada proposta de projeto
    for doc_projeto in range(1,16):
        if doc_projeto == 1:
            file_name = str("\\"+ UC + "_CARTA.pdf")
        if doc_projeto == 2:
            file_name = str("\\"+ UC + "_DIAGNOSTICO.pdf")
        if doc_projeto == 3:
            file_name = str("\\"+ UC + "_CATALOGO.pdf")
        if doc_projeto == 4:
            file_name = str("\\"+ UC + "_MEMORIA.xlsm")
        if doc_projeto == 5:
            file_name = str("\\"+ UC + "_EXPERIENCIA.pdf")
        if doc_projeto == 6:
            file_name = str("\\"+ UC + "_ORÇAMENTO.pdf")
        if doc_projeto == 7:
            file_name = str("\\"+ UC + "_ART.pdf")
        if doc_projeto == 8:
            file_name = str("\\"+ UC + "_CONTR ou ESTAT SOCIAL.pdf")
        if doc_projeto == 9:
            file_name = str("\\"+ UC + "_CNPJ.pdf")
        if doc_projeto == 10:
            file_name = str("\\"+ UC + "_INSS.pdf")
        if doc_projeto == 11:
            file_name = str("\\"+ UC + "_FGTS.pdf")
        if doc_projeto == 12:
            file_name = str("\\"+ UC + "_SIMPLES NACIONAL.pdf")
        if doc_projeto == 13:
            file_name = str("\\"+ UC + "_FORMUL_CLIENTE.xlsx")
        if doc_projeto == 14:
            file_name = str("\\"+ UC + "_DEMONSTR ou FILANTR.pdf")
        if doc_projeto == 15:
            file_name = str("\\"+ UC + "_CMVP.pdf")      

        doc_path = str(path + file_name)                
        doc= driver.find_element_by_xpath("/html/body/div/div/div/div/div[2]/div[1]/div[2]/fieldset/div/div["+str(doc_projeto)+"]/div[2]/div/a")
        file_URL = doc.get_attribute('href')
        r = requests.get(file_URL, stream = True)
        with open(doc_path,"wb") as file:
            file.write(r.content) 
    
    k=k+1

ROOT = tk.Tk()

ROOT.withdraw()

messagebox.showinfo("Status do download", "Ufa! Acabei aqui! Fiz o download de "+ str(k)+" propostas de projeto! Boa avaliação!")

driver.close()
