# -*- coding: utf-8 -*-
"""
Created on Thu Apr  2 18:27:45 2020

@author: Aline Tenório
"""
import  xlrd, requests, pandas as pd
from bs4 import BeautifulSoup
import urllib.request, re

#filename = 'DIARIO_01-04-2020.xlsx'
url_xlsx = 'http://sdro.ons.org.br/SDRO/DIARIO/2020_04_01/Html/DIARIO_01-04-2020.xlsx'
url = 'http://sdro.ons.org.br/SDRO/DIARIO/2020_04_01/index.htm'

#load local .xlsx and store specific cell values
def load_local_xlsx(filename, sheetname):
    workbook = xlrd.open_workbook(filename)
    
    #banlanco de energia
    sheets = workbook.sheet_names()
    
    worksheet = workbook.sheet_by_name(sheetname)
    #active_sheet = excel1_arquivo_ob.active
    # reading a cell
    #print(worksheet["A1"].value)
    
    #1 PLANILHA - 

    # GERAÇÃO
    #sin
    #print(worksheet.cell_value(13,3 ))
    
    dados = {}
    
    
    dados["sin_hidro_nacional"] = worksheet.cell_value(5,3)
    dados["sin_hidro_nacional_percent"] = worksheet.cell_value(5,4)*100
    dados["sin_itaipu_binacional"] = worksheet.cell_value(6,3)
    


    return dados

def load_url_xlsx(url):
    
    #Setting sheet_name as None will load all sheets
    df = pd.read_excel(url,sheet_name=None, header=None)
    
    
    #To access specific sheet: df[sheetname]
    sheet = df["01-Balanço de Energia"] 
    
    dados = {}
    
    #sheet.values: get cell values
    dados["sin_hidro_nacional"] = sheet.values[5,3]
    dados["sin_hidro_nacional_percent"] = sheet.values[5,4]*100
    dados["sin_itaipu_binacional"] = sheet.values[6,3]
    
 
    sheet = df["20-Variação Energia Armazenada"] 

    dados["sul_cap_max_arm"] = sheet.values[5,1]
    dados["seco_cap_max_arm"] = sheet.values[5,2]
    
 
    return dados

def get_href_html(url):
   # html = open(url).read()
    html = urllib.request.urlopen(url)
    #return html
    soup = BeautifulSoup(html,features="lxml")
    all = soup.find_all('a')
    
    for link in soup.findAll('a'):
        doc = link.get('href')
        
        if("DIARIO" in doc and ".xlsx" in doc):
           xlsx = doc
           break
       
    

#url_xlsx = get_href_html(url)

content = load_url_xlsx(url_xlsx)

#load_local_xlsx(filename)