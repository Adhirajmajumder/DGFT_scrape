# -*- coding: utf-8 -*-
"""
Created on Thu Jul 16 10:40:22 2020

@author: ADHIRAJ MAJUMDAR
"""
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options
import re
import csv
import pandas as pd
import time


options = Options()
chromedriver = "chromedriver.exe"
browser = webdriver.Chrome(chromedriver,options=options)
data = pd.read_excel('Input/input_company.xlsx')
with open('Output/IEC_Details_file.csv', 'w', newline='') as outcsv:
    writer = csv.writer(outcsv)
    writer.writerow(["IEC","IEC_Allotment_Date","File_Number","Party_Name_and_Address","Phone_No","e_mail","Exporter_Type","Date_of_Establishment","PAN_ISSUE_DATE","BIN","PAN_ISSUED_BY","Nature_Of_Concern","Bank","Dirct1","Dirct2"])
for index,row in data.iterrows():
    browser.get("http://dgft.delhi.nic.in:8100/dgft/IecPrint")
    IEC = browser.find_element_by_xpath('/html/body/form/input[1]')
    if len(str(row["IEC"]))<=9:
        IEC.send_keys('0'+str(row["IEC"]))
    else:
        IEC.send_keys(row["IEC"])
    company = browser.find_element_by_xpath('/html/body/form/input[2]')
    company.send_keys(row["Party Name"])
    button = browser.find_element_by_xpath('/html/body/form/input[3]')
    button.click()
    Dirct1=None
    Dirct1=None
    
    page_data = browser.page_source
    soup = BeautifulSoup(page_data,"html.parser")
    if soup.find('td',text=re.compile('IEC'))!=None:
        IEC = soup.find('td',text=re.compile('IEC')).find_next("td").find_next("td").text
    else:
        IEC=None
    if soup.find('td',text=re.compile('IEC Allotment Date')) != None:
        IEC_Allotment_Date = soup.find('td',text=re.compile('IEC Allotment Date')).find_next("td").find_next("td").text
    else:
        IEC_Allotment_Date=None
    if soup.find('td',text=re.compile('File Number')) != None:
        File_Number = soup.find('td',text=re.compile('File Number')).find_next("td").find_next("td").text
    else:
        File_Number=None
    if soup.find('td',text=re.compile('Party Name')) != None:
        Party_Name_Address = soup.find('td',text=re.compile('Party Name')).find_next("td").find_next("td").text
    else:
        Party_Name_Address=None
    if soup.find('td',text=re.compile('Phone No')) != None:
        Phone_No = soup.find('td',text=re.compile('Phone No')).find_next("td").find_next("td").text
    else:
        Phone_No=None
    if soup.find('td',text=re.compile('e_mail')) != None:
        e_mail = soup.find('td',text=re.compile('e_mail')).find_next("td").find_next("td").text
    else:
        e_mail=None
    if soup.find('td',text=re.compile('Exporter Type')) != None:
        Exporter_Type = soup.find('td',text=re.compile('Exporter Type')).find_next("td").find_next("td").text
    else:
        Exporter_Type=None
    if soup.find('td',text=re.compile('Date of Establishment')) != None:
        Date_of_Establishment = soup.find('td',text=re.compile('Date of Establishment')).find_next("td").find_next("td").text
    else:
        Date_of_Establishment=None
    if soup.find('td',text=re.compile('BIN')) != None:
        BIN = soup.find('td',text=re.compile('BIN')).find_next("td").find_next("td").text
    else:
        BIN=None
    if soup.find('td',text=re.compile('PAN ISSUE DATE')) != None:
        PAN_ISSUE_DATE = soup.find('td',text=re.compile('PAN ISSUE DATE')).find_next("td").find_next("td").text
    else:
        PAN_ISSUE_DATE=None
    if soup.find('td',text=re.compile('PAN ISSUED BY')) != None:
        PAN_ISSUED_BY = soup.find('td',text=re.compile('PAN ISSUED BY')).find_next("td").find_next("td").text
    else:
        PAN_ISSUED_BY=None
    if soup.find('td',text=re.compile('Nature Of Concern')) != None:
        Nature_Of_Concern = soup.find('td',text=re.compile('Nature Of Concern')).find_next("td").find_next("td").text
    else:
        Nature_Of_Concern=None
    if soup.find('td',text=re.compile('Banker Detail')) != None:
        Bank = soup.find('td',text=re.compile('Banker Detail')).find_next("td").find_next("td").text
    else:
        Bank=None
    if soup.find('td',text=re.compile('^1.$'))!= None:
        Dirct1 = soup.find('td',text=re.compile('^1.$')).find_next("td").text
    else:
        Dirct1=None
    if soup.find('td',text=re.compile("^2.$"))!= None:
        Dirct2 = soup.find('td',text=re.compile("^2.$")).find_next("td").text
    else:
        Dirct2=None
    with open('Output/IEC_Details_file.csv', 'a', newline='') as outcsv:
        writer = csv.writer(outcsv)
        writer.writerow([IEC,IEC_Allotment_Date,File_Number,Party_Name_Address,Phone_No,e_mail,Exporter_Type,Date_of_Establishment,PAN_ISSUE_DATE,BIN,PAN_ISSUED_BY,Nature_Of_Concern,Bank,Dirct1,Dirct2])
    print(IEC,"\n",IEC_Allotment_Date,"\n",File_Number,"\n",Party_Name_Address,"\n",Phone_No,"\n",e_mail,"\n",Exporter_Type,"\n",Date_of_Establishment,"\n",PAN_ISSUE_DATE,"\n",BIN,"\n",PAN_ISSUED_BY,"\n",Nature_Of_Concern,"\n",Bank,"\n",Dirct1,"\n",Dirct2)
    time.sleep(5)
browser.close()