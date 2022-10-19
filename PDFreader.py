import shutil
from PyPDF2 import PdfFileReader
import re
import requests
import os
import xlrd
import pandas as pd

listDir= os.listdir('files')
dati=[]

def reader() :
  
    key = '/Annots'
    uri = '/URI'
    ank = '/A'
    listaURL=[]
   
    with open("files\\dato.pdf", 'rb') as f:
        
        pdf = PdfFileReader(f)
        
        for numeroDiPagina in range(pdf.numPages):
            pagePDF = pdf.getPage(numeroDiPagina)
            pageElements = pagePDF.getObject()

            if key in pageElements.keys():
                ann = pageElements[key]
                for a in ann:
                    u = a.getObject()
                    if uri in u[ank].keys():
                        if(re.search(r'raw_data',u[ank][uri]) != None):
                            listaURL.append(str(u[ank][uri]))
    f.close()
    for i in listDir:
        for j in listaURL:
            if(re.search(i,j)):
                apriXls('files\\'+i)
                listaURL.remove(j)
                break
    download(listaURL)

def download(listaURL) :
    for url in listaURL:
        src_path = os.path.join(download_file(url))
        dst_path = os.path.join('files')
        try:
            shutil.move(src_path,dst_path)
        except:
            os.remove(src_path)
        
        print(src_path)
        apriXls('files\\'+src_path)


def download_file(url):
    local_filename = url.split('/')[-1]
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(local_filename, 'wb') as f:
            for chunk in r.iter_content(chunk_size=1024): 
                f.write(chunk)
    r.close()
    return local_filename       

def apriXls(file):
    loadXls = xlrd.open_workbook(file)
    sheet = loadXls.sheet_by_index(0)
    for i in range(sheet.nrows):
        
        date = sheet.cell_value(rowx=i,colx=0)
        if isinstance(date,str):
            continue
        prezzo =  sheet.cell_value(rowx=i,colx=7)
        datetime_date = xlrd.xldate_as_datetime(date, loadXls.datemode)

        if (sheet.cell_value(rowx=i,colx=3) == "Euro-super 95"):
            if(type(prezzo) == str):
                prezzo = prezzo.replace(",","")
            dati.append([datetime_date.date(),sheet.cell_value(rowx=i,colx=1),float(prezzo)/1000,"Benzina"])
        
        if(sheet.cell_value(rowx=i,colx=3) == "Automotive gas oil"):
            if(type(prezzo) == str):
                prezzo = prezzo.replace(",","")
            dati.append([datetime_date.date(),sheet.cell_value(rowx=i,colx=1),float(prezzo)/1000,"Diesel"])
    

reader()
df = pd.DataFrame(dati,columns=['Date','Country','Price','Type'])

print(df)
#apriXls('C:\\Users\\Amministratore\\Desktop\\python workspace\\files\\2020_04_27_raw_data_1997.xlsx')