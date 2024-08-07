from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Border, Side, Alignment
import pandas as pd
import numpy as np
import os

f = open("dosyaYolu.txt", "r")
dosyaYolu = r'%s'%(f.read())
f.close()
yolFatura = r'%s\Su\Fatura Listesi.xlsx'%(dosyaYolu)
yolTakip = r'%s\Taslaklar\suTakip.xlsx'%(dosyaYolu)

wb = load_workbook(yolTakip)
ws = wb['2024']

col = 7

ws.cell(row=2,column=col+4).value = float(20)

wb.save(yolTakip)

df_mebbis = pd.read_excel(yolFatura, header=0)
cols_to_keep_meb = [
                    'KURUM ADI',
                    'ABONE NUMARASI', 
                    'VERGİ NO', 
                    'FATURA NUMARASI', 
                    'FATURA TARİHİ',
                    'İCMAL NO',
                    'TÜKETİM MİKTARI',
                    'FATURA TUTARI'
                ]
df_mebbis = df_mebbis.reindex(cols_to_keep_meb, axis=1)
df_mebbis = df_mebbis.drop(df_mebbis.index[-1]) #tablodaki son satırı siliyor
df_mebbis['ABONE NUMARASI']=df_mebbis['ABONE NUMARASI'].astype(int) #abone numaralarını mys dosyasıyla eşlemek için tam sayı haline getiriyor
df_mebbis['VERGİ NO']=df_mebbis['VERGİ NO'].astype(str) #ilerde hata almamak için kolonun veri tipini değiştiriyor
df_mebbis['İCMAL NO']=df_mebbis['İCMAL NO'].astype(str) #böylece normalde sayısal veriler kayıpsız şekilde metin olarak saklanabiliyor





df_takip = pd.read_excel(yolTakip, header=0)
df_takip = df_takip.fillna(0)

#print(df_mebbis)
