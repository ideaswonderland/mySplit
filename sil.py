import pandas as pd
f = open("dosyaYolu.txt", "r")
dosyaYolu = r'%s'%(f.read())
f.close()
yolFatura = r'%s\Doğalgaz\Fatura Listesi.xlsx'%(dosyaYolu)
yolMYS = r'%s\Doğalgaz\MYS.xlsx'%(dosyaYolu)  
