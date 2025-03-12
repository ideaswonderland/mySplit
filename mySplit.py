import pandas as pd
import numpy as np
import os
import warnings
from paracevir import ParaCevir
from datetime import date
from myS import MYS

f = open("dosyaYolu.txt", "r")
dosyaYolu = r'%s'%(f.read())
f.close()
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 

class Fatura():
    def __init__(self,ay,yıl):
        
        #Genel dosya yolları
        self.dönem = r'%s/%s'%(yıl,ay)
        self.yolKurum = r'kurum.xlsx'
        self.yolİmza = r'imza.xlsx'
        #Elektrik dosya yolları
        self.yolFatura_E = r'%s\Elektrik\Fatura Listesi.xlsx'%(dosyaYolu)
        self.yolMYS_E = r'%s\Elektrik\MYS.xlsx'%(dosyaYolu)
        #ADM dosya yolları
        self.yolFatura_A = r'%s\ADM\Fatura Listesi.xlsx'%(dosyaYolu)
        self.yolMYS_A = r'%s\ADM\MYS.xlsx'%(dosyaYolu)
        #İnternet dosya yolları
        self.yolFatura_I = r'%s\İnternet\Fatura Listesi.xlsx'%(dosyaYolu)
        self.yolMYS_I = r'%s\İnternet\MYS.xlsx'%(dosyaYolu)
        #MEM İnternet dosya yolları
        self.yolFatura_MI = r'%s\MEM\İnternet\Fatura Listesi.xlsx'%(dosyaYolu)
        self.yolMYS_MI = r'%s\MEM\İnternet\MYS.xlsx'%(dosyaYolu)
        #Telefon dosya yolları
        self.yolFatura_T = r'%s\Telefon\Fatura Listesi.xlsx'%(dosyaYolu)
        self.yolMYS_T = r'%s\Telefon\MYS.xlsx'%(dosyaYolu)  
        #MEM Telefon dosya yolları
        self.yolFatura_MT = r'%s\MEM\Telefon\Fatura Listesi.xlsx'%(dosyaYolu)
        self.yolMYS_MT = r'%s\MEM\Telefon\MYS.xlsx'%(dosyaYolu)  
        #Doğalgaz dosya yolları      
        self.yolFatura_D = r'%s\Doğalgaz\Fatura Listesi.xlsx'%(dosyaYolu)
        self.yolMYS_D = r'%s\Doğalgaz\MYS.xlsx'%(dosyaYolu)  


    def Elektrik(self):
        yolFatura = self.yolFatura_E
        yolMYS = self.yolMYS_E
        yolKurum = self.yolKurum
        yolİmza = self.yolİmza
        
        df_firma = pd.read_excel(io='firma.xlsx',sheet_name='Elektrik', header=0)

        firma = df_firma.values[0][1]
        faturaTür = 'Elektrik Aboneliği Ödemesi'

        warnings.simplefilter(action='ignore', category=UserWarning)
        
        if os.path.exists(yolFatura):
            if os.path.exists(yolMYS):
                df_imza = pd.read_excel(yolİmza, header=0) #mys'den indirilen dosyayı okuyor
                imzaListe = [
                    df_imza.values[0][0],
                    df_imza.values[1][0],
                    df_imza.values[0][1],
                    df_imza.values[1][1]
                ]

                df_mys = pd.read_excel(yolMYS, header=0) #mys'den indirilen dosyayı okuyor
                cols_to_keep_mys = [
                    'Fatura No',
                    'Harcama Birimi',
                    'Fatura Tarihi',
                    'Ödenecek Tutar',
                    'Müşteri Kimlik Bilgisi'
                ]
                df_mys = df_mys.reindex(cols_to_keep_mys, axis=1) #mys dosyasından verileri çekiyor
                df_mys[['VKN', 'Okul']] = df_mys['Harcama Birimi'].str.split('-', expand=True) #Harcama Birimi kısmını VKN ve Okul olarak ayırıyor
                #df_mys[['Tarih', 'Boş']] = df_mys['Fatura Tarihi'].str.split(' ', expand=True) #Fatura tarihi kısımını ayırıyor
                #df_mys['Tarih'] = df_mys['Tarih'].str.replace('-', '/') #Fatura tarihini mebbis'ten alınan dosyanın formatına dönüştürüyor
                df_mys['Tarih'] = df_mys['Fatura Tarihi'].str.replace('-', '/') #Fatura tarihini mebbis'ten alınan dosyanın formatına dönüştürüyor
                df_mys[['Sayaç No','Müşteri No','Tesisat No','Abone']]=df_mys['Müşteri Kimlik Bilgisi'].str.split('-', expand=True) #Müşteri kimlik bilgisi kısmını bölüyor
                df_mys['Abone']=df_mys['Abone'].astype(int) #abone numaralarını mebbis dosyasıyla eşlemek için tam sayı haline getiriyor
                #gereksiz sütunları temizliyor
                df_mys = df_mys.drop(columns=['Harcama Birimi',
                                            'Fatura Tarihi',
                                            'Müşteri Kimlik Bilgisi',
                                            'Sayaç No',
                                            'Müşteri No',
                                            'Tesisat No',
                                            'Okul'])
                df_mys = df_mys[['Fatura No','Ödenecek Tutar','Tarih','Abone','VKN']]

                df_mebbis = pd.read_excel(yolFatura, header=0) #mebbis dosyasından verileri çekiyor
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
                
                df_mebbis_dummy = pd.read_excel(yolFatura, header=0, dtype=str) #aynı mebbis dosyasını okuyor ancak vergi no ve icmal metin olarak saklanmalı
                cols_to_keep_dummy = ['VERGİ NO','İCMAL NO'] #bu veriler sayısal olarak saklanırsa başında 0 varsa siliyor
                df_mebbis_dummy = df_mebbis_dummy.reindex(cols_to_keep_dummy, axis=1)
                df_mebbis_dummy = df_mebbis_dummy.drop(df_mebbis_dummy.index[-1])    

                df_kurum = pd.read_excel(yolKurum, dtype=str, header=0) #kurumlar dosyasından verileri çekiyor
                cols_to_keep_kurum = [
                            'VERGİ KİMLİK NO', 
                            'KURUM TÜRÜ'
                        ]
                df_kurum = df_kurum.reindex(cols_to_keep_kurum, axis=1)
                df_kurum =df_kurum.drop(df_kurum.index[-1])
                df_kurum_ana = df_kurum[df_kurum['KURUM TÜRÜ']=='Okul Öncesi'] #anaokullarını ayırıyor
                df_kurum_mem = df_kurum[df_kurum['KURUM TÜRÜ']=='MEM'] #MEM'i ayırıyor
                
                #Tek kaynak formu işlemleri
                toplam = "{:,.2f}".format(df_mebbis['FATURA TUTARI'].sum())
                toplam = toplam.replace(',','-')
                toplam = toplam.replace('.',',')
                toplam = toplam.replace('-','.')
                top2 = df_mebbis['FATURA TUTARI'].sum()
                tutar = f'{toplam} TL ({ParaCevir(top2)})'
                unik = len(df_mebbis['KURUM ADI'].unique())
                if unik<4:
                    kurumlar = ", ".join(str(element) for element in df_mebbis['KURUM ADI'].unique())
                else:
                    kurumlar = f'{len(df_mebbis)} Adet Fatura'
                icmaller = ", ".join(str(element) for element in df_mebbis_dummy['İCMAL NO'].unique())
                ihtiyaç = f'TEMEL EĞİTİM OKULLARI {self.dönem} DÖNEM ELEKTRİK ÖDEMESİ {kurumlar} (İcmal No: {icmaller})'
                
                
                tekKaynak = {
                    'firma' : df_firma.values[0][1],
                    'tebligat' : df_firma.values[1][1],
                    'vergi' : df_firma.values[2][1],
                    'telefon' : df_firma.values[3][1],
                    'eposta' : df_firma.values[4][1],
                    'tutar' : tutar,
                    'ihtiyaç' : ihtiyaç,
                    'harcama' : imzaListe[2],
                    'unvan' : imzaListe[3]
                }
                
                tür = 'elektrik'
                ilkTertip = f'40.149.423.292.13.68.01.03.02'
                anaTertip = f'40.149.422.291.13.68.01.03.02'
                memTertip = f'98.900.9006.306.13.67.01.03.02'
                nitelik = f'Temel Eğitim Okulları {self.dönem} Dönem Elektrik Ödemesi' ###
                metin = f'      Müdürlüğümüz  Temel Eğitim Okullarının ({kurumlar}) {self.dönem} dönem {tür} aboneliklerine ait toplam {tutar} borç ödemesi hususunu Onaylarınıza arz ederim.'
                
                harcamaTalimatı = {
                    'tarih' : date.today().strftime("%d.%m.%Y"),
                    'tanım' : faturaTür,
                    'nitelik' : nitelik,
                    'miktar' : tutar,
                    'ödenek1' : '',
                    'ödenek2' : '',
                    'metin' : metin,
                    'ilkTertip': ilkTertip,
                    'anaTertip': anaTertip,
                    'memTertip': memTertip
                }                
                

                MYS(df_mys,df_mebbis,df_mebbis_dummy, df_kurum_ana, df_kurum_mem, firma, faturaTür, imzaListe, dosyaYolu, tekKaynak, harcamaTalimatı).MYS()

    def ADM(self):
        yolFatura = self.yolFatura_A
        yolMYS = self.yolMYS_A
        yolKurum = self.yolKurum
        yolİmza = self.yolİmza
        
        df_firma = pd.read_excel(io='firma.xlsx',sheet_name='ADM', header=0)

        firma = df_firma.values[0][1]
        faturaTür = 'ADM Elektrik Aboneliği Ödemesi'

        warnings.simplefilter(action='ignore', category=UserWarning)
        
        if os.path.exists(yolFatura):
            if os.path.exists(yolMYS):
                df_imza = pd.read_excel(yolİmza, header=0) #mys'den indirilen dosyayı okuyor
                imzaListe = [
                    df_imza.values[0][0],
                    df_imza.values[1][0],
                    df_imza.values[0][1],
                    df_imza.values[1][1]
                ]

                df_mys = pd.read_excel(yolMYS, header=0) #mys'den indirilen dosyayı okuyor
                cols_to_keep_mys = [
                    'Fatura No',
                    'Harcama Birimi',
                    'Fatura Tarihi',
                    'Ödenecek Tutar',
                    'Müşteri Kimlik Bilgisi'
                ]
                df_mys = df_mys.reindex(cols_to_keep_mys, axis=1) #mys dosyasından verileri çekiyor
                df_mys[['VKN', 'Okul']] = df_mys['Harcama Birimi'].str.split('-', expand=True) #Harcama Birimi kısmını VKN ve Okul olarak ayırıyor
                #df_mys[['Tarih', 'Boş']] = df_mys['Fatura Tarihi'].str.split(' ', expand=True) #Fatura tarihi kısımını ayırıyor
                #df_mys['Tarih'] = df_mys['Tarih'].str.replace('-', '/') #Fatura tarihini mebbis'ten alınan dosyanın formatına dönüştürüyor
                df_mys['Tarih'] = df_mys['Fatura Tarihi'].str.replace('-', '/')
                #df_mys[['Sayaç No','Müşteri No','Tesisat No','Abone']]=df_mys['Müşteri Kimlik Bilgisi'].str.split('-', expand=True) #Müşteri kimlik bilgisi kısmını bölüyor
                #df_mys['Abone']=df_mys['Abone'].astype(int) #abone numaralarını mebbis dosyasıyla eşlemek için tam sayı haline getiriyor
                df_mys['Abone'] = 100207225527
                df_mys['Abone']=df_mys['Abone'].astype(float)

                #gereksiz sütunları temizliyor
                """df_mys = df_mys.drop(columns=['Harcama Birimi',
                                            'Fatura Tarihi',
                                            'Müşteri Kimlik Bilgisi',
                                            'Okul'])"""
                df_mys = df_mys[['Fatura No','Ödenecek Tutar','Tarih','Abone','VKN']]

                df_mebbis = pd.read_excel(yolFatura, header=0) #mebbis dosyasından verileri çekiyor
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
                
                df_mebbis_dummy = pd.read_excel(yolFatura, header=0, dtype=str) #aynı mebbis dosyasını okuyor ancak vergi no ve icmal metin olarak saklanmalı
                cols_to_keep_dummy = ['VERGİ NO','İCMAL NO'] #bu veriler sayısal olarak saklanırsa başında 0 varsa siliyor
                df_mebbis_dummy = df_mebbis_dummy.reindex(cols_to_keep_dummy, axis=1)
                df_mebbis_dummy = df_mebbis_dummy.drop(df_mebbis_dummy.index[-1])    

                df_kurum = pd.read_excel(yolKurum, dtype=str, header=0) #kurumlar dosyasından verileri çekiyor
                cols_to_keep_kurum = [
                            'VERGİ KİMLİK NO', 
                            'KURUM TÜRÜ'
                        ]
                df_kurum = df_kurum.reindex(cols_to_keep_kurum, axis=1)
                df_kurum =df_kurum.drop(df_kurum.index[-1])
                df_kurum_ana = df_kurum[df_kurum['KURUM TÜRÜ']=='Okul Öncesi'] #anaokullarını ayırıyor
                df_kurum_mem = df_kurum[df_kurum['KURUM TÜRÜ']=='MEM'] #MEM'i ayırıyor
                
                #Tek kaynak formu işlemleri
                toplam = "{:,.2f}".format(df_mebbis['FATURA TUTARI'].sum())
                toplam = toplam.replace(',','-')
                toplam = toplam.replace('.',',')
                toplam = toplam.replace('-','.')
                top2 = df_mebbis['FATURA TUTARI'].sum()
                tutar = f'{toplam} TL ({ParaCevir(top2)})'
                unik = len(df_mebbis['KURUM ADI'].unique())
                if unik<4:
                    kurumlar = ", ".join(str(element) for element in df_mebbis['KURUM ADI'].unique())
                else:
                    kurumlar = f'{len(df_mebbis)} Adet Fatura'
                icmaller = ", ".join(str(element) for element in df_mebbis_dummy['İCMAL NO'].unique())
                ihtiyaç = f'TEMEL EĞİTİM OKULLARI {self.dönem} DÖNEM ELEKTRİK ÖDEMESİ {kurumlar} (İcmal No: {icmaller})'
                
                
                tekKaynak = {
                    'firma' : df_firma.values[0][1],
                    'tebligat' : df_firma.values[1][1],
                    'vergi' : df_firma.values[2][1],
                    'telefon' : df_firma.values[3][1],
                    'eposta' : df_firma.values[4][1],
                    'tutar' : tutar,
                    'ihtiyaç' : ihtiyaç,
                    'harcama' : imzaListe[2],
                    'unvan' : imzaListe[3]
                }
                
                tür = 'elektrik'
                ilkTertip = f'40.149.423.292.13.68.01.03.02'
                anaTertip = f'40.149.422.291.13.68.01.03.02'
                memTertip = f'98.900.9006.306.13.67.01.03.02'
                nitelik = f'Temel Eğitim Okulları {self.dönem} Dönem Elektrik Ödemesi' ###
                metin = f'      Müdürlüğümüz  Temel Eğitim Okullarının ({kurumlar}) {self.dönem} dönem {tür} aboneliklerine ait toplam {tutar} borç ödemesi hususunu Onaylarınıza arz ederim.'
                
                harcamaTalimatı = {
                    'tarih' : date.today().strftime("%d.%m.%Y"),
                    'tanım' : faturaTür,
                    'nitelik' : nitelik,
                    'miktar' : tutar,
                    'ödenek1' : '',
                    'ödenek2' : '',
                    'metin' : metin,
                    'ilkTertip': ilkTertip,
                    'anaTertip': anaTertip,
                    'memTertip': memTertip
                }                
                

                MYS(df_mys,df_mebbis,df_mebbis_dummy, df_kurum_ana, df_kurum_mem, firma, faturaTür, imzaListe, dosyaYolu, tekKaynak, harcamaTalimatı).MYS()            
                
    def İnternet(self):
        yolFatura = self.yolFatura_I
        yolMYS = self.yolMYS_I
        yolKurum = self.yolKurum
        yolİmza = self.yolİmza

        df_firma = pd.read_excel(io='firma.xlsx',sheet_name='İnternet', header=0)

        firma = df_firma.values[0][1]
        faturaTür = 'İnternet Aboneliği Ödemesi'

        warnings.simplefilter(action='ignore', category=UserWarning)
        
        if os.path.exists(yolFatura):
            if os.path.exists(yolMYS):
                df_imza = pd.read_excel(yolİmza, header=0) #mys'den indirilen dosyayı okuyor
                imzaListe = [
                    df_imza.values[0][0],
                    df_imza.values[1][0],
                    df_imza.values[0][1],
                    df_imza.values[1][1]
                ]

                df_mys = pd.read_excel(yolMYS, header=0) #mys'den indirilen dosyayı okuyor
                cols_to_keep_mys = [
                    'Fatura No',
                    'Harcama Birimi',
                    'Fatura Tarihi',
                    'Ödenecek Tutar',
                    'Müşteri Kimlik Bilgisi'
                ]
                df_mys = df_mys.reindex(cols_to_keep_mys, axis=1) #mys dosyasından verileri çekiyor
                df_mys[['VKN', 'Okul']] = df_mys['Harcama Birimi'].str.split('-', expand=True) #Harcama Birimi kısmını VKN ve Okul olarak ayırıyor
                #df_mys[['Tarih', 'Boş']] = df_mys['Fatura Tarihi'].str.split(' ', expand=True) #Fatura tarihi kısımını ayırıyor
                #df_mys['Tarih'] = df_mys['Tarih'].str.replace('-', '/') #Fatura tarihini mebbis'ten alınan dosyanın formatına dönüştürüyor
                df_mys['Tarih'] = df_mys['Fatura Tarihi'].str.replace('-', '/')
                df_mys[['Hizmet No','Abone']]=df_mys['Müşteri Kimlik Bilgisi'].str.split('-', expand=True) #Müşteri kimlik bilgisi kısmını bölüyor
                df_mys['Abone']=df_mys['Abone'].astype(float) #abone numaralarını mebbis dosyasıyla eşlemek için tam sayı haline getiriyor
                #gereksiz sütunları temizliyor
                df_mys = df_mys.drop(columns=['Harcama Birimi',
                                            'Fatura Tarihi',
                                            'Müşteri Kimlik Bilgisi',
                                            'Hizmet No',
                                            'Okul'])
                df_mys = df_mys[['Fatura No','Ödenecek Tutar','Tarih','Abone','VKN']]

                df_mebbis = pd.read_excel(yolFatura, header=0) #mebbis dosyasından verileri çekiyor
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
                
                df_mebbis_dummy = pd.read_excel(yolFatura, header=0, dtype=str) #aynı mebbis dosyasını okuyor ancak vergi no ve icmal metin olarak saklanmalı
                cols_to_keep_dummy = ['VERGİ NO','İCMAL NO'] #bu veriler sayısal olarak saklanırsa başında 0 varsa siliyor
                df_mebbis_dummy = df_mebbis_dummy.reindex(cols_to_keep_dummy, axis=1)
                df_mebbis_dummy = df_mebbis_dummy.drop(df_mebbis_dummy.index[-1])    

                df_kurum = pd.read_excel(yolKurum, dtype=str, header=0) #kurumlar dosyasından verileri çekiyor
                cols_to_keep_kurum = [
                            'VERGİ KİMLİK NO', 
                            'KURUM TÜRÜ'
                        ]
                df_kurum = df_kurum.reindex(cols_to_keep_kurum, axis=1)
                df_kurum =df_kurum.drop(df_kurum.index[-1])
                df_kurum_ana = df_kurum[df_kurum['KURUM TÜRÜ']=='Okul Öncesi'] #anaokullarını ayırıyor
                df_kurum_mem = df_kurum[df_kurum['KURUM TÜRÜ']=='MEM'] #MEM'i ayırıyor
                
                #Tek kaynak formu işlemleri
                toplam = "{:,.2f}".format(df_mebbis['FATURA TUTARI'].sum())
                toplam = toplam.replace(',','-')
                toplam = toplam.replace('.',',')
                toplam = toplam.replace('-','.')
                top2 = df_mebbis['FATURA TUTARI'].sum()
                tutar = f'{toplam} TL ({ParaCevir(top2)})'
                unik = len(df_mebbis['KURUM ADI'].unique())
                if unik<4:
                    kurumlar = ", ".join(str(element) for element in df_mebbis['KURUM ADI'].unique())
                else:
                    kurumlar = f'{len(df_mebbis)} Adet Fatura'
                icmaller = ", ".join(str(element) for element in df_mebbis_dummy['İCMAL NO'].unique())
                ihtiyaç = f'TEMEL EĞİTİM OKULLARI {self.dönem} DÖNEM İNTERNET ÖDEMESİ {kurumlar} (İcmal No: {icmaller})'
                tekKaynak = {
                    'firma' : df_firma.values[0][1],
                    'tebligat' : df_firma.values[1][1],
                    'vergi' : df_firma.values[2][1],
                    'telefon' : df_firma.values[3][1],
                    'eposta' : df_firma.values[4][1],
                    'tutar' : tutar,
                    'ihtiyaç' : ihtiyaç,
                    'harcama' : imzaListe[2],
                    'unvan' : imzaListe[3]
                }
                
                tür = 'internet'
                ilkTertip = f'40.149.423.292.13.68.01.03.05'
                anaTertip = f'40.149.422.291.13.68.01.03.05'
                memTertip = f'98.900.9006.306.13.67.01.03.05'
                nitelik = f'Temel Eğitim Okulları {self.dönem} Dönem İnternet Ödemesi' ###
                metin = f'      Müdürlüğümüz  Temel Eğitim Okullarının ({kurumlar}) {self.dönem} dönem {tür} aboneliklerine ait toplam {tutar} borç ödemesi hususunu Onaylarınıza arz ederim.'
                
                harcamaTalimatı = {
                    'tarih' : date.today().strftime("%d.%m.%Y"),
                    'tanım' : faturaTür,
                    'nitelik' : nitelik,
                    'miktar' : tutar,
                    'ödenek1' : '',
                    'ödenek2' : '',
                    'metin' : metin,
                    'ilkTertip': ilkTertip,
                    'anaTertip': anaTertip,
                    'memTertip': memTertip
                }       

                MYS(df_mys,df_mebbis,df_mebbis_dummy, df_kurum_ana, df_kurum_mem, firma, faturaTür, imzaListe, dosyaYolu, tekKaynak, harcamaTalimatı).MYS()

    def MEMİnternet(self):
        yolFatura = self.yolFatura_MI
        yolMYS = self.yolMYS_MI
        yolKurum = self.yolKurum
        yolİmza = self.yolİmza

        df_firma = pd.read_excel(io='firma.xlsx',sheet_name='İnternet', header=0)

        firma = df_firma.values[0][1]
        faturaTür = 'İnternet Aboneliği Ödemesi'

        warnings.simplefilter(action='ignore', category=UserWarning)
        
        if os.path.exists(yolFatura):
            if os.path.exists(yolMYS):
                df_imza = pd.read_excel(yolİmza, header=0) #mys'den indirilen dosyayı okuyor
                imzaListe = [
                    df_imza.values[0][0],
                    df_imza.values[1][0],
                    df_imza.values[0][1],
                    df_imza.values[1][1]
                ]

                df_mys = pd.read_excel(yolMYS, header=0) #mys'den indirilen dosyayı okuyor
                cols_to_keep_mys = [
                    'Fatura No',
                    'Harcama Birimi',
                    'Fatura Tarihi',
                    'Ödenecek Tutar',
                    'Müşteri Kimlik Bilgisi'
                ]
                df_mys = df_mys.reindex(cols_to_keep_mys, axis=1) #mys dosyasından verileri çekiyor
                df_mys[['VKN', 'Okul']] = df_mys['Harcama Birimi'].str.split('-', expand=True) #Harcama Birimi kısmını VKN ve Okul olarak ayırıyor
                #df_mys[['Tarih', 'Boş']] = df_mys['Fatura Tarihi'].str.split(' ', expand=True) #Fatura tarihi kısımını ayırıyor
                #df_mys['Tarih'] = df_mys['Tarih'].str.replace('-', '/') #Fatura tarihini mebbis'ten alınan dosyanın formatına dönüştürüyor
                df_mys['Tarih'] = df_mys['Fatura Tarihi'].str.replace('-', '/')
                df_mys[['Hizmet No','Abone']]=df_mys['Müşteri Kimlik Bilgisi'].str.split('-', expand=True) #Müşteri kimlik bilgisi kısmını bölüyor
                df_mys['Abone']=df_mys['Abone'].astype(float) #abone numaralarını mebbis dosyasıyla eşlemek için tam sayı haline getiriyor
                #gereksiz sütunları temizliyor
                df_mys = df_mys.drop(columns=['Harcama Birimi',
                                            'Fatura Tarihi',
                                            'Müşteri Kimlik Bilgisi',
                                            'Hizmet No',
                                            'Okul'])
                df_mys = df_mys[['Fatura No','Ödenecek Tutar','Tarih','Abone','VKN']]

                df_mebbis = pd.read_excel(yolFatura, header=0) #mebbis dosyasından verileri çekiyor

                df_mebbis['KURUM ADI'] = 'İlçe Milli Eğitim Müdürlüğü'
                df_mebbis['VERGİ NO'] = 7820458686
                df_mebbis = df_mebbis.rename(columns={'İCMAL NUMARASI':'İCMAL NO'})

                cols_to_keep_meb = [
                    'KURUM ADI', #eklenecek
                    'ABONE NUMARASI', 
                    'VERGİ NO', #eklenecek
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
                
                df_mebbis_dummy = pd.read_excel(yolFatura, header=0, dtype=str) #aynı mebbis dosyasını okuyor ancak vergi no ve icmal metin olarak saklanmalı
                df_mebbis_dummy['VERGİ NO'] = '7820458686'
                df_mebbis_dummy = df_mebbis_dummy.rename(columns={'İCMAL NUMARASI':'İCMAL NO'})
                cols_to_keep_dummy = ['VERGİ NO','İCMAL NO'] #bu veriler sayısal olarak saklanırsa başında 0 varsa siliyor
                df_mebbis_dummy = df_mebbis_dummy.reindex(cols_to_keep_dummy, axis=1)
                df_mebbis_dummy = df_mebbis_dummy.drop(df_mebbis_dummy.index[-1])    

                df_kurum = pd.read_excel(yolKurum, dtype=str, header=0) #kurumlar dosyasından verileri çekiyor
                cols_to_keep_kurum = [
                            'VERGİ KİMLİK NO', 
                            'KURUM TÜRÜ'
                        ]
                df_kurum = df_kurum.reindex(cols_to_keep_kurum, axis=1)
                df_kurum =df_kurum.drop(df_kurum.index[-1])
                df_kurum_ana = df_kurum[df_kurum['KURUM TÜRÜ']=='Okul Öncesi'] #anaokullarını ayırıyor
                df_kurum_mem = df_kurum[df_kurum['KURUM TÜRÜ']=='MEM'] #MEM'i ayırıyor
                
                #Tek kaynak formu işlemleri
                toplam = "{:,.2f}".format(df_mebbis['FATURA TUTARI'].sum())
                toplam = toplam.replace(',','-')
                toplam = toplam.replace('.',',')
                toplam = toplam.replace('-','.')
                top2 = df_mebbis['FATURA TUTARI'].sum()
                tutar = f'{toplam} TL ({ParaCevir(top2)})'
                unik = len(df_mebbis['KURUM ADI'].unique())
                if unik<4:
                    kurumlar = ", ".join(str(element) for element in df_mebbis['KURUM ADI'].unique())
                else:
                    kurumlar = f'{len(df_mebbis)} Adet Fatura'
                icmaller = ", ".join(str(element) for element in df_mebbis_dummy['İCMAL NO'].unique())
                ihtiyaç = f'İLÇE MİLLİ EĞİTİM MÜDÜRLÜĞÜ {self.dönem} DÖNEM İNTERNET ÖDEMESİ (İcmal No: {icmaller})'
                tekKaynak = {
                    'firma' : df_firma.values[0][1],
                    'tebligat' : df_firma.values[1][1],
                    'vergi' : df_firma.values[2][1],
                    'telefon' : df_firma.values[3][1],
                    'eposta' : df_firma.values[4][1],
                    'tutar' : tutar,
                    'ihtiyaç' : ihtiyaç,
                    'harcama' : imzaListe[2],
                    'unvan' : imzaListe[3]
                }
                
                tür = 'internet'
                ilkTertip = f'40.149.423.292.13.68.01.03.05'
                anaTertip = f'40.149.422.291.13.68.01.03.05'
                memTertip = f'98.900.9006.306.13.67.01.03.05'
                nitelik = f'İlçe Milli Eğitim Müdürlüğü {self.dönem} Dönem İnternet Ödemesi' ###
                metin = f'      Müdürlüğümüze ait {self.dönem} dönem {tür} abonelikleri toplam {tutar} borç ödemesi hususunu Onaylarınıza arz ederim.'
                
                harcamaTalimatı = {
                    'tarih' : date.today().strftime("%d.%m.%Y"),
                    'tanım' : faturaTür,
                    'nitelik' : nitelik,
                    'miktar' : tutar,
                    'ödenek1' : '',
                    'ödenek2' : '',
                    'metin' : metin,
                    'ilkTertip': ilkTertip,
                    'anaTertip': anaTertip,
                    'memTertip': memTertip
                }       

                MYS(df_mys,df_mebbis,df_mebbis_dummy, df_kurum_ana, df_kurum_mem, firma, faturaTür, imzaListe, dosyaYolu, tekKaynak, harcamaTalimatı).MYS()

    def Telefon(self):
        yolFatura = self.yolFatura_T
        yolMYS = self.yolMYS_T
        yolKurum = self.yolKurum
        yolİmza = self.yolİmza

        df_firma = pd.read_excel(io='firma.xlsx',sheet_name='Telefon', header=0)

        firma = df_firma.values[0][1]
        faturaTür = 'Telefon Aboneliği Ödemesi'

        warnings.simplefilter(action='ignore', category=UserWarning)
        
        if os.path.exists(yolFatura):
            if os.path.exists(yolMYS):
                df_imza = pd.read_excel(yolİmza, header=0) #mys'den indirilen dosyayı okuyor
                imzaListe = [
                    df_imza.values[0][0],
                    df_imza.values[1][0],
                    df_imza.values[0][1],
                    df_imza.values[1][1]
                ]

                df_mys = pd.read_excel(yolMYS, header=0) #mys'den indirilen dosyayı okuyor
                cols_to_keep_mys = [
                    'Fatura No',
                    'Harcama Birimi',
                    'Fatura Tarihi',
                    'Ödenecek Tutar',
                    'Müşteri Kimlik Bilgisi'
                ]
                df_mys = df_mys.reindex(cols_to_keep_mys, axis=1) #mys dosyasından verileri çekiyor
                df_mys[['VKN', 'Okul']] = df_mys['Harcama Birimi'].str.split('-', expand=True) #Harcama Birimi kısmını VKN ve Okul olarak ayırıyor
                #df_mys[['Tarih', 'Boş']] = df_mys['Fatura Tarihi'].str.split(' ', expand=True) #Fatura tarihi kısımını ayırıyor
                #df_mys['Tarih'] = df_mys['Tarih'].str.replace('-', '/') #Fatura tarihini mebbis'ten alınan dosyanın formatına dönüştürüyor
                df_mys['Tarih'] = df_mys['Fatura Tarihi'].str.replace('-', '/')
                df_mys[['Hizmet No','Abone']]=df_mys['Müşteri Kimlik Bilgisi'].str.split('-', expand=True) #Müşteri kimlik bilgisi kısmını bölüyor
                df_mys['Abone'] = df_mys['Abone'].str.replace(" ","")
                df_mys['Abone']=df_mys['Abone'].astype(float) #abone numaralarını mebbis dosyasıyla eşlemek için tam sayı haline getiriyor
                #gereksiz sütunları temizliyor
                df_mys = df_mys.drop(columns=['Harcama Birimi',
                                            'Fatura Tarihi',
                                            'Müşteri Kimlik Bilgisi',
                                            'Hizmet No',
                                            'Okul'])
                df_mys = df_mys[['Fatura No','Ödenecek Tutar','Tarih','Abone','VKN']]

                df_mebbis = pd.read_excel(yolFatura, header=0) #mebbis dosyasından verileri çekiyor
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
                df_mebbis['ABONE NUMARASI']=df_mebbis['ABONE NUMARASI'].astype(float) #abone numaralarını mys dosyasıyla eşlemek için tam sayı haline getiriyor
                df_mebbis['VERGİ NO']=df_mebbis['VERGİ NO'].astype(str) #ilerde hata almamak için kolonun veri tipini değiştiriyor
                df_mebbis['İCMAL NO']=df_mebbis['İCMAL NO'].astype(str) #böylece normalde sayısal veriler kayıpsız şekilde metin olarak saklanabiliyor
                
                df_mebbis_dummy = pd.read_excel(yolFatura, header=0, dtype=str) #aynı mebbis dosyasını okuyor ancak vergi no ve icmal metin olarak saklanmalı
                cols_to_keep_dummy = ['VERGİ NO','İCMAL NO'] #bu veriler sayısal olarak saklanırsa başında 0 varsa siliyor
                df_mebbis_dummy = df_mebbis_dummy.reindex(cols_to_keep_dummy, axis=1)
                df_mebbis_dummy = df_mebbis_dummy.drop(df_mebbis_dummy.index[-1])    

                df_kurum = pd.read_excel(yolKurum, dtype=str, header=0) #kurumlar dosyasından verileri çekiyor
                cols_to_keep_kurum = [
                            'VERGİ KİMLİK NO', 
                            'KURUM TÜRÜ'
                        ]
                df_kurum = df_kurum.reindex(cols_to_keep_kurum, axis=1)
                df_kurum =df_kurum.drop(df_kurum.index[-1])
                df_kurum_ana = df_kurum[df_kurum['KURUM TÜRÜ']=='Okul Öncesi'] #anaokullarını ayırıyor
                df_kurum_mem = df_kurum[df_kurum['KURUM TÜRÜ']=='MEM'] #MEM'i ayırıyor
                
                #Tek kaynak formu işlemleri
                toplam = "{:,.2f}".format(df_mebbis['FATURA TUTARI'].sum())
                toplam = toplam.replace(',','-')
                toplam = toplam.replace('.',',')
                toplam = toplam.replace('-','.')
                top2 = df_mebbis['FATURA TUTARI'].sum()
                tutar = f'{toplam} TL ({ParaCevir(top2)})'
                unik = len(df_mebbis['KURUM ADI'].unique())
                if unik<4:
                    kurumlar = ", ".join(str(element) for element in df_mebbis['KURUM ADI'].unique())
                else:
                    kurumlar = f'{len(df_mebbis)} Adet Fatura'
                icmaller = ", ".join(str(element) for element in df_mebbis_dummy['İCMAL NO'].unique())
                ihtiyaç = f'TEMEL EĞİTİM OKULLARI {self.dönem} DÖNEM TELEFON ÖDEMESİ {kurumlar} (İcmal No: {icmaller})'
                tekKaynak = {
                    'firma' : df_firma.values[0][1],
                    'tebligat' : df_firma.values[1][1],
                    'vergi' : df_firma.values[2][1],
                    'telefon' : df_firma.values[3][1],
                    'eposta' : df_firma.values[4][1],
                    'tutar' : tutar,
                    'ihtiyaç' : ihtiyaç,
                    'harcama' : imzaListe[2],
                    'unvan' : imzaListe[3]
                }
                
                tür = 'telefon'
                ilkTertip = f'40.149.423.292.13.68.01.03.05'
                anaTertip = f'40.149.422.291.13.68.01.03.05'
                memTertip = f'98.900.9006.306.13.67.01.03.05'
                nitelik = f'Temel Eğitim Okulları {self.dönem} Dönem Telefon Ödemesi' ###
                metin = f'      Müdürlüğümüz  Temel Eğitim Okullarının ({kurumlar}) {self.dönem} dönem {tür} aboneliklerine ait toplam {tutar} borç ödemesi hususunu Onaylarınıza arz ederim.'
                
                harcamaTalimatı = {
                    'tarih' : date.today().strftime("%d.%m.%Y"),
                    'tanım' : faturaTür,
                    'nitelik' : nitelik,
                    'miktar' : tutar,
                    'ödenek1' : '',
                    'ödenek2' : '',
                    'metin' : metin,
                    'ilkTertip': ilkTertip,
                    'anaTertip': anaTertip,
                    'memTertip': memTertip
                }       


                MYS(df_mys,df_mebbis,df_mebbis_dummy, df_kurum_ana, df_kurum_mem, firma, faturaTür, imzaListe, dosyaYolu, tekKaynak, harcamaTalimatı).MYS()

    def MEMTelefon(self):
        yolFatura = self.yolFatura_MT
        yolMYS = self.yolMYS_MT
        yolKurum = self.yolKurum
        yolİmza = self.yolİmza

        df_firma = pd.read_excel(io='firma.xlsx',sheet_name='Telefon', header=0)

        firma = df_firma.values[0][1]
        faturaTür = 'Telefon Aboneliği Ödemesi'

        warnings.simplefilter(action='ignore', category=UserWarning)
        
        if os.path.exists(yolFatura):
            if os.path.exists(yolMYS):
                df_imza = pd.read_excel(yolİmza, header=0) #mys'den indirilen dosyayı okuyor
                imzaListe = [
                    df_imza.values[0][0],
                    df_imza.values[1][0],
                    df_imza.values[0][1],
                    df_imza.values[1][1]
                ]

                df_mys = pd.read_excel(yolMYS, header=0) #mys'den indirilen dosyayı okuyor
                cols_to_keep_mys = [
                    'Fatura No',
                    'Harcama Birimi',
                    'Fatura Tarihi',
                    'Ödenecek Tutar',
                    'Müşteri Kimlik Bilgisi'
                ]
                df_mys = df_mys.reindex(cols_to_keep_mys, axis=1) #mys dosyasından verileri çekiyor
                df_mys[['VKN', 'Okul']] = df_mys['Harcama Birimi'].str.split('-', expand=True) #Harcama Birimi kısmını VKN ve Okul olarak ayırıyor
                #df_mys[['Tarih', 'Boş']] = df_mys['Fatura Tarihi'].str.split(' ', expand=True) #Fatura tarihi kısımını ayırıyor
                #df_mys['Tarih'] = df_mys['Tarih'].str.replace('-', '/') #Fatura tarihini mebbis'ten alınan dosyanın formatına dönüştürüyor
                df_mys['Tarih'] = df_mys['Fatura Tarihi'].str.replace('-', '/')
                df_mys[['Hizmet No','Abone']]=df_mys['Müşteri Kimlik Bilgisi'].str.split('-', expand=True) #Müşteri kimlik bilgisi kısmını bölüyor
                df_mys['Abone'] = df_mys['Abone'].str.replace(" ","")
                df_mys['Abone']=df_mys['Abone'].astype(float) #abone numaralarını mebbis dosyasıyla eşlemek için tam sayı haline getiriyor
                #gereksiz sütunları temizliyor
                df_mys = df_mys.drop(columns=['Harcama Birimi',
                                            'Fatura Tarihi',
                                            'Müşteri Kimlik Bilgisi',
                                            'Hizmet No',
                                            'Okul'])
                df_mys = df_mys[['Fatura No','Ödenecek Tutar','Tarih','Abone','VKN']]

                df_mebbis = pd.read_excel(yolFatura, header=0) #mebbis dosyasından verileri çekiyor

                df_mebbis['KURUM ADI'] = 'İlçe Milli Eğitim Müdürlüğü'
                df_mebbis['VERGİ NO'] = 7820458686
                df_mebbis = df_mebbis.rename(columns={'İCMAL NUMARASI':'İCMAL NO'})

                cols_to_keep_meb = [
                    'KURUM ADI', #eklenecek
                    'ABONE NUMARASI', 
                    'VERGİ NO', #eklenecek
                    'FATURA NUMARASI', 
                    'FATURA TARİHİ',
                    'İCMAL NO',
                    'TÜKETİM MİKTARI',
                    'FATURA TUTARI'
                ]
                df_mebbis = df_mebbis.reindex(cols_to_keep_meb, axis=1)
                df_mebbis = df_mebbis.drop(df_mebbis.index[-1]) #tablodaki son satırı siliyor
                #df_mebbis['ABONE NUMARASI']=df_mebbis['ABONE NUMARASI'].astype(int) #abone numaralarını mys dosyasıyla eşlemek için tam sayı haline getiriyor
                df_mebbis['VERGİ NO']=df_mebbis['VERGİ NO'].astype(str) #ilerde hata almamak için kolonun veri tipini değiştiriyor
                df_mebbis['İCMAL NO']=df_mebbis['İCMAL NO'].astype(str) #böylece normalde sayısal veriler kayıpsız şekilde metin olarak saklanabiliyor
                
                df_mebbis_dummy = pd.read_excel(yolFatura, header=0, dtype=str) #aynı mebbis dosyasını okuyor ancak vergi no ve icmal metin olarak saklanmalı
                df_mebbis_dummy['VERGİ NO'] = '7820458686'
                df_mebbis_dummy = df_mebbis_dummy.rename(columns={'İCMAL NUMARASI':'İCMAL NO'})
                cols_to_keep_dummy = ['VERGİ NO','İCMAL NO'] #bu veriler sayısal olarak saklanırsa başında 0 varsa siliyor
                df_mebbis_dummy = df_mebbis_dummy.reindex(cols_to_keep_dummy, axis=1)
                df_mebbis_dummy = df_mebbis_dummy.drop(df_mebbis_dummy.index[-1])    

                df_kurum = pd.read_excel(yolKurum, dtype=str, header=0) #kurumlar dosyasından verileri çekiyor
                cols_to_keep_kurum = [
                            'VERGİ KİMLİK NO', 
                            'KURUM TÜRÜ'
                        ]
                df_kurum = df_kurum.reindex(cols_to_keep_kurum, axis=1)
                df_kurum =df_kurum.drop(df_kurum.index[-1])
                df_kurum_ana = df_kurum[df_kurum['KURUM TÜRÜ']=='Okul Öncesi'] #anaokullarını ayırıyor
                df_kurum_mem = df_kurum[df_kurum['KURUM TÜRÜ']=='MEM'] #MEM'i ayırıyor
                
                #Tek kaynak formu işlemleri
                toplam = "{:,.2f}".format(df_mebbis['FATURA TUTARI'].sum())
                toplam = toplam.replace(',','-')
                toplam = toplam.replace('.',',')
                toplam = toplam.replace('-','.')
                top2 = df_mebbis['FATURA TUTARI'].sum()
                tutar = f'{toplam} TL ({ParaCevir(top2)})'
                unik = len(df_mebbis['KURUM ADI'].unique())
                if unik<4:
                    kurumlar = ", ".join(str(element) for element in df_mebbis['KURUM ADI'].unique())
                else:
                    kurumlar = f'{len(df_mebbis)} Adet Fatura'
                icmaller = ", ".join(str(element) for element in df_mebbis_dummy['İCMAL NO'].unique())
                ihtiyaç = f'İLÇE MİLLİ EĞİTİM MÜDÜRLÜĞÜ {self.dönem} DÖNEM TELEFON ÖDEMESİ (İcmal No: {icmaller})'
                tekKaynak = {
                    'firma' : df_firma.values[0][1],
                    'tebligat' : df_firma.values[1][1],
                    'vergi' : df_firma.values[2][1],
                    'telefon' : df_firma.values[3][1],
                    'eposta' : df_firma.values[4][1],
                    'tutar' : tutar,
                    'ihtiyaç' : ihtiyaç,
                    'harcama' : imzaListe[2],
                    'unvan' : imzaListe[3]
                }
                
                tür = 'telefon'
                ilkTertip = f'40.149.423.292.13.68.01.03.05'
                anaTertip = f'40.149.422.291.13.68.01.03.05'
                memTertip = f'98.900.9006.306.13.67.01.03.05'
                nitelik = f'İlçe Milli Eğitim Müdürlüğü {self.dönem} Dönem Telefon Ödemesi' ###
                metin = f'      Müdürlüğümüze ait {self.dönem} dönem {tür} abonelikleri toplam {tutar} borç ödemesi hususunu Onaylarınıza arz ederim.'
                
                harcamaTalimatı = {
                    'tarih' : date.today().strftime("%d.%m.%Y"),
                    'tanım' : faturaTür,
                    'nitelik' : nitelik,
                    'miktar' : tutar,
                    'ödenek1' : '',
                    'ödenek2' : '',
                    'metin' : metin,
                    'ilkTertip': ilkTertip,
                    'anaTertip': anaTertip,
                    'memTertip': memTertip
                }       

                MYS(df_mys,df_mebbis,df_mebbis_dummy, df_kurum_ana, df_kurum_mem, firma, faturaTür, imzaListe, dosyaYolu, tekKaynak, harcamaTalimatı).MYS()

    def Doğalgaz(self):
        yolFatura = self.yolFatura_D
        yolMYS = self.yolMYS_D
        yolKurum = self.yolKurum
        yolİmza = self.yolİmza

        df_firma = pd.read_excel(io='firma.xlsx',sheet_name='Gaz', header=0)

        firma = df_firma.values[0][1]
        faturaTür = 'Doğalgaz Aboneliği Ödemesi'
        
        warnings.simplefilter(action='ignore', category=UserWarning)
        
        if os.path.exists(yolFatura):
            if os.path.exists(yolMYS):
                df_imza = pd.read_excel(yolİmza, header=0) #mys'den indirilen dosyayı okuyor
                imzaListe = [
                    df_imza.values[0][0],
                    df_imza.values[1][0],
                    df_imza.values[0][1],
                    df_imza.values[1][1]
                ]

                df_mys = pd.read_excel(yolMYS, header=0) #mys'den indirilen dosyayı okuyor
                cols_to_keep_mys = [
                    'Fatura No',
                    'Harcama Birimi',
                    'Fatura Tarihi',
                    'Ödenecek Tutar',
                    'Müşteri Kimlik Bilgisi'
                ]
                df_mys = df_mys.reindex(cols_to_keep_mys, axis=1) #mys dosyasından verileri çekiyor
                df_mys[['VKN', 'Okul']] = df_mys['Harcama Birimi'].str.split('-', expand=True) #Harcama Birimi kısmını VKN ve Okul olarak ayırıyor
                #df_mys[['Tarih', 'Boş']] = df_mys['Fatura Tarihi'].str.split(' ', expand=True) #Fatura tarihi kısımını ayırıyor
                df_mys['Tarih'] = df_mys['Fatura Tarihi'].str.replace('-', '/') #Fatura tarihini mebbis'ten alınan dosyanın formatına dönüştürüyor
                df_mys['Abone']="" #Boş bir abone sütunu açıyor

                #gereksiz sütunları temizliyor
                df_mys = df_mys.drop(columns=['Harcama Birimi',
                                            'Fatura Tarihi',
                                            'Müşteri Kimlik Bilgisi',
                                            'Okul'])
                df_mys = df_mys[['Fatura No','Ödenecek Tutar','Tarih','Abone','VKN']]

                df_mebbis = pd.read_excel(yolFatura, header=0) #mebbis dosyasından verileri çekiyor
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
    
                df_mebbis['ABONE NUMARASI']=df_mebbis['ABONE NUMARASI'].astype(float) 

                meb_abo_col = df_mebbis.columns.get_loc('ABONE NUMARASI')
                mys_abo_col = df_mys.columns.get_loc('Abone')
                meb_fat_col = df_mebbis.columns.get_loc('FATURA NUMARASI')
                
                for i in range(len(df_mebbis.index)):
                    indices_fatura = np.where(df_mys['Fatura No'] == df_mebbis.iat[i,meb_fat_col])
                    row_indices_fatura = indices_fatura[0]
                    if row_indices_fatura.size != 0:
                        fatNo = df_mys['Fatura No'].values == df_mebbis.iat[i,meb_fat_col]
                        if max(fatNo)==True: 
                            row_indices_df_mys = np.where(fatNo == True)
                            df_mys.iat[row_indices_df_mys[0][0],mys_abo_col] = df_mebbis.iat[i,meb_abo_col]

                
                df_mebbis['VERGİ NO']=df_mebbis['VERGİ NO'].astype(str) #ilerde hata almamak için kolonun veri tipini değiştiriyor
                df_mebbis['İCMAL NO']=df_mebbis['İCMAL NO'].astype(str) #böylece normalde sayısal veriler kayıpsız şekilde metin olarak saklanabiliyor
                
                df_mebbis_dummy = pd.read_excel(yolFatura, header=0, dtype=str) #aynı mebbis dosyasını okuyor ancak vergi no ve icmal metin olarak saklanmalı
                cols_to_keep_dummy = ['VERGİ NO','İCMAL NO'] #bu veriler sayısal olarak saklanırsa başında 0 varsa siliyor
                df_mebbis_dummy = df_mebbis_dummy.reindex(cols_to_keep_dummy, axis=1)
                df_mebbis_dummy = df_mebbis_dummy.drop(df_mebbis_dummy.index[-1])    

                df_kurum = pd.read_excel(yolKurum, dtype=str, header=0) #kurumlar dosyasından verileri çekiyor
                cols_to_keep_kurum = [
                            'VERGİ KİMLİK NO', 
                            'KURUM TÜRÜ'
                        ]
                df_kurum = df_kurum.reindex(cols_to_keep_kurum, axis=1)
                df_kurum =df_kurum.drop(df_kurum.index[-1])
                df_kurum_ana = df_kurum[df_kurum['KURUM TÜRÜ']=='Okul Öncesi'] #anaokullarını ayırıyor
                df_kurum_mem = df_kurum[df_kurum['KURUM TÜRÜ']=='MEM'] #MEM'i ayırıyor
                
                #Tek kaynak formu işlemleri
                toplam = "{:,.2f}".format(df_mebbis['FATURA TUTARI'].sum())
                toplam = toplam.replace(',','-')
                toplam = toplam.replace('.',',')
                toplam = toplam.replace('-','.')
                top2 = df_mebbis['FATURA TUTARI'].sum()
                tutar = f'{toplam} TL ({ParaCevir(top2)})'
                unik = len(df_mebbis['KURUM ADI'].unique())
                if unik<4:
                    kurumlar = ", ".join(str(element) for element in df_mebbis['KURUM ADI'].unique())
                else:
                    kurumlar = f'{len(df_mebbis)} Adet Fatura'
                icmaller = ", ".join(str(element) for element in df_mebbis_dummy['İCMAL NO'].unique())
                ihtiyaç = f'TEMEL EĞİTİM OKULLARI {self.dönem} DÖNEM DOĞALGAZ ÖDEMESİ {kurumlar} (İcmal No: {icmaller})'
                tekKaynak = {
                    'firma' : df_firma.values[0][1],
                    'tebligat' : df_firma.values[1][1],
                    'vergi' : df_firma.values[2][1],
                    'telefon' : df_firma.values[3][1],
                    'eposta' : df_firma.values[4][1],
                    'tutar' : tutar,
                    'ihtiyaç' : ihtiyaç,
                    'harcama' : imzaListe[2],
                    'unvan' : imzaListe[3]
                }
                
                tür = 'doğalgaz'
                ilkTertip = f'40.149.423.292.13.68.01.03.02'
                anaTertip = f'40.149.422.291.13.68.01.03.02'
                memTertip = f'98.900.9006.306.13.67.01.03.02'
                nitelik = f'Temel Eğitim Okulları {self.dönem} Dönem Doğalgaz Ödemesi' ###
                metin = f'      Müdürlüğümüz  Temel Eğitim Okullarının ({kurumlar}) {self.dönem} dönem {tür} aboneliklerine ait toplam {tutar} borç ödemesi hususunu Onaylarınıza arz ederim.'
                
                harcamaTalimatı = {
                    'tarih' : date.today().strftime("%d.%m.%Y"),
                    'tanım' : faturaTür,
                    'nitelik' : nitelik,
                    'miktar' : tutar,
                    'ödenek1' : '',
                    'ödenek2' : '',
                    'metin' : metin,
                    'ilkTertip': ilkTertip,
                    'anaTertip': anaTertip,
                    'memTertip': memTertip
                }       

                MYS(df_mys,df_mebbis,df_mebbis_dummy, df_kurum_ana, df_kurum_mem, firma, faturaTür, imzaListe, dosyaYolu, tekKaynak, harcamaTalimatı).MYS()




