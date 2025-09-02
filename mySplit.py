
import pandas as pd
import numpy as np
import os
from datetime import date
from paracevir import ParaCevir
from myS import MYS


class Fatura:
    def __init__(self, ay, yıl):
        self.dönem = f"{yıl}/{ay}"
        base = self._read_dosya_yolu()
        # Ortak yollar
        self.paths = {
            "kurum": "kurum.xlsx",
            "imza": "imza.xlsx",
            "Elektrik": (f"{base}\\Elektrik\\Fatura Listesi.xlsx", f"{base}\\Elektrik\\MYS.xlsx", ("firma.xlsx", "Elektrik")),
            "ADM": (f"{base}\\ADM\\Fatura Listesi.xlsx", f"{base}\\ADM\\MYS.xlsx", ("firma.xlsx", "ADM")),
            "İnternet": (f"{base}\\İnternet\\Fatura Listesi.xlsx", f"{base}\\İnternet\\MYS.xlsx", ("firma.xlsx", "İnternet")),
            "MEMİnternet": (f"{base}\\MEM\\İnternet\\Fatura Listesi.xlsx", f"{base}\\MEM\\İnternet\\MYS.xlsx", ("firma.xlsx", "İnternet")),
            "Telefon": (f"{base}\\Telefon\\Fatura Listesi.xlsx", f"{base}\\Telefon\\MYS.xlsx", ("firma.xlsx", "Telefon")),
            "MEMTelefon": (f"{base}\\MEM\\Telefon\\Fatura Listesi.xlsx", f"{base}\\MEM\\Telefon\\MYS.xlsx", ("firma.xlsx", "Telefon")),
            "Doğalgaz": (f"{base}\\Doğalgaz\\Fatura Listesi.xlsx", f"{base}\\Doğalgaz\\MYS.xlsx", ("firma.xlsx", "Gaz")),
        }

    # ---------- IO helpers ----------
    def _read_dosya_yolu(self):
        try:
            with open("dosyaYolu.txt", "r", encoding="utf-8") as f:
                return f.read().strip()
        except UnicodeDecodeError:
            with open("dosyaYolu.txt", "r", encoding="cp1254") as f:
                return f.read().strip()

    def _read_excel_cols(self, path, cols, dtype=None, drop_last=True):
        df = pd.read_excel(path, header=0, dtype=dtype)
        df = df.reindex(cols, axis=1)
        if drop_last and not df.empty:
            df = df.drop(df.index[-1])
        return df

    def _read_firma(self, path, sheet):
        return pd.read_excel(io=path, sheet_name=sheet, header=0)

    def _read_imza(self):
        df = pd.read_excel(self.paths["imza"], header=0)
        # [harcamaYetkilisiAd, gerçekleştirmeGörevlisiAd, harcamaYetkilisiUnvan, gerçekleştirmeGörevlisiUnvan]
        return [df.values[0][0], df.values[1][0], df.values[0][1], df.values[1][1]]

    # ---------- Processing helpers ----------
    def _process_mys(self, path, split_cols=None, drop_cols=None, abone_type=None,
                     abone_fill=None, abone_strip=False):
        df = pd.read_excel(path, header=0)

        # Kolon bölmeler
        if split_cols:
            for col, (new_cols, expand) in split_cols.items():
                if col in df.columns:
                    df[new_cols] = df[col].astype(str).str.split('-', expand=expand)

        # Her zaman Tarih sütunu ekle (varsa Fatura Tarihi'nden üret)
        if "Fatura Tarihi" in df.columns:
            df["Tarih"] = df["Fatura Tarihi"].astype(str).str.replace('-', '/', regex=False)

        # Abone temizlik / doldurma
        if abone_strip and "Abone" in df.columns:
            df['Abone'] = df['Abone'].astype(str).str.replace(" ", "", regex=False)

        if abone_fill is not None:
            if "Abone" not in df.columns:
                df["Abone"] = None
            df['Abone'] = abone_fill

        if abone_type and "Abone" in df.columns:
            # Güvenli dönüştürme
            try:
                df['Abone'] = df['Abone'].astype(abone_type)
            except Exception:
                pass

        # Gereksiz kolonları at
        if drop_cols:
            df = df.drop(columns=drop_cols, errors="ignore")

        return df

    def _build_tek_kaynak(self, df_mebbis, df_mebbis_dummy, df_firma, imzaListe, ihtiyaç_fmt):
        toplam_tutar = float(df_mebbis['FATURA TUTARI'].sum())
        toplam = "{:,.2f}".format(toplam_tutar).replace(',', '-').replace('.', ',').replace('-', '.')
        tutar = f'{toplam} TL ({ParaCevir(toplam_tutar)})'

        unik = len(df_mebbis['KURUM ADI'].unique())
        kurumlar = ", ".join(str(e) for e in df_mebbis['KURUM ADI'].unique()) if unik < 4 else f'{len(df_mebbis)} Adet Fatura'
        icmaller = ", ".join(str(e) for e in df_mebbis_dummy['İCMAL NO'].unique())
        ihtiyaç = ihtiyaç_fmt.format(dönem=self.dönem, kurumlar=kurumlar, icmaller=icmaller)

        return {
            'firma': df_firma.values[0][1],
            'tebligat': df_firma.values[1][1],
            'vergi': df_firma.values[2][1],
            'telefon': df_firma.values[3][1],
            'eposta': df_firma.values[4][1],
            'tutar': tutar,
            'ihtiyaç': ihtiyaç,
            'harcama': imzaListe[2],
            'unvan': imzaListe[3]
        }, tutar, kurumlar

    def _build_harcama_talimati(self, faturaTür, nitelik, tutar, kurumlar, tür, tertipler, mem=False):
        if mem:
            metin = f'      Müdürlüğümüze ait {self.dönem} dönem {tür} abonelikleri toplam {tutar} borç ödemesi hususunu Onaylarınıza arz ederim.'
        else:
            metin = f'      Müdürlüğümüz  Temel Eğitim Okullarının ({kurumlar}) {self.dönem} dönem {tür} aboneliklerine ait toplam {tutar} borç ödemesi hususunu Onaylarınıza arz ederim.'
        return {
            'tarih': date.today().strftime("%d.%m.%Y"),
            'tanım': faturaTür,
            'nitelik': nitelik,
            'miktar': tutar,
            'ödenek1': '',
            'ödenek2': '',
            'metin': metin,
            **tertipler
        }

    def _process_fatura(self, kategori, tür, tertipler, mys_opts=None, abone_type=None, abone_fill=None,
                         abone_strip=False, special_mebbis=None, ihtiyaç_fmt=None, mem=False,
                         update_fatura_tutari=False):
        yolFatura, yolMYS, (firma_path, firma_sheet) = self.paths[kategori]

        if not (os.path.exists(yolFatura) and os.path.exists(yolMYS)):
            print(f"{kategori} dosyaları bulunamadı!")
            return

        df_firma = self._read_firma(firma_path, firma_sheet)
        imzaListe = self._read_imza()

        # MYS
        df_mys = self._process_mys(
            yolMYS, **(mys_opts or {}),
            abone_type=abone_type, abone_fill=abone_fill, abone_strip=abone_strip
        )

        # MEBBİS
        meb_cols = [
            'KURUM ADI', 'ABONE NUMARASI', 'VERGİ NO', 'FATURA NUMARASI',
            'FATURA TARİHİ', 'İCMAL NO', 'TÜKETİM MİKTARI', 'FATURA TUTARI'
        ]
        df_mebbis = self._read_excel_cols(yolFatura, meb_cols)

        if special_mebbis:
            for k, v in special_mebbis.items():
                df_mebbis[k] = v

        # Elektrik için MYS -> MEBBİS tutar güncellemesi
        if update_fatura_tutari:
            if {'Fatura No', 'Ödenecek Tutar'}.issubset(set(df_mys.columns)):
                mapping = dict(zip(df_mys['Fatura No'], df_mys['Ödenecek Tutar']))
                df_mebbis['FATURA TUTARI'] = df_mebbis['FATURA NUMARASI'].map(mapping).fillna(df_mebbis['FATURA TUTARI'])

        # Tip dönüşümleri
        df_mebbis['VERGİ NO'] = df_mebbis['VERGİ NO'].astype(str)
        df_mebbis['İCMAL NO'] = df_mebbis['İCMAL NO'].astype(str)

        # Dummy ve kurum listeleri
        df_mebbis_dummy = self._read_excel_cols(yolFatura, ['VERGİ NO', 'İCMAL NO'], dtype=str)
        df_kurum = self._read_excel_cols(self.paths["kurum"], ['VERGİ KİMLİK NO', 'KURUM TÜRÜ'], dtype=str)
        df_kurum_ana = df_kurum[df_kurum['KURUM TÜRÜ'] == 'Okul Öncesi']
        df_kurum_mem = df_kurum[df_kurum['KURUM TÜRÜ'] == 'MEM']

        # Belgeler
        tekKaynak, tutar, kurumlar = self._build_tek_kaynak(
            df_mebbis, df_mebbis_dummy, df_firma, imzaListe, ihtiyaç_fmt
        )
        harcamaTalimatı = self._build_harcama_talimati(
            f"{tür.capitalize()} Aboneliği Ödemesi",
            f'Temel Eğitim Okulları {self.dönem} Dönem {tür.capitalize()} Ödemesi',
            tutar, kurumlar, tür, tertipler, mem=mem
        )

        # MYS çıktısı
        MYS(
            df_mys, df_mebbis, df_mebbis_dummy,
            df_kurum_ana, df_kurum_mem,
            df_firma.values[0][1],
            f"{tür.capitalize()} Aboneliği Ödemesi",
            imzaListe, self._read_dosya_yolu(),
            tekKaynak, harcamaTalimatı
        ).MYS()

    # ----------------- Fatura Tipleri -----------------
    def Elektrik(self):
        self._process_fatura(
            "Elektrik", "elektrik",
            {
                'ilkTertip': '40.149.423.292.13.68.01.03.02',
                'anaTertip': '40.149.422.291.13.68.01.03.02',
                'memTertip': '98.900.9006.306.13.67.01.03.02'
            },
            mys_opts={
                'split_cols': {
                    "Harcama Birimi": (["VKN", "Okul"], True),
                    'Müşteri Kimlik Bilgisi': (["Tesisat No", "Abone"], True)
                },
                'drop_cols': ['Harcama Birimi', 'Fatura Tarihi', 'Müşteri Kimlik Bilgisi', 'Tesisat No', 'Okul']
            },
            abone_type=int,
            ihtiyaç_fmt='TEMEL EĞİTİM OKULLARI {dönem} DÖNEM ELEKTRİK ÖDEMESİ {kurumlar} (İcmal No: {icmaller})',
            update_fatura_tutari=True
        )

    def ADM(self):
        self._process_fatura(
            "ADM", "elektrik",
            {
                'ilkTertip': '40.149.423.292.13.68.01.03.02',
                'anaTertip': '40.149.422.291.13.68.01.03.02',
                'memTertip': '98.900.9006.306.13.67.01.03.02'
            },
            mys_opts={
                'split_cols': {"Harcama Birimi": (["VKN", "Okul"], True)},
                'drop_cols': ['Harcama Birimi', 'Fatura Tarihi', 'Müşteri Kimlik Bilgisi', 'Okul']
            },
            abone_fill=str(100207225527),
            ihtiyaç_fmt='TEMEL EĞİTİM OKULLARI {dönem} DÖNEM ELEKTRİK ÖDEMESİ {kurumlar} (İcmal No: {icmaller})'
        )

    def İnternet(self):
        self._process_fatura(
            "İnternet", "internet",
            {
                'ilkTertip': '40.149.423.292.13.68.01.03.05',
                'anaTertip': '40.149.422.291.13.68.01.03.05',
                'memTertip': '98.900.9006.306.13.67.01.03.05'
            },
            mys_opts={
                'split_cols': {
                    "Harcama Birimi": (["VKN", "Okul"], True),
                    'Müşteri Kimlik Bilgisi': (["Hizmet No", "Abone"], True)
                },
                'drop_cols': ['Harcama Birimi', 'Fatura Tarihi', 'Müşteri Kimlik Bilgisi', 'Hizmet No', 'Okul']
            },
            abone_type=float,
            ihtiyaç_fmt='TEMEL EĞİTİM OKULLARI {dönem} DÖNEM İNTERNET ÖDEMESİ {kurumlar} (İcmal No: {icmaller})'
        )

    def MEMİnternet(self):
        self._process_fatura(
            "MEMİnternet", "internet",
            {
                'ilkTertip': '40.149.423.292.13.68.01.03.05',
                'anaTertip': '40.149.422.291.13.68.01.03.05',
                'memTertip': '98.900.9006.306.13.67.01.03.05'
            },
            mys_opts={
                'split_cols': {
                    "Harcama Birimi": (["VKN", "Okul"], True),
                    'Müşteri Kimlik Bilgisi': (["Hizmet No", "Abone"], True)
                },
                'drop_cols': ['Harcama Birimi', 'Fatura Tarihi', 'Müşteri Kimlik Bilgisi', 'Hizmet No', 'Okul']
            },
            abone_type=float,
            special_mebbis={'KURUM ADI': 'İlçe Milli Eğitim Müdürlüğü', 'VERGİ NO': 7820458686},
            ihtiyaç_fmt='İLÇE MİLLİ EĞİTİM MÜDÜRLÜĞÜ {dönem} DÖNEM İNTERNET ÖDEMESİ (İcmal No: {icmaller})',
            mem=True
        )

    def Telefon(self):
        self._process_fatura(
            "Telefon", "telefon",
            {
                'ilkTertip': '40.149.423.292.13.68.01.03.05',
                'anaTertip': '40.149.422.291.13.68.01.03.05',
                'memTertip': '98.900.9006.306.13.67.01.03.05'
            },
            mys_opts={
                'split_cols': {
                    "Harcama Birimi": (["VKN", "Okul"], True),
                    'Müşteri Kimlik Bilgisi': (["Hizmet No", "Abone"], True)
                },
                'drop_cols': ['Harcama Birimi', 'Fatura Tarihi', 'Müşteri Kimlik Bilgisi', 'Hizmet No', 'Okul']
            },
            abone_type=float,
            abone_strip=True,
            ihtiyaç_fmt='TEMEL EĞİTİM OKULLARI {dönem} DÖNEM TELEFON ÖDEMESİ {kurumlar} (İcmal No: {icmaller})'
        )

    def MEMTelefon(self):
        self._process_fatura(
            "MEMTelefon", "telefon",
            {
                'ilkTertip': '40.149.423.292.13.68.01.03.05',
                'anaTertip': '40.149.422.291.13.68.01.03.05',
                'memTertip': '98.900.9006.306.13.67.01.03.05'
            },
            mys_opts={
                'split_cols': {
                    "Harcama Birimi": (["VKN", "Okul"], True),
                    'Müşteri Kimlik Bilgisi': (["Hizmet No", "Abone"], True)
                },
                'drop_cols': ['Harcama Birimi', 'Fatura Tarihi', 'Müşteri Kimlik Bilgisi', 'Hizmet No', 'Okul']
            },
            abone_type=float,
            abone_strip=True,
            special_mebbis={'KURUM ADI': 'İlçe Milli Eğitim Müdürlüğü', 'VERGİ NO': 7820458686},
            ihtiyaç_fmt='İLÇE MİLLİ EĞİTİM MÜDÜRLÜĞÜ {dönem} DÖNEM TELEFON ÖDEMESİ (İcmal No: {icmaller})',
            mem=True
        )

    def Doğalgaz(self):
        self._process_fatura(
            "Doğalgaz", "doğalgaz",
            {
                'ilkTertip': '40.149.423.292.13.68.01.03.02',
                'anaTertip': '40.149.422.291.13.68.01.03.02',
                'memTertip': '98.900.9006.306.13.67.01.03.02'
            },
            mys_opts={
                'split_cols': {"Harcama Birimi": (["VKN", "Okul"], True)},
                'drop_cols': ['Harcama Birimi', 'Fatura Tarihi', 'Müşteri Kimlik Bilgisi', 'Okul']
            },
            abone_fill="",
            abone_type=float,
            ihtiyaç_fmt='TEMEL EĞİTİM OKULLARI {dönem} DÖNEM DOĞALGAZ ÖDEMESİ {kurumlar} (İcmal No: {icmaller})'
        )
