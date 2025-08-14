import xml.etree.ElementTree as ET
import pandas as pd
import os

def parse_invoice(xml_path):
    ns = {"cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"}
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()

        fatura_numarasi = root.findtext(".//cbc:ID", namespaces=ns)
        fatura_tutari = root.findtext(".//cbc:PayableAmount", namespaces=ns)
        notlar = [note.text for note in root.findall(".//cbc:Note", namespaces=ns)]

        fatura_tarihi = next((n.split(": ")[1] for n in notlar if n and n.startswith("#BLDAT")), None)
        tesisat_no = next((n.split(": ")[1] for n in notlar if n and n.startswith("#TESISAT")), None)
        fatura_donemi = next((n.split(": ")[1] for n in notlar if n and n.startswith("#DONEM")), None)
        vkn = next((n.split(": ")[1] for n in notlar if n and n.startswith("#VNO")), None)

        return {
            "Dosya": os.path.basename(xml_path),
            "Fatura Numarası": fatura_numarasi,
            "Fatura Tutarı": fatura_tutari,
            "Fatura Tarihi": fatura_tarihi,
            "Tesisat No": str(tesisat_no) if tesisat_no else None,
            "Fatura Dönemi": fatura_donemi,
            "VKN": vkn
        }

    except Exception as e:
        print(f"HATA ({xml_path}): {e}")
        return None

def read_all_invoices():
    folder_path = r"C:\Users\Levent Aydın\Desktop\Faturalar\Elektrik\xml"
    excel_path = r"C:\Users\Levent Aydın\Desktop\Faturalar\Elektrik\Fatura Listesi.xlsx"

    data = []
    for file in os.listdir(folder_path):
        if file.lower().endswith(".xml"):
            xml_path = os.path.join(folder_path, file)
            parsed = parse_invoice(xml_path)
            if parsed:
                data.append(parsed)

    df_xml = pd.DataFrame(data)

    # Excel dosyasını oku
    df_excel = pd.read_excel(excel_path)[["ABONE NUMARASI", "İCMAL NO", "KURUM ADI"]]

    # NaN temizliği ve .0 uzantılarını kaldırma
    df_excel["ABONE NUMARASI"] = df_excel["ABONE NUMARASI"].fillna("").astype(str).str.split(".").str[0]
    df_excel["İCMAL NO"] = df_excel["İCMAL NO"].fillna("").astype(str).str.split(".").str[0]
    df_xml["Tesisat No"] = df_xml["Tesisat No"].fillna("").astype(str)

    # Birleştirme (merge)
    merged_df = df_xml.merge(
        df_excel,
        left_on="Tesisat No",
        right_on="ABONE NUMARASI",
        how="left"
    )

    # Boş eşleşmeleri filtrele
    bos_eslesmeler = merged_df[merged_df["İCMAL NO"].isna()]

    # Sonuçları yazdır
    print("✅ TÜM VERİLER:")
    print(merged_df)

    print("\n❌ BOŞ EŞLEŞMELER (Excel'de bulunmayan Tesisat No'lar):")
    print(bos_eslesmeler)

    # Excel çıktısı
    output_path = r"C:\Users\Levent Aydın\Desktop\eslesme_sonuclari.xlsx"
    with pd.ExcelWriter(output_path) as writer:
        merged_df.to_excel(writer, sheet_name="Tüm Veriler", index=False)
        bos_eslesmeler.to_excel(writer, sheet_name="Boş Eşleşmeler", index=False)

    print(f"\n📁 Sonuçlar kaydedildi: {output_path}")

# Çalıştır
read_all_invoices()
