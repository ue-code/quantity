import comtypes.client
import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog
import time

REBAR_UNIT_WEIGHTS = {
    8: 0.395,
    10: 0.617,
    12: 0.888,
    14: 1.210,
    16: 1.580,
    18: 2.000,
    20: 2.470,
    25: 3.850,
    32: 6.310,
}

def select_dwg_file():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="DWG Dosyasını Seç",
        filetypes=[("DWG Dosyaları", "*.dwg"), ("Tüm Dosyalar", "*.*")]
    )

def start_autocad():
    try:
        acad = comtypes.client.GetActiveObject("AutoCAD.Application")
        print("AutoCAD aktif bir oturuma bağlandı.")
    except Exception:
        print("Aktif bir AutoCAD oturumu bulunamadı. AutoCAD başlatılıyor...")
        acad = comtypes.client.CreateObject("AutoCAD.Application")
        acad.Visible = True
        time.sleep(5)
    return acad

def wait_for_autocad_ready(acad):
    for i in range(10):
        try:
            doc = acad.ActiveDocument
            print(f"AutoCAD aktif belgeye erişildi: {doc.Name}")
            return doc
        except Exception:
            print("AutoCAD henüz hazır değil. Bekleniyor...")
            time.sleep(2)
    raise RuntimeError("AutoCAD bağlantısı zaman aşımına uğradı.")

def open_dwg_file(acad, dwg_file_path):
    try:
        if os.path.exists(dwg_file_path):
            acad.Documents.Open(dwg_file_path)
            print(f"DWG dosyası başarıyla açıldı: {dwg_file_path}")
            return wait_for_autocad_ready(acad)
        else:
            print(f"Belirtilen dosya bulunamadı: {dwg_file_path}")
            return None
    except Exception as e:
        print(f"DWG dosyası açılırken bir hata oluştu: {e}")
        return None

def list_modelspace_entities_safe(doc):
    try:
        print("Model Space'teki nesneler:")
        for entity in doc.ModelSpace:
            print(f"Entity Name: {entity.EntityName}")
    except Exception as e:
        print(f"Model Space'teki nesneler okunurken bir hata oluştu: {e}")

def parse_text_content(text):
    """Metin içeriğini analiz ederek adet, çap ve uzunluk bilgilerini döndürür."""
    # Adet, çap ve uzunluğu yakalamaya çalışıyoruz
    match = re.search(r"(\d+x\d+|\d+)?Φ(\d+)(?:/\d+)?\s+l=(\d+)", text, re.IGNORECASE)
    if match:
        # Adet bilgisi varsa ayıkla, yoksa 1 olarak varsay
        count_str = match.group(1)
        if count_str:
            if "x" in count_str.lower():
                # Çarpım ifadesini hesapla (örneğin, "2x2" -> 4)
                count_parts = count_str.lower().split("x")
                count = int(count_parts[0]) * int(count_parts[1])
            else:
                count = int(count_str)
        else:
            count = 1

        diameter = int(match.group(2))  # Çap
        length = int(match.group(3)) / 100  # Boy (metre)
        return count, diameter, length
    return None, None, None


def extract_text_from_blocktablerecord(block_ref):
    """BlockTableRecord içindeki tüm nesneleri dolaşarak metinleri çıkarır."""
    block_texts = []
    try:
        # BlockTableRecord'ı al
        block_table_record = block_ref.BlockTableRecord

        # İçerikteki nesneleri dolaş
        for item in block_table_record:
            if item.EntityName in ["AcDbText", "AcDbMText"]:
                text = item.TextString.strip()
                block_texts.append(text)
                print(f"Blok içindeki metin: {text}")
    except Exception as e:
        print(f"Blok nesneleri okunurken bir hata oluştu: {e}")
    return block_texts

def process_blocks_with_table_record(doc):
    """BlockTableRecord kullanarak tüm blok referanslarını işler."""
    rebar_data = []

    try:
        for entity in doc.ModelSpace:
            if entity.EntityName == "AcDbBlockReference":
                block_name = entity.Name
                print(f"Blok bulundu: {block_name}")

                # Blok içindeki metinleri al
                block_texts = extract_text_from_blocktablerecord(entity)
                for text in block_texts:
                    count, diameter, length = parse_text_content(text)
                    if count and diameter and length:
                        unit_weight = REBAR_UNIT_WEIGHTS.get(diameter, 0)
                        if unit_weight == 0:
                            print(f"Uyarı: Çap {diameter} için birim ağırlık tanımlı değil.")
                            continue
                        total_weight = count * length * unit_weight
                        rebar_data.append({
                            "Count": count,
                            "Diameter (mm)": diameter,
                            "Length (m)": length,
                            "Unit Weight (kg/m)": unit_weight,
                            "Total Weight (kg)": total_weight,
                        })
    except Exception as e:
        print(f"Blok verileri işlenirken bir hata oluştu: {e}")
        return []

    return rebar_data

def process_rebars(doc):
    rebar_data = []
    try:
        for entity in doc.ModelSpace:
            if entity.EntityName in ["AcDbText", "AcDbMText"]:
                text_content = entity.TextString.strip()
                print(f"Analiz edilen metin: {text_content}")
                count, diameter, length = parse_text_content(text_content)
                if count and diameter and length:
                    unit_weight = REBAR_UNIT_WEIGHTS.get(diameter, 0)
                    if unit_weight == 0:
                        print(f"Uyarı: Çap {diameter} için birim ağırlık tanımlı değil.")
                        continue
                    total_weight = count * length * unit_weight
                    rebar_data.append({
                        "Count": count,
                        "Diameter (mm)": diameter,
                        "Length (m)": length,
                        "Unit Weight (kg/m)": unit_weight,
                        "Total Weight (kg)": total_weight,
                    })
    except Exception as e:
        print(f"Donatı verileri okunurken bir hata oluştu: {e}")
        return []
    return rebar_data

def save_to_excel(rebar_data, dwg_file_path):
    """Donatı verilerini bir Excel dosyasına kaydeder."""
    if not rebar_data:
        print("Analiz edilecek donatı verisi bulunamadı.")
        return

    # DataFrame oluşturma
    df = pd.DataFrame(rebar_data)

    # Toplam ağırlık hesaplama
    total_rebar_weight = df["Total Weight (kg)"].sum()
    print(f"Toplam donatı ağırlığı: {total_rebar_weight:.2f} kg")
    
  # Toplam satırını oluşturma
    total_row = pd.DataFrame([{
        "Count": "Toplam",
        "Diameter (mm)": "",
        "Length (m)": "",
        "Unit Weight (kg/m)": "",
        "Total Weight (kg)": total_rebar_weight,
    }])

    # Toplam satırını veri çerçevesine ekleme
    df = pd.concat([df, total_row], ignore_index=True)

    # Excel dosyasını kaydetme
    output_file_excel = os.path.join(os.path.dirname(dwg_file_path), "donati_metraj.xlsx")
    df.to_excel(output_file_excel, index=False, sheet_name="Donati Metraj")
    print(f"Donatı metrajı Excel dosyası şu dosyaya kaydedildi: {output_file_excel}")

if __name__ == "__main__":
    acad = start_autocad()
    dwg_file_path = select_dwg_file()
    if dwg_file_path:
        acad.Documents.Open(dwg_file_path)
        doc = acad.ActiveDocument
        rebar_data = process_rebars(doc)
        save_to_excel(rebar_data, dwg_file_path)
    else:
        print("Hiçbir dosya seçilmedi. İşlem iptal edildi.")