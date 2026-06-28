import pandas as pd
import requests
from typing import List, Dict
import time
from datetime import datetime, timedelta, timezone

# Konfigurasi
FONNTE_API_KEY = "9eHaFAQMb9pqoGABACCb"
FONNTE_API_URL = "https://api.fonnte.com/send"

# Google Sheets ID (dari URL spreadsheet kamu)
SHEET_ID = "1ZlPaZQwz89UkjABHgxqmcYnTklB5yK2V"
SHEET_NAME = "Laporan"

def read_laporan_from_google_sheets(sheet_id: str, sheet_name: str) -> Dict:
    """
    Baca data langsung dari Google Sheets
    """
    try:
        # URL untuk export sheet sebagai CSV
        csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
        
        print(f"✓ Mengambil data dari Google Sheets...")
        
        # Baca sebagai pandas DataFrame
        df = pd.read_csv(csv_url, header=None)
        
        print(f"✓ Data berhasil diambil ({len(df)} baris)")
        
        # Helper function untuk ambil cell value
        def get_cell_value(row, col, default=''):
            try:
                val = df.iloc[row, col]
                if pd.isna(val):
                    return default
                return str(val).strip()
            except:
                return default
        
        def format_tahun(tahun):
            try:
                if isinstance(tahun, (int, float)):
                    return str(int(tahun))
                else:
                    tahun_str = str(tahun).split('.')[0]
                    return tahun_str
            except:
                return str(tahun)
        
        def format_date(date_str):
            if not date_str:
                return ''
            try:
                if isinstance(date_str, str):
                    if ' ' in date_str:
                        date_part = date_str.split(' ')[0]
                    else:
                        date_part = date_str
                    
                    if '-' in date_part:
                        parts = date_part.split('-')
                        if len(parts) == 3:
                            return f"{parts[2]}-{parts[1]}-{parts[0]}"
                
                return date_str
            except:
                return date_str
        
        # Mapping baris dan kolom (index mulai dari 0)
        # B3 = row 2, col 1
        # C5 = row 4, col 2
        # dst
        
        data = {
            # Header & Informasi Dasar
            'perusahaan': get_cell_value(2, 1, 'PT GI REJOSO PASURUAN'),        # B3
            'judul_laporan': get_cell_value(4, 2, 'LAPORAN RUTIN DGA'),          # C5
            'tahun': format_tahun(get_cell_value(5, 2, '2025')),                 # C6
            'tanggal_pengukuran': format_date(get_cell_value(5, 1, '')),         # B6
            'tanggal_laporan': format_date(get_cell_value(7, 1, '')),            # B8
            'unit': get_cell_value(4, 6, ''),                                    # G5
            'asset': get_cell_value(5, 6, ''),                                   # G6
            
            # Spesifikasi Trafo Kolom C
            'manufacture': get_cell_value(10, 2, ''),                            # C11
            'type': get_cell_value(11, 2, ''),                                   # C12
            'kapasitas': get_cell_value(12, 2, ''),                              # C13
            'serial_number': get_cell_value(13, 2, ''),                          # C14
            'tegangan_primer': get_cell_value(14, 2, ''),                        # C15
            'tegangan_sekunder': get_cell_value(15, 2, ''),                      # C16
            
            # Spesifikasi Trafo Kolom E
            'arus_primer': get_cell_value(10, 4, ''),                            # E11
            'arus_sekunder': get_cell_value(11, 4, ''),                          # E12
            'frequency': get_cell_value(12, 4, ''),                              # E13
            'phase': get_cell_value(13, 4, ''),                                  # E14
            'temp_rise': get_cell_value(14, 4, ''),                              # E15
            'berat_oli': get_cell_value(15, 4, ''),                              # E16
            
            # Spesifikasi Trafo Kolom G
            'hubungan': get_cell_value(10, 6, ''),                               # G11
            'cooling_type': get_cell_value(11, 6, ''),                           # G12
            'tahun_pembuatan': format_tahun(get_cell_value(12, 6, '')),          # G13
            'impedansi': get_cell_value(13, 6, ''),                              # G14
            'berat_total': get_cell_value(14, 6, ''),                            # G15
            
            # Analisis Konten Gas (baris mulai dari 18 = index 18)
            'h2_hasil': get_cell_value(18, 3, 0),          # D19
            'h2_kondisi': get_cell_value(18, 4, 0),        # E19
            'h2_prealarm': get_cell_value(18, 5, 0),       # F19
            'h2_alarm': get_cell_value(18, 6, 0),          # G19
            
            'h2o_hasil': get_cell_value(19, 3, 0),         # D20
            'h2o_kondisi': get_cell_value(19, 4, 0),       # E20
            'h2o_prealarm': get_cell_value(19, 5, 0),      # F20
            'h2o_alarm': get_cell_value(19, 6, 0),         # G20
            
            'co2_hasil': get_cell_value(21, 3, 0),         # D22
            'co2_kondisi': get_cell_value(21, 4, 0),       # E22
            'co2_prealarm': get_cell_value(21, 5, 0),      # F22
            'co2_alarm': get_cell_value(21, 6, 0),         # G22
            
            'co_hasil': get_cell_value(22, 3, 0),          # D23
            'co_kondisi': get_cell_value(22, 4, 0),        # E23
            'co_prealarm': get_cell_value(22, 5, 0),       # F23
            'co_alarm': get_cell_value(22, 6, 0),          # G23
            
            'c2h4_hasil': get_cell_value(23, 3, 0),        # D24
            'c2h4_kondisi': get_cell_value(23, 4, 0),      # E24
            'c2h4_prealarm': get_cell_value(23, 5, 0),     # F24
            'c2h4_alarm': get_cell_value(23, 6, 0),        # G24
            
            'c2h6_hasil': get_cell_value(24, 3, 0),        # D25
            'c2h6_kondisi': get_cell_value(24, 4, 0),      # E25
            'c2h6_prealarm': get_cell_value(24, 5, 0),     # F25
            'c2h6_alarm': get_cell_value(24, 6, 0),        # G25
            
            'ch4_hasil': get_cell_value(25, 3, 0),         # D26
            'ch4_kondisi': get_cell_value(25, 4, 0),       # E26
            'ch4_prealarm': get_cell_value(25, 5, 0),      # F26
            'ch4_alarm': get_cell_value(25, 6, 0),         # G26
            
            'c2h2_hasil': get_cell_value(26, 3, 0),        # D27
            'c2h2_kondisi': get_cell_value(26, 4, 0),      # E27
            'c2h2_prealarm': get_cell_value(26, 5, 0),     # F27
            'c2h2_alarm': get_cell_value(26, 6, 0),        # G27
            
            'tdcg_hasil': get_cell_value(27, 3, 0),        # D28
            'tdcg_kondisi': get_cell_value(27, 4, 0),      # E28
            'tdcg_prealarm': get_cell_value(27, 5, 0),     # F28
            'tdcg_alarm': get_cell_value(27, 6, 0),        # G28
            
            # Analisa
            'tdcg_analisa': get_cell_value(31, 2, 'KONDISI 1'),                          # C32
            'keygas_analisa': get_cell_value(32, 2, 'OVERHEATING OF CELLULOSE'),         # C33
            'roger_ratio_analisa': get_cell_value(33, 2, 'TIDAK TERDEFINISIKAN'),        # C34
            'duval_triangle_analisa': get_cell_value(34, 2, 'T3'),                       # C35
            
            # Rekomendasi
            'tdcg_rekomendasi': get_cell_value(36, 2, 'Sampling 12 Bulan + Operasi Normal'),           # C37
            'keygas_rekomendasi': get_cell_value(37, 2, 'Uji PD, Furan, Kadar Air + Cek Umur Isolasi'), # C38
            'roger_ratio_rekomendasi': get_cell_value(38, 2, 'GUNAKAN METODE KEYGAS DAN DUVAL TRIANGLE'), # C39
            'duval_triangle_rekomendasi': get_cell_value(39, 2, 'LEPAS BEBAN + INSPEKSI INTERNAL TRAFO'), # C40
        }
        
        return data
        
    except Exception as e:
        print(f"✗ Error: {str(e)}")
        return None

def get_current_time_wib():
    """Ambil waktu saat ini dalam timezone Indonesia (WIB = UTC+7)"""
    tz_wib = timezone(timedelta(hours=7))
    now = datetime.now(tz_wib)
    return now.strftime('%d-%m-%Y %H:%M')

def format_laporan_to_whatsapp_single(data: Dict) -> str:
    message = f"*🏢 {data['perusahaan']}*\n"
    message += f"*📊 {data['judul_laporan']} - {data['tahun']}*\n"
    message += "=" * 50 + "\n\n"

    message += "*📅 INFORMASI DASAR*\n"
    message += f"Tanggal Pengukuran: {data['tanggal_pengukuran']}\n"
    message += f"Tanggal Laporan: {data['tanggal_laporan']}\n"
    message += f"Unit: {data['unit']}\n"
    message += f"Asset: {data['asset']}\n\n"

    message += "*⚙️ SPESIFIKASI TRAFO*\n"
    message += f"Manufacture: {data['manufacture']}\n"
    message += f"Kapasitas: {data['kapasitas']}\n"
    message += f"Serial Number: {data['serial_number']}\n"
    message += f"Tegangan Primer: {data['tegangan_primer']}\n"
    message += f"Tegangan Sekunder: {data['tegangan_sekunder']}\n"

    if data['arus_primer']:
        message += f"Arus Primer: {data['arus_primer']}\n"
    if data['arus_sekunder']:
        message += f"Arus Sekunder: {data['arus_sekunder']}\n"

    message += f"Frequency: {data['frequency']}\n"
    message += f"Phase: {data['phase']}\n"

    if data['temp_rise']:
        message += f"Temp Rise: {data['temp_rise']}\n"

    message += f"Kapasitas Oil: {data['berat_oli']}\n"
    message += f"Konfigurasi: {data['hubungan']}\n"
    message += f"Cooling : {data['cooling_type']}\n"
    message += f"Tahun Pembuatan: {data['tahun_pembuatan']}\n"

    message += "\n"

    message += "*🔬 CONTENT ANALYSIS*\n\n"

    gas_data = [
        ("Hydrogen", "H2", data['h2_hasil'], data['h2_kondisi'], data['h2_prealarm'], data['h2_alarm']),
        ("Air", "H2O", data['h2o_hasil'], data['h2o_kondisi'], data['h2o_prealarm'], data['h2o_alarm']),
        ("Karbon Dioksida", "CO2", data['co2_hasil'], data['co2_kondisi'], data['co2_prealarm'], data['co2_alarm']),
        ("Karbon Monoksida", "CO", data['co_hasil'], data['co_kondisi'], data['co_prealarm'], data['co_alarm']),
        ("Etilen", "C2H4", data['c2h4_hasil'], data['c2h4_kondisi'], data['c2h4_prealarm'], data['c2h4_alarm']),
        ("Etana", "C2H6", data['c2h6_hasil'], data['c2h6_kondisi'], data['c2h6_prealarm'], data['c2h6_alarm']),
        ("Metana", "CH4", data['ch4_hasil'], data['ch4_kondisi'], data['ch4_prealarm'], data['ch4_alarm']),
        ("Asetilen", "C2H2", data['c2h2_hasil'], data['c2h2_kondisi'], data['c2h2_prealarm'], data['c2h2_alarm']),
        ("Total Gas", "TDCG", data['tdcg_hasil'], data['tdcg_kondisi'], data['tdcg_prealarm'], data['tdcg_alarm']),
    ]

    for nama_gas, simbol, hasil, kondisi, prealarm, alarm in gas_data:
        message += f"{nama_gas} ({simbol})\n"
        message += f"  • Hasil Pengujian: {hasil}  | Kondisi 2: {kondisi}  | Pre-Alarm: {prealarm}  | Alarm: {alarm}\n"
        message += "\n"

    message += "*🔍 HASIL ANALISA*\n\n"
    message += f"TDCG:\n{data['tdcg_analisa']}\n\n"
    message += f"KEY GAS:\n{data['keygas_analisa']}\n\n"
    message += f"ROGER RATIO:\n{data['roger_ratio_analisa']}\n\n"
    message += f"DUVAL TRIANGLE:\n{data['duval_triangle_analisa']}\n\n"

    message += "*📌 REKOMENDASI*\n\n"
    message += f"TDCG:\n{data['tdcg_rekomendasi']}\n\n"
    message += f"KEY GAS:\n{data['keygas_rekomendasi']}\n\n"
    message += f"ROGER RATIO:\n{data['roger_ratio_rekomendasi']}\n\n"
    message += f"DUVAL TRIANGLE:\n{data['duval_triangle_rekomendasi']}\n\n"

    message += "=" * 50 + "\n"
    message += f"Generated: {get_current_time_wib()}\n"

    return message

def send_whatsapp_message(phone_number: str, message: str, is_group: bool = False) -> bool:
    try:
        headers = {
            "Authorization": FONNTE_API_KEY,
            "Content-Type": "application/json"
        }

        if is_group:
            payload = {
                "target": phone_number,
                "message": message
            }
        else:
            payload = {
                "target": phone_number,
                "message": message,
                "countryCode": "62"
            }

        response = requests.post(
            FONNTE_API_URL,
            headers=headers,
            json=payload,
            timeout=10
        )

        if response.status_code == 200:
            result = response.json()
            if result.get("status"):
                device_type = "grup" if is_group else "personal"
                print(f"  ✓ Pesan berhasil dikirim ke {device_type}: {phone_number}")
                return True
            else:
                print(f"  ✗ Gagal: {result.get('reason', 'Unknown error')}")
                return False
        else:
            print(f"  ✗ Error API: Status {response.status_code}")
            return False

    except Exception as e:
        print(f"  ✗ Error: {str(e)}")
        return False

# ===== MAIN =====

if __name__ == "__main__":
    print("\n" + "="*70)
    print("📋 PROGRAM KIRIM LAPORAN DGA KE WHATSAPP")
    print(f"⏰ Waktu: {get_current_time_wib()}")
    print("="*70 + "\n")

    print("1️⃣  Membaca data dari Google Sheets...")
    data = read_laporan_from_google_sheets(SHEET_ID, SHEET_NAME)

    if data is None:
        print("\n❌ Program dihentikan!")
        exit(1)

    print("2️⃣  Memformat laporan...")
    message = format_laporan_to_whatsapp_single(data)
    print(f"  ✓ Laporan siap: {len(message)} karakter\n")

    phone_numbers = [
        {"number": "6282157905834", "type": "personal"},
        {"number": "120363426691870718@g.us", "type": "group"},
    ]

    print(f"3️⃣  Mengirim ke {len(phone_numbers)} nomor/grup WhatsApp...")
    print("-" * 70)

    success_total = 0
    for item in phone_numbers:
        is_group = item.get('type') == 'group'
        phone_number = item.get('number')
        device_type = "Grup" if is_group else "Personal"

        print(f"\n📱 [{device_type}] Mengirim ke: {phone_number}")
        if send_whatsapp_message(phone_number, message, is_group):
            success_total += 1
        time.sleep(1)

    print("\n" + "="*70)
    if success_total == len(phone_numbers):
        print(f"✅ BERHASIL! {success_total}/{len(phone_numbers)} pesan terkirim")
    else:
        print(f"⚠️  {success_total}/{len(phone_numbers)} pesan terkirim")
    print("="*70 + "\n")
