import pandas as pd
import requests
from typing import Dict, Any
import time
from datetime import datetime, timedelta, timezone

# ================== KONFIGURASI ==================
FONNTE_API_KEY = "9eHaFAQMb9pqoGABACCb"
FONNTE_API_URL = "https://api.fonnte.com/send"

SHEET_ID = "1ZlPaZQwz89UkjABHgxqmcYnTklB5yK2V"
SHEET_NAME = "Laporan"

def read_laporan_from_google_sheets(sheet_id: str, sheet_name: str) -> Dict:
    try:
        csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
        
        print("✓ Mengambil data dari Google Sheets...")
        df = pd.read_csv(csv_url, header=None)
        print(f"✓ Data diambil: {df.shape[0]} baris x {df.shape[1]} kolom")
        
        def get_cell(row: int, col: int, default: Any = ""):
            try:
                val = df.iloc[row, col]
                if pd.isna(val) or str(val).strip() in ["", "nan", "NaN"]:
                    return default
                return str(val).strip()
            except:
                return default

        def clean_number(val):
            if not val or val == "":
                return "0"
            val = str(val).replace("ppm", "").replace("°C", "").replace(",", ".").strip()
            return val

        data = {
            # Header
            'perusahaan': get_cell(0, 1, 'PT GI REJOSO PASURUAN'),
            'judul_laporan': get_cell(0, 3, 'LAPORAN RUTIN DGA'),
            'tahun': get_cell(1, 2, '2025'),
            'tanggal_pengukuran': get_cell(1, 1, '28-Mei-25'),
            'tanggal_laporan': get_cell(3, 1, '28-Mei-25'),
            'unit': get_cell(0, 6, 'Unit Trafo A'),
            'asset': get_cell(1, 5, ''),

            # Spesifikasi Trafo
            'manufacture': get_cell(5, 1, 'ILJUN'),
            'type': get_cell(6, 1),
            'kapasitas': get_cell(7, 1, '30 MVA'),
            'serial_number': get_cell(8, 1),
            'tegangan_primer': get_cell(9, 1, '150 KV'),
            'tegangan_sekunder': get_cell(10, 1, '6.6 KV'),

            'arus_primer': get_cell(5, 2),
            'arus_sekunder': get_cell(6, 2),
            'frequency': get_cell(7, 2),
            'phase': get_cell(8, 2),
            'temp_rise': get_cell(9, 2),
            'berat_oli': get_cell(10, 2),

            'hubungan': get_cell(5, 6),
            'cooling_type': get_cell(6, 6),
            'tahun_pembuatan': get_cell(7, 6, '1995'),
            'impedansi': get_cell(8, 6),
            'berat_total': get_cell(9, 6),

            # Gas Analysis (baris mulai dari 12)
            'h2_hasil': clean_number(get_cell(12, 2)),
            'h2_kondisi': get_cell(12, 3),
            'h2_prealarm': get_cell(12, 4),
            'h2_alarm': get_cell(12, 5),

            'h2o_hasil': clean_number(get_cell(13, 2)),
            'h2o_kondisi': get_cell(13, 3),
            'h2o_prealarm': get_cell(13, 4),
            'h2o_alarm': get_cell(13, 5),

            'co2_hasil': clean_number(get_cell(14, 2)),
            'co2_kondisi': get_cell(14, 3),
            'co2_prealarm': get_cell(14, 4),
            'co2_alarm': get_cell(14, 5),

            'co_hasil': clean_number(get_cell(15, 2)),
            'co_kondisi': get_cell(15, 3),
            'co_prealarm': get_cell(15, 4),
            'co_alarm': get_cell(15, 5),

            'c2h4_hasil': clean_number(get_cell(16, 2)),
            'c2h4_kondisi': get_cell(16, 3),
            'c2h4_prealarm': get_cell(16, 4),
            'c2h4_alarm': get_cell(16, 5),

            'c2h6_hasil': clean_number(get_cell(17, 2)),
            'c2h6_kondisi': get_cell(17, 3),
            'c2h6_prealarm': get_cell(17, 4),
            'c2h6_alarm': get_cell(17, 5),

            'ch4_hasil': clean_number(get_cell(18, 2)),
            'ch4_kondisi': get_cell(18, 3),
            'ch4_prealarm': get_cell(18, 4),
            'ch4_alarm': get_cell(18, 5),

            'c2h2_hasil': clean_number(get_cell(19, 2)),
            'c2h2_kondisi': get_cell(19, 3),
            'c2h2_prealarm': get_cell(19, 4),
            'c2h2_alarm': get_cell(19, 5),

            'tdcg_hasil': clean_number(get_cell(20, 2)),
            'tdcg_kondisi': get_cell(20, 3),
            'tdcg_prealarm': get_cell(20, 4),
            'tdcg_alarm': get_cell(20, 5),

            # Analisa
            'tdcg_analisa': get_cell(23, 1),
            'keygas_analisa': get_cell(24, 1),
            'roger_ratio_analisa': get_cell(25, 1),
            'duval_triangle_analisa': get_cell(26, 1),

            # Rekomendasi
            'tdcg_rekomendasi': get_cell(28, 1),
            'keygas_rekomendasi': get_cell(29, 1),
            'roger_ratio_rekomendasi': get_cell(30, 1),
            'duval_triangle_rekomendasi': get_cell(31, 1),
        }

        print("✓ Data berhasil dipetakan")
        return data

    except Exception as e:
        print(f"✗ Error: {str(e)}")
        return None


# (Bagian format_laporan_to_whatsapp_single, send_whatsapp_message, dan main tetap sama seperti kode sebelumnya)

def get_current_time_wib():
    tz_wib = timezone(timedelta(hours=7))
    return datetime.now(tz_wib).strftime('%d-%m-%Y %H:%M')

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
    message += f"Frequency: {data['frequency']} | Phase: {data['phase']}\n"
    message += f"Cooling: {data['cooling_type']} | Hubungan: {data['hubungan']}\n"
    message += f"Tahun Pembuatan: {data['tahun_pembuatan']}\n\n"

    message += "*🔬 CONTENT ANALYSIS*\n\n"
    gas_list = [
        ("Hydrogen", "H2", data['h2_hasil'], data['h2_kondisi'], data['h2_prealarm'], data['h2_alarm']),
        ("Water", "H2O", data['h2o_hasil'], data['h2o_kondisi'], data['h2o_prealarm'], data['h2o_alarm']),
        ("Carbon Dioxide", "CO2", data['co2_hasil'], data['co2_kondisi'], data['co2_prealarm'], data['co2_alarm']),
        ("Carbon Monoxide", "CO", data['co_hasil'], data['co_kondisi'], data['co_prealarm'], data['co_alarm']),
        ("Ethylene", "C2H4", data['c2h4_hasil'], data['c2h4_kondisi'], data['c2h4_prealarm'], data['c2h4_alarm']),
        ("Ethane", "C2H6", data['c2h6_hasil'], data['c2h6_kondisi'], data['c2h6_prealarm'], data['c2h6_alarm']),
        ("Methane", "CH4", data['ch4_hasil'], data['ch4_kondisi'], data['ch4_prealarm'], data['ch4_alarm']),
        ("Acetylene", "C2H2", data['c2h2_hasil'], data['c2h2_kondisi'], data['c2h2_prealarm'], data['c2h2_alarm']),
        ("Total Dissolved Gas", "TDCG", data['tdcg_hasil'], data['tdcg_kondisi'], data['tdcg_prealarm'], data['tdcg_alarm']),
    ]

    for nama, simbol, hasil, kondisi, pre, alarm in gas_list:
        message += f"**{nama} ({simbol})**\n"
        message += f"Hasil: {hasil} | Kondisi: {kondisi} | Pre-Alarm: {pre} | Alarm: {alarm}\n\n"

    message += "*🔍 HASIL ANALISA*\n\n"
    message += f"TDCG          : {data['tdcg_analisa']}\n"
    message += f"KEYGAS        : {data['keygas_analisa']}\n"
    message += f"ROGER RATIO   : {data['roger_ratio_analisa']}\n"
    message += f"DUVAL TRIANGLE: {data['duval_triangle_analisa']}\n\n"

    message += "*📌 REKOMENDASI*\n\n"
    message += f"TDCG          : {data['tdcg_rekomendasi']}\n"
    message += f"KEYGAS        : {data['keygas_rekomendasi']}\n"
    message += f"ROGER RATIO   : {data['roger_ratio_rekomendasi']}\n"
    message += f"DUVAL TRIANGLE: {data['duval_triangle_rekomendasi']}\n\n"

    message += "=" * 50 + "\n"
    message += f"Generated: {get_current_time_wib()}\n"

    return message


def send_whatsapp_message(phone_number: str, message: str, is_group: bool = False) -> bool:
    try:
        headers = {"Authorization": FONNTE_API_KEY, "Content-Type": "application/json"}
        payload = {"target": phone_number, "message": message}
        if not is_group:
            payload["countryCode"] = "62"

        response = requests.post(FONNTE_API_URL, headers=headers, json=payload, timeout=15)
        if response.status_code == 200 and response.json().get("status"):
            print(f"✓ Berhasil dikirim ke {'Group' if is_group else 'Personal'}")
            return True
        else:
            print(f"✗ Gagal: {response.text}")
            return False
    except Exception as e:
        print(f"✗ Error kirim WA: {e}")
        return False


# ================== MAIN ==================
if __name__ == "__main__":
    print("="*80)
    print("📋 PROGRAM KIRIM LAPORAN DGA")
    print(f"⏰ Waktu: {get_current_time_wib()}")
    print("="*80)

    data = read_laporan_from_google_sheets(SHEET_ID, SHEET_NAME)
    if not data:
        print("❌ Gagal membaca data!")
        exit(1)

    message = format_laporan_to_whatsapp_single(data)
    print(f"✓ Pesan siap ({len(message)} karakter)")

    phone_numbers = [
        {"number": "6282157905834", "type": "personal"},
        {"number": "120363426691870718@g.us", "type": "group"},
    ]

    success = 0
    for item in phone_numbers:
        print(f"\n📱 Mengirim ke {item['number']} ({item['type']})")
        if send_whatsapp_message(item['number'], message, item['type'] == "group"):
            success += 1
        time.sleep(2)

    print("\n" + "="*80)
    print(f"✅ SELESAI - Berhasil: {success}/{len(phone_numbers)}")
    print("="*80)
