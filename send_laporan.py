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
        df = pd.read_csv(csv_url, header=None)

        def get_cell(row: int, col: int, default: Any = ""):
            try:
                val = df.iloc[row, col]
                if pd.isna(val) or str(val).strip() in ["", "nan", "NaN"]:
                    return default
                return str(val).strip()
            except:
                return default

        def clean_number(val):
            if not val:
                return "0"
            return str(val).replace("ppm", "").replace("ltr", "").replace(",", ".").strip()

        data = {
            'perusahaan': get_cell(2, 1),
            'judul_laporan': get_cell(4, 3),
            'tahun': get_cell(5, 2),
            'tanggal_pengukuran': get_cell(5, 1),
            'tanggal_laporan': get_cell(7, 1),
            'unit': get_cell(4, 6),
            'asset': get_cell(5, 5),

            'manufacture': get_cell(9, 2),
            'kapasitas': get_cell(11, 2),
            'serial_number': get_cell(12, 1),
            'tegangan_primer': get_cell(13, 2),
            'tegangan_sekunder': get_cell(14, 2),
            'frequency': get_cell(11, 4),
            'phase': get_cell(12, 4),
            'cooling_type': get_cell(10, 6),
            'hubungan': get_cell(9, 6),
            'tahun_pembuatan': get_cell(11, 6),
            'berat_oli': get_cell(14, 4),

            # === GAS ANALYSIS - Mapping akurat berdasarkan CSV ===
            'h2_hasil': clean_number(get_cell(16, 3)),
            'h2_kondisi': get_cell(16, 4),
            'h2_prealarm': get_cell(16, 5),
            'h2_alarm': get_cell(16, 6),

            'h2o_hasil': clean_number(get_cell(17, 3)),
            'h2o_kondisi': get_cell(17, 4),
            'h2o_prealarm': get_cell(17, 5),
            'h2o_alarm': get_cell(17, 6),

            'co2_hasil': clean_number(get_cell(18, 3)),
            'co2_kondisi': get_cell(18, 4),
            'co2_prealarm': get_cell(18, 5),
            'co2_alarm': get_cell(18, 6),

            'co_hasil': clean_number(get_cell(19, 3)),
            'co_kondisi': get_cell(19, 4),
            'co_prealarm': get_cell(19, 5),
            'co_alarm': get_cell(19, 6),

            'c2h4_hasil': clean_number(get_cell(20, 3)),
            'c2h4_kondisi': get_cell(20, 4),
            'c2h4_prealarm': get_cell(20, 5),
            'c2h4_alarm': get_cell(20, 6),

            'c2h6_hasil': clean_number(get_cell(21, 3)),
            'c2h6_kondisi': get_cell(21, 4),
            'c2h6_prealarm': get_cell(21, 5),
            'c2h6_alarm': get_cell(21, 6),

            'ch4_hasil': clean_number(get_cell(22, 3)),
            'ch4_kondisi': get_cell(22, 4),
            'ch4_prealarm': get_cell(22, 5),
            'ch4_alarm': get_cell(22, 6),

            'c2h2_hasil': clean_number(get_cell(23, 3)),
            'c2h2_kondisi': get_cell(23, 4),
            'c2h2_prealarm': get_cell(23, 5),
            'c2h2_alarm': get_cell(23, 6),

            'tdcg_hasil': clean_number(get_cell(24, 3)),
            'tdcg_kondisi': get_cell(24, 4),
            'tdcg_prealarm': get_cell(24, 5),
            'tdcg_alarm': get_cell(24, 6),

            # Analisa & Rekomendasi
            'tdcg_analisa': get_cell(27, 1),
            'keygas_analisa': get_cell(28, 1),
            'roger_ratio_analisa': get_cell(29, 1),
            'duval_triangle_analisa': get_cell(30, 1),

            'tdcg_rekomendasi': get_cell(32, 1),
            'keygas_rekomendasi': get_cell(33, 1),
            'roger_ratio_rekomendasi': get_cell(34, 1),
            'duval_triangle_rekomendasi': get_cell(35, 1),
        }

        return data

    except Exception as e:
        print(f"Error: {e}")
        return None


# ================== FORMAT MESSAGE ==================
def get_current_time_wib():
    tz_wib = timezone(timedelta(hours=7))
    return datetime.now(tz_wib).strftime('%d-%m-%Y %H:%M')

def format_laporan_to_whatsapp_single(data: Dict) -> str:
    message = f"*🏢 {data['perusahaan']}*\n"
    message += f"*📊 {data['judul_laporan']} - {data['tahun']}*\n"
    message += "=" * 55 + "\n\n"

    message += "*📅 INFORMASI DASAR*\n"
    message += f"Tanggal Pengukuran: {data['tanggal_pengukuran']}\n"
    message += f"Tanggal Laporan: {data['tanggal_laporan']}\n"
    message += f"Unit: {data['unit']}\n\n"

    message += "*⚙️ SPESIFIKASI TRAFO*\n"
    message += f"Manufacture: {data['manufacture']}\n"
    message += f"Kapasitas: {data['kapasitas']}\n"
    message += f"Serial Number: {data['serial_number']}\n"
    message += f"Tegangan Primer: {data['tegangan_primer']}\n"
    message += f"Tegangan Sekunder: {data['tegangan_sekunder']}\n"
    message += f"Frequency: {data['frequency']} | Phase: {data['phase']}\n"
    message += f"Cooling: {data['cooling_type']} | Hubungan: {data['hubungan']}\n"
    message += f"Tahun Pembuatan: {data['tahun_pembuatan']}\n"
    message += f"Berat Oli: {data['berat_oli']}\n\n"

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

    message += "=" * 55 + "\n"
    message += f"Generated: {get_current_time_wib()}\n"
    return message


# ================== SEND WA ==================
def send_whatsapp_message(phone_number: str, message: str, is_group: bool = False) -> bool:
    try:
        headers = {"Authorization": FONNTE_API_KEY, "Content-Type": "application/json"}
        payload = {"target": phone_number, "message": message}
        if not is_group:
            payload["countryCode"] = "62"

        r = requests.post(FONNTE_API_URL, headers=headers, json=payload, timeout=12)
        return r.status_code == 200 and r.json().get("status") is True
    except:
        return False


# ================== MAIN ==================
if __name__ == "__main__":
    print("="*70)
    print("📋 PROGRAM KIRIM LAPORAN DGA")
    print(f"⏰ Waktu: {get_current_time_wib()}")
    print("="*70)

    data = read_laporan_from_google_sheets(SHEET_ID, SHEET_NAME)
    if not data:
        print("❌ Gagal membaca data")
        exit(1)

    message = format_laporan_to_whatsapp_single(data)
    print(f"✓ Pesan siap ({len(message)} karakter)")

    phone_numbers = [
        {"number": "6282157905834", "type": "personal"},
        {"number": "120363426691870718@g.us", "type": "group"},
    ]

    success = 0
    for item in phone_numbers:
        print(f"\n📱 Mengirim ke {item['number']}...")
        if send_whatsapp_message(item['number'], message, item['type'] == "group"):
            success += 1
        time.sleep(2)

    print(f"\n✅ SELESAI - Berhasil: {success}/2")
