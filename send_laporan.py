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
        print(f"✓ Data diambil: {df.shape[0]} baris x {df.shape[1]} kolom\n")

        # Karena header=None, baris ke-N di Google Sheets = df.iloc[N-1].
        # TIDAK perlu offset tambahan. (Offset 5 sebelumnya adalah penyebab
        # data tergeser/acak.)
        def get_cell(row: int, col: int, default: Any = ""):
            try:
                val = df.iloc[row, col]
                if pd.isna(val) or str(val).strip() == "":
                    return default
                return str(val).strip()
            except Exception:
                return default

        def clean_number(val):
            if not val:
                return "0"
            return str(val).replace("ppm", "").replace("°C", "").replace(",", ".").strip()

        data = {
            # Header (index baris/kolom = posisi asli di sheet - 1)
            'perusahaan': get_cell(2, 1, 'PT GI REJOSO PASURUAN'),       # B3
            'judul_laporan': get_cell(4, 2, 'LAPORAN RUTIN DGA'),        # C5
            'tahun': get_cell(5, 2, '2025'),                             # C6
            'tanggal_pengukuran': get_cell(5, 1),                        # B6
            'tanggal_laporan': get_cell(7, 1),                           # B8
            'unit': get_cell(4, 6),                                      # G5
            'asset': get_cell(5, 6),                                     # G6

            # Spesifikasi Trafo
            'manufacture': get_cell(10, 2),         # C11
            'type': get_cell(11, 2),                # C12
            'kapasitas': get_cell(12, 2),            # C13
            'serial_number': get_cell(13, 2),        # C14
            'tegangan_primer': get_cell(14, 2),       # C15
            'tegangan_sekunder': get_cell(15, 2),     # C16

            'arus_primer': get_cell(10, 4),          # E11
            'arus_sekunder': get_cell(11, 4),         # E12
            'frequency': get_cell(12, 4),             # E13
            'phase': get_cell(13, 4),                 # E14
            'temp_rise': get_cell(14, 4),             # E15
            'berat_oli': get_cell(15, 4),             # E16

            'konfigurasi': get_cell(10, 6),          # G11
            'cooling_type': get_cell(11, 6),          # G12
            'tahun_pembuatan': get_cell(12, 6),       # G13
            'impedansi': get_cell(13, 6),             # G14
            'berat_total': get_cell(14, 6),           # G15

            # Gas Analysis -> kolom: D=Hasil, E=Kondisi2, F=PreAlarm, G=Alarm
            'h2_hasil': clean_number(get_cell(18, 3)),     # baris 19: Hydrogen
            'h2_kondisi': get_cell(18, 4),
            'h2_prealarm': get_cell(18, 5),
            'h2_alarm': get_cell(18, 6),

            'h2o_hasil': clean_number(get_cell(19, 3)),    # baris 20: Water
            'h2o_kondisi': get_cell(19, 4),
            'h2o_prealarm': get_cell(19, 5),
            'h2o_alarm': get_cell(19, 6),

            'co2_hasil': clean_number(get_cell(20, 3)),    # baris 21: Carbon Dioxide
            'co2_kondisi': get_cell(20, 4),
            'co2_prealarm': get_cell(20, 5),
            'co2_alarm': get_cell(20, 6),

            'co_hasil': clean_number(get_cell(21, 3)),     # baris 22: Carbon Monoxide
            'co_kondisi': get_cell(21, 4),
            'co_prealarm': get_cell(21, 5),
            'co_alarm': get_cell(21, 6),

            'c2h4_hasil': clean_number(get_cell(22, 3)),   # baris 23: Ethylene
            'c2h4_kondisi': get_cell(22, 4),
            'c2h4_prealarm': get_cell(22, 5),
            'c2h4_alarm': get_cell(22, 6),

            'c2h6_hasil': clean_number(get_cell(23, 3)),   # baris 24: Ethane
            'c2h6_kondisi': get_cell(23, 4),
            'c2h6_prealarm': get_cell(23, 5),
            'c2h6_alarm': get_cell(23, 6),

            'ch4_hasil': clean_number(get_cell(24, 3)),    # baris 25: Methane
            'ch4_kondisi': get_cell(24, 4),
            'ch4_prealarm': get_cell(24, 5),
            'ch4_alarm': get_cell(24, 6),

            'c2h2_hasil': clean_number(get_cell(25, 3)),   # baris 26: Acetylene
            'c2h2_kondisi': get_cell(25, 4),
            'c2h2_prealarm': get_cell(25, 5),
            'c2h2_alarm': get_cell(25, 6),

            'tdcg_hasil': clean_number(get_cell(26, 3)),   # baris 27: Total Dissolved Gas
            'tdcg_kondisi': get_cell(26, 4),
            'tdcg_prealarm': get_cell(26, 5),
            'tdcg_alarm': get_cell(26, 6),

            # Analisa & Rekomendasi (kolom C)
            'tdcg_analisa': get_cell(31, 2),                  # baris 32
            'keygas_analisa': get_cell(32, 2),                # baris 33
            'roger_ratio_analisa': get_cell(33, 2),           # baris 34
            'duval_triangle_analisa': get_cell(34, 2),        # baris 35

            'tdcg_rekomendasi': get_cell(36, 2),              # baris 37
            'keygas_rekomendasi': get_cell(37, 2),            # baris 38
            'roger_ratio_rekomendasi': get_cell(38, 2),       # baris 39
            'duval_triangle_rekomendasi': get_cell(39, 2),    # baris 40
        }

        # Debug mapping
        print("🔍 HASIL MAPPING (beberapa contoh):")
        print(f"Manufacture : {data['manufacture']}")
        print(f"Kapasitas : {data['kapasitas']}")
        print(f"H2 Hasil : {data['h2_hasil']}")
        print(f"CO2 Hasil : {data['co2_hasil']}")
        print(f"TDCG Hasil : {data['tdcg_hasil']}")
        print(f"Key Gas Analisa : {data['keygas_analisa']}")

        return data

    except Exception as e:
        print(f"✗ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return None


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
    message += f"Type: {data.get('type', '')}\n"
    message += f"Kapasitas: {data['kapasitas']}\n"
    message += f"Serial Number: {data['serial_number']}\n"
    message += f"Tegangan Primer: {data['tegangan_primer']}\n"
    message += f"Tegangan Sekunder: {data['tegangan_sekunder']}\n"
    message += f"Frequency: {data['frequency']} | Phase: {data['phase']}\n"
    message += f"Cooling: {data['cooling_type']} | Konfigurasi: {data['konfigurasi']}\n"
    message += f"Tahun Pembuatan: {data['tahun_pembuatan']}\n\n"

    message += "*🔬 CONTENT ANALYSIS*\n\n"
    gas_list = [
        ("Hydrogen", "H2", data['h2_hasil'], data['h2_kondisi'], data['h2_prealarm'], data['h2_alarm']),
        ("Air", "H2O", data['h2o_hasil'], data['h2o_kondisi'], data['h2o_prealarm'], data['h2o_alarm']),
        ("CO2", "CO2", data['co2_hasil'], data['co2_kondisi'], data['co2_prealarm'], data['co2_alarm']),
        ("CO", "CO", data['co_hasil'], data['co_kondisi'], data['co_prealarm'], data['co_alarm']),
        ("Etilen", "C2H4", data['c2h4_hasil'], data['c2h4_kondisi'], data['c2h4_prealarm'], data['c2h4_alarm']),
        ("Etana", "C2H6", data['c2h6_hasil'], data['c2h6_kondisi'], data['c2h6_prealarm'], data['c2h6_alarm']),
        ("Metana", "CH4", data['ch4_hasil'], data['ch4_kondisi'], data['ch4_prealarm'], data['ch4_alarm']),
        ("Asetilen", "C2H2", data['c2h2_hasil'], data['c2h2_kondisi'], data['c2h2_prealarm'], data['c2h2_alarm']),
        ("Total Dissolved Combustible Gas", "TDCG", data['tdcg_hasil'], data['tdcg_kondisi'], data['tdcg_prealarm'], data['tdcg_alarm']),
    ]
    for nama, simbol, hasil, kondisi, pre, alarm in gas_list:
        message += f"**{nama} ({simbol})**\n"
        message += f"Hasil: {hasil} | Kondisi: {kondisi} | PreAlarm: {pre} | Alarm: {alarm}\n\n"

    message += "*🔍 HASIL ANALISA*\n\n"
    message += f"TDCG : {data['tdcg_analisa']}\n"
    message += f"Key Gas : {data['keygas_analisa']}\n"
    message += f"Roger Ratio : {data['roger_ratio_analisa']}\n"
    message += f"Duval Triangle: {data['duval_triangle_analisa']}\n\n"

    message += "*📌 REKOMENDASI*\n\n"
    message += f"TDCG : {data['tdcg_rekomendasi']}\n"
    message += f"Key Gas : {data['keygas_rekomendasi']}\n"
    message += f"Roger Ratio : {data['roger_ratio_rekomendasi']}\n"
    message += f"Duval Triangle: {data['duval_triangle_rekomendasi']}\n\n"

    message += "=" * 50 + "\n"
    message += f"Generated: {get_current_time_wib()}\n"
    return message


# ================== SEND WHATSAPP ==================
def send_whatsapp_message(phone_number: str, message: str, is_group: bool = False) -> bool:
    try:
        headers = {
            "Authorization": FONNTE_API_KEY,
            "Content-Type": "application/json"
        }
        payload = {
            "target": phone_number,
            "message": message,
            "countryCode": "62" if not is_group else None
        }
        if is_group:
            payload.pop("countryCode", None)

        response = requests.post(FONNTE_API_URL, headers=headers, json=payload, timeout=15)
        if response.status_code == 200:
            result = response.json()
            print(f"✓ Berhasil ke {'Group' if is_group else 'Personal'}: {phone_number}")
            return True
        else:
            print(f"✗ Gagal ({response.status_code}): {response.text}")
            return False
    except Exception as e:
        print(f"✗ Exception: {e}")
        return False


# ================== MAIN ==================
if __name__ == "__main__":
    print("=" * 80)
    print("📋 PROGRAM KIRIM LAPORAN DGA")
    print(f"⏰ Waktu: {get_current_time_wib()}")
    print("=" * 80)

    data = read_laporan_from_google_sheets(SHEET_ID, SHEET_NAME)
    if not data:
        print("❌ Gagal membaca data sheet!")
        exit(1)

    message = format_laporan_to_whatsapp_single(data)
    print(f"✓ Pesan siap ({len(message)} karakter)")

    phone_numbers = [
        {"number": "6282157905834", "type": "personal"},
    ]

    success = 0
    for item in phone_numbers:
        print(f"\n📱 Mengirim ke {item['number']} ({item['type']})")
        if send_whatsapp_message(item['number'], message, item['type'] == "group"):
            success += 1
        time.sleep(2)

    print("\n" + "=" * 80)
    print(f"✅ SELESAI - Berhasil: {success}/{len(phone_numbers)}")
    print("=" * 80)
