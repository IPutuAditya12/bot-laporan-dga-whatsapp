import pandas as pd
import requests
from typing import Dict, Any, Optional, Tuple
import time
from datetime import datetime, timedelta, timezone
import os

# ================== KONFIGURASI ==================
FONNTE_API_KEY = os.environ.get("FONNTE_API_KEY", "")
FONNTE_API_URL = "https://api.fonnte.com/send"
SHEET_ID = "1ZlPaZQwz89UkjABHgxqmcYnTklB5yK2V"
SHEET_NAME = "Laporan"


# ================== HELPER: PENCARIAN BERBASIS LABEL ==================
def clean_cell_str(val) -> str:
    """Konversi nilai sel ke string, hilangkan '.0' di belakang angka bulat
    (pandas suka membaca '3' sebagai 3.0 kalau kolomnya numerik)."""
    s = str(val).strip()
    if s.endswith(".0"):
        try:
            f = float(s)
            if f == int(f):
                return str(int(f))
        except ValueError:
            pass
    return s


def norm(s: Any) -> str:
    """Normalisasi teks untuk pencocokan label (hilangkan semua whitespace tersembunyi, lowercase)."""
    if s is None:
        return ""
    # Ganti non-breaking space, tab, dan karakter whitespace unicode lainnya
    s = str(s).replace('\xa0', ' ').replace('\t', ' ')
    # Hilangkan semua karakter non-printable
    s = ''.join(c if c.isprintable() else ' ' for c in s)
    s = ' '.join(s.split())  # normalisasi whitespace ganda
    if s.endswith(":"):
        s = s[:-1].strip()
    return s.lower()


def find_label(df: pd.DataFrame, label: str, row_start: int = 0, row_end: Optional[int] = None,
               exact: bool = True) -> Optional[Tuple[int, int]]:
    """
    Cari sel yang isinya cocok dengan `label` di dalam rentang baris [row_start, row_end).
    Mengembalikan (row, col) dari kemunculan PERTAMA, atau None jika tidak ditemukan.
    Ini membuat script tahan terhadap pergeseran baris -- selama labelnya masih ada
    di sheet, datanya akan tetap ketemu walau posisinya pindah.
    """
    target = norm(label)
    if row_end is None:
        row_end = df.shape[0]
    for r in range(row_start, min(row_end, df.shape[0])):
        for c in range(df.shape[1]):
            cell = norm(df.iat[r, c])
            if not cell:
                continue
            if (exact and cell == target) or (not exact and target in cell):
                return (r, c)
    return None


def find_label_with_value(df: pd.DataFrame, label: str, row_start: int = 0,
                          row_end: Optional[int] = None, max_scan_cols: int = 4) -> Optional[Tuple[int, int]]:
    """
    Seperti find_label, tapi HANYA mengembalikan posisi jika minimal ada satu sel
    tidak kosong di sebelah kanannya (offset 1..max_scan_cols).
    Berguna saat label muncul lebih dari sekali di sheet tapi hanya satu yang
    punya nilai nyata di sampingnya.
    """
    target = norm(label)
    if row_end is None:
        row_end = df.shape[0]
    for r in range(row_start, min(row_end, df.shape[0])):
        for c in range(df.shape[1]):
            cell = norm(df.iat[r, c])
            if not cell or cell != target:
                continue
            # Cek apakah ada nilai non-kosong di sebelah kanan
            for o in range(1, max_scan_cols + 1):
                c2 = c + o
                if c2 >= df.shape[1]:
                    break
                val = df.iat[r, c2]
                if pd.notna(val) and str(val).strip().replace(' ', '') != '':
                    return (r, c)
    return None


def value_right(df: pd.DataFrame, pos: Optional[Tuple[int, int]], offset: int = 1, default: str = "",
                 scan: bool = False, max_scan: int = 3) -> str:
    """Ambil nilai di sebelah kanan posisi label.

    Jika scan=True dan sel pada `offset` kosong, lanjut mencari sel tidak-kosong
    berikutnya ke kanan (sampai max_scan kolom). Ini untuk berjaga-jaga kalau ada
    kolom tersembunyi/merge yang membuat posisi nilai sebenarnya geser sedikit
    dari offset yang diharapkan.
    """
    if pos is None:
        return default
    r, c = pos
    tries = range(offset, offset + max_scan) if scan else [offset]
    for o in tries:
        c2 = c + o
        if c2 >= df.shape[1]:
            continue
        val = df.iat[r, c2]
        if pd.notna(val) and str(val).strip() != "":
            return clean_cell_str(val)
    return default


def value_below(df: pd.DataFrame, pos: Optional[Tuple[int, int]], offset: int = 1, default: str = "") -> str:
    """Ambil nilai di bawah posisi label (dipakai untuk tanggal pengukuran/laporan)."""
    if pos is None:
        return default
    r, c = pos
    r2 = r + offset
    if r2 >= df.shape[0]:
        return default
    val = df.iat[r2, c]
    if pd.isna(val) or str(val).strip() == "":
        return default
    return str(val).strip()


def first_nonempty_above(df: pd.DataFrame, before_row: int, col_candidates=(1, 2)) -> str:
    """Cari teks pertama yang tidak kosong di baris-baris sebelum `before_row` (dipakai untuk nama perusahaan,
    karena cell ini biasanya tidak punya label sendiri)."""
    for r in range(before_row - 1, -1, -1):
        for c in col_candidates:
            val = df.iat[r, c] if c < df.shape[1] else None
            if pd.notna(val) and str(val).strip() != "":
                return str(val).strip()
    return ""


def clean_number(val):
    if not val:
        return "0"
    return str(val).replace("ppm", "").replace("°C", "").replace(",", ".").strip()


def find_cell_containing(df: pd.DataFrame, keyword: str) -> str:
    """Cari sel pertama yang isinya mengandung kata kunci (case-insensitive), kembalikan nilainya."""
    kw = keyword.lower()
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            val = df.iat[r, c]
            if pd.isna(val):
                continue
            s = str(val).strip()
            if kw in s.lower():
                return s
    return ""


# ================== AMBIL & PARSE DATA ==================
def read_laporan_from_google_sheets(sheet_id: str, sheet_name: str) -> Optional[Dict]:
    try:
        # GID tab "Laporan" = 688001833 (dari URL sheet)
        csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid=688001833"
        print("Mengambil data dari Google Sheets...")
        df = pd.read_csv(csv_url, header=None)
        print(f"Data diambil: {df.shape[0]} baris x {df.shape[1]} kolom\n")

        # ---------- Anchor / batas section (dicari berdasarkan teks, BUKAN nomor baris tetap) ----------
        pos_tgl_ukur_label = find_label(df, "Date of Measurement")
        pos_tgl_lapor_label = find_label(df, "Date of Report")
        pos_unit = find_label(df, "Unit")
        pos_asset = find_label(df, "Asset")  # hati-hati: "Asset Description" mengandung kata "Asset" juga
        pos_spek = find_label(df, "a. Spesifikasi", exact=False)
        pos_content = find_label(df, "CONTENT ANALYSIS", exact=False)
        pos_analisa_header = find_label(df, "c. Analisa", exact=False)
        pos_rekom_header = find_label(df, "d. Rekomendasi", exact=False)
        pos_tool = find_label(df, "Tool", exact=False)

        analisa_start = pos_analisa_header[0] if pos_analisa_header else 0
        rekom_start = pos_rekom_header[0] if pos_rekom_header else df.shape[0]
        rekom_end = pos_tool[0] if pos_tool else df.shape[0]

        # ---------- Header umum ----------
        judul_laporan = value_right(df, pos_tgl_ukur_label, offset=1, default="LAPORAN RUTIN DGA")
        tahun = value_below(df, pos_tgl_ukur_label, offset=1, default="")  # baris di bawah judul = tahun (kolom C)
        # tahun sebenarnya ada 1 kolom ke kanan dari posisi label, 1 baris di bawah -> ambil manual:
        if pos_tgl_ukur_label:
            r0, c0 = pos_tgl_ukur_label
            tahun = df.iat[r0 + 1, c0 + 1] if r0 + 1 < df.shape[0] and c0 + 1 < df.shape[1] else ""
            tahun = "" if pd.isna(tahun) else str(tahun).strip()

        data = {
            'perusahaan': "GI REJOSO INDONESIA",
            'judul_laporan': judul_laporan,
            'tahun': tahun or "2025",
            'tanggal_pengukuran': value_below(df, pos_tgl_ukur_label, offset=1),
            'tanggal_laporan': value_below(df, pos_tgl_lapor_label, offset=1),
            'unit': value_right(df, pos_unit, offset=1),
            'asset': value_right(df, pos_asset, offset=1),

            # ---------- Spesifikasi trafo (label : nilai = sel sebelah kanan, scan jaga-jaga jika kolom geser) ----------
            'manufacture': value_right(df, find_label(df, "Manufacture", pos_spek[0] if pos_spek else 0), scan=True),
            'type': value_right(df, find_label(df, "Type", pos_spek[0] if pos_spek else 0), scan=True),
            'kapasitas': value_right(df, find_label(df, "Kapasitas", pos_spek[0] if pos_spek else 0), scan=True),
            'serial_number': value_right(df, find_label(df, "Serial No", pos_spek[0] if pos_spek else 0, exact=False), scan=True),
            'tegangan_primer': value_right(df, find_label(df, "Tegangan Primer", pos_spek[0] if pos_spek else 0), scan=True),
            'tegangan_sekunder': value_right(df, find_label(df, "Tegangan Sekunder", pos_spek[0] if pos_spek else 0), scan=True),

            'arus_primer': value_right(df, find_label(df, "Arus Primer", pos_spek[0] if pos_spek else 0), scan=True),
            'arus_sekunder': value_right(df, find_label(df, "Arus Sekunder", pos_spek[0] if pos_spek else 0), scan=True),
            'frequency': find_cell_containing(df, "Hz"),
            'phase': value_right(df, find_label(df, "Phase", pos_spek[0] if pos_spek else 0), scan=True),
            'temp_rise': value_right(df, find_label(df, "Temp Rise", pos_spek[0] if pos_spek else 0), scan=True),
            'berat_oli': find_cell_containing(df, "ltr"),

            'konfigurasi': value_right(df, find_label(df, "Konfigurasi", pos_spek[0] if pos_spek else 0), scan=True),
            'cooling_type': value_right(df, find_label(df, "Cooling Type", pos_spek[0] if pos_spek else 0), scan=True),
            'tahun_pembuatan': value_right(df, find_label(df, "Tahun Pembuatan", pos_spek[0] if pos_spek else 0), scan=True),
            'impedansi': value_right(df, find_label(df, "Impedansi", pos_spek[0] if pos_spek else 0, exact=False), scan=True),
            'berat_total': value_right(df, find_label(df, "Berat Total", pos_spek[0] if pos_spek else 0), scan=True),
        }

        # ---------- Tabel gas (CONTENT ANALYSIS) ----------
        # Untuk tiap gas: cari simbolnya (H2, CO2, dst), lalu Hasil/Kondisi/PreAlarm/Alarm
        # ada di kolom-kolom berikutnya secara berurutan, berapapun pergeseran barisnya.
        content_row_start = pos_content[0] if pos_content else 0
        gas_symbols = [
            ("h2", "H2"), ("h2o", "H2O"), ("co2", "CO2"), ("co", "CO"),
            ("c2h4", "C2H4"), ("c2h6", "C2H6"), ("ch4", "CH4"),
            ("c2h2", "C2H2"), ("tdcg", "TDCG"),
        ]
        for key, symbol in gas_symbols:
            pos = find_label(df, symbol, row_start=content_row_start)
            data[f'{key}_hasil'] = clean_number(value_right(df, pos, offset=1))
            data[f'{key}_kondisi'] = value_right(df, pos, offset=2)
            data[f'{key}_prealarm'] = value_right(df, pos, offset=3)
            data[f'{key}_alarm'] = value_right(df, pos, offset=4)

        # ---------- Analisa & Rekomendasi ----------
        # Label TDCG/KEYGAS/ROGER RATIO/DUVAL TRIANGLE muncul 2x di sheet (analisa & rekomendasi),
        # jadi pencarian harus dibatasi per-section (row_start/row_end) supaya tidak ketuker.
        analisa_labels = {
            'tdcg_analisa': "TDCG",
            'keygas_analisa': "KEYGAS",
            'roger_ratio_analisa': "ROGER RATIO",
            'duval_triangle_analisa': "DUVAL TRIANGLE",
        }
        for key, label in analisa_labels.items():
            pos = find_label(df, label, row_start=analisa_start, row_end=rekom_start)
            data[key] = value_right(df, pos)

        rekom_labels = {
            'tdcg_rekomendasi': "TDCG",
            'keygas_rekomendasi': "KEYGAS",
            'roger_ratio_rekomendasi': "ROGER RATIO",
            'duval_triangle_rekomendasi': "DUVAL TRIANGLE",
        }
        for key, label in rekom_labels.items():
            pos = find_label(df, label, row_start=rekom_start, row_end=rekom_end)
            data[key] = value_right(df, pos)

        # ===== DIAGNOSTIC DUMP =====
        print(f"\nTotal kolom: {df.shape[1]}")
        print("\n=== SEMUA KOLOM row 7 (Frequency) ===")
        for c in range(df.shape[1]):
            val = df.iat[7, c]
            print(f"  col{c}: {repr(val)}")
        print("\n=== SEMUA KOLOM row 10 (Berat Oli) ===")
        for c in range(df.shape[1]):
            val = df.iat[10, c]
            print(f"  col{c}: {repr(val)}")
        # =====
        print("\n=== RAW CSV ROWS 0-20 ===")
        for r in range(min(21, df.shape[0])):
            row_vals = [str(df.iat[r, c])[:15] if not pd.isna(df.iat[r, c]) else "." for c in range(min(8, df.shape[1]))]
            print(f"  row{r:02d}: {row_vals}")
        print("=========================\n")
        # ===== END DIAGNOSTIC =====

        # Debug khusus untuk Frequency & Berat Oli
        pos_freq_dbg = find_label_with_value(df, "Frequency", pos_spek[0] if pos_spek else 0)
        pos_boli_dbg = find_label_with_value(df, "Berat Oli", pos_spek[0] if pos_spek else 0)
        print(f"[DEBUG] pos 'Frequency' : {pos_freq_dbg}")
        if pos_freq_dbg:
            r2, c2 = pos_freq_dbg
            print(f"  -> label=[{r2},{c2}] = {repr(df.iat[r2,c2])}")
            print(f"  -> nilai=[{r2},{c2+1}] = {repr(df.iat[r2,c2+1]) if c2+1 < df.shape[1] else 'OUT'}")
        print(f"[DEBUG] pos 'Berat Oli' : {pos_boli_dbg}")
        if pos_boli_dbg:
            r2, c2 = pos_boli_dbg
            print(f"  -> label=[{r2},{c2}] = {repr(df.iat[r2,c2])}")
            print(f"  -> nilai=[{r2},{c2+1}] = {repr(df.iat[r2,c2+1]) if c2+1 < df.shape[1] else 'OUT'}")

        # Debug mapping
        print("\nHASIL MAPPING:")
        print(f"Perusahaan      : {data['perusahaan']}")
        print(f"Manufacture     : {data['manufacture']}")
        print(f"Kapasitas       : {data['kapasitas']}")
        print(f"Frequency       : {data['frequency']}")
        print(f"Berat Oli       : {data['berat_oli']}")
        print(f"H2 Hasil        : {data['h2_hasil']}")
        print(f"CO2 Hasil       : {data['co2_hasil']}")
        print(f"TDCG Hasil      : {data['tdcg_hasil']}")
        print(f"TDCG Analisa    : {data['tdcg_analisa']}")
        print(f"TDCG Rekomendasi: {data['tdcg_rekomendasi']}")
        print(f"Key Gas Analisa : {data['keygas_analisa']}")

        return data

    except Exception as e:
        print(f"Error: {str(e)}")
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
    message += f"Frequency: {data['frequency']}\n"
    message += f"Phase: {data['phase']}\n"
    message += f"Cooling: {data['cooling_type']}\n" 
    message += f"Konfigurasi: {data['konfigurasi']}\n"
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
        message += f"Hasil Pengujian: {hasil} | Kondisi 2: {kondisi} | Pre-ALARM: {pre} | ALARM: {alarm}\n\n"

    message += "*🔍 HASIL ANALISA*\n\n"
    message += f"TDCG : {data['tdcg_analisa']}\n"
    message += f"KEYGAS : {data['keygas_analisa']}\n"
    message += f"ROGER RATIO : {data['roger_ratio_analisa']}\n"
    message += f"DUVAL TRIANGLE: {data['duval_triangle_analisa']}\n\n"

    message += "*📌 REKOMENDASI*\n\n"
    message += f"TDCG : {data['tdcg_rekomendasi']}\n"
    message += f"KEYGAS : {data['keygas_rekomendasi']}\n"
    message += f"ROGER RATIO : {data['roger_ratio_rekomendasi']}\n"
    message += f"DUVAL TRIANGLE: {data['duval_triangle_rekomendasi']}\n\n"

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
            print(f"Berhasil ke {'Group' if is_group else 'Personal'}: {phone_number}")
            return True
        else:
            print(f"Gagal ({response.status_code}): {response.text}")
            return False
    except Exception as e:
        print(f"Exception: {e}")
        return False


# ================== MAIN ==================
if __name__ == "__main__":
    print("=" * 80)
    print("PROGRAM KIRIM LAPORAN DGA")
    print(f"Waktu: {get_current_time_wib()}")
    print("=" * 80)

    data = read_laporan_from_google_sheets(SHEET_ID, SHEET_NAME)
    if not data:
        print("Gagal membaca data sheet!")
        exit(1)

    message = format_laporan_to_whatsapp_single(data)
    print(f"Pesan siap ({len(message)} karakter)")

    phone_numbers = [
        {"number": "6282157905834", "type": "personal"},
        {"number": "120363426691870718@g.us", "type": "group"},
    ]

    success = 0
    for item in phone_numbers:
        print(f"\nMengirim ke {item['number']} ({item['type']})")
        if send_whatsapp_message(item['number'], message, item['type'] == "group"):
            success += 1
        time.sleep(2)

    print("\n" + "=" * 80)
    print(f"SELESAI - Berhasil: {success}/{len(phone_numbers)}")
    print("=" * 80)
