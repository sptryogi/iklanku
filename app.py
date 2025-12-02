import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
import xlsxwriter

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="IklanKu (Laporan Harian)", layout="wide")

# --- FUNGSI UTAMA (HELPER) ---

def clean_nama_iklan(text):
    if not isinstance(text, str):
        return str(text)
    # Hapus [angka] di belakang, misal "Produk A [26]" -> "Produk A"
    return re.sub(r'\s*\[\d+\]\s*$', '', text).strip()

def extract_time_hour(dt):
    try:
        # Asumsi format timestamp pandas
        return dt.hour
    except:
        return 0

def extract_eksemplar(variasi_text):
    if not isinstance(variasi_text, str):
        return 1
    # Cari angka dalam variasi, misal "PAKET ISI 3" -> 3
    match = re.search(r'(\d+)', variasi_text)
    if match:
        return int(match.group(1))
    return 1 # Default jika tidak ada angka

def clean_variasi(text):
    if not isinstance(text, str) or pd.isna(text) or text == '':
        return ''
    # Ambil value di belakang koma, uppercase
    if ',' in text:
        parts = text.split(',')
        return parts[-1].strip().upper()
    return text.strip().upper()

# --- LOGIKA PROSES DATA ---

def process_data(store_name, file_order, file_iklan, file_seller):
    # 1. LOAD DATA
    df_order = pd.read_excel(file_order)
    # df_iklan_raw = pd.read_csv(file_iklan, header=None) # Load tanpa header dulu untuk skip baris
    df_iklan_raw = pd.read_csv(file_iklan, header=None, sep=None, engine='python')
    df_seller = pd.read_csv(file_seller)

    # 2. PRE-PROCESS ORDER-ALL
    # Filter Status Pesanan != Batal
    if 'Status Pesanan' in df_order.columns:
        df_order = df_order[df_order['Status Pesanan'] != 'Batal'].copy()
    
    # Konversi kolom waktu
    if 'Waktu Pesanan Dibuat' in df_order.columns:
        df_order['Waktu Pesanan Dibuat'] = pd.to_datetime(df_order['Waktu Pesanan Dibuat'])
        df_order['Jam'] = df_order['Waktu Pesanan Dibuat'].dt.hour
        # Ambil tanggal untuk header laporan
        report_date = df_order['Waktu Pesanan Dibuat'].dt.strftime('%A, %d-%m-%Y').iloc[0] if not df_order.empty else "TANGGAL TIDAK DIKETAHUI"
    else:
        st.error("Kolom 'Waktu Pesanan Dibuat' tidak ditemukan di Order-all")
        return None

    if 'Total Harga Produk' in df_order.columns:
        df_order['Total Harga Produk'] = pd.to_numeric(df_order['Total Harga Produk'], errors='coerce').fillna(0)

    # 3. PRE-PROCESS IKLAN (Sheet 'Iklan klik')
    # Hapus 7 baris pertama (index 0-6), baris ke-8 (index 7) jadi header
    new_header = df_iklan_raw.iloc[7]
    df_iklan = df_iklan_raw[8:].copy()
    df_iklan.columns = new_header
    
    # Bersihkan Nama Iklan
    if 'Nama Iklan' in df_iklan.columns:
        df_iklan['Nama Iklan'] = df_iklan['Nama Iklan'].apply(clean_nama_iklan)
        # Hapus Duplikat Nama Iklan
        df_iklan = df_iklan.drop_duplicates(subset=['Nama Iklan'])
    
    # Konversi kolom numerik di iklan
    cols_to_num = ['Dilihat', 'Jumlah Klik', 'Omzet Penjualan', 'Biaya']
    for col in cols_to_num:
        if col in df_iklan.columns:
            # Hapus simbol mata uang atau pemisah ribuan jika ada
            # df_iklan[col] = pd.to_numeric(df_iklan[col].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce').fillna(0)
            df_iklan[col] = df_iklan[col].astype(str).str.replace('Rp', '', regex=False).str.strip().str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
            df_iklan[col] = pd.to_numeric(df_iklan[col], errors='coerce').fillna(0)

    # 4. KATEGORISASI DATA (AFFILIATE, IKLAN, ORGANIK) & HIGHLIGHTING
    # Setup list untuk tracking
    list_affiliate_ids = df_seller['Kode Pesanan'].astype(str).tolist() if 'Kode Pesanan' in df_seller.columns else []
    list_iklan_names = df_iklan['Nama Iklan'].tolist() if 'Nama Iklan' in df_iklan.columns else []

    # Buat kolom helper di df_order
    df_order['is_affiliate'] = df_order['No. Pesanan'].astype(str).isin(list_affiliate_ids)
    df_order['is_iklan_product'] = df_order['Nama Produk'].apply(lambda x: clean_nama_iklan(x) in list_iklan_names)
    
    # Prioritas: Affiliate > Iklan (match product) > Organik
    # Namun prompt meminta: "Order all yg termasuk seller conversion (Affiliate)" dan "Diluar Seller Conversion dan Diluar Nama Iklan (Organik)"
    # Maka Sisanya (Diluar Seller tapi ADA di Nama Iklan) adalah Pesanan Iklan.
    
    df_affiliate = df_order[df_order['is_affiliate']].copy()
    df_organic = df_order[(~df_order['is_affiliate']) & (~df_order['is_iklan_product'])].copy()
    df_ads_orders = df_order[(~df_order['is_affiliate']) & (df_order['is_iklan_product'])].copy()

    # --- MEMBUAT DATA UNTUK LAPORAN ---

    # A. TABEL PESANAN IKLAN (Jam 0-23)
    # Prompt: "diambil dari order all dilihat di rentang jam... sum No Pesanan... Kuantitas... Omzet"
    # Menggunakan df_ads_orders (Pesanan yg berasal dari produk yg diiklanin)
    
    hours_df = pd.DataFrame({'Jam': range(24)})
    
    def agg_by_hour(df_source):
        if df_source.empty:
            return pd.DataFrame({'Jam': range(24), 'PESANAN': 0, 'KUANTITAS': 0, 'OMZET PENJUALAN': 0})
        
        # Hitung Pesanan (Unique No. Pesanan)
        grp_pesanan = df_source.groupby('Jam')['No. Pesanan'].nunique().reset_index(name='PESANAN')
        # Hitung Kuantitas & Omzet (Sum)
        grp_metrics = df_source.groupby('Jam')[['Jumlah', 'Total Harga Produk']].sum().reset_index()
        grp_metrics.rename(columns={'Jumlah': 'KUANTITAS', 'Total Harga Produk': 'OMZET PENJUALAN'}, inplace=True)
        
        merged = hours_df.merge(grp_pesanan, on='Jam', how='left').merge(grp_metrics, on='Jam', how='left')
        return merged.fillna(0)

    # Tabel 1 Data (Pesanan Iklan)
    tbl_iklan_data = agg_by_hour(df_ads_orders) # Menggunakan pesanan kategori Iklan
    # Note: Jika user ingin SEMUA pesanan masuk sini, ganti df_ads_orders dengan df_order. 
    # Tapi berdasarkan logika tabel Organik, harusnya ini dipisah. Saya gunakan df_ads_orders.

    # B. TABEL RINCIAN IKLAN KLIK
    total_dilihat = df_iklan['Dilihat'].sum()
    total_klik = df_iklan['Jumlah Klik'].sum()
    # Prompt: "Presentase Klik... dari 'Total Iklan Klik' dibagi 'Total Jumlah Klik'". 
    # Asumsi typo user: maksudnya Klik / Dilihat (CTR) atau Dilihat/Klik. 
    # Saya gunakan (Klik / Dilihat) * 100 karena %
    persentase_klik = (total_klik / total_dilihat) if total_dilihat > 0 else 0
    
    penjualan_iklan = tbl_iklan_data['OMZET PENJUALAN'].sum()

    # Hitung Biaya Iklan Spesifik
    # A5 Koran (Kapital di prompt berarti cek spesifik string atau case sensitive? Prompt: "tapi yang kapital")
    # Interpretasi: Mengandung "A5 Koran" DAN (Original String mengandung substring kapital atau case sensitive match)
    # Saya akan filter case sensitive untuk "A5 Koran" vs lower.
    
    # Helper filter contains
    def get_omzet_contains(query, case_sensitive=False):
        if case_sensitive:
            mask = df_iklan['Nama Iklan'].str.contains(query, case=True, regex=False)
        else:
            mask = df_iklan['Nama Iklan'].str.contains(query, case=False, regex=False)
        return df_iklan[mask]['Biaya'].sum()
        
    def get_biaya_regex(pattern, case_sensitive=False):
        if 'Biaya' not in df_iklan.columns:
            return 0
        # regex=True dan '.*' memungkinkan ada kata di tengah (misal: A5 Kertas Koran)
        mask = df_iklan['Nama Iklan'].str.contains(pattern, case=case_sensitive, regex=True, na=False)
        return df_iklan[mask]['Biaya'].sum()

    # "A5 Koran" (Kapital logic - asumsikan mengandung 'A5 KORAN' atau 'A5 Koran' vs 'a5 koran')
    # Prompt agak ambigu, saya gunakan pendekatan: Mengandung "A5 Koran" (Case sensitive)
    # biaya_a5_koran = get_omzet_contains("A5 Koran", case_sensitive=True) 
    
    # # "A5 Koran Paket 7" (Lowercase logic? Prompt: "tapii yang lowercase")
    # # Saya gunakan pendekatan: Mengandung "a5 koran" (lowercase) tapi TIDAK mengandung "A5 Koran"
    # mask_lower = (df_iklan['Nama Iklan'].str.contains("A5 koran", case=False)) & (~df_iklan['Nama Iklan'].str.contains("A5 Koran", case=True))
    # biaya_a5_koran_pkt7 = df_iklan[mask_lower]['Biaya'].sum()
    
    # # "A6 Pastel" (Case insensitive)
    # biaya_a6_pastel = get_omzet_contains("A6 Pastel", case_sensitive=False)
    # A5 Koran (versi WAKAF / uppercase)
    biaya_a5_koran = get_biaya_regex(r"A5.*KORAN", case_sensitive=True)
    
    # 2. Biaya Iklan A6 Pastel (Bebas besar/kecil)
    # Mencari "A6" diikuti "Pastel" (Case Sensitive = False)
    # Cocok dengan "...A6 ... WARNA PASTEL"
    biaya_a6_pastel = get_biaya_regex(r"A6.*Pastel", case_sensitive=False)
    
    # 3. Biaya Iklan A5 Koran Paket 7 (LOWERCASE / Title Case)
    # Logic: Cari yang mengandung "A5...Koran" (secara umum), TAPI kurangi yang sudah masuk kategori KAPITAL diatas.
    # Ini menangkap "Paket ... A5 Kertas Koran" (karena 'Koran' tidak sama dengan 'KORAN' di mode case sensitive)
    total_a5_general = get_biaya_regex(r"A5.*Koran", case_sensitive=False)
    biaya_a5_koran_pkt7 = total_a5_general - biaya_a5_koran
    
    # 4. Biaya Komik Pahlawan
    biaya_komik = get_biaya_regex(r"Komik Pahlawan", case_sensitive=False)
    
    # Total dan ROASI
    total_biaya_iklan_rinci = biaya_a5_koran + biaya_a5_koran_pkt7 + biaya_a6_pastel + biaya_komik
    roasi = (penjualan_iklan / total_biaya_iklan_rinci) if total_biaya_iklan_rinci > 0 else 0

    # C. TABEL PESANAN AFFILIATE
    tbl_affiliate_data = agg_by_hour(df_affiliate)
    # Tambah kolom Komisi (Perlu mapping jam ke komisi)
    # Komisi ada di df_seller, tapi df_seller tidak punya 'Jam'. 
    # Kita harus join df_affiliate (yang punya jam) dengan df_seller (yang punya komisi) based on Order ID
    df_aff_merged = df_affiliate.merge(df_seller[['Kode Pesanan', 'Pengeluaran(Rp)']], left_on='No. Pesanan', right_on='Kode Pesanan', how='left')
    # Konversi Pengeluaran ke float
    df_aff_merged['Pengeluaran(Rp)'] = pd.to_numeric(df_aff_merged['Pengeluaran(Rp)'].astype(str).str.replace('.','').str.replace(',','.'), errors='coerce').fillna(0)
    
    # Group komisi by jam
    komisi_per_jam = df_aff_merged.groupby('Jam')['Pengeluaran(Rp)'].sum().reset_index()
    tbl_affiliate_data = tbl_affiliate_data.merge(komisi_per_jam, on='Jam', how='left').fillna(0)
    tbl_affiliate_data.rename(columns={'Pengeluaran(Rp)': 'KOMISI'}, inplace=True)

    # D. TABEL PESANAN ORGANIK
    tbl_organik_data = agg_by_hour(df_organic)

    # E. TABEL RINCIAN SELURUH PESANAN (Product Level)
    # Logika: "jika dalam 1 No. Pesanan ada banyak Nama produk... ambil paling atas saja"
    # Ini berarti kita dedup by No. Pesanan dulu? Tapi tabel ini adalah rincian produk.
    # Jika kita dedup by Order ID, kita akan kehilangan data produk lain dalam order yang sama.
    # NAMUN, prompt bilang "Jumlah Pesanan diambil dari berapa pesanan...".
    # Interpretasi: Hitung frekuensi unik Order ID per Produk.
    # TAPI instruksi "ambil paling atas saja" sangat spesifik.
    # Saya akan lakukan: Group by Order ID -> Ambil first row -> Baru hitung stats produk dari hasil filter ini.
    
    df_unique_orders = df_order.sort_values(['No. Pesanan', 'Nama Produk']).drop_duplicates(subset=['No. Pesanan'], keep='first').copy()
    
    # Siapkan kolom variasi bersih
    df_unique_orders['Variasi_Clean'] = df_unique_orders['Nama Variasi'].apply(clean_variasi)
    
    # Group by Nama Produk & Variasi
    grp_rincian = df_unique_orders.groupby(['Nama Produk', 'Variasi_Clean']).agg(
        Jumlah_Pesanan=('No. Pesanan', 'count')
    ).reset_index()
    
    # Hitung Eksemplar
    grp_rincian['Jumlah Eksemplar'] = grp_rincian.apply(
        lambda row: extract_eksemplar(row['Variasi_Clean']) * row['Jumlah_Pesanan'], axis=1
    )

    # F. TABEL SUMMARY
    # Hitung total-total untuk summary
    total_omzet_all = tbl_iklan_data['OMZET PENJUALAN'].sum() + tbl_affiliate_data['OMZET PENJUALAN'].sum() + tbl_organik_data['OMZET PENJUALAN'].sum()
    total_komisi_aff = tbl_affiliate_data['KOMISI'].sum()
    roasf = total_omzet_all / (total_biaya_iklan_rinci + total_komisi_aff) if (total_biaya_iklan_rinci + total_komisi_aff) > 0 else 0

    # --- MEMBUAT FILE EXCEL ---
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    
    # FORMATS
    fmt_header_main = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': '#ADD8E6', 'align': 'left', 'valign': 'vcenter'})
    fmt_header_table = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    fmt_date = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    fmt_col_name = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#f0f0f0'})
    fmt_num = workbook.add_format({'border': 1, 'align': 'center'})
    fmt_curr = workbook.add_format({'border': 1, 'num_format': '#,##0', 'align': 'center'})
    fmt_percent = workbook.add_format({'border': 1, 'num_format': '0.00%', 'align': 'center'})
    fmt_text_left = workbook.add_format({'border': 1, 'align': 'left'})
    
    # --- SHEET 1: LAPORAN IKLAN ---
    ws_lap = workbook.add_worksheet('LAPORAN IKLAN')
    
    # Judul Utama
    ws_lap.merge_range('A1:P2', f'LAPORAN IKLAN {store_name.upper()}', fmt_header_main)
    
    # --- TABEL 1: PESANAN IKLAN (A-E) ---
    start_row = 3 # Row 4
    ws_lap.merge_range(start_row, 0, start_row, 4, 'PESANAN IKLAN', fmt_header_table)
    ws_lap.merge_range(start_row+1, 0, start_row+2, 4, report_date, fmt_date)
    
    cols_t1 = ['LIHAT', 'KLIK', 'PESANAN', 'KUANTITAS', 'OMZET PENJUALAN']
    for i, col in enumerate(cols_t1):
        ws_lap.write(start_row+3, i, col, fmt_col_name)
        
    # Isi Data Tabel 1
    row_cursor = start_row + 4
    for idx, row in tbl_iklan_data.iterrows():
        ws_lap.write(row_cursor, 0, "", fmt_num) # Lihat (Kosong)
        ws_lap.write(row_cursor, 1, "", fmt_num) # Klik (Kosong)
        ws_lap.write(row_cursor, 2, row['PESANAN'], fmt_num)
        ws_lap.write(row_cursor, 3, row['KUANTITAS'], fmt_num)
        ws_lap.write(row_cursor, 4, row['OMZET PENJUALAN'], fmt_curr)
        row_cursor += 1
        
    # Total Tabel 1
    ws_lap.write(row_cursor, 0, "TOTAL", fmt_col_name)
    ws_lap.write(row_cursor, 1, "", fmt_col_name)
    ws_lap.write(row_cursor, 2, tbl_iklan_data['PESANAN'].sum(), fmt_col_name)
    ws_lap.write(row_cursor, 3, tbl_iklan_data['KUANTITAS'].sum(), fmt_col_name)
    ws_lap.write(row_cursor, 4, tbl_iklan_data['OMZET PENJUALAN'].sum(), fmt_col_name)

    # --- TABEL 2: RINCIAN IKLAN KLIK (G-H) ---
    # Posisi sejajar dengan PESANAN IKLAN
    t2_row = start_row
    ws_lap.merge_range(t2_row, 6, t2_row, 7, 'RINCIAN IKLAN KLIK', fmt_header_table)
    # Item rincian
    rincian_items = [
        ('Total Iklan Dilihat', total_dilihat, fmt_num),
        ('Total Jumlah Klik', total_klik, fmt_num),
        ('Presentase Klik', persentase_klik, fmt_percent),
        ('Penjualan Iklan', penjualan_iklan, fmt_curr),
        ('Biaya Iklan A5 Koran', biaya_a5_koran, fmt_curr),
        ('Biaya Iklan A5 Koran Paket 7', biaya_a5_koran_pkt7, fmt_curr),
        ('Biaya Iklan A6 Pastel', biaya_a6_pastel, fmt_curr),
        ('Biaya Iklan Komik Pahlawan', biaya_komik, fmt_curr),
        ('ROASI', roasi, fmt_num) # ROAS biasanya desimal/ratio
    ]
    
    curr_t2_row = t2_row + 1
    for label, val, fmt in rincian_items:
        ws_lap.write(curr_t2_row, 6, label, fmt_text_left)
        ws_lap.write(curr_t2_row, 7, val, fmt)
        curr_t2_row += 1

    # --- TABEL 3: PESANAN AFFILIATE (L-P) ---
    t3_row = start_row
    ws_lap.merge_range(t3_row, 11, t3_row, 15, 'PESANAN AFFILIATE', fmt_header_table)
    cols_t3 = ['Jam', 'Pesanan', 'Kuantitas', 'Omzet Penjualan', 'Komisi']
    for i, col in enumerate(cols_t3):
        ws_lap.write(t3_row+1, 11+i, col, fmt_col_name)
        
    curr_t3_row = t3_row + 2
    for idx, row in tbl_affiliate_data.iterrows():
        ws_lap.write(curr_t3_row, 11, f"{int(row['Jam']):02d}:00", fmt_num)
        ws_lap.write(curr_t3_row, 12, row['PESANAN'], fmt_num)
        ws_lap.write(curr_t3_row, 13, row['KUANTITAS'], fmt_num)
        ws_lap.write(curr_t3_row, 14, row['OMZET PENJUALAN'], fmt_curr)
        ws_lap.write(curr_t3_row, 15, row['KOMISI'], fmt_curr)
        curr_t3_row += 1
        
    # Total T3
    total_omzet_aff = tbl_affiliate_data['OMZET PENJUALAN'].sum()
    total_komisi_aff_val = tbl_affiliate_data['KOMISI'].sum()
    ws_lap.write(curr_t3_row, 11, "TOTAL", fmt_col_name)
    ws_lap.write(curr_t3_row, 12, tbl_affiliate_data['PESANAN'].sum(), fmt_col_name)
    ws_lap.write(curr_t3_row, 13, tbl_affiliate_data['KUANTITAS'].sum(), fmt_col_name)
    ws_lap.write(curr_t3_row, 14, total_omzet_aff, fmt_col_name)
    ws_lap.write(curr_t3_row, 15, total_komisi_aff_val, fmt_col_name)
    
    # ROASA
    roasa = total_omzet_aff / total_komisi_aff_val if total_komisi_aff_val > 0 else 0
    curr_t3_row += 1
    ws_lap.write(curr_t3_row, 11, "ROASA", fmt_col_name)
    ws_lap.merge_range(curr_t3_row, 12, curr_t3_row, 13, "", fmt_num)
    ws_lap.write(curr_t3_row, 14, roasa, fmt_num)
    ws_lap.write(curr_t3_row, 15, "", fmt_num)
    
    last_row_affiliate = curr_t3_row

    # --- TABEL 4: PESANAN ORGANIK (L-O) ---
    # Dibawah Affiliate
    t4_row = last_row_affiliate + 2
    ws_lap.merge_range(t4_row, 11, t4_row, 14, 'PESANAN ORGANIK', fmt_header_table)
    cols_t4 = ['Jam', 'Pesanan', 'Kuantitas', 'Omzet Penjualan']
    for i, col in enumerate(cols_t4):
        ws_lap.write(t4_row+1, 11+i, col, fmt_col_name)
        
    curr_t4_row = t4_row + 2
    for idx, row in tbl_organik_data.iterrows():
        ws_lap.write(curr_t4_row, 11, f"{int(row['Jam']):02d}:00", fmt_num)
        ws_lap.write(curr_t4_row, 12, row['PESANAN'], fmt_num)
        ws_lap.write(curr_t4_row, 13, row['KUANTITAS'], fmt_num)
        ws_lap.write(curr_t4_row, 14, row['OMZET PENJUALAN'], fmt_curr)
        curr_t4_row += 1
        
    # Total T4
    ws_lap.write(curr_t4_row, 11, "TOTAL", fmt_col_name)
    ws_lap.write(curr_t4_row, 12, tbl_organik_data['PESANAN'].sum(), fmt_col_name)
    ws_lap.write(curr_t4_row, 13, tbl_organik_data['KUANTITAS'].sum(), fmt_col_name)
    ws_lap.write(curr_t4_row, 14, tbl_organik_data['OMZET PENJUALAN'].sum(), fmt_col_name)
    
    last_row_organik = curr_t4_row

    # --- TABEL 5: RINCIAN SELURUH PESANAN (G-J) ---
    # Posisi: Sejajar PESANAN ORGANIK (Row 13-14 dari L, berarti t4_row)
    # Prompt: "sejajar dengan tabel PESANAN ORGANIK... baris 13-14" (asumsi relatif terhadap layout)
    t5_row = t4_row 
    
    total_seluruh_pesanan_val = tbl_iklan_data['PESANAN'].sum() + tbl_affiliate_data['PESANAN'].sum() + tbl_organik_data['PESANAN'].sum()
    
    ws_lap.write(t5_row, 6, 'RINCIAN SELURUH PESANAN', fmt_header_table)
    ws_lap.write(t5_row, 7, total_seluruh_pesanan_val, fmt_header_table) # Col H
    ws_lap.merge_range(t5_row, 8, t5_row, 9, "", fmt_header_table)
    
    cols_t5 = ['Nama Produk', 'Variasi', 'Jumlah Pesanan', 'Jumlah Eksemplar']
    for i, col in enumerate(cols_t5):
        ws_lap.write(t5_row+1, 6+i, col, fmt_col_name)
        
    curr_t5_row = t5_row + 2
    for idx, row in grp_rincian.iterrows():
        ws_lap.write(curr_t5_row, 6, row['Nama Produk'], fmt_text_left)
        ws_lap.write(curr_t5_row, 7, row['Variasi_Clean'], fmt_num)
        ws_lap.write(curr_t5_row, 8, row['Jumlah_Pesanan'], fmt_num)
        ws_lap.write(curr_t5_row, 9, row['Jumlah Eksemplar'], fmt_num)
        curr_t5_row += 1
    
    # Total Eksemplar
    ws_lap.write(curr_t5_row, 8, "TOTAL EKSEMPLAR", fmt_col_name)
    ws_lap.write(curr_t5_row, 9, grp_rincian['Jumlah Eksemplar'].sum(), fmt_col_name)
    
    # --- TABEL 6: SUMMARY (L-P) ---
    # Posisi: 2 baris spasi dibawah Organik
    t6_row = last_row_organik + 3
    
    summary_data = [
        ('Penjualan Keseluruhan', total_omzet_all, fmt_curr),
        ('Total Biaya Iklan Klik', total_biaya_iklan_rinci, fmt_curr),
        ('Total Komisi Affiliate', total_komisi_aff, fmt_curr),
        ('ROASF', roasf, fmt_num)
    ]
    
    for label, val, fmt in summary_data:
        ws_lap.merge_range(t6_row, 11, t6_row, 14, label, fmt_text_left)
        ws_lap.write(t6_row, 15, val, fmt)
        t6_row += 1

    # --- SIMPAN SHEET LAINNYA ---
    # 1. order-all (dengan highlight)
    df_order.to_excel(writer, sheet_name='order-all', index=False)
    ws_order = workbook.get_worksheet_by_name('order-all')
    
    # Format Highlight
    fmt_yellow = workbook.add_format({'bg_color': '#FFFF00'})
    fmt_pink = workbook.add_format({'bg_color': '#FFC0CB'})
    
    # Iterasi untuk highlight
    # Note: XlsxWriter tidak bisa overwrite format cell dengan mudah tanpa menulis ulang
    # Kita loop df_order untuk nulis ulang baris dengan format yang sesuai
    
    # Get columns for rewrite
    columns = df_order.columns.tolist()
    
    for row_idx, row_data in df_order.iterrows():
        row_fmt = None
        if row_data['is_affiliate']:
            row_fmt = fmt_yellow
        # Kondisi Pink: Diluar Affiliate (sudah checked via elif/logic) DAN Diluar Iklan
        # Logika Highlight Pink: "diluar yang termasuk Seller conversion dan diluar dari Nama Produk yang tidak ada di Nama Iklan"
        elif not row_data['is_iklan_product']:
            row_fmt = fmt_pink
            
        if row_fmt:
            for col_idx, col_name in enumerate(columns):
                # Tulis ulang cell dengan format. Handle datetime convert
                val = row_data[col_name]
                # Sederhanakan penulisan (XlsxWriter handle basic types)
                if pd.isna(val):
                    ws_order.write(row_idx + 1, col_idx, "", row_fmt)
                else:
                     # Check if datetime
                    if isinstance(val, pd.Timestamp):
                        ws_order.write_datetime(row_idx + 1, col_idx, val, workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm', 'bg_color': row_fmt.bg_color}))
                    else:
                        ws_order.write(row_idx + 1, col_idx, val, row_fmt)

    # 2. Iklan klik (Cleaned)
    df_iklan.to_excel(writer, sheet_name='Iklan klik', index=False)
    
    # 3. Seller conversion (Raw)
    df_seller.to_excel(writer, sheet_name='Seller conversion', index=False)

    writer.close()
    output.seek(0)
    return output

# --- INTERFACE STREAMLIT ---

st.title("🛒 IklanKu - Generator Laporan Otomatis")
st.markdown("---")

# Input Toko
toko = st.selectbox("Pilih Toko:", ["Human Store", "Pasific BookStore", "Dama Store"])

# Input File
col1, col2, col3 = st.columns(3)
with col1:
    f_order = st.file_uploader("Upload 'Order-all' (xlsx)", type=['xlsx'])
with col2:
    f_iklan = st.file_uploader("Upload 'Iklan Keseluruhan' (csv)", type=['csv'])
with col3:
    f_seller = st.file_uploader("Upload 'Seller conversion' (csv)", type=['csv'])

if st.button("Mulai Proses", type="primary"):
    if f_order and f_iklan and f_seller:
        with st.spinner('Sedang memproses data... Tunggu sebentar ya...'):
            try:
                excel_file = process_data(toko, f_order, f_iklan, f_seller)
                
                if excel_file:
                    st.success("Proses Selesai!")
                    st.download_button(
                        label="📥 Download Laporan Excel",
                        data=excel_file,
                        file_name=f"LAPORAN_IKLAN_{toko.replace(' ', '_').upper()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses: {e}")
                st.write(e) # Debug info
    else:
        st.warning("Mohon upload ketiga file terlebih dahulu.")
