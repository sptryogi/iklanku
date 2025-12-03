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

    # Dictionary untuk menyimpan panjang maksimum setiap kolom
    col_widths = {}
    
    def update_width(col_idx, value):
        # Hitung panjang string valuenya
        width = len(str(value)) if value is not None else 0
        current_max = col_widths.get(col_idx, 0)
        if width > current_max:
            col_widths[col_idx] = width
            
    # 1. LOAD DATA
    df_order = pd.read_excel(file_order)
    df_iklan = pd.read_csv(file_iklan, skiprows=7)
    # df_seller = pd.read_csv(file_seller)
    # Cek apakah file_seller ada isinya (Optional)
    if file_seller is not None:
        df_seller = pd.read_csv(file_seller)
    else:
        # Jika tidak ada, buat DataFrame kosong dengan kolom minimal agar tidak error saat merge
        df_seller = pd.DataFrame(columns=['Kode Pesanan', 'Pengeluaran(Rp)'])

    # 2. PRE-PROCESS ORDER-ALL
    # Filter Status Pesanan != Batal dan Belum Bayar
    if 'Status Pesanan' in df_order.columns:
        df_order = df_order[df_order['Status Pesanan'] != 'Batal'].copy()
        df_order = df_order[df_order['Status Pesanan'] != 'Belum Bayar'].copy()
    
    # Konversi kolom waktu
    if 'Waktu Pesanan Dibuat' in df_order.columns:
        df_order['Waktu Pesanan Dibuat'] = pd.to_datetime(df_order['Waktu Pesanan Dibuat'])
        df_order['Jam'] = df_order['Waktu Pesanan Dibuat'].dt.hour
        # Ambil tanggal untuk header laporan
        report_date = df_order['Waktu Pesanan Dibuat'].dt.strftime('%A, %d-%m-%Y').iloc[0] if not df_order.empty else "TANGGAL TIDAK DIKETAHUI"
    else:
        st.error("Kolom 'Waktu Pesanan Dibuat' tidak ditemukan di Order-all")
        return None

    def clean_indo_currency(x):
        if pd.isna(x): return 0
        if isinstance(x, (int, float)): return x # Jika sudah angka, kembalikan langsung
        # Jika string, bersihkan format Rp dan titik ribuan
        x = str(x).replace('Rp', '').replace(' ', '').replace('.', '') # Hapus titik ribuan
        x = x.replace(',', '.') # Ganti koma desimal jadi titik
        try:
            return float(x)
        except:
            return 0

    # Terapkan pembersihan hanya ke kolom 'Total Harga Produk' dan 'Jumlah'
    if 'Total Harga Produk' in df_order.columns:
        df_order['Total Harga Produk'] = df_order['Total Harga Produk'].apply(clean_indo_currency)
    
    if 'Jumlah' in df_order.columns:
        df_order['Jumlah'] = df_order['Jumlah'].apply(clean_indo_currency)

    if 'Harga Satuan' in df_order.columns:
        df_order['Harga Satuan'] = df_order['Harga Satuan'].apply(clean_indo_currency)

    # 3. PRE-PROCESS IKLAN (Sheet 'Iklan klik')
    # Hapus 7 baris pertama (index 0-6), baris ke-8 (index 7) jadi header
    # new_header = df_iklan_raw.iloc[7]
    # df_iklan = df_iklan_raw[8:].copy()
    # df_iklan.columns = new_header
    df_iklan.columns = df_iklan.columns.str.strip()
    
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
    
    # --- AGGREGATION FUNCTIONS ---

    # A. TABEL PESANAN IKLAN (Fixed 24 Jam)
    hours_fixed = pd.DataFrame({'Jam': range(24)})
    
    def agg_fixed_hours(df_source):
        if df_source.empty:
            return pd.DataFrame({'Jam': range(24), 'PESANAN': 0, 'KUANTITAS': 0, 'OMZET PENJUALAN': 0})
        grp_pesanan = df_source.groupby('Jam')['No. Pesanan'].nunique().reset_index(name='PESANAN')
        grp_metrics = df_source.groupby('Jam')[['Jumlah', 'Total Harga Produk']].sum().reset_index()
        grp_metrics.rename(columns={'Jumlah': 'KUANTITAS', 'Total Harga Produk': 'OMZET PENJUALAN'}, inplace=True)
        merged = hours_fixed.merge(grp_pesanan, on='Jam', how='left').merge(grp_metrics, on='Jam', how='left')
        return merged.fillna(0)

    tbl_iklan_data = agg_fixed_hours(df_ads_orders)

    # B. TABEL DINAMIS (AFFILIATE & ORGANIK)
    def agg_dynamic_hours(df_source, context=""):
        # PERBAIKAN: Definisikan kolom wajib agar tidak Error saat data kosong
        expected_cols = ['Jam', 'PESANAN', 'KUANTITAS', 'OMZET PENJUALAN']
        
        if df_source.empty:
            # Kembalikan DataFrame kosong TAPI dengan nama kolom yang sudah disiapkan
            return pd.DataFrame(columns=expected_cols) 
        
        grp_pesanan = df_source.groupby('Jam')['No. Pesanan'].nunique().reset_index(name='PESANAN')
        grp_metrics = df_source.groupby('Jam')[['Jumlah', 'Total Harga Produk']].sum().reset_index()
        grp_metrics.rename(columns={'Jumlah': 'KUANTITAS', 'Total Harga Produk': 'OMZET PENJUALAN'}, inplace=True)
        
        merged = grp_pesanan.merge(grp_metrics, on='Jam', how='left').fillna(0)
        merged = merged.sort_values('Jam')
        return merged

    # Note: Jika user ingin SEMUA pesanan masuk sini, ganti df_ads_orders dengan df_order. 
    # Tapi berdasarkan logika tabel Organik, harusnya ini dipisah. Saya gunakan df_ads_orders.

    # C. TABEL RINCIAN IKLAN KLIK
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

    # # A5 Koran (versi WAKAF / uppercase)
    # biaya_a5_koran = get_biaya_regex(r"A5.*KORAN", case_sensitive=True)
    
    # # 2. Biaya Iklan A6 Pastel (Bebas besar/kecil)
    # # Mencari "A6" diikuti "Pastel" (Case Sensitive = False)
    # # Cocok dengan "...A6 ... WARNA PASTEL"
    # biaya_a6_pastel = get_biaya_regex(r"A6.*Pastel", case_sensitive=False)
    
    # # 3. Biaya Iklan A5 Koran Paket 7 (LOWERCASE / Title Case)
    # # Logic: Cari yang mengandung "A5...Koran" (secara umum), TAPI kurangi yang sudah masuk kategori KAPITAL diatas.
    # # Ini menangkap "Paket ... A5 Kertas Koran" (karena 'Koran' tidak sama dengan 'KORAN' di mode case sensitive)
    # total_a5_general = get_biaya_regex(r"A5.*Koran", case_sensitive=False)
    # biaya_a5_koran_pkt7 = total_a5_general - biaya_a5_koran
    
    # # 4. Biaya Komik Pahlawan
    # biaya_komik = get_biaya_regex(r"Komik Pahlawan", case_sensitive=False)

    # Hitung Variabel Biaya
    biaya_a5_koran = get_biaya_regex(r"A5.*KORAN", case_sensitive=True)
    biaya_a6_pastel = get_biaya_regex(r"A6.*Pastel", case_sensitive=False)
    total_a5_general = get_biaya_regex(r"A5.*Koran", case_sensitive=False)
    biaya_a5_koran_pkt7 = total_a5_general - biaya_a5_koran
    biaya_komik = get_biaya_regex(r"Komik Pahlawan", case_sensitive=False)
    
    # --- LOGIKA TOKO UNTUK ITEM RINCIAN ---
    rincian_items = [
        ('Total Iklan Dilihat', total_dilihat),
        ('Total Jumlah Klik', total_klik),
        ('Presentase Klik', persentase_klik),
        ('Penjualan Iklan', penjualan_iklan)
    ]

    total_biaya_iklan_rinci = 0
    
    if store_name == 'Human Store':
        rincian_items.append(('Biaya Iklan A5 Koran', biaya_a5_koran))
        rincian_items.append(('Biaya Iklan A5 Koran Paket 7', biaya_a5_koran_pkt7))
        rincian_items.append(('Biaya Iklan A6 Pastel', biaya_a6_pastel))
        rincian_items.append(('Biaya Iklan Komik Pahlawan', biaya_komik))
        total_biaya_iklan_rinci = biaya_a5_koran + biaya_a5_koran_pkt7 + biaya_a6_pastel + biaya_komik
        
    elif store_name == 'Pasific BookStore':
        rincian_items.append(('Biaya Iklan A5 Koran', biaya_a5_koran))
        rincian_items.append(('Biaya Iklan A6 Pastel', biaya_a6_pastel))
        total_biaya_iklan_rinci = biaya_a5_koran + biaya_a6_pastel
        
    elif store_name == 'Dama Store':
        rincian_items.append(('Biaya Iklan A5 Koran', biaya_a5_koran))
        rincian_items.append(('Biaya Iklan A5 Koran Paket 7', biaya_a5_koran_pkt7))
        rincian_items.append(('Biaya Iklan A6 Pastel', biaya_a6_pastel))
        total_biaya_iklan_rinci = biaya_a5_koran + biaya_a5_koran_pkt7 + biaya_a6_pastel
    
    # # Total dan ROASI
    # total_biaya_iklan_rinci = biaya_a5_koran + biaya_a5_koran_pkt7 + biaya_a6_pastel + biaya_komik
    # roasi = (penjualan_iklan / total_biaya_iklan_rinci) if total_biaya_iklan_rinci > 0 else 0

    # Hitung ROASI
    roasi = (penjualan_iklan / total_biaya_iklan_rinci) if total_biaya_iklan_rinci > 0 else 0
    rincian_items.append(('ROASI', roasi))

    # # D. TABEL PESANAN AFFILIATE
    # tbl_affiliate_data = agg_by_hour(df_affiliate)
    # # Tambah kolom Komisi (Perlu mapping jam ke komisi)
    # # Komisi ada di df_seller, tapi df_seller tidak punya 'Jam'. 
    # # Kita harus join df_affiliate (yang punya jam) dengan df_seller (yang punya komisi) based on Order ID
    # df_aff_merged = df_affiliate.merge(df_seller[['Kode Pesanan', 'Pengeluaran(Rp)']], left_on='No. Pesanan', right_on='Kode Pesanan', how='left')
    # # Konversi Pengeluaran ke float
    # df_aff_merged['Pengeluaran(Rp)'] = pd.to_numeric(df_aff_merged['Pengeluaran(Rp)'].astype(str).str.replace('.','').str.replace(',','.'), errors='coerce').fillna(0)
    
    # # Group komisi by jam
    # komisi_per_jam = df_aff_merged.groupby('Jam')['Pengeluaran(Rp)'].sum().reset_index()
    # tbl_affiliate_data = tbl_affiliate_data.merge(komisi_per_jam, on='Jam', how='left').fillna(0)
    # tbl_affiliate_data.rename(columns={'Pengeluaran(Rp)': 'KOMISI'}, inplace=True)

    # # F. TABEL PESANAN ORGANIK
    # tbl_organik_data = agg_by_hour(df_organic)
    
    # D. PREP AFFILIATE & ORGANIK DATA
    tbl_affiliate_data = agg_dynamic_hours(df_affiliate)
    # Tambah komisi untuk affiliate
    if not tbl_affiliate_data.empty:
        if 'Kode Pesanan' in df_seller.columns and 'Pengeluaran(Rp)' in df_seller.columns:
            df_aff_merged = df_affiliate.merge(df_seller[['Kode Pesanan', 'Pengeluaran(Rp)']], left_on='No. Pesanan', right_on='Kode Pesanan', how='left')
            df_aff_merged['Pengeluaran(Rp)'] = df_aff_merged['Pengeluaran(Rp)'].apply(clean_indo_currency)
            komisi_per_jam = df_aff_merged.groupby('Jam')['Pengeluaran(Rp)'].sum().reset_index()
            tbl_affiliate_data = tbl_affiliate_data.merge(komisi_per_jam, on='Jam', how='left').fillna(0)
            tbl_affiliate_data.rename(columns={'Pengeluaran(Rp)': 'KOMISI'}, inplace=True)
        else:
            tbl_affiliate_data['KOMISI'] = 0
            
    tbl_organik_data = agg_dynamic_hours(df_organic)

    # E. TABEL RINCIAN SELURUH PESANAN (Product Level)
    # Logika: "jika dalam 1 No. Pesanan ada banyak Nama produk... ambil paling atas saja"
    # Ini berarti kita dedup by No. Pesanan dulu? Tapi tabel ini adalah rincian produk.
    # Jika kita dedup by Order ID, kita akan kehilangan data produk lain dalam order yang sama.
    # NAMUN, prompt bilang "Jumlah Pesanan diambil dari berapa pesanan...".
    # Interpretasi: Hitung frekuensi unik Order ID per Produk.
    # TAPI instruksi "ambil paling atas saja" sangat spesifik.
    # Saya akan lakukan: Group by Order ID -> Ambil first row -> Baru hitung stats produk dari hasil filter ini.
    
    # df_unique_orders = df_order.sort_values(['No. Pesanan', 'Nama Produk']).drop_duplicates(subset=['No. Pesanan'], keep='first').copy()
    
    # # Siapkan kolom variasi bersih
    # df_unique_orders['Variasi_Clean'] = df_unique_orders['Nama Variasi'].apply(clean_variasi)
    
    # # Group by Nama Produk & Variasi
    # grp_rincian = df_unique_orders.groupby(['Nama Produk', 'Variasi_Clean']).agg(
    #     Jumlah_Pesanan=('No. Pesanan', 'count')
    # ).reset_index()
    
    # # Hitung Eksemplar
    # grp_rincian['Jumlah Eksemplar'] = grp_rincian.apply(
    #     lambda row: extract_eksemplar(row['Variasi_Clean']) * row['Jumlah_Pesanan'], axis=1
    # )
    df_unique_orders = df_order.drop_duplicates(subset=['No. Pesanan'], keep='first').copy()
    
    if 'Nama Variasi' in df_unique_orders.columns:
        df_unique_orders['Variasi_Clean'] = df_unique_orders['Nama Variasi'].apply(clean_variasi)
    else:
        df_unique_orders['Variasi_Clean'] = ''

    grp_rincian = df_unique_orders.groupby(['Nama Produk', 'Variasi_Clean']).agg(
        Jumlah_Pesanan=('No. Pesanan', 'count')
    ).reset_index()
    
    grp_rincian['Jumlah Eksemplar'] = grp_rincian.apply(
        lambda row: extract_eksemplar(row['Variasi_Clean']) * row['Jumlah_Pesanan'], axis=1
    )

    # F. TABEL SUMMARY
    # Hitung total-total untuk summary
    # total_omzet_all = tbl_iklan_data['OMZET PENJUALAN'].sum() + tbl_affiliate_data['OMZET PENJUALAN'].sum() + tbl_organik_data['OMZET PENJUALAN'].sum()
    # total_komisi_aff = tbl_affiliate_data['KOMISI'].sum()
    # roasf = total_omzet_all / (total_biaya_iklan_rinci + total_komisi_aff) if (total_biaya_iklan_rinci + total_komisi_aff) > 0 else 0
    total_omzet_all = 0
    
    # 1. Tambah Iklan
    if 'OMZET PENJUALAN' in tbl_iklan_data.columns:
        total_omzet_all += tbl_iklan_data['OMZET PENJUALAN'].sum()
        
    # 2. Tambah Affiliate (Cek empty dan kolom)
    if not tbl_affiliate_data.empty and 'OMZET PENJUALAN' in tbl_affiliate_data.columns:
        total_omzet_all += tbl_affiliate_data['OMZET PENJUALAN'].sum()
        
    # 3. Tambah Organik (Cek empty dan kolom)
    if not tbl_organik_data.empty and 'OMZET PENJUALAN' in tbl_organik_data.columns:
        total_omzet_all += tbl_organik_data['OMZET PENJUALAN'].sum()
    
    # Hitung Komisi
    total_komisi_aff = 0
    if not tbl_affiliate_data.empty and 'KOMISI' in tbl_affiliate_data.columns:
        total_komisi_aff = tbl_affiliate_data['KOMISI'].sum()
        
    # Hitung ROASF
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

    # --- FORMAT WARNA HEADER BARU ---
    # Tabel 1: Orange (#FFA500)
    fmt_head_orange = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#FFA500'})
    # Tabel 2: Coklat (#D2691E - Chocolate, agar tulisan hitam masih terbaca)
    fmt_head_brown = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#D2691E', 'font_color': 'white'})
    # Tabel 3: Kuning (#FFFF00)
    fmt_head_yellow = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#FFFF00'})
    # Tabel 4: Pink (#FFC0CB)
    fmt_head_pink = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#FFC0CB'})
    # Tabel 5: Hijau (#90EE90)
    fmt_head_green = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#90EE90'})

    # --- FORMAT BOLD UNTUK SUMMARY ---
    fmt_text_left_bold = workbook.add_format({'border': 1, 'align': 'left', 'bold': True})
    fmt_curr_bold = workbook.add_format({'border': 1, 'num_format': '#,##0', 'align': 'center', 'bold': True})
    fmt_num_bold = workbook.add_format({'border': 1, 'align': 'center', 'bold': True})

    # TAMBAHKAN INI: Format angka dengan 2 desimal
    fmt_decimal = workbook.add_format({'border': 1, 'num_format': '0.00', 'align': 'center'})
    
    # --- SHEET 1: LAPORAN IKLAN ---
    ws_lap = workbook.add_worksheet('LAPORAN IKLAN')
    
    # Judul Utama
    ws_lap.merge_range('A1:Q2', f'LAPORAN IKLAN {store_name.upper()}', fmt_header_main)
    
    # --- TABEL 1: PESANAN IKLAN (A-F) ---
    start_row = 3 # Row 4
    ws_lap.merge_range(start_row, 0, start_row, 5, 'PESANAN IKLAN', fmt_head_orange)
    ws_lap.merge_range(start_row+1, 0, start_row+2, 5, report_date, fmt_date)
    
    cols_t1 = ['JAM', 'LIHAT', 'KLIK', 'PESANAN', 'KUANTITAS', 'OMZET PENJUALAN']
    for i, col in enumerate(cols_t1):
        ws_lap.write(start_row+3, i, col, fmt_col_name)
        update_width(i, col)
        
    # Isi Data Tabel 1
    row_cursor = start_row + 4
    for idx, row in tbl_iklan_data.iterrows():
        # Kolom Jam 00:00 - 23:00
        jam_str = f"{int(row['Jam']):02d}:00"
        ws_lap.write(row_cursor, 0, jam_str, fmt_num)
        update_width(0, jam_str)
        
        ws_lap.write(row_cursor, 1, "", fmt_num) # Lihat (Kosong)
        ws_lap.write(row_cursor, 2, "", fmt_num) # Klik (Kosong)
        ws_lap.write(row_cursor, 3, row['PESANAN'], fmt_num)
        ws_lap.write(row_cursor, 4, row['KUANTITAS'], fmt_num)
        ws_lap.write(row_cursor, 5, row['OMZET PENJUALAN'], fmt_curr)

        # Update widths
        update_width(3, row['PESANAN'])
        update_width(4, row['KUANTITAS'])
        update_width(5, f"{row['OMZET PENJUALAN']:,}")
        
        row_cursor += 1
        
    # Total Tabel 1
    ws_lap.write(row_cursor, 0, "TOTAL", fmt_col_name)
    ws_lap.merge_range(row_cursor, 1, row_cursor, 2, "", fmt_col_name)
    ws_lap.write(row_cursor, 3, tbl_iklan_data['PESANAN'].sum(), fmt_col_name)
    ws_lap.write(row_cursor, 4, tbl_iklan_data['KUANTITAS'].sum(), fmt_col_name)
    ws_lap.write(row_cursor, 5, tbl_iklan_data['OMZET PENJUALAN'].sum(), fmt_col_name)
    update_width(5, f"{tbl_iklan_data['OMZET PENJUALAN'].sum():,}")
    
    # --- TABEL 2: RINCIAN IKLAN KLIK (G-H) ---
    # Posisi sejajar dengan PESANAN IKLAN
    # t2_row = start_row
    # ws_lap.merge_range(t2_row, 6, t2_row, 7, 'RINCIAN IKLAN KLIK', fmt_head_brown)
    # # Item rincian
    # rincian_items = [
    #     ('Total Iklan Dilihat', total_dilihat, fmt_num),
    #     ('Total Jumlah Klik', total_klik, fmt_num),
    #     ('Presentase Klik', persentase_klik, fmt_percent),
    #     ('Penjualan Iklan', penjualan_iklan, fmt_curr),
    #     ('Biaya Iklan A5 Koran', biaya_a5_koran, fmt_curr),
    #     ('Biaya Iklan A5 Koran Paket 7', biaya_a5_koran_pkt7, fmt_curr),
    #     ('Biaya Iklan A6 Pastel', biaya_a6_pastel, fmt_curr),
    #     ('Biaya Iklan Komik Pahlawan', biaya_komik, fmt_curr),
    #     ('ROASI', roasi, fmt_num) # ROAS biasanya desimal/ratio
    # ]
    
    # curr_t2_row = t2_row + 1
    # for label, val, fmt in rincian_items:
    #     ws_lap.write(curr_t2_row, 6, label, fmt_text_left)
    #     ws_lap.write(curr_t2_row, 7, val, fmt)
    #     curr_t2_row += 1
    t2_col_start = 7 # H
    t2_row = start_row
    
    ws_lap.merge_range(t2_row, t2_col_start, t2_row, t2_col_start+1, 'RINCIAN IKLAN KLIK', fmt_head_brown)
    
    curr_t2_row = t2_row + 1
    for label, val in rincian_items:
        # Tentukan format
        if 'Presentase' in label: fmt = fmt_percent
        elif 'ROAS' in label: fmt = fmt_decimal
        elif 'Total' in label and 'Dilihat' in label: fmt = fmt_num
        elif 'Total' in label and 'Klik' in label: fmt = fmt_num
        else: fmt = fmt_curr # Default currency
        
        ws_lap.write(curr_t2_row, t2_col_start, label, fmt_text_left)
        ws_lap.write(curr_t2_row, t2_col_start+1, val, fmt)
        
        update_width(t2_col_start, label)
        update_width(t2_col_start+1, str(val))
        curr_t2_row += 1

    # --- TABEL 3: PESANAN AFFILIATE (L-P) ---
    # t3_row = start_row
    # ws_lap.merge_range(t3_row, 11, t3_row, 15, 'PESANAN AFFILIATE', fmt_head_yellow)
    # cols_t3 = ['Jam', 'Pesanan', 'Kuantitas', 'Omzet Penjualan', 'Komisi']
    # for i, col in enumerate(cols_t3):
    #     ws_lap.write(t3_row+1, 11+i, col, fmt_col_name)
        
    # curr_t3_row = t3_row + 2
    # for idx, row in tbl_affiliate_data.iterrows():
    #     ws_lap.write(curr_t3_row, 11, f"{int(row['Jam']):02d}:00", fmt_num)
    #     ws_lap.write(curr_t3_row, 12, row['PESANAN'], fmt_num)
    #     ws_lap.write(curr_t3_row, 13, row['KUANTITAS'], fmt_num)
    #     ws_lap.write(curr_t3_row, 14, row['OMZET PENJUALAN'], fmt_curr)
    #     ws_lap.write(curr_t3_row, 15, row['KOMISI'], fmt_curr)
    #     curr_t3_row += 1
        
    # # Total T3
    # total_omzet_aff = tbl_affiliate_data['OMZET PENJUALAN'].sum()
    # total_komisi_aff_val = tbl_affiliate_data['KOMISI'].sum()
    # ws_lap.write(curr_t3_row, 11, "TOTAL", fmt_col_name)
    # ws_lap.write(curr_t3_row, 12, tbl_affiliate_data['PESANAN'].sum(), fmt_col_name)
    # ws_lap.write(curr_t3_row, 13, tbl_affiliate_data['KUANTITAS'].sum(), fmt_col_name)
    # ws_lap.write(curr_t3_row, 14, total_omzet_aff, fmt_col_name)
    # ws_lap.write(curr_t3_row, 15, total_komisi_aff_val, fmt_col_name)
    
    # # ROASA
    # roasa = total_omzet_aff / total_komisi_aff_val if total_komisi_aff_val > 0 else 0
    # curr_t3_row += 1
    # ws_lap.write(curr_t3_row, 11, "ROASA", fmt_col_name)
    # ws_lap.merge_range(curr_t3_row, 12, curr_t3_row, 13, "", fmt_num)
    # ws_lap.write(curr_t3_row, 14, roasa, fmt_num)
    # ws_lap.write(curr_t3_row, 15, "", fmt_num)
    
    # last_row_affiliate = curr_t3_row

    t3_col_start = 12 # M
    t3_row = start_row
    t3_cols = ['Jam', 'Pesanan', 'Kuantitas', 'Omzet Penjualan', 'Komisi']
    
    ws_lap.merge_range(t3_row, t3_col_start, t3_row, t3_col_start+4, 'PESANAN AFFILIATE', fmt_head_yellow)
    for i, col in enumerate(t3_cols):
        ws_lap.write(t3_row+1, t3_col_start+i, col, fmt_col_name)
        update_width(t3_col_start+i, col)
        
    curr_t3_row = t3_row + 2
    
    # Jika Data Kosong, buat 5 baris kosong
    if tbl_affiliate_data.empty:
        for _ in range(5):
             for i in range(5):
                 ws_lap.write(curr_t3_row, t3_col_start+i, "", fmt_num)
             curr_t3_row += 1
    else:
        for idx, row in tbl_affiliate_data.iterrows():
            ws_lap.write(curr_t3_row, t3_col_start, f"{int(row['Jam']):02d}:00", fmt_num)
            ws_lap.write(curr_t3_row, t3_col_start+1, row['PESANAN'], fmt_num)
            ws_lap.write(curr_t3_row, t3_col_start+2, row['KUANTITAS'], fmt_num)
            ws_lap.write(curr_t3_row, t3_col_start+3, row['OMZET PENJUALAN'], fmt_curr)
            ws_lap.write(curr_t3_row, t3_col_start+4, row['KOMISI'], fmt_curr)
            
            update_width(t3_col_start, f"{int(row['Jam']):02d}:00")
            update_width(t3_col_start+3, f"{row['OMZET PENJUALAN']:,}")
            curr_t3_row += 1

    # Total T3
    if not tbl_affiliate_data.empty:
        ws_lap.write(curr_t3_row, t3_col_start, "TOTAL", fmt_col_name)
        ws_lap.write(curr_t3_row, t3_col_start+1, tbl_affiliate_data['PESANAN'].sum(), fmt_col_name)
        ws_lap.write(curr_t3_row, t3_col_start+2, tbl_affiliate_data['KUANTITAS'].sum(), fmt_col_name)
        ws_lap.write(curr_t3_row, t3_col_start+3, tbl_affiliate_data['OMZET PENJUALAN'].sum(), fmt_col_name)
        ws_lap.write(curr_t3_row, t3_col_start+4, tbl_affiliate_data['KOMISI'].sum(), fmt_col_name)
        total_omzet_aff = tbl_affiliate_data['OMZET PENJUALAN'].sum()
        total_komisi_aff_val = tbl_affiliate_data['KOMISI'].sum()
        roasa = total_omzet_aff / total_komisi_aff_val if total_komisi_aff_val > 0 else 0
        curr_t3_row += 1
        
        # ROASA
        ws_lap.write(curr_t3_row, t3_col_start, "ROASA", fmt_col_name)
        ws_lap.merge_range(curr_t3_row, t3_col_start+1, curr_t3_row, t3_col_start+2, "", fmt_num)
        ws_lap.write(curr_t3_row, t3_col_start+3, roasa, fmt_decimal)
        ws_lap.write(curr_t3_row, t3_col_start+4, "", fmt_num)
        curr_t3_row += 1
    
    last_row_affiliate = curr_t3_row

    # --- TABEL 4: PESANAN ORGANIK (M-P) ---
    # Dibawah Affiliate
    # t4_row = last_row_affiliate + 2
    # ws_lap.merge_range(t4_row, 11, t4_row, 14, 'PESANAN ORGANIK', fmt_head_pink)
    # cols_t4 = ['Jam', 'Pesanan', 'Kuantitas', 'Omzet Penjualan']
    # for i, col in enumerate(cols_t4):
    #     ws_lap.write(t4_row+1, 11+i, col, fmt_col_name)
        
    # curr_t4_row = t4_row + 2
    # for idx, row in tbl_organik_data.iterrows():
    #     ws_lap.write(curr_t4_row, 11, f"{int(row['Jam']):02d}:00", fmt_num)
    #     ws_lap.write(curr_t4_row, 12, row['PESANAN'], fmt_num)
    #     ws_lap.write(curr_t4_row, 13, row['KUANTITAS'], fmt_num)
    #     ws_lap.write(curr_t4_row, 14, row['OMZET PENJUALAN'], fmt_curr)
    #     curr_t4_row += 1
        
    # # Total T4
    # ws_lap.write(curr_t4_row, 11, "TOTAL", fmt_col_name)
    # ws_lap.write(curr_t4_row, 12, tbl_organik_data['PESANAN'].sum(), fmt_col_name)
    # ws_lap.write(curr_t4_row, 13, tbl_organik_data['KUANTITAS'].sum(), fmt_col_name)
    # ws_lap.write(curr_t4_row, 14, tbl_organik_data['OMZET PENJUALAN'].sum(), fmt_col_name)
    
    # last_row_organik = curr_t4_row

    t4_row = last_row_affiliate + 2
    t4_col_start = t3_col_start # M
    t4_cols = ['Jam', 'Pesanan', 'Kuantitas', 'Omzet Penjualan']
    
    ws_lap.merge_range(t4_row, t4_col_start, t4_row, t4_col_start+3, 'PESANAN ORGANIK', fmt_head_pink)
    for i, col in enumerate(t4_cols):
        ws_lap.write(t4_row+1, t4_col_start+i, col, fmt_col_name)
        update_width(t4_col_start+i, col)

    curr_t4_row = t4_row + 2
    if tbl_organik_data.empty:
        for _ in range(5):
             for i in range(4):
                 ws_lap.write(curr_t4_row, t4_col_start+i, "", fmt_num)
             curr_t4_row += 1
    else:
        for idx, row in tbl_organik_data.iterrows():
            ws_lap.write(curr_t4_row, t4_col_start, f"{int(row['Jam']):02d}:00", fmt_num)
            ws_lap.write(curr_t4_row, t4_col_start+1, row['PESANAN'], fmt_num)
            ws_lap.write(curr_t4_row, t4_col_start+2, row['KUANTITAS'], fmt_num)
            ws_lap.write(curr_t4_row, t4_col_start+3, row['OMZET PENJUALAN'], fmt_curr)
            update_width(t4_col_start, f"{int(row['Jam']):02d}:00")
            update_width(t4_col_start+3, f"{row['OMZET PENJUALAN']:,}")
            curr_t4_row += 1
            
    if not tbl_organik_data.empty:
        ws_lap.write(curr_t4_row, t4_col_start, "TOTAL", fmt_col_name)
        ws_lap.write(curr_t4_row, t4_col_start+1, tbl_organik_data['PESANAN'].sum(), fmt_col_name)
        ws_lap.write(curr_t4_row, t4_col_start+2, tbl_organik_data['KUANTITAS'].sum(), fmt_col_name)
        ws_lap.write(curr_t4_row, t4_col_start+3, tbl_organik_data['OMZET PENJUALAN'].sum(), fmt_col_name)
        curr_t4_row += 1
        
    last_row_organik = curr_t4_row

    # --- TABEL 5: RINCIAN SELURUH PESANAN (H-K) ---
    # Posisi: Sejajar PESANAN ORGANIK (Row 13-14 dari L, berarti t4_row)
    # Prompt: "sejajar dengan tabel PESANAN ORGANIK... baris 13-14" (asumsi relatif terhadap layout)
    # t5_row = t4_row 
    
    # total_seluruh_pesanan_val = tbl_iklan_data['PESANAN'].sum() + tbl_affiliate_data['PESANAN'].sum() + tbl_organik_data['PESANAN'].sum()
    
    # ws_lap.write(t5_row, 6, 'RINCIAN SELURUH PESANAN', fmt_head_green)
    # ws_lap.write(t5_row, 7, total_seluruh_pesanan_val, fmt_header_table) # Col H
    # ws_lap.merge_range(t5_row, 8, t5_row, 9, "", fmt_header_table)
    
    # cols_t5 = ['Nama Produk', 'Variasi', 'Jumlah Pesanan', 'Jumlah Eksemplar']
    # for i, col in enumerate(cols_t5):
    #     ws_lap.write(t5_row+1, 6+i, col, fmt_col_name)
        
    # curr_t5_row = t5_row + 2
    # for idx, row in grp_rincian.iterrows():
    #     ws_lap.write(curr_t5_row, 6, row['Nama Produk'], fmt_text_left)
    #     ws_lap.write(curr_t5_row, 7, row['Variasi_Clean'], fmt_num)
    #     ws_lap.write(curr_t5_row, 8, row['Jumlah_Pesanan'], fmt_num)
    #     ws_lap.write(curr_t5_row, 9, row['Jumlah Eksemplar'], fmt_num)
    #     curr_t5_row += 1
    
    # # Total Eksemplar
    # ws_lap.write(curr_t5_row, 8, "TOTAL EKSEMPLAR", fmt_col_name)
    # ws_lap.write(curr_t5_row, 9, grp_rincian['Jumlah Eksemplar'].sum(), fmt_col_name)
    t5_row = curr_t2_row + 2 
    t5_col_start = 7 # H
    
    total_seluruh_pesanan_val = tbl_iklan_data['PESANAN'].sum()
    if not tbl_affiliate_data.empty: total_seluruh_pesanan_val += tbl_affiliate_data['PESANAN'].sum()
    if not tbl_organik_data.empty: total_seluruh_pesanan_val += tbl_organik_data['PESANAN'].sum()
    
    ws_lap.write(t5_row, t5_col_start, 'RINCIAN SELURUH PESANAN', fmt_head_green)
    ws_lap.write(t5_row, t5_col_start+1, total_seluruh_pesanan_val, fmt_header_table)
    ws_lap.merge_range(t5_row, t5_col_start+2, t5_row, t5_col_start+3, "", fmt_header_table)
    
    t5_cols = ['Nama Produk', 'Variasi', 'Jumlah Pesanan', 'Jumlah Eksemplar']
    for i, col in enumerate(t5_cols):
        ws_lap.write(t5_row+1, t5_col_start+i, col, fmt_col_name)
        update_width(t5_col_start+i, col)
        
    curr_t5_row = t5_row + 2
    for idx, row in grp_rincian.iterrows():
        ws_lap.write(curr_t5_row, t5_col_start, row['Nama Produk'], fmt_text_left)
        ws_lap.write(curr_t5_row, t5_col_start+1, row['Variasi_Clean'], fmt_num)
        ws_lap.write(curr_t5_row, t5_col_start+2, row['Jumlah_Pesanan'], fmt_num)
        ws_lap.write(curr_t5_row, t5_col_start+3, row['Jumlah Eksemplar'], fmt_num)
        
        update_width(t5_col_start, row['Nama Produk'])
        update_width(t5_col_start+1, row['Variasi_Clean'])
        curr_t5_row += 1
        
    ws_lap.write(curr_t5_row, t5_col_start+2, "TOTAL EKSEMPLAR", fmt_col_name)
    ws_lap.write(curr_t5_row, t5_col_start+3, grp_rincian['Jumlah Eksemplar'].sum(), fmt_col_name)
    update_width(t5_col_start+2, "TOTAL EKSEMPLAR")
    
    # --- TABEL 6: SUMMARY (M-Q) ---
    # Posisi: 2 baris spasi dibawah Organik
    t6_row = curr_t5_row + 2
    t6_col_start = 7 # H
    
    summary_data = [
        ('Penjualan Keseluruhan', total_omzet_all, fmt_curr),
        ('Total Biaya Iklan Klik', total_biaya_iklan_rinci, fmt_curr),
        ('Total Komisi Affiliate', total_komisi_aff, fmt_curr),
        ('ROASF', roasf, fmt_decimal)
    ]
    
    for label, val, fmt in summary_data:
        # Tentukan format nilai (Currency atau Number/Percent) tapi versi BOLD
        if fmt == fmt_curr:
            use_fmt = fmt_curr_bold
        else:
            use_fmt = fmt_num_bold # Default ke num bold untuk ROASF
            
        ws_lap.merge_range(t6_row, t6_col_start, t6_row, t6_col_start+1, label, fmt_text_left)
        ws_lap.write(t6_row, t6_col_start+2, val, fmt)
        update_width(t6_col_start, label)
        update_width(t6_col_start+2, str(val))
        t6_row += 1

    # --- APPLY AUTO WIDTH ---
    for col_idx, max_len in col_widths.items():
        # Set minimal width 10, max width 50, buffer +2 char
        width = max(10, min(max_len + 2, 50))
        ws_lap.set_column(col_idx, col_idx, width)

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
# with col3:
#     f_seller = st.file_uploader("Upload 'Seller conversion' (csv)", type=['csv'])
with col3:
    # Tambahkan label (Opsional)
    f_seller = st.file_uploader("Upload 'Seller conversion' (csv) - Opsional", type=['csv'])

if st.button("Mulai Proses", type="primary"):
    if f_order and f_iklan:
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
