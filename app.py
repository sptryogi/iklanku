# import streamlit as st
# import pandas as pd
# import io
# import re
# from datetime import datetime

# # --- KONFIGURASI HALAMAN ---
# st.set_page_config(page_title="IklanKu Processor", layout="wide")

# # --- FUNGSI UTAMA (HELPER) ---

# def clean_nama_iklan(text):
#     if not isinstance(text, str):
#         return str(text)
#     # Hapus [angka] di belakang, misal "Produk A [26]" -> "Produk A"
#     return re.sub(r'\s*\[\d+\]\s*$', '', text).strip()

# def extract_time_hour(dt):
#     try:
#         # Asumsi format timestamp pandas
#         return dt.hour
#     except:
#         return 0

# def extract_eksemplar(variasi_text):
#     if not isinstance(variasi_text, str):
#         return 1
#     # Cari angka dalam variasi, misal "PAKET ISI 3" -> 3
#     match = re.search(r'(\d+)', variasi_text)
#     if match:
#         return int(match.group(1))
#     return 1 # Default jika tidak ada angka

# def clean_variasi(text):
#     if not isinstance(text, str) or pd.isna(text) or text == '':
#         return ''
#     # Ambil value di belakang koma, uppercase
#     if ',' in text:
#         parts = text.split(',')
#         return parts[-1].strip().upper()
#     return text.strip().upper()

# # --- LOGIKA PROSES DATA ---

# def process_data(store_name, file_order, file_iklan, file_seller):
#     # 1. LOAD DATA
#     df_order = pd.read_excel(file_order)
#     df_iklan_raw = pd.read_csv(file_iklan, header=None) # Load tanpa header dulu untuk skip baris
#     df_seller = pd.read_csv(file_seller)

#     # 2. PRE-PROCESS ORDER-ALL
#     # Filter Status Pesanan != Batal
#     if 'Status Pesanan' in df_order.columns:
#         df_order = df_order[df_order['Status Pesanan'] != 'Batal'].copy()
    
#     # Konversi kolom waktu
#     if 'Waktu Pesanan Dibuat' in df_order.columns:
#         df_order['Waktu Pesanan Dibuat'] = pd.to_datetime(df_order['Waktu Pesanan Dibuat'])
#         df_order['Jam'] = df_order['Waktu Pesanan Dibuat'].dt.hour
#         # Ambil tanggal untuk header laporan
#         report_date = df_order['Waktu Pesanan Dibuat'].dt.strftime('%A, %d-%m-%Y').iloc[0] if not df_order.empty else "TANGGAL TIDAK DIKETAHUI"
#     else:
#         st.error("Kolom 'Waktu Pesanan Dibuat' tidak ditemukan di Order-all")
#         return None

#     # 3. PRE-PROCESS IKLAN (Sheet 'Iklan klik')
#     # Hapus 7 baris pertama (index 0-6), baris ke-8 (index 7) jadi header
#     new_header = df_iklan_raw.iloc[7]
#     df_iklan = df_iklan_raw[8:].copy()
#     df_iklan.columns = new_header
    
#     # Bersihkan Nama Iklan
#     if 'Nama Iklan' in df_iklan.columns:
#         df_iklan['Nama Iklan'] = df_iklan['Nama Iklan'].apply(clean_nama_iklan)
#         # Hapus Duplikat Nama Iklan
#         df_iklan = df_iklan.drop_duplicates(subset=['Nama Iklan'])
    
#     # Konversi kolom numerik di iklan
#     cols_to_num = ['Dilihat', 'Jumlah Klik', 'Omzet']
#     for col in cols_to_num:
#         if col in df_iklan.columns:
#             # Hapus simbol mata uang atau pemisah ribuan jika ada
#             df_iklan[col] = pd.to_numeric(df_iklan[col].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce').fillna(0)

#     # 4. KATEGORISASI DATA (AFFILIATE, IKLAN, ORGANIK) & HIGHLIGHTING
#     # Setup list untuk tracking
#     list_affiliate_ids = df_seller['Kode Pesanan'].astype(str).tolist() if 'Kode Pesanan' in df_seller.columns else []
#     list_iklan_names = df_iklan['Nama Iklan'].tolist() if 'Nama Iklan' in df_iklan.columns else []

#     # Buat kolom helper di df_order
#     df_order['is_affiliate'] = df_order['No. Pesanan'].astype(str).isin(list_affiliate_ids)
#     df_order['is_iklan_product'] = df_order['Nama Produk'].apply(lambda x: clean_nama_iklan(x) in list_iklan_names)
    
#     # Prioritas: Affiliate > Iklan (match product) > Organik
#     # Namun prompt meminta: "Order all yg termasuk seller conversion (Affiliate)" dan "Diluar Seller Conversion dan Diluar Nama Iklan (Organik)"
#     # Maka Sisanya (Diluar Seller tapi ADA di Nama Iklan) adalah Pesanan Iklan.
    
#     df_affiliate = df_order[df_order['is_affiliate']].copy()
#     df_organic = df_order[(~df_order['is_affiliate']) & (~df_order['is_iklan_product'])].copy()
#     df_ads_orders = df_order[(~df_order['is_affiliate']) & (df_order['is_iklan_product'])].copy()

#     # --- MEMBUAT DATA UNTUK LAPORAN ---

#     # A. TABEL PESANAN IKLAN (Jam 0-23)
#     # Prompt: "diambil dari order all dilihat di rentang jam... sum No Pesanan... Kuantitas... Omzet"
#     # Menggunakan df_ads_orders (Pesanan yg berasal dari produk yg diiklanin)
    
#     hours_df = pd.DataFrame({'Jam': range(24)})
    
#     def agg_by_hour(df_source):
#         if df_source.empty:
#             return pd.DataFrame({'Jam': range(24), 'PESANAN': 0, 'KUANTITAS': 0, 'OMZET PENJUALAN': 0})
        
#         # Hitung Pesanan (Unique No. Pesanan)
#         grp_pesanan = df_source.groupby('Jam')['No. Pesanan'].nunique().reset_index(name='PESANAN')
#         # Hitung Kuantitas & Omzet (Sum)
#         grp_metrics = df_source.groupby('Jam')[['Jumlah', 'Total Harga Produk']].sum().reset_index()
#         grp_metrics.rename(columns={'Jumlah': 'KUANTITAS', 'Total Harga Produk': 'OMZET PENJUALAN'}, inplace=True)
        
#         merged = hours_df.merge(grp_pesanan, on='Jam', how='left').merge(grp_metrics, on='Jam', how='left')
#         return merged.fillna(0)

#     # Tabel 1 Data (Pesanan Iklan)
#     tbl_iklan_data = agg_by_hour(df_ads_orders) # Menggunakan pesanan kategori Iklan
#     # Note: Jika user ingin SEMUA pesanan masuk sini, ganti df_ads_orders dengan df_order. 
#     # Tapi berdasarkan logika tabel Organik, harusnya ini dipisah. Saya gunakan df_ads_orders.

#     # B. TABEL RINCIAN IKLAN KLIK
#     total_dilihat = df_iklan['Dilihat'].sum()
#     total_klik = df_iklan['Jumlah Klik'].sum()
#     # Prompt: "Presentase Klik... dari 'Total Iklan Klik' dibagi 'Total Jumlah Klik'". 
#     # Asumsi typo user: maksudnya Klik / Dilihat (CTR) atau Dilihat/Klik. 
#     # Saya gunakan (Klik / Dilihat) * 100 karena %
#     persentase_klik = (total_klik / total_dilihat) if total_dilihat > 0 else 0
    
#     penjualan_iklan = tbl_iklan_data['OMZET PENJUALAN'].sum()

#     # Hitung Biaya Iklan Spesifik
#     # A5 Koran (Kapital di prompt berarti cek spesifik string atau case sensitive? Prompt: "tapi yang kapital")
#     # Interpretasi: Mengandung "A5 Koran" DAN (Original String mengandung substring kapital atau case sensitive match)
#     # Saya akan filter case sensitive untuk "A5 Koran" vs lower.
    
#     # Helper filter contains
#     def get_omzet_contains(query, case_sensitive=False):
#         if case_sensitive:
#             mask = df_iklan['Nama Iklan'].str.contains(query, case=True, regex=False)
#         else:
#             mask = df_iklan['Nama Iklan'].str.contains(query, case=False, regex=False)
#         return df_iklan[mask]['Omzet'].sum()

#     # "A5 Koran" (Kapital logic - asumsikan mengandung 'A5 KORAN' atau 'A5 Koran' vs 'a5 koran')
#     # Prompt agak ambigu, saya gunakan pendekatan: Mengandung "A5 Koran" (Case sensitive)
#     biaya_a5_koran = get_omzet_contains("A5 Koran", case_sensitive=True) 
    
#     # "A5 Koran Paket 7" (Lowercase logic? Prompt: "tapii yang lowercase")
#     # Saya gunakan pendekatan: Mengandung "a5 koran" (lowercase) tapi TIDAK mengandung "A5 Koran"
#     mask_lower = (df_iklan['Nama Iklan'].str.contains("a5 koran", case=False)) & (~df_iklan['Nama Iklan'].str.contains("A5 Koran", case=True))
#     biaya_a5_koran_pkt7 = df_iklan[mask_lower]['Omzet'].sum()
    
#     # "A6 Pastel" (Case insensitive)
#     biaya_a6_pastel = get_omzet_contains("A6 Pastel", case_sensitive=False)
    
#     # "Komik Pahlawan" (Case insensitive)
#     biaya_komik = get_omzet_contains("Komik Pahlawan", case_sensitive=False)
    
#     total_biaya_iklan_rinci = biaya_a5_koran + biaya_a5_koran_pkt7 + biaya_a6_pastel + biaya_komik
#     roasi = (penjualan_iklan / total_biaya_iklan_rinci) if total_biaya_iklan_rinci > 0 else 0

#     # C. TABEL PESANAN AFFILIATE
#     tbl_affiliate_data = agg_by_hour(df_affiliate)
#     # Tambah kolom Komisi (Perlu mapping jam ke komisi)
#     # Komisi ada di df_seller, tapi df_seller tidak punya 'Jam'. 
#     # Kita harus join df_affiliate (yang punya jam) dengan df_seller (yang punya komisi) based on Order ID
#     df_aff_merged = df_affiliate.merge(df_seller[['Kode Pesanan', 'Pengeluaran(Rp)']], left_on='No. Pesanan', right_on='Kode Pesanan', how='left')
#     # Konversi Pengeluaran ke float
#     df_aff_merged['Pengeluaran(Rp)'] = pd.to_numeric(df_aff_merged['Pengeluaran(Rp)'].astype(str).str.replace('.','').str.replace(',','.'), errors='coerce').fillna(0)
    
#     # Group komisi by jam
#     komisi_per_jam = df_aff_merged.groupby('Jam')['Pengeluaran(Rp)'].sum().reset_index()
#     tbl_affiliate_data = tbl_affiliate_data.merge(komisi_per_jam, on='Jam', how='left').fillna(0)
#     tbl_affiliate_data.rename(columns={'Pengeluaran(Rp)': 'KOMISI'}, inplace=True)

#     # D. TABEL PESANAN ORGANIK
#     tbl_organik_data = agg_by_hour(df_organic)

#     # E. TABEL RINCIAN SELURUH PESANAN (Product Level)
#     # Logika: "jika dalam 1 No. Pesanan ada banyak Nama produk... ambil paling atas saja"
#     # Ini berarti kita dedup by No. Pesanan dulu? Tapi tabel ini adalah rincian produk.
#     # Jika kita dedup by Order ID, kita akan kehilangan data produk lain dalam order yang sama.
#     # NAMUN, prompt bilang "Jumlah Pesanan diambil dari berapa pesanan...".
#     # Interpretasi: Hitung frekuensi unik Order ID per Produk.
#     # TAPI instruksi "ambil paling atas saja" sangat spesifik.
#     # Saya akan lakukan: Group by Order ID -> Ambil first row -> Baru hitung stats produk dari hasil filter ini.
    
#     df_unique_orders = df_order.sort_values(['No. Pesanan', 'Nama Produk']).drop_duplicates(subset=['No. Pesanan'], keep='first').copy()
    
#     # Siapkan kolom variasi bersih
#     df_unique_orders['Variasi_Clean'] = df_unique_orders['Variasi'].apply(clean_variasi)
    
#     # Group by Nama Produk & Variasi
#     grp_rincian = df_unique_orders.groupby(['Nama Produk', 'Variasi_Clean']).agg(
#         Jumlah_Pesanan=('No. Pesanan', 'count')
#     ).reset_index()
    
#     # Hitung Eksemplar
#     grp_rincian['Jumlah Eksemplar'] = grp_rincian.apply(
#         lambda row: extract_eksemplar(row['Variasi_Clean']) * row['Jumlah_Pesanan'], axis=1
#     )

#     # F. TABEL SUMMARY
#     # Hitung total-total untuk summary
#     total_omzet_all = tbl_iklan_data['OMZET PENJUALAN'].sum() + tbl_affiliate_data['OMZET PENJUALAN'].sum() + tbl_organik_data['OMZET PENJUALAN'].sum()
#     total_komisi_aff = tbl_affiliate_data['KOMISI'].sum()
#     roasf = total_omzet_all / (total_biaya_iklan_rinci + total_komisi_aff) if (total_biaya_iklan_rinci + total_komisi_aff) > 0 else 0

#     # --- MEMBUAT FILE EXCEL ---
#     output = io.BytesIO()
#     workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
#     # FORMATS
#     fmt_header_main = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': '#ADD8E6', 'align': 'left', 'valign': 'vcenter'})
#     fmt_header_table = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
#     fmt_date = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
#     fmt_col_name = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#f0f0f0'})
#     fmt_num = workbook.add_format({'border': 1, 'align': 'center'})
#     fmt_curr = workbook.add_format({'border': 1, 'num_format': '#,##0', 'align': 'center'})
#     fmt_percent = workbook.add_format({'border': 1, 'num_format': '0.00%', 'align': 'center'})
    
#     # --- SHEET 1: LAPORAN IKLAN ---
#     ws_lap = workbook.add_worksheet('LAPORAN IKLAN')
    
#     # Judul Utama
#     ws_lap.merge_range('A1:P2', f'LAPORAN IKLAN {store_name.upper()}', fmt_header_main)
    
#     # --- TABEL 1: PESANAN IKLAN (A-E) ---
#     start_row = 3 # Row 4
#     ws_lap.merge_range(start_row, 0, start_row, 4, 'PESANAN IKLAN', fmt_header_table)
#     ws_lap.merge_range(start_row+1, 0, start_row+2, 4, report_date, fmt_date)
    
#     cols_t1 = ['LIHAT', 'KLIK', 'PESANAN', 'KUANTITAS', 'OMZET PENJUALAN']
#     for i, col in enumerate(cols_t1):
#         ws_lap.write(start_row+3, i, col, fmt_col_name)
        
#     # Isi Data Tabel 1
#     row_cursor = start_row + 4
#     for idx, row in tbl_iklan_data.iterrows():
#         ws_lap.write(row_cursor, 0, "", fmt_num) # Lihat (Kosong)
#         ws_lap.write(row_cursor, 1, "", fmt_num) # Klik (Kosong)
#         ws_lap.write(row_cursor, 2, row['PESANAN'], fmt_num)
#         ws_lap.write(row_cursor, 3, row['KUANTITAS'], fmt_num)
#         ws_lap.write(row_cursor, 4, row['OMZET PENJUALAN'], fmt_curr)
#         row_cursor += 1
        
#     # Total Tabel 1
#     ws_lap.write(row_cursor, 0, "TOTAL", fmt_col_name)
#     ws_lap.write(row_cursor, 1, "", fmt_col_name)
#     ws_lap.write(row_cursor, 2, tbl_iklan_data['PESANAN'].sum(), fmt_col_name)
#     ws_lap.write(row_cursor, 3, tbl_iklan_data['KUANTITAS'].sum(), fmt_col_name)
#     ws_lap.write(row_cursor, 4, tbl_iklan_data['OMZET PENJUALAN'].sum(), fmt_col_name)

#     # --- TABEL 2: RINCIAN IKLAN KLIK (G-H) ---
#     # Posisi sejajar dengan PESANAN IKLAN
#     t2_row = start_row
#     ws_lap.merge_range(t2_row, 6, t2_row, 7, 'RINCIAN IKLAN KLIK', fmt_header_table)
#     # Item rincian
#     rincian_items = [
#         ('Total Iklan Dilihat', total_dilihat, fmt_num),
#         ('Total Jumlah Klik', total_klik, fmt_num),
#         ('Presentase Klik', persentase_klik, fmt_percent),
#         ('Penjualan Iklan', penjualan_iklan, fmt_curr),
#         ('Biaya Iklan A5 Koran', biaya_a5_koran, fmt_curr),
#         ('Biaya Iklan A5 Koran Paket 7', biaya_a5_koran_pkt7, fmt_curr),
#         ('Biaya Iklan A6 Pastel', biaya_a6_pastel, fmt_curr),
#         ('Biaya Iklan Komik Pahlawan', biaya_komik, fmt_curr),
#         ('ROASI', roasi, fmt_num) # ROAS biasanya desimal/ratio
#     ]
    
#     curr_t2_row = t2_row + 1
#     for label, val, fmt in rincian_items:
#         ws_lap.write(curr_t2_row, 6, label, fmt_num)
#         ws_lap.write(curr_t2_row, 7, val, fmt)
#         curr_t2_row += 1

#     # --- TABEL 3: PESANAN AFFILIATE (L-P) ---
#     t3_row = start_row
#     ws_lap.merge_range(t3_row, 11, t3_row, 15, 'PESANAN AFFILIATE', fmt_header_table)
#     cols_t3 = ['Jam', 'Pesanan', 'Kuantitas', 'Omzet Penjualan', 'Komisi']
#     for i, col in enumerate(cols_t3):
#         ws_lap.write(t3_row+1, 11+i, col, fmt_col_name)
        
#     curr_t3_row = t3_row + 2
#     for idx, row in tbl_affiliate_data.iterrows():
#         ws_lap.write(curr_t3_row, 11, f"{int(row['Jam']):02d}:00", fmt_num)
#         ws_lap.write(curr_t3_row, 12, row['PESANAN'], fmt_num)
#         ws_lap.write(curr_t3_row, 13, row['KUANTITAS'], fmt_num)
#         ws_lap.write(curr_t3_row, 14, row['OMZET PENJUALAN'], fmt_curr)
#         ws_lap.write(curr_t3_row, 15, row['KOMISI'], fmt_curr)
#         curr_t3_row += 1
        
#     # Total T3
#     total_omzet_aff = tbl_affiliate_data['OMZET PENJUALAN'].sum()
#     total_komisi_aff_val = tbl_affiliate_data['KOMISI'].sum()
#     ws_lap.write(curr_t3_row, 11, "TOTAL", fmt_col_name)
#     ws_lap.write(curr_t3_row, 12, tbl_affiliate_data['PESANAN'].sum(), fmt_col_name)
#     ws_lap.write(curr_t3_row, 13, tbl_affiliate_data['KUANTITAS'].sum(), fmt_col_name)
#     ws_lap.write(curr_t3_row, 14, total_omzet_aff, fmt_col_name)
#     ws_lap.write(curr_t3_row, 15, total_komisi_aff_val, fmt_col_name)
    
#     # ROASA
#     roasa = total_omzet_aff / total_komisi_aff_val if total_komisi_aff_val > 0 else 0
#     curr_t3_row += 1
#     ws_lap.write(curr_t3_row, 11, "ROASA", fmt_col_name)
#     ws_lap.merge_range(curr_t3_row, 12, curr_t3_row, 13, "", fmt_num)
#     ws_lap.write(curr_t3_row, 14, roasa, fmt_num)
#     ws_lap.write(curr_t3_row, 15, "", fmt_num)
    
#     last_row_affiliate = curr_t3_row

#     # --- TABEL 4: PESANAN ORGANIK (L-O) ---
#     # Dibawah Affiliate
#     t4_row = last_row_affiliate + 2
#     ws_lap.merge_range(t4_row, 11, t4_row, 14, 'PESANAN ORGANIK', fmt_header_table)
#     cols_t4 = ['Jam', 'Pesanan', 'Kuantitas', 'Omzet Penjualan']
#     for i, col in enumerate(cols_t4):
#         ws_lap.write(t4_row+1, 11+i, col, fmt_col_name)
        
#     curr_t4_row = t4_row + 2
#     for idx, row in tbl_organik_data.iterrows():
#         ws_lap.write(curr_t4_row, 11, f"{int(row['Jam']):02d}:00", fmt_num)
#         ws_lap.write(curr_t4_row, 12, row['PESANAN'], fmt_num)
#         ws_lap.write(curr_t4_row, 13, row['KUANTITAS'], fmt_num)
#         ws_lap.write(curr_t4_row, 14, row['OMZET PENJUALAN'], fmt_curr)
#         curr_t4_row += 1
        
#     # Total T4
#     ws_lap.write(curr_t4_row, 11, "TOTAL", fmt_col_name)
#     ws_lap.write(curr_t4_row, 12, tbl_organik_data['PESANAN'].sum(), fmt_col_name)
#     ws_lap.write(curr_t4_row, 13, tbl_organik_data['KUANTITAS'].sum(), fmt_col_name)
#     ws_lap.write(curr_t4_row, 14, tbl_organik_data['OMZET PENJUALAN'].sum(), fmt_col_name)
    
#     last_row_organik = curr_t4_row

#     # --- TABEL 5: RINCIAN SELURUH PESANAN (G-J) ---
#     # Posisi: Sejajar PESANAN ORGANIK (Row 13-14 dari L, berarti t4_row)
#     # Prompt: "sejajar dengan tabel PESANAN ORGANIK... baris 13-14" (asumsi relatif terhadap layout)
#     t5_row = t4_row 
    
#     total_seluruh_pesanan_val = tbl_iklan_data['PESANAN'].sum() + tbl_affiliate_data['PESANAN'].sum() + tbl_organik_data['PESANAN'].sum()
    
#     ws_lap.write(t5_row, 6, 'RINCIAN SELURUH PESANAN', fmt_header_table)
#     ws_lap.write(t5_row, 7, total_seluruh_pesanan_val, fmt_header_table) # Col H
#     ws_lap.merge_range(t5_row, 8, t5_row, 9, "", fmt_header_table)
    
#     cols_t5 = ['Nama Produk', 'Variasi', 'Jumlah Pesanan', 'Jumlah Eksemplar']
#     for i, col in enumerate(cols_t5):
#         ws_lap.write(t5_row+1, 6+i, col, fmt_col_name)
        
#     curr_t5_row = t5_row + 2
#     for idx, row in grp_rincian.iterrows():
#         ws_lap.write(curr_t5_row, 6, row['Nama Produk'], fmt_num)
#         ws_lap.write(curr_t5_row, 7, row['Variasi_Clean'], fmt_num)
#         ws_lap.write(curr_t5_row, 8, row['Jumlah_Pesanan'], fmt_num)
#         ws_lap.write(curr_t5_row, 9, row['Jumlah Eksemplar'], fmt_num)
#         curr_t5_row += 1
    
#     # Total Eksemplar
#     ws_lap.write(curr_t5_row, 8, "TOTAL EKSEMPLAR", fmt_col_name)
#     ws_lap.write(curr_t5_row, 9, grp_rincian['Jumlah Eksemplar'].sum(), fmt_col_name)
    
#     # --- TABEL 6: SUMMARY (L-P) ---
#     # Posisi: 2 baris spasi dibawah Organik
#     t6_row = last_row_organik + 3
    
#     summary_data = [
#         ('Penjualan Keseluruhan', total_omzet_all, fmt_curr),
#         ('Total Biaya Iklan Klik', total_biaya_iklan_rinci, fmt_curr),
#         ('Total Komisi Affiliate', total_komisi_aff, fmt_curr),
#         ('ROASF', roasf, fmt_num)
#     ]
    
#     for label, val, fmt in summary_data:
#         ws_lap.merge_range(t6_row, 11, t6_row, 14, label, fmt_num)
#         ws_lap.write(t6_row, 15, val, fmt)
#         t6_row += 1

#     # --- SIMPAN SHEET LAINNYA ---
#     # 1. order-all (dengan highlight)
#     df_order.to_excel(workbook, sheet_name='order-all', index=False)
#     ws_order = workbook.get_worksheet_by_name('order-all')
    
#     # Format Highlight
#     fmt_yellow = workbook.add_format({'bg_color': '#FFFF00'})
#     fmt_pink = workbook.add_format({'bg_color': '#FFC0CB'})
    
#     # Iterasi untuk highlight
#     # Note: XlsxWriter tidak bisa overwrite format cell dengan mudah tanpa menulis ulang
#     # Kita loop df_order untuk nulis ulang baris dengan format yang sesuai
    
#     # Get columns for rewrite
#     columns = df_order.columns.tolist()
    
#     for row_idx, row_data in df_order.iterrows():
#         row_fmt = None
#         if row_data['is_affiliate']:
#             row_fmt = fmt_yellow
#         # Kondisi Pink: Diluar Affiliate (sudah checked via elif/logic) DAN Diluar Iklan
#         # Logika Highlight Pink: "diluar yang termasuk Seller conversion dan diluar dari Nama Produk yang tidak ada di Nama Iklan"
#         elif not row_data['is_iklan_product']:
#             row_fmt = fmt_pink
            
#         if row_fmt:
#             for col_idx, col_name in enumerate(columns):
#                 # Tulis ulang cell dengan format. Handle datetime convert
#                 val = row_data[col_name]
#                 # Sederhanakan penulisan (XlsxWriter handle basic types)
#                 if pd.isna(val):
#                     ws_order.write(row_idx + 1, col_idx, "", row_fmt)
#                 else:
#                      # Check if datetime
#                     if isinstance(val, pd.Timestamp):
#                         ws_order.write_datetime(row_idx + 1, col_idx, val, workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm', 'bg_color': row_fmt.bg_color}))
#                     else:
#                         ws_order.write(row_idx + 1, col_idx, val, row_fmt)

#     # 2. Iklan klik (Cleaned)
#     df_iklan.to_excel(workbook, sheet_name='Iklan klik', index=False)
    
#     # 3. Seller conversion (Raw)
#     df_seller.to_excel(workbook, sheet_name='Seller conversion', index=False)

#     workbook.close()
#     output.seek(0)
#     return output

# # --- INTERFACE STREAMLIT ---

# st.title("🛒 IklanKu - Generator Laporan Otomatis")
# st.markdown("---")

# # Input Toko
# toko = st.selectbox("Pilih Toko:", ["Human Store", "Pasific BookStore", "Dama Store"])

# # Input File
# col1, col2, col3 = st.columns(3)
# with col1:
#     f_order = st.file_uploader("Upload 'Order-all' (xlsx)", type=['xlsx'])
# with col2:
#     f_iklan = st.file_uploader("Upload 'Iklan Keseluruhan' (csv)", type=['csv'])
# with col3:
#     f_seller = st.file_uploader("Upload 'Seller conversion' (csv)", type=['csv'])

# if st.button("Mulai Proses", type="primary"):
#     if f_order and f_iklan and f_seller:
#         with st.spinner('Sedang memproses data... Tunggu sebentar ya...'):
#             try:
#                 excel_file = process_data(toko, f_order, f_iklan, f_seller)
                
#                 if excel_file:
#                     st.success("Proses Selesai!")
#                     st.download_button(
#                         label="📥 Download Laporan Excel",
#                         data=excel_file,
#                         file_name=f"LAPORAN_IKLAN_{toko.replace(' ', '_').upper()}.xlsx",
#                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#                     )
#             except Exception as e:
#                 st.error(f"Terjadi kesalahan saat memproses: {e}")
#                 st.write(e) # Debug info
#     else:
#         st.warning("Mohon upload ketiga file terlebih dahulu.")

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime
from dateutil import parser

st.set_page_config(page_title="IklanKu", layout="wide")

st.title("IklanKu — Laporan Iklan Harian (1-Click)")

# ---------- Helper: flexible column lookup ----------
def find_col(df, candidates):
    """Return first existing column name from candidates list or None."""
    for c in candidates:
        if c in df.columns:
            return c
    # try case-insensitive
    lowcols = {col.lower(): col for col in df.columns}
    for c in candidates:
        if c.lower() in lowcols:
            return lowcols[c.lower()]
    return None

def clean_ad_name(name):
    if pd.isna(name):
        return ""
    s = str(name).strip()
    # remove trailing [digits] or (digits)
    s = re.sub(r'\s*\[\d+\]\s*$', '', s)
    s = re.sub(r'\s*\(\d+\)\s*$', '', s)
    return s.strip()

def safe_to_numeric_series(s):
    return (s.astype(str)
            .str.replace(r'[^\d\-,\.]', '', regex=True)
            .str.replace(',', '.', regex=False)
            .replace('', '0')
            .astype(float).fillna(0))

# ---------- UI: select store + upload ----------
store = st.selectbox("Pilih Toko", ["", "Human Store", "Pasific BookStore", "Dama Store"])

col1, col2, col3 = st.columns(3)
with col1:
    uploaded_order = st.file_uploader("1️⃣ Upload file Order-all (xlsx)", type=["xlsx"])
with col2:
    uploaded_iklan = st.file_uploader("2️⃣ Upload file Iklan Keseluruhan (CSV/XLSX)", type=["csv","xlsx"])
with col3:
    uploaded_seller = st.file_uploader("3️⃣ Upload file Seller conversion (CSV/XLSX)", type=["csv","xlsx"])

st.markdown("---")
if not store:
    st.info("Pilih toko dulu.")
    st.stop()

if not (uploaded_order and uploaded_iklan and uploaded_seller):
    st.info("Unggah ketiga file yang diminta (Order-all, Iklan, Seller conversion).")
    st.stop()

# ---------- Read uploaded files (robust) ----------
@st.cache_data
def read_excel_maybe(file):
    try:
        return pd.read_excel(file)
    except Exception:
        try:
            return pd.read_excel(file, sheet_name=0)
        except Exception:
            return pd.DataFrame()

@st.cache_data
def read_csv_maybe(file):
    try:
        return pd.read_csv(file)
    except Exception:
        try:
            return pd.read_csv(file, encoding='latin1')
        except Exception:
            return pd.DataFrame()

# load order-all
try:
    df_order = read_excel_maybe(uploaded_order)
except Exception as e:
    st.error(f"Gagal baca order-all: {e}")
    st.stop()

# load iklan
if uploaded_iklan.name.lower().endswith(".csv"):
    df_iklan = read_csv_maybe(uploaded_iklan)
else:
    df_iklan = read_excel_maybe(uploaded_iklan)

# load seller conversion
if uploaded_seller.name.lower().endswith(".csv"):
    df_seller = read_csv_maybe(uploaded_seller)
else:
    df_seller = read_excel_maybe(uploaded_seller)

st.success("File berhasil dimuat. Periksa mapping kolom bila perlu di log.")

# ---------- Column mapping (fallbacks) ----------
# Jika header file berbeda, sesuaikan candidates list
COL_MAP = {
    "no_pesanan": ["No. Pesanan", "Order ID", "Order ID ", "No Pesanan", "OrderID", "Order_ID"],
    "waktu_pesanan": ["Waktu Pesanan Dibuat", "Waktu Pesanan", "create_time", "Waktu", "Waktu Dibuat", "Waktu Pesanan Dibuat "],
    "jumlah": ["Jumlah", "Quantity", "Qty", "Jumlah Terjual"],
    "harga_setelah_diskon": ["Harga Setelah Diskon", "SKU Subtotal Before Discount", "Harga"],
    "total_harga_produk": ["Total Harga Produk", "Total Harga", "Total"],
    "nama_produk": ["Nama Produk", "Product Name", "ProductName", "Nama Produk "],
    "variasi": ["Variation", "Variasi", "Variant", "Varian"],
    "status": ["Status Pesanan", "Order Status", "Order_Status", "Status"]
}

# find actual col names for order df
order_cols = {}
for key, cand in COL_MAP.items():
    order_cols[key] = find_col(df_order, cand)

# Log missing columns
missing = [k for k,v in order_cols.items() if v is None and k in ["no_pesanan","waktu_pesanan","jumlah","total_harga_produk","nama_produk"]]
if missing:
    st.warning(f"Kemungkinan kolom penting tidak ditemukan di order-all: {missing}. Silakan cek nama kolom file kamu.")

# find iklan cols
iklan_map = {
    "nama_iklan": ["Nama Iklan","Ad name","Ad Name","Campaign Name"],
    "dilihat": ["Dilihat","Impressions","Impression"],
    "jumlah_klik": ["Jumlah Klik","Clicks","Click"],
    "biaya": ["Biaya","Cost","Biaya Iklan","Biaya (Rp)"],
    "produk_terjual": ["Produk Terjual","Products Sold"],
    "omzet": ["Omzet Penjualan","Omzet","Omzet Penjualan "]
}
iklan_cols = {}
for k,cand in iklan_map.items():
    iklan_cols[k] = find_col(df_iklan, cand)

# seller conversion mapping
seller_map = {
    "kode_pesanan": ["Kode Pesanan","Order ID","OrderID","No. Pesanan"],
    "pengeluaran": ["Pengeluaran(Rp)","Pengeluaran (Rp)","Pengeluaran","Amount","Pengeluaran(Rp) "]
}
seller_cols = {}
for k,cand in seller_map.items():
    seller_cols[k] = find_col(df_seller, cand)

# ------- Preprocessing order-all -------
# drop canceled orders if status column exists
if order_cols.get("status"):
    df_order = df_order[~df_order[order_cols["status"]].astype(str).str.contains("Batal", case=False, na=False)]

# ensure essential columns exist; create defaults if not
if order_cols.get("no_pesanan") is None:
    df_order.insert(0, "No. Pesanan", np.nan)
    order_cols["no_pesanan"] = "No. Pesanan"
if order_cols.get("waktu_pesanan") is None:
    df_order.insert(1, "Waktu Pesanan Dibuat", np.nan)
    order_cols["waktu_pesanan"] = "Waktu Pesanan Dibuat"
if order_cols.get("jumlah") is None:
    df_order["Jumlah"] = 1
    order_cols["jumlah"] = "Jumlah"
if order_cols.get("total_harga_produk") is None:
    df_order["Total Harga Produk"] = 0
    order_cols["total_harga_produk"] = "Total Harga Produk"
if order_cols.get("nama_produk") is None:
    df_order["Nama Produk"] = ""
    order_cols["nama_produk"] = "Nama Produk"
if order_cols.get("variasi") is None:
    df_order["Variasi"] = ""
    order_cols["variasi"] = "Variasi"

# normalize datetime & extract hour
def parse_datetime_safe(x):
    if pd.isna(x): return pd.NaT
    if isinstance(x, (pd.Timestamp, datetime)): return pd.to_datetime(x)
    try:
        return pd.to_datetime(x)
    except Exception:
        try:
            return parser.parse(str(x))
        except Exception:
            return pd.NaT

df_order[order_cols["waktu_pesanan"]] = df_order[order_cols["waktu_pesanan"]].apply(parse_datetime_safe)
df_order["__hour"] = df_order[order_cols["waktu_pesanan"]].dt.hour.fillna(-1).astype(int)
df_order["__hour_label"] = df_order[order_cols["waktu_pesanan"]].dt.strftime("%H:%M").fillna("")

# ensure numeric columns
df_order[order_cols["jumlah"]] = pd.to_numeric(df_order[order_cols["jumlah"]].astype(str).str.replace(r'[^\d\-\.]', '', regex=True), errors='coerce').fillna(0).astype(int)
df_order[order_cols["total_harga_produk"]] = pd.to_numeric(df_order[order_cols["total_harga_produk"]].astype(str).str.replace(r'[^\d\-\.]', '', regex=True), errors='coerce').fillna(0)

# ------- Preprocessing iklan -------
# if first rows are headers to drop (7), remove them when csv has header rows of metadata
# Heuristic: if iklan file has top rows that are not numeric in Dilihat, drop first 7 if necessary
if iklan_cols.get("dilihat") and df_iklan.shape[0] > 10:
    # if first real row contains non-numeric in numeric column, drop first 7
    first_val = str(df_iklan.iloc[0][iklan_cols["dilihat"]])
    if not re.search(r'\d', first_val):
        df_iklan = df_iklan.iloc[7:].reset_index(drop=True)

# clean nama iklan and drop duplicates by cleaned name (keep first)
if iklan_cols.get("nama_iklan"):
    df_iklan["__nama_iklan_clean"] = df_iklan[iklan_cols["nama_iklan"]].astype(str).apply(clean_ad_name)
    # aggregate summing numeric metrics by cleaned name
    # unify numeric columns
    if iklan_cols.get("dilihat"):
        df_iklan["__dilihat_num"] = pd.to_numeric(df_iklan[iklan_cols["dilihat"]].astype(str).str.replace(r'[^\d\-\.]', '', regex=True), errors='coerce').fillna(0)
    else:
        df_iklan["__dilihat_num"] = 0
    if iklan_cols.get("jumlah_klik"):
        df_iklan["__klik_num"] = pd.to_numeric(df_iklan[iklan_cols["jumlah_klik"]].astype(str).str.replace(r'[^\d\-\.]', '', regex=True), errors='coerce').fillna(0)
    else:
        df_iklan["__klik_num"] = 0
    if iklan_cols.get("omzet"):
        df_iklan["__omzet_num"] = pd.to_numeric(df_iklan[iklan_cols["omzet"]].astype(str).str.replace(r'[^\d\-\.]', '', regex=True), errors='coerce').fillna(0)
    else:
        df_iklan["__omzet_num"] = 0
    # group by clean name
    iklan_agg = df_iklan.groupby("__nama_iklan_clean").agg({
        "__dilihat_num": "sum",
        "__klik_num": "sum",
        "__omzet_num": "sum"
    }).reset_index().rename(columns={
        "__nama_iklan_clean":"Nama Iklan",
        "__dilihat_num":"Dilihat",
        "__klik_num":"Jumlah Klik",
        "__omzet_num":"Omzet"
    })
else:
    iklan_agg = pd.DataFrame(columns=["Nama Iklan","Dilihat","Jumlah Klik","Omzet"])

# ------- Preprocessing seller conversion -------
# normalize seller columns
if seller_cols.get("pengeluaran"):
    df_seller[seller_cols["pengeluaran"]] = pd.to_numeric(df_seller[seller_cols["pengeluaran"]].astype(str).str.replace(r'[^\d\-\.]', '', regex=True), errors='coerce').fillna(0)
else:
    df_seller["Pengeluaran(Rp)"] = 0
    seller_cols["pengeluaran"] = "Pengeluaran(Rp)"
if seller_cols.get("kode_pesanan") is None:
    # try common column names
    possible = ["Order ID","Order ID ","No. Pesanan","OrderID","Kode Pesanan"]
    c = find_col(df_seller, possible)
    if c:
        seller_cols["kode_pesanan"] = c
    else:
        df_seller["Kode Pesanan"] = np.nan
        seller_cols["kode_pesanan"] = "Kode Pesanan"

# standardize order/no columns to string for matching
df_order[order_cols["no_pesanan"]] = df_order[order_cols["no_pesanan"]].astype(str).str.strip()
df_seller[seller_cols["kode_pesanan"]] = df_seller[seller_cols["kode_pesanan"]].astype(str).str.strip()

# ---------- Compute tables ----------
# LAPORAN IKLAN: PESANAN IKLAN (24 rows)
hours = list(range(24))
pesanan_data = []
for h in hours:
    mask = df_order["__hour"] == h
    # count distinct No. Pesanan in that hour
    pesanan_count = df_order[mask][order_cols["no_pesanan"]].nunique()
    kuantitas_sum = df_order[mask][order_cols["jumlah"]].sum()
    omzet_sum = df_order[mask][order_cols["total_harga_produk"]].sum()
    pesanan_data.append({"Jam": f"{h:02d}:00", "LIHAT": "", "KLIK":"", "PESANAN": int(pesanan_count), "KUANTITAS": int(kuantitas_sum), "OMZET PENJUALAN": float(omzet_sum)})

df_pesanan_iklan = pd.DataFrame(pesanan_data)
total_row = {
    "Jam":"TOTAL",
    "LIHAT":"",
    "KLIK":"",
    "PESANAN": df_pesanan_iklan["PESANAN"].sum(),
    "KUANTITAS": df_pesanan_iklan["KUANTITAS"].sum(),
    "OMZET PENJUALAN": df_pesanan_iklan["OMZET PENJUALAN"].sum()
}
df_pesanan_iklan = pd.concat([df_pesanan_iklan, pd.DataFrame([total_row])], ignore_index=True)

# RINCIAN IKLAN KLIK (G-H)
total_dilihat = iklan_agg["Dilihat"].sum() if not iklan_agg.empty else 0
total_klik = iklan_agg["Jumlah Klik"].sum() if not iklan_agg.empty else 0
presentase_klik = (total_klik / total_dilihat * 100) if total_dilihat else 0
penjualan_iklan = df_pesanan_iklan.loc[df_pesanan_iklan['Jam']!="TOTAL","OMZET PENJUALAN"].sum()
# product-specific sums by keyword
def sum_omzet_by_keyword(keyword):
    if iklan_agg.empty: return 0
    mask = iklan_agg["Nama Iklan"].str.contains(keyword, case=False, na=False)
    return iklan_agg.loc[mask,"Omzet"].sum()
biaya_a5_koran = sum_omzet_by_keyword(r'\bA5 KORAN\b')  # exact caps - try word boundary
biaya_a5_koran_pkg7 = sum_omzet_by_keyword(r'\ba5 koran\b')  # lowercase variant
biaya_a6_pastel = sum_omzet_by_keyword(r'A6 Pastel')
biaya_komik_pahlawan = sum_omzet_by_keyword(r'Komik Pahlawan')

roasi_n = penjualan_iklan
roasi_d = (biaya_a5_koran + biaya_a5_koran_pkg7 + biaya_a6_pastel + biaya_komik_pahlawan)
roasi = (roasi_n/roasi_d) if roasi_d else np.nan

df_rincian_iklan = pd.DataFrame({
    "Label": [
        "Total Iklan Dilihat", "Total Jumlah Klik", "Presentase Klik",
        "Penjualan Iklan", "Biaya Iklan A5 Koran","Biaya Iklan A5 Koran Paket 7",
        "Biaya Iklan A6 Pastel","Biaya Iklan Komik Pahlawan","ROASI"
    ],
    "Value": [
        int(total_dilihat), int(total_klik), f"{presentase_klik:.2f}%", float(penjualan_iklan),
        float(biaya_a5_koran), float(biaya_a5_koran_pkg7),
        float(biaya_a6_pastel), float(biaya_komik_pahlawan), float(roasi) if not np.isnan(roasi) else ""
    ]
})

# PESANAN AFFILIATE: orders whose No. Pesanan in seller conversion
affiliate_orders = df_order[df_order[order_cols["no_pesanan"]].isin(df_seller[seller_cols["kode_pesanan"]])]
# join with seller to get Pengeluaran
affiliate_orders = affiliate_orders.merge(df_seller[[seller_cols["kode_pesanan"], seller_cols["pengeluaran"]]].rename(columns={seller_cols["kode_pesanan"]:"Kode Pesanan", seller_cols["pengeluaran"]:"Pengeluaran(Rp)"}),
                                          left_on=order_cols["no_pesanan"], right_on="Kode Pesanan", how="left")
# prepare affiliate table by hour
aff_pivot = affiliate_orders.groupby("__hour").agg({
    order_cols["no_pesanan"] : pd.Series.nunique,
    order_cols["jumlah"]: "sum",
    order_cols["total_harga_produk"]: "sum",
    "Pengeluaran(Rp)": "sum"
}).reindex(hours, fill_value=0).reset_index().rename(columns={
    "__hour":"Hour",
    order_cols["no_pesanan"]:"Pesanan",
    order_cols["jumlah"]:"Kuantitas",
    order_cols["total_harga_produk"]:"Omzet Penjualan",
    "Pengeluaran(Rp)":"Komisi"
})
aff_pivot["Jam"] = aff_pivot["Hour"].apply(lambda x: f"{int(x):02d}:00")
aff_table = aff_pivot[["Jam","Pesanan","Kuantitas","Omzet Penjualan","Komisi"]]
# append total
aff_total = {"Jam":"TOTAL","Pesanan":int(aff_table["Pesanan"].sum()),"Kuantitas":int(aff_table["Kuantitas"].sum()),"Omzet Penjualan":float(aff_table["Omzet Penjualan"].sum()),"Komisi":float(aff_table["Komisi"].sum())}
aff_table = pd.concat([aff_table, pd.DataFrame([aff_total])], ignore_index=True)

# PESANAN ORGANIK: orders not in affiliate and not belonging to iklan product names
iklan_product_names = set()
# attempt to collect product names from iklan (if available in iklan file)
# in many flows there's no product name in iklan file; we fallback to matching by keywords later
if "Produk Terjual" in df_iklan.columns:
    # not reliable; leave empty
    pass

# define iklan product keywords from iklan_agg names: take tokens except common words
iklan_keywords = [name for name in iklan_agg["Nama Iklan"].tolist() if name]

def is_order_in_iklan_products(prod_name):
    if not prod_name or not iklan_keywords:
        return False
    s = str(prod_name).lower()
    return any(kw.lower() in s for kw in iklan_keywords)

mask_affiliate = df_order[order_cols["no_pesanan"]].isin(df_seller[seller_cols["kode_pesanan"]])
mask_iklanprod = df_order[order_cols["nama_produk"]].apply(is_order_in_iklan_products)
organic_orders = df_order[~mask_affiliate & ~mask_iklanprod]

org_pivot = organic_orders.groupby("__hour").agg({
    order_cols["no_pesanan"]: pd.Series.nunique,
    order_cols["jumlah"]:"sum",
    order_cols["total_harga_produk"]:"sum"
}).reindex(hours, fill_value=0).reset_index().rename(columns={
    "__hour":"Hour",
    order_cols["no_pesanan"]:"Pesanan",
    order_cols["jumlah"]:"Kuantitas",
    order_cols["total_harga_produk"]:"Omzet Penjualan"
})
org_pivot["Jam"] = org_pivot["Hour"].apply(lambda x: f"{int(x):02d}:00")
org_table = org_pivot[["Jam","Pesanan","Kuantitas","Omzet Penjualan"]]
org_total = {"Jam":"TOTAL","Pesanan":int(org_table["Pesanan"].sum()),"Kuantitas":int(org_table["Kuantitas"].sum()),"Omzet Penjualan":float(org_table["Omzet Penjualan"].sum())}
org_table = pd.concat([org_table, pd.DataFrame([org_total])], ignore_index=True)

# RINCIAN SELURUH PESANAN (G-J): group by Nama Produk + Variasi (only orders)
# For variation, keep substring after comma and uppercase per spec
def extract_variation_suffix(v):
    if pd.isna(v): return ""
    s = str(v)
    if "," in s:
        suffix = s.split(",")[-1].strip().upper()
        return suffix
    return ""  # empty if no variation

df_order["_VAR_SUFFIX"] = df_order[order_cols["variasi"]].apply(extract_variation_suffix)
rincian = df_order.groupby([order_cols["nama_produk"], "_VAR_SUFFIX"]).agg({
    order_cols["no_pesanan"]: pd.Series.nunique,
    order_cols["jumlah"]:"sum"
}).reset_index().rename(columns={
    order_cols["nama_produk"]:"Nama Produk",
    "_VAR_SUFFIX":"Variasi",
    order_cols["no_pesanan"]:"Jumlah Pesanan",
    order_cols["jumlah"]:"Jumlah Eksemplar"
})
# For Jumlah Eksemplar: if variation contains 'PAKET ISI N' etc, extract number multiplier
def extract_pkg_multiplier(var_text):
    if not var_text: return 1
    m = re.search(r'(\d+)', var_text)
    return int(m.group(1)) if m else 1

rincian["Multiplier"] = rincian["Variasi"].apply(extract_pkg_multiplier)
rincian["Jumlah Eksemplar"] = rincian["Jumlah Pesanan"] * rincian["Multiplier"]
rincian_total = {"Nama Produk":"TOTAL","Variasi":"","Jumlah Pesanan":int(rincian["Jumlah Pesanan"].sum()),"Jumlah Eksemplar":int(rincian["Jumlah Eksemplar"].sum()),"Multiplier":""}
rincian_output = rincian[["Nama Produk","Variasi","Jumlah Pesanan","Jumlah Eksemplar"]]
rincian_output = pd.concat([rincian_output, pd.DataFrame([rincian_total])], ignore_index=True)

# SUMMARY table under L-P
penjualan_keseluruhan = df_pesanan_iklan.loc[df_pesanan_iklan['Jam']!="TOTAL","OMZET PENJUALAN"].sum() + aff_table.loc[aff_table['Jam']!="TOTAL","Omzet Penjualan"].sum() + org_table.loc[org_table['Jam']!="TOTAL","Omzet Penjualan"].sum()
total_biaya_iklan_klik = df_rincian_iklan.loc[df_rincian_iklan["Label"].isin(["Biaya Iklan A5 Koran","Biaya Iklan A5 Koran Paket 7","Biaya Iklan A6 Pastel","Biaya Iklan Komik Pahlawan"]),"Value"].sum()
total_komisi_affiliate = aff_table.loc[aff_table["Jam"]=="TOTAL","Komisi"].sum()
roasf = (penjualan_keseluruhan / (total_biaya_iklan_klik + total_komisi_affiliate)) if (total_biaya_iklan_klik + total_komisi_affiliate) else np.nan

summary_rows = [
    ("Penjualan Keseluruhan", float(penjualan_keseluruhan)),
    ("Total Biaya Iklan Klik", float(total_biaya_iklan_klik)),
    ("Total Komisi Affiliate", float(total_komisi_affiliate)),
    ("ROASF", float(roasf) if not np.isnan(roasf) else "")
]

# Highlight rules for order-all: (1) rows where No. Pesanan in Seller Conversion -> yellow
# (2) rows outside seller conv AND whose Nama Produk not in iklan product names -> pink
df_order["_HIGHLIGHT"] = ""
in_seller_set = set(df_seller[seller_cols["kode_pesanan"]].astype(str))
iklan_name_set = set(iklan_agg["Nama Iklan"].astype(str))
for idx, row in df_order.iterrows():
    no = str(row[order_cols["no_pesanan"]])
    prod = str(row[order_cols["nama_produk"]])
    if no in in_seller_set:
        df_order.at[idx, "_HIGHLIGHT"] = "YELLOW"
    else:
        # if product not matched to any iklan keywords -> pink
        if not any(k.lower() in prod.lower() for k in iklan_name_set if k):
            df_order.at[idx, "_HIGHLIGHT"] = "PINK"

# ---------- Build Excel file using xlsxwriter ----------
output = io.BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    workbook = writer.book

    # 1) LAPORAN IKLAN sheet - we'll write multiple ranges
    sheet_name = "LAPORAN IKLAN"
    # create empty sheet and write headers/tables at desired positions
    # We'll position tables roughly; user may adjust positions if needed
    # Write main title merged A1:P2
    ws = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws

    # Formats
    fmt_title = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': '#cfe8ff', 'align':'left'})
    fmt_table_title = workbook.add_format({'bold': True, 'align':'center', 'bg_color':'#e6f2ff'})
    fmt_center = workbook.add_format({'align':'center'})
    fmt_bold = workbook.add_format({'bold':True})
    fmt_currency = workbook.add_format({'num_format':'#,##0','align':'right'})
    fmt_percent = workbook.add_format({'num_format':'0.00%','align':'right'})

    # Title
    title_text = f"LAPORAN IKLAN {store.upper()}"
    ws.merge_range('A1:E1', title_text, fmt_title)

    # PESANAN IKLAN table - starting at A3
    start_row = 2
    start_col = 0
    ws.merge_range(start_row, start_col, start_row, start_col+4, "PESANAN IKLAN", fmt_table_title)
    # day label row (merged)
    day_label = datetime.now().strftime("%A, %d-%m-%Y")
    ws.merge_range(start_row+1, start_col, start_row+1, start_col+4, day_label, fmt_center)
    # header row for columns
    headers = ["LIHAT","KLIK","PESANAN","KUANTITAS","OMZET PENJUALAN"]
    for i,h in enumerate(headers):
        ws.write(start_row+2, start_col+i, h, fmt_bold)
    # write 24 rows + total
    for i, r in df_pesanan_iklan.iterrows():
        for j, col in enumerate(["LIHAT","KLIK","PESANAN","KUANTITAS","OMZET PENJUALAN"]):
            val = r[col]
            rowpos = start_row+3+i
            if col == "OMZET PENJUALAN":
                ws.write_number(rowpos, start_col+j, float(val), fmt_currency)
            elif col in ["PESANAN","KUANTITAS"]:
                ws.write_number(rowpos, start_col+j, int(val))
            else:
                ws.write(rowpos, start_col+j, val)

    # RINCIAN IKLAN KLIK table - put in G3 (col 6)
    rincian_col = 6
    ws.merge_range(start_row, rincian_col, start_row, rincian_col+1, "RINCIAN IKLAN KLIK", fmt_table_title)
    # write labels/values from df_rincian_iklan
    rstart = start_row+2
    for i, row in df_rincian_iklan.iterrows():
        ws.write(rstart+i, rincian_col, row["Label"])
        v = row["Value"]
        # format percent
        if isinstance(v, str) and v.endswith("%"):
            try:
                pct = float(v.strip("%"))/100.0
                ws.write_number(rstart+i, rincian_col+1, pct, fmt_percent)
            except:
                ws.write(rstart+i, rincian_col+1, v)
        elif isinstance(v, (int, float)) and not pd.isna(v):
            if "Biaya" in row["Label"] or "Penjualan" in row["Label"] or "ROASI" in row["Label"]:
                ws.write_number(rstart+i, rincian_col+1, float(v), fmt_currency)
            else:
                ws.write(rstart+i, rincian_col+1, float(v))
        else:
            ws.write(rstart+i, rincian_col+1, v)

    # PESANAN AFFILIATE table - put at L3 (col 11)
    aff_col = 11
    ws.merge_range(start_row, aff_col, start_row, aff_col+4, "PESANAN AFFILIATE", fmt_table_title)
    # header row (Jam, Pesanan, Kuantitas, Omzet Penjualan, Komisi)
    for i,h in enumerate(["Jam","Pesanan","Kuantitas","Omzet Penjualan","Komisi"]):
        ws.write(start_row+2, aff_col+i, h, fmt_bold)
    for i, r in aff_table.iterrows():
        for j, key in enumerate(["Jam","Pesanan","Kuantitas","Omzet Penjualan","Komisi"]):
            val = r[key]
            rowpos = start_row+3+i
            if key in ["Omzet Penjualan","Komisi"] and val!="" and not pd.isna(val):
                ws.write_number(rowpos, aff_col+j, float(val), fmt_currency)
            elif key in ["Pesanan","Kuantitas"]:
                ws.write_number(rowpos, aff_col+j, int(val))
            else:
                ws.write(rowpos, aff_col+j, val)

    # PESANAN ORGANIK table - right after affiliate, but we wrote aff with total row; put org at rows below or different columns
    org_col = aff_col
    org_row_start = start_row + 3 + len(aff_table) + 1
    ws.merge_range(org_row_start-1, org_col, org_row_start-1, org_col+3, "PESANAN ORGANIK", fmt_table_title)
    for i,h in enumerate(["Jam","Pesanan","Kuantitas","Omzet Penjualan"]):
        ws.write(org_row_start, org_col+i, h, fmt_bold)
    for i, r in org_table.iterrows():
        for j, key in enumerate(["Jam","Pesanan","Kuantitas","Omzet Penjualan"]):
            val = r[key]
            rowpos = org_row_start+1+i
            if key == "Omzet Penjualan" and val!="" and not pd.isna(val):
                ws.write_number(rowpos, org_col+j, float(val), fmt_currency)
            elif key in ["Pesanan","Kuantitas"]:
                ws.write_number(rowpos, org_col+j, int(val))
            else:
                ws.write(rowpos, org_col+j, val)

    # RINCIAN SELURUH PESANAN (G - J) place below main PESANAN IKLAN table, around row 35
    rinc_row_start = start_row + 3 + len(df_pesanan_iklan) + 2
    rinc_col = 6
    ws.merge_range(rinc_row_start-1, rinc_col, rinc_row_start-1, rinc_col+3, "RINCIAN SELURUH PESANAN", fmt_table_title)
    # write headers
    for i,h in enumerate(["Nama Produk","Variasi","Jumlah Pesanan","Jumlah Eksemplar"]):
        ws.write(rinc_row_start, rinc_col+i, h, fmt_bold)
    for i, r in rincian_output.iterrows():
        for j, key in enumerate(["Nama Produk","Variasi","Jumlah Pesanan","Jumlah Eksemplar"]):
            val = r[key]
            rowpos = rinc_row_start+1+i
            if key in ["Jumlah Pesanan","Jumlah Eksemplar"] and val!="" and not pd.isna(val):
                ws.write_number(rowpos, rinc_col+j, int(val))
            else:
                ws.write(rowpos, rinc_col+j, val)

    # SUMMARY table at columns L-P below org table with 2-row space
    summary_start_row = org_row_start + len(org_table) + 3
    summary_col = 11
    for i, (label, val) in enumerate(summary_rows):
        # merge L-O for label
        ws.merge_range(summary_start_row + i, summary_col, summary_start_row + i, summary_col+3, label)
        if isinstance(val, (int,float)):
            ws.write_number(summary_start_row + i, summary_col+4, val, fmt_currency)
        else:
            ws.write(summary_start_row + i, summary_col+4, val)

    # 2) order-all sheet
    df_order_to_write = df_order.copy()
    # drop internal helper columns in write
    drop_cols = ["__hour","__hour_label","_VAR_SUFFIX","_HIGHLIGHT"]
    for c in drop_cols:
        if c in df_order_to_write.columns:
            df_order_to_write[c] = df_order_to_write[c]

    # Write order-all sheet
    df_order_to_write.to_excel(writer, sheet_name="order-all", index=False)
    ws2 = writer.sheets["order-all"]
    # Apply highlight formatting
    fmt_yellow = workbook.add_format({'bg_color':'#fff2cc'})
    fmt_pink = workbook.add_format({'bg_color':'#ffd6e7'})
    # find column index for No. Pesanan
    ord_sheet_cols = list(df_order_to_write.columns)
    for i, r in enumerate(df_order_to_write.itertuples(), start=1):
        hl = getattr(r, "_HIGHLIGHT") if "_HIGHLIGHT" in df_order_to_write.columns else ""
        if hl == "YELLOW":
            ws2.set_row(i, cell_format=fmt_yellow)
        elif hl == "PINK":
            ws2.set_row(i, cell_format=fmt_pink)

    # 3) Iklan klik sheet (write aggregated iklan_agg)
    iklan_agg.to_excel(writer, sheet_name="Iklan klik", index=False)
    # 4) Seller conversion sheet
    df_seller.to_excel(writer, sheet_name="Seller conversion", index=False)

    # close writer
    writer.close()

output.seek(0)

st.success("Proses selesai. File siap didownload.")
st.download_button("📥 Download File Output (IklanKu_Output.xlsx)", data=output, file_name="IklanKu_Output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

