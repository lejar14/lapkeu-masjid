import streamlit as st
import pandas as pd
import datetime
import base64
from io import BytesIO
import calendar
import locale

# Set locale ke Indonesia untuk format tanggal
try:
    locale.setlocale(locale.LC_ALL, 'id_ID.UTF-8')
except:
    pass

# Konfigurasi halaman
st.set_page_config(
    page_title="Laporan Keuangan Kas Masjid",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Fungsi untuk mendapatkan hari terakhir bulan
def get_last_day_of_month(year, month):
    return calendar.monthrange(year, month)[1]

# Fungsi untuk format tanggal misalnya "1 Maret 2025"
def format_tanggal_indonesia(tanggal):
    bulan_indonesia = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
        7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }
    return f"{tanggal.day} {bulan_indonesia[tanggal.month]} {tanggal.year}"

# Fungsi untuk membuat file Excel laporan dengan merge untuk tanggal yang sama, orientasi landscape,
# dan mengatur alignment sel yang di-merge ke tengah (middle) serta prefix "Rp. " untuk nominal
def to_excel(df, nama_ketua, nama_bendahara, bulan, tahun):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    # Siapkan data untuk Excel
    df_excel = df.copy()
    # Ubah nama kolom (Tanggal -> tanggal)
    df_excel = df_excel.rename(columns={
        'Tanggal': 'tanggal',
        'Keterangan': 'KETERANGAN',
        'Pemasukan': 'Debet',
        'Pengeluaran': 'Kredit',
        'Saldo': 'Saldo'
    })
    # Format kolom tanggal dengan format "1 Maret 2025"
    df_excel['tanggal'] = pd.to_datetime(df_excel['tanggal']).apply(lambda x: format_tanggal_indonesia(x))
    
    # Tulis data mulai baris ke-5 (indeks 4) setelah judul dan header
    df_excel.to_excel(writer, sheet_name='Sheet1', startrow=4, index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Atur orientasi halaman ke landscape dan ukuran kertas A4
    worksheet.set_landscape()
    worksheet.set_paper(9)
    
    # Definisi format sel
    money_format = workbook.add_format({
        'num_format': '"Rp. " #,##0',  # format dengan prefix "Rp. "
        'border': 1,
        'align': 'right'
    })
    header_format = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#D3D3D3'
    })
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 16,
        'align': 'center'
    })
    date_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })
    text_format = workbook.add_format({
        'border': 1,
        'align': 'left',
        'text_wrap': True
    })
    
    # Judul laporan tiga baris (judul tetap)
    nama_bulan_str = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
        7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }[bulan]
    judul_line1 = "LAPORAN KEUANGAN KAS MASJID JAM'I AL FAIZIN"
    judul_line2 = "LINGKUNGAN RT 010-RW 005 KEL BENDUNGAN KEC. CILEGON"
    judul_line3 = f"Periode Bulan {nama_bulan_str} {tahun}"
    
    worksheet.merge_range('A1:E1', judul_line1, title_format)
    worksheet.merge_range('A2:E2', judul_line2, title_format)
    worksheet.merge_range('A3:E3', judul_line3, title_format)
    
    # Tulis header tabel pada baris ke-5 (indeks 4)
    headers = ['tanggal', 'KETERANGAN', 'Debet', 'Kredit', 'Saldo']
    for col_num, value in enumerate(headers):
        worksheet.write(4, col_num, value, header_format)
    
    # Tulis data mulai dari baris ke-6 (indeks 5)
    data_start = 5
    # Lakukan merge untuk kolom tanggal jika nilai sama secara berurutan
    date_values = df_excel['tanggal'].tolist()
    i = 0
    while i < len(date_values):
        start_excel_row = data_start + i
        current_date = date_values[i]
        j = i
        while j + 1 < len(date_values) and date_values[j+1] == current_date:
            j += 1
        end_excel_row = data_start + j
        if start_excel_row == end_excel_row:
            worksheet.write(start_excel_row, 0, current_date, date_format)
        else:
            worksheet.merge_range(start_excel_row, 0, end_excel_row, 0, current_date, date_format)
        i = j + 1
    
    # Tulis data untuk kolom selain tanggal (kolom 1 s/d 4)
    for row_num, row in enumerate(df_excel.values):
        excel_row = data_start + row_num
        worksheet.write(excel_row, 1, row[1], text_format)
        worksheet.write(excel_row, 2, row[2], money_format)
        worksheet.write(excel_row, 3, row[3], money_format)
        worksheet.write(excel_row, 4, row[4], money_format)
    
    # Sesuaikan lebar kolom
    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 40)
    worksheet.set_column('C:E', 20)
    
    # Hitung jumlah total
    total_debet = df['Pemasukan'].sum()
    total_kredit = df['Pengeluaran'].sum()
    last_row = len(df) + data_start
    jumlah_format = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'center',
        'bg_color': '#D3D3D3'
    })
    jumlah_angka_format = workbook.add_format({
        'bold': True,
        'border': 1,
        'num_format': '"Rp. " #,##0',
        'align': 'right',
        'bg_color': '#D3D3D3'
    })
    
    worksheet.merge_range(last_row, 0, last_row, 1, "JUMLAH", jumlah_format)
    worksheet.write(last_row, 2, total_debet, jumlah_angka_format)
    worksheet.write(last_row, 3, total_kredit, jumlah_angka_format)
    worksheet.write(last_row, 4, df['Saldo'].iloc[-1] if not df.empty else 0, jumlah_angka_format)
    
    # Tambahkan area tanda tangan
    sign_row = last_row + 3
    signature_header = workbook.add_format({
        'align': 'center',
        'bold': True
    })
    signature_name = workbook.add_format({
        'align': 'center',
        'bottom': 1
    })
    
    last_day = get_last_day_of_month(tahun, bulan)
    tgl_ttd = datetime.date(tahun, bulan, last_day)
    tgl_ttd_str = format_tanggal_indonesia(tgl_ttd)
    
    worksheet.merge_range(f'A{sign_row}:B{sign_row}', "Mengetahui,", signature_header)
    worksheet.merge_range(f'D{sign_row}:E{sign_row}', tgl_ttd_str, signature_header)
    
    sign_row += 1
    worksheet.merge_range(f'A{sign_row}:B{sign_row}', "Ketua DKM", signature_header)
    worksheet.merge_range(f'D{sign_row}:E{sign_row}', "Bendahara", signature_header)
    
    sign_row += 3
    worksheet.merge_range(f'A{sign_row}:B{sign_row}', nama_ketua, signature_name)
    worksheet.merge_range(f'D{sign_row}:E{sign_row}', nama_bendahara, signature_name)
    
    writer.close()
    processed_data = output.getvalue()
    return processed_data

# Sidebar Pengaturan
with st.sidebar:
    st.header("Pengaturan Laporan")
    col1, col2 = st.columns(2)
    with col1:
        bulan = st.selectbox("Bulan", range(1, 13), index=datetime.date.today().month - 1)
    with col2:
        tahun = st.selectbox("Tahun", range(datetime.date.today().year - 5, datetime.date.today().year + 1), index=5)
    
    nama_bulan = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
        7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }[bulan]
    bulan_tahun = f"{nama_bulan} {tahun}"
    st.write(f"**Periode Laporan: {bulan_tahun}**")
    
    st.markdown("---")
    # Judul laporan sudah tetap, sehingga input judul dihilangkan
    nama_ketua = st.text_input("Nama Ketua DKM", "Nama Ketua DKM")
    nama_bendahara = st.text_input("Nama Bendahara", "Nama Bendahara")
    
    st.markdown("---")
    saldo_awal = st.number_input("Saldo Awal Bulan (Rp)", min_value=0, value=0, step=1000)
    
    st.markdown("---")
    if st.button("Reset Semua Data", use_container_width=True):
        st.session_state.transaksi = pd.DataFrame({
            'Tanggal': [datetime.date(tahun, bulan, 1)],
            'Keterangan': [f'Saldo Awal {bulan_tahun}'],
            'Pemasukan': [saldo_awal],
            'Pengeluaran': [0],
            'Saldo': [saldo_awal]
        })
        st.success("Data berhasil direset!")
        st.rerun()

# Inisialisasi data transaksi
if 'transaksi' not in st.session_state:
    st.session_state.transaksi = pd.DataFrame({
        'Tanggal': [datetime.date(tahun, bulan, 1)],
        'Keterangan': [f'Saldo Awal {bulan_tahun}'],
        'Pemasukan': [saldo_awal],
        'Pengeluaran': [0],
        'Saldo': [saldo_awal]
    })

if 'bulan_tahun_aktif' not in st.session_state:
    st.session_state.bulan_tahun_aktif = bulan_tahun
if st.session_state.bulan_tahun_aktif != bulan_tahun:
    st.session_state.bulan_tahun_aktif = bulan_tahun
    st.session_state.transaksi = pd.DataFrame({
        'Tanggal': [datetime.date(tahun, bulan, 1)],
        'Keterangan': [f'Saldo Awal {bulan_tahun}'],
        'Pemasukan': [saldo_awal],
        'Pengeluaran': [0],
        'Saldo': [saldo_awal]
    })

# Jika saldo awal berubah, perbarui baris pertama dan re-kalkulasi saldo
if not st.session_state.transaksi.empty and st.session_state.transaksi.iloc[0]['Keterangan'].startswith("Saldo Awal"):
    current_saldo_awal = st.session_state.transaksi.iloc[0]['Pemasukan']
    if current_saldo_awal != saldo_awal:
        st.session_state.transaksi.at[0, 'Tanggal'] = datetime.date(tahun, bulan, 1)
        st.session_state.transaksi.at[0, 'Keterangan'] = f'Saldo Awal {bulan_tahun}'
        st.session_state.transaksi.at[0, 'Pemasukan'] = saldo_awal
        st.session_state.transaksi.at[0, 'Saldo'] = saldo_awal
        for i in range(1, len(st.session_state.transaksi)):
            prev_balance = st.session_state.transaksi.at[i-1, 'Saldo']
            pemasukan = st.session_state.transaksi.at[i, 'Pemasukan']
            pengeluaran = st.session_state.transaksi.at[i, 'Pengeluaran']
            st.session_state.transaksi.at[i, 'Saldo'] = prev_balance + pemasukan - pengeluaran

# Pastikan flag notifikasi ada di session_state
if "transaction_added" not in st.session_state:
    st.session_state.transaction_added = False

# Callback untuk menghapus notifikasi saat ada perubahan input
def clear_success():
    st.session_state.transaction_added = False

# Tabs utama
tab1, tab3 = st.tabs(["📝 Input Transaksi", "📑 Preview & Unduh Laporan"])

# Tab 1: Input Transaksi
with tab1:
    st.header("Tambah Transaksi")
    # Menggunakan kolom untuk mengatur tampilan input
    col1, col2 = st.columns(2)

    with col1:
        tanggal = st.date_input(
            "Tanggal",
            datetime.date(tahun, bulan, 1),
            min_value=datetime.date(tahun, bulan, 1),
            max_value=datetime.date(tahun, bulan, get_last_day_of_month(tahun, bulan)),
            key="tanggal_input",
            on_change=clear_success
        )
        keterangan = st.text_area(
            "Keterangan",
            height=100,
            key="keterangan_input",
            on_change=clear_success
        )
    with col2:
        jenis_transaksi = st.radio(
            "Jenis Transaksi",
            ["Pemasukan", "Pengeluaran"],
            key="jenis_input",
            on_change=clear_success
        )
        nominal = st.number_input(
            "Nominal (Rp)",
            min_value=0,
            step=1000,
            key="nominal_input",
            on_change=clear_success
        )

    # Tombol tambah transaksi
    if st.button("Tambah Transaksi"):
        if not keterangan:
            st.error("Keterangan tidak boleh kosong!")
        elif nominal == 0:
            st.error("Nominal tidak boleh 0!")
        else:
            pemasukan = nominal if jenis_transaksi == "Pemasukan" else 0
            pengeluaran = nominal if jenis_transaksi == "Pengeluaran" else 0
            saldo_terakhir = st.session_state.transaksi['Saldo'].iloc[-1] if not st.session_state.transaksi.empty else 0
            saldo_baru = saldo_terakhir + pemasukan - pengeluaran
            transaksi_baru = pd.DataFrame({
                'Tanggal': [tanggal],
                'Keterangan': [keterangan],
                'Pemasukan': [pemasukan],
                'Pengeluaran': [pengeluaran],
                'Saldo': [saldo_baru]
            })
            st.session_state.transaksi = pd.concat([st.session_state.transaksi, transaksi_baru], ignore_index=True)
            st.session_state.transaction_added = True

    # Tampilkan notifikasi jika transaksi baru telah ditambahkan
    if st.session_state.transaction_added:
        st.success("Transaksi berhasil ditambahkan!")
        
# Tab 2: Daftar Transaksi
    st.header(f"Daftar Transaksi - {bulan_tahun}")
    if not st.session_state.transaksi.empty:
        df_display = st.session_state.transaksi.copy()
        df_display = df_display.rename(columns={"Tanggal": "tanggal"})
        df_display['tanggal'] = pd.to_datetime(df_display['tanggal']).apply(lambda x: format_tanggal_indonesia(x))
        df_display['Pemasukan'] = df_display['Pemasukan'].apply(lambda x: f"{x:,.0f}".replace(',', '.'))
        df_display['Pengeluaran'] = df_display['Pengeluaran'].apply(lambda x: f"{x:,.0f}".replace(',', '.'))
        df_display['Saldo'] = df_display['Saldo'].apply(lambda x: f"{x:,.0f}".replace(',', '.'))
        
        st.dataframe(
            df_display,
            column_config={
                "tanggal": st.column_config.TextColumn("tanggal"),
                "Keterangan": st.column_config.TextColumn("KETERANGAN", width="large"),
                "Pemasukan": st.column_config.TextColumn("Debet"),
                "Pengeluaran": st.column_config.TextColumn("Kredit"),
                "Saldo": st.column_config.TextColumn("Saldo"),
            },
            hide_index=True,
            use_container_width=True
        )
        
        total_pemasukan = st.session_state.transaksi['Pemasukan'].sum()
        total_pengeluaran = st.session_state.transaksi['Pengeluaran'].sum()
        saldo_akhir = st.session_state.transaksi['Saldo'].iloc[-1]
        
        st.markdown("### Ringkasan")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Debet", f"Rp {total_pemasukan:,.0f}".replace(',', '.'))
        col2.metric("Total Kredit", f"Rp {total_pengeluaran:,.0f}".replace(',', '.'))
        col3.metric("Saldo Akhir", f"Rp {saldo_akhir:,.0f}".replace(',', '.'))
        
        st.markdown("### Kelola Transaksi")
        # Jika terdapat transaksi selain saldo awal, berikan pilihan untuk menghapus transaksi pada baris tertentu.
        if len(st.session_state.transaksi) > 1:
            # Buat list opsi untuk transaksi yang dapat dihapus (indeks 1 s/d n)
            df_temp = st.session_state.transaksi.iloc[1:].reset_index(drop=True)
            options = [f"Transaksi ke-{i+2}: {df_temp.loc[i, 'Keterangan']} ({format_tanggal_indonesia(df_temp.loc[i, 'Tanggal'])})" 
                       for i in range(len(df_temp))]
            selected_option = st.selectbox("Pilih transaksi yang akan dihapus", options)
            if st.button("Hapus Transaksi", use_container_width=True):
                # Tentukan indeks asli: opsi ke-i berarti indeks = i+1
                index_to_delete = options.index(selected_option) + 1
                st.session_state.transaksi = st.session_state.transaksi.drop(st.session_state.transaksi.index[index_to_delete]).reset_index(drop=True)
                # Re-kalkulasi saldo
                for i in range(1, len(st.session_state.transaksi)):
                    st.session_state.transaksi.at[i, 'Saldo'] = st.session_state.transaksi.at[i-1, 'Saldo'] + st.session_state.transaksi.at[i, 'Pemasukan'] - st.session_state.transaksi.at[i, 'Pengeluaran']
                st.success("Transaksi berhasil dihapus!")
                st.rerun()
    else:
        st.info("Belum ada transaksi. Silakan tambahkan transaksi baru di tab Input Transaksi.")

# Tab 3: Preview & Unduh Laporan
with tab3:
    st.markdown("<h2 style='text-align: center;'>LAPORAN KEUANGAN KAS MASJID JAM'I AL FAIZIN</h2>", unsafe_allow_html=True)
    st.markdown("<h2 style='text-align: center;'>LINGKUNGAN RT 010-RW 005 KEL BENDUNGAN KEC. CILEGON</h2>", unsafe_allow_html=True)
    st.markdown(f"<h2 style='text-align: center;'>Periode Bulan {nama_bulan} {tahun}</h2>", unsafe_allow_html=True)
    
    if not st.session_state.transaksi.empty:
        df_preview = st.session_state.transaksi.copy()
        df_preview = df_preview.rename(columns={"Tanggal": "tanggal"})
        df_preview['tanggal'] = pd.to_datetime(df_preview['tanggal']).apply(lambda x: format_tanggal_indonesia(x))
        df_preview['Pemasukan'] = df_preview['Pemasukan'].apply(lambda x: f"{x:,.0f}".replace(',', '.'))
        df_preview['Pengeluaran'] = df_preview['Pengeluaran'].apply(lambda x: f"{x:,.0f}".replace(',', '.'))
        df_preview['Saldo'] = df_preview['Saldo'].apply(lambda x: f"{x:,.0f}".replace(',', '.'))
        
        st.table(df_preview)
        
        total_pemasukan = st.session_state.transaksi['Pemasukan'].sum()
        total_pengeluaran = st.session_state.transaksi['Pengeluaran'].sum()
        saldo_akhir = st.session_state.transaksi['Saldo'].iloc[-1]
        
        st.markdown("**JUMLAH:**")
        jumlah_df = pd.DataFrame({
            'Debet': [f"{total_pemasukan:,.0f}".replace(',', '.')],
            'Kredit': [f"{total_pengeluaran:,.0f}".replace(',', '.')],
            'Saldo': [f"{saldo_akhir:,.0f}".replace(',', '.')]
        })
        st.table(jumlah_df)
        
        st.markdown("")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Mengetahui,**")
            st.markdown("**Ketua DKM**")
            st.markdown("")
            st.markdown("")
            st.markdown(f"**{nama_ketua}**")
        with col2:
            last_day = get_last_day_of_month(tahun, bulan)
            tgl_ttd = datetime.date(tahun, bulan, last_day)
            tgl_ttd_str = format_tanggal_indonesia(tgl_ttd)
            st.markdown(f"**{tgl_ttd_str}**")
            st.markdown("**Bendahara**")
            st.markdown("")
            st.markdown("")
            st.markdown(f"**{nama_bendahara}**")
        
        st.markdown("---")
        st.subheader("Unduh Laporan")
        excel_data = to_excel(st.session_state.transaksi, nama_ketua, nama_bendahara, bulan, tahun)
        excel_b64 = base64.b64encode(excel_data).decode()
        download_filename = f"Laporan_Keuangan_Masjid_{bulan_tahun.replace(' ', '_')}.xlsx"
        
        download_button = f"""
        <div style="text-align: center; margin-top: 20px;">
            <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}" 
               download="{download_filename}" 
               style="background-color: #4CAF50; color: white; padding: 12px 20px; text-align: center; 
                      text-decoration: none; display: inline-block; font-size: 16px; margin: 4px 2px; 
                      cursor: pointer; border-radius: 4px;">
                📥 Download Laporan Excel
            </a>
        </div>
        """
        st.markdown(download_button, unsafe_allow_html=True)
    else:
        st.info("Belum ada transaksi. Silakan tambahkan transaksi baru di tab Input Transaksi.")

# Footer
st.markdown("---")
st.markdown("<div style='text-align: center;'>© Sistem Laporan Keuangan Kas Masjid</div>", unsafe_allow_html=True)
