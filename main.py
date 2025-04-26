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

# Fungsi untuk format tanggal, misalnya "1 Maret 2025"
def format_tanggal_indonesia(tanggal):
    bulan_indonesia = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
        7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }
    return f"{tanggal.day} {bulan_indonesia[tanggal.month]} {tanggal.year}"

# Fungsi untuk membuat file Excel laporan
def to_excel(df, nama_ketua, nama_bendahara, bulan, tahun):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    df_excel = df.copy()
    df_excel = df_excel.rename(columns={
        'Tanggal': 'tanggal',
        'Keterangan': 'KETERANGAN',
        'Pemasukan': 'Debet',
        'Pengeluaran': 'Kredit',
        'Saldo': 'Saldo'
    })
    df_excel['tanggal'] = pd.to_datetime(df_excel['tanggal']).apply(lambda x: format_tanggal_indonesia(x))
    df_excel.to_excel(writer, sheet_name='Sheet1', startrow=4, index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    worksheet.set_landscape()
    worksheet.set_paper(9)
    
    money_format = workbook.add_format({'num_format': '"Rp. " #,##0', 'border': 1, 'align': 'right'})
    header_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D3D3D3'})
    title_format = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center'})
    date_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    text_format = workbook.add_format({'border': 1, 'align': 'left', 'text_wrap': True})
    
    nama_bulan_str = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
                      7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}[bulan]
    judul_line1 = "LAPORAN KEUANGAN KAS MASJID JAM'I AL FAIZIN"
    judul_line2 = "LINGKUNGAN RT 010-RW 005 KEL BENDUNGAN KEC. CILEGON"
    judul_line3 = f"Periode Bulan {nama_bulan_str} {tahun}"
    
    worksheet.merge_range('A1:E1', judul_line1, title_format)
    worksheet.merge_range('A2:E2', judul_line2, title_format)
    worksheet.merge_range('A3:E3', judul_line3, title_format)
    
    headers = ['tanggal', 'KETERANGAN', 'Debet', 'Kredit', 'Saldo']
    for col_num, value in enumerate(headers):
        worksheet.write(4, col_num, value, header_format)
    
    data_start = 5
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
    
    for row_num, row in enumerate(df_excel.values):
        excel_row = data_start + row_num
        worksheet.write(excel_row, 1, row[1], text_format)
        worksheet.write(excel_row, 2, row[2], money_format)
        worksheet.write(excel_row, 3, row[3], money_format)
        worksheet.write(excel_row, 4, row[4], money_format)
    
    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 40)
    worksheet.set_column('C:E', 20)
    
    total_debet = df['Pemasukan'].sum()
    total_kredit = df['Pengeluaran'].sum()
    last_row = len(df) + data_start
    jumlah_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#D3D3D3'})
    jumlah_angka_format = workbook.add_format({'bold': True, 'border': 1, 'num_format': '"Rp. " #,##0', 'align': 'right', 'bg_color': '#D3D3D3'})
    
    worksheet.merge_range(last_row, 0, last_row, 1, "JUMLAH", jumlah_format)
    worksheet.write(last_row, 2, total_debet, jumlah_angka_format)
    worksheet.write(last_row, 3, total_kredit, jumlah_angka_format)
    worksheet.write(last_row, 4, df['Saldo'].iloc[-1] if not df.empty else 0, jumlah_angka_format)
    
    sign_row = last_row + 3
    signature_header = workbook.add_format({'align': 'center', 'bold': True})
    signature_name = workbook.add_format({'align': 'center', 'bottom': 1})
    
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
    
    nama_bulan = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
                  7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}[bulan]
    bulan_tahun = f"{nama_bulan} {tahun}"
    st.write(f"**Periode Laporan: {bulan_tahun}**")
    
    st.markdown("---")
    nama_ketua = st.text_input("Nama Ketua DKM", "H Didi Rosyadi, ST")
    nama_bendahara = st.text_input("Nama Bendahara", "Sudiro")
    
    st.markdown("---")
    saldo_awal = st.number_input("Saldo Awal Bulan (Rp)", min_value=0, value=0, step=1000)
    
    st.markdown("---")
    if st.button("Reset Semua Data", use_container_width=True):
        st.session_state.transaksi = pd.DataFrame(columns=['Tanggal', 'Keterangan', 'Pemasukan', 'Pengeluaran'])
        st.success("Data berhasil direset!")
        st.rerun()

# Inisialisasi data transaksi
if 'transaksi' not in st.session_state:
    st.session_state.transaksi = pd.DataFrame(columns=['Tanggal', 'Keterangan', 'Pemasukan', 'Pengeluaran'])

if 'bulan_tahun_aktif' not in st.session_state:
    st.session_state.bulan_tahun_aktif = bulan_tahun
if st.session_state.bulan_tahun_aktif != bulan_tahun:
    st.session_state.bulan_tahun_aktif = bulan_tahun
    st.session_state.transaksi = pd.DataFrame(columns=['Tanggal', 'Keterangan', 'Pemasukan', 'Pengeluaran'])

# Inisialisasi session state untuk input form
if 'tanggal_input' not in st.session_state:
    st.session_state.tanggal_input = datetime.date(tahun, bulan, 1)
if 'keterangan_input' not in st.session_state:
    st.session_state.keterangan_input = ""
if 'jenis_input' not in st.session_state:
    st.session_state.jenis_input = "Pemasukan"
if 'nominal_input' not in st.session_state:
    st.session_state.nominal_input = 0

# Tabs utama
tab1, tab3 = st.tabs(["📝 Input Transaksi", "📑 Preview & Unduh Laporan"])

# Tab 1: Input Transaksi
with tab1:
    st.header("Tambah Transaksi")
    col1, col2 = st.columns(2)
    with col1:
        min_date = datetime.date(tahun, bulan, 1)
        max_date = datetime.date(tahun, bulan, get_last_day_of_month(tahun, bulan))

        # Validasi agar default tanggal tetap dalam rentang min-max
        default_date = st.session_state.get("tanggal_input", datetime.date.today())
        if default_date < min_date:
            default_date = min_date
        elif default_date > max_date:
            default_date = max_date

        tanggal = st.date_input(
            "Tanggal",
            value=default_date,
            min_value=min_date,
            max_value=max_date
        )

    with col2:
        jenis_transaksi = st.radio(
            "Jenis Transaksi",
            ["Pemasukan", "Pengeluaran"],
            index=0 if st.session_state.jenis_input == "Pemasukan" else 1,
            key="jenis_input_widget"
        )
        nominal = st.number_input(
            "Nominal (Rp)",
            min_value=0,
            step=1000,
            value=st.session_state.nominal_input,
            key="nominal_input_widget"
        )
    # Input Keterangan
    keterangan = st.text_area("Keterangan")

    # Input Nominal
    nominal = st.number_input("Nominal", min_value=0)

    if st.button("Tambah Transaksi"):
        if not keterangan:
            st.error("Keterangan tidak boleh kosong!")
        elif nominal == 0:
            st.error("Nominal tidak boleh 0!")
        else:
            pemasukan = nominal if jenis_transaksi == "Pemasukan" else 0
            pengeluaran = nominal if jenis_transaksi == "Pengeluaran" else 0
            transaksi_baru = pd.DataFrame({
                'Tanggal': [tanggal],
                'Keterangan': [keterangan],
                'Pemasukan': [pemasukan],
                'Pengeluaran': [pengeluaran]
            })
            st.session_state.transaksi = pd.concat([st.session_state.transaksi, transaksi_baru], ignore_index=True)
            st.success("Transaksi berhasil ditambahkan!")
            
            # Kosongkan input form
            st.session_state.tanggal_input = datetime.date(tahun, bulan, 1)
            st.session_state.keterangan_input = ""
            st.session_state.jenis_input = "Pemasukan"
            st.session_state.nominal_input = 0
            st.rerun()  # Refresh halaman untuk memperbarui tampilan

    st.header("Kelola Transaksi")
    if not st.session_state.transaksi.empty:
        edited_df = st.data_editor(
            st.session_state.transaksi,
            column_config={
                "Tanggal": st.column_config.DateColumn(
                    "Tanggal",
                    min_value=datetime.date(tahun, bulan, 1),
                    max_value=datetime.date(tahun, bulan, get_last_day_of_month(tahun, bulan)),
                    format="DD/MM/YYYY",
                    required=True
                ),
                "Keterangan": st.column_config.TextColumn("Keterangan", required=True),
                "Pemasukan": st.column_config.NumberColumn("Pemasukan", min_value=0, step=1000, default=0),
                "Pengeluaran": st.column_config.NumberColumn("Pengeluaran", min_value=0, step=1000, default=0),
            },
            hide_index=True,
            use_container_width=True,
            num_rows="dynamic",
            key="transaction_editor"
        )
        st.session_state.transaksi = edited_df
    else:
        st.info("Belum ada transaksi. Silakan tambahkan transaksi baru.")

    # Buat dataframe lengkap dengan saldo awal untuk ditampilkan
    initial_row = pd.DataFrame({
        'Tanggal': [datetime.date(tahun, bulan, 1)],
        'Keterangan': [f'Saldo Awal {bulan_tahun}'],
        'Pemasukan': [saldo_awal],
        'Pengeluaran': [0],
        'Saldo': [saldo_awal]
    })
    full_df = pd.concat([initial_row, st.session_state.transaksi], ignore_index=True)
    full_df = full_df.sort_values('Tanggal').reset_index(drop=True)
    full_df['Saldo'] = 0.0
    full_df.loc[0, 'Saldo'] = saldo_awal
    for i in range(1, len(full_df)):
        full_df.loc[i, 'Saldo'] = full_df.loc[i-1, 'Saldo'] + full_df.loc[i, 'Pemasukan'] - full_df.loc[i, 'Pengeluaran']
    
    st.header(f"Daftar Transaksi - {bulan_tahun}")
    df_display = full_df.copy()
    df_display['Tanggal'] = pd.to_datetime(df_display['Tanggal']).apply(lambda x: format_tanggal_indonesia(x))
    df_display['Pemasukan'] = df_display['Pemasukan'].apply(lambda x: f"{x:,.0f}".replace(',', '.'))
    df_display['Pengeluaran'] = df_display['Pengeluaran'].apply(lambda x: f"{x:,.0f}".replace(',', '.'))
    df_display['Saldo'] = df_display['Saldo'].apply(lambda x: f"{x:,.0f}".replace(',', '.'))
    st.dataframe(df_display, hide_index=True, use_container_width=True)
    
    # Ringkasan
    total_pemasukan = full_df['Pemasukan'].sum()
    total_pengeluaran = full_df['Pengeluaran'].sum()
    saldo_akhir = full_df['Saldo'].iloc[-1]
    st.markdown("### Ringkasan")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Debet", f"Rp {total_pemasukan:,.0f}".replace(',', '.'))
    col2.metric("Total Kredit", f"Rp {total_pengeluaran:,.0f}".replace(',', '.'))
    col3.metric("Saldo Akhir", f"Rp {saldo_akhir:,.0f}".replace(',', '.'))

# Tab 3: Preview & Unduh Laporan
with tab3:
    st.markdown("<h2 style='text-align: center;'>LAPORAN KEUANGAN KAS MASJID JAM'I AL FAIZIN</h2>", unsafe_allow_html=True)
    st.markdown("<h2 style='text-align: center;'>LINGKUNGAN RT 010-RW 005 KEL BENDUNGAN KEC. CILEGON</h2>", unsafe_allow_html=True)
    st.markdown(f"<h2 style='text-align: center;'>Periode Bulan {nama_bulan} {tahun}</h2>", unsafe_allow_html=True)
    
    # Buat dataframe lengkap untuk preview dan unduhan
    initial_row = pd.DataFrame({
        'Tanggal': [datetime.date(tahun, bulan, 1)],
        'Keterangan': [f'Saldo Awal {bulan_tahun}'],
        'Pemasukan': [saldo_awal],
        'Pengeluaran': [0],
        'Saldo': [saldo_awal]
    })
    full_df = pd.concat([initial_row, st.session_state.transaksi], ignore_index=True)
    full_df = full_df.sort_values('Tanggal').reset_index(drop=True)
    full_df['Saldo'] = 0.0
    full_df.loc[0, 'Saldo'] = saldo_awal
    for i in range(1, len(full_df)):
        full_df.loc[i, 'Saldo'] = full_df.loc[i-1, 'Saldo'] + full_df.loc[i, 'Pemasukan'] - full_df.loc[i, 'Pengeluaran']
    
    if not full_df.empty:
        df_preview = full_df.copy()
        df_preview['Tanggal'] = pd.to_datetime(df_preview['Tanggal']).apply(lambda x: format_tanggal_indonesia(x))
        df_preview['Pemasukan'] = df_preview['Pemasukan'].apply(lambda x: f"{x:,.0f}".replace(',', '.'))
        df_preview['Pengeluaran'] = df_preview['Pengeluaran'].apply(lambda x: f"{x:,.0f}".replace(',', '.'))
        df_preview['Saldo'] = df_preview['Saldo'].apply(lambda x: f"{x:,.0f}".replace(',', '.'))
        st.table(df_preview)  # st.table tidak mendukung hide_index, tetapi sudah tanpa indeks default
        
        total_pemasukan = full_df['Pemasukan'].sum()
        total_pengeluaran = full_df['Pengeluaran'].sum()
        saldo_akhir = full_df['Saldo'].iloc[-1]
        
        st.markdown("**JUMLAH:**")
        jumlah_df = pd.DataFrame({
            'Debet': [f"{total_pemasukan:,.0f}".replace(',', '.')],
            'Kredit': [f"{total_pengeluaran:,.0f}".replace(',', '.')],
            'Saldo': [f"{saldo_akhir:,.0f}".replace(',', '.')]
        })
        st.table(jumlah_df)  # st.table tidak mendukung hide_index
        
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
        excel_data = to_excel(full_df, nama_ketua, nama_bendahara, bulan, tahun)
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