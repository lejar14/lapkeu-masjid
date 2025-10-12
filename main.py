import streamlit as st
import pandas as pd
import sqlite3
import datetime
import calendar
import os
from io import BytesIO
import xlsxwriter

# ---------- Setup ----------
DB_PATH = "data/masjid.db"
os.makedirs("data", exist_ok=True)
st.set_page_config(page_title="Laporan Keuangan Kas Masjid", layout="wide")

def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS transaksi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            periode TEXT,
            tanggal TEXT,
            keterangan TEXT,
            pemasukan REAL,
            pengeluaran REAL
        )
    """)
    return conn

conn = get_conn()

# ---------- Helper ----------
def get_last_day_of_month(year, month):
    return calendar.monthrange(year, month)[1]

def format_tanggal_indonesia(tgl):
    bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
             "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    return f"{tgl.day} {bulan[tgl.month-1]} {tgl.year}"

# ---------- DB ----------
def get_available_periods():
    df = pd.read_sql_query("SELECT DISTINCT periode FROM transaksi ORDER BY periode DESC", conn)
    return df["periode"].tolist()

def load_data(periode):
    df = pd.read_sql_query("SELECT * FROM transaksi WHERE periode=? ORDER BY id ASC", conn, params=(periode,))
    if df.empty:
        return pd.DataFrame(columns=["id", "periode", "tanggal", "keterangan", "pemasukan", "pengeluaran"])
    df["tanggal"] = pd.to_datetime(df["tanggal"]).dt.date
    return df

def insert_data(periode, tanggal, ket, masuk, keluar):
    conn.execute(
        "INSERT INTO transaksi (periode,tanggal,keterangan,pemasukan,pengeluaran) VALUES (?,?,?,?,?)",
        (periode, tanggal.isoformat(), ket, masuk, keluar)
    )
    conn.commit()

# ---------- Excel Export ----------
def to_excel(df, nama_ketua, nama_bendahara, bulan, tahun):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    df_export = df.copy()

    # pastikan urutan dan kolom sesuai
    df_export = df_export[["Tanggal", "Keterangan", "Pemasukan", "Pengeluaran", "Saldo"]]
    df_export["Tanggal"] = pd.to_datetime(df_export["Tanggal"]).dt.date.apply(format_tanggal_indonesia)
    df_export = df_export.rename(columns={
        "Tanggal": "tanggal",
        "Keterangan": "KETERANGAN",
        "Pemasukan": "Debet",
        "Pengeluaran": "Kredit",
        "Saldo": "Saldo"
    })

    wb = writer.book
    ws = wb.add_worksheet("Laporan")

    # page setup
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(5)
    ws.set_paper(9)

    # format
    money_fmt = wb.add_format({'num_format': '"Rp. " #,##0', 'border': 1, 'align': 'right'})
    text_fmt = wb.add_format({'border': 1, 'align': 'left', 'text_wrap': True})
    date_fmt = wb.add_format({'border': 1, 'align': 'center'})
    hdr_fmt = wb.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D3D3D3'})
    ttl_fmt = wb.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
    sum_fmt = wb.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#D3D3D3'})
    sum_money = wb.add_format({'bold': True, 'border': 1, 'align': 'right', 'num_format': '"Rp. " #,##0', 'bg_color': '#D3D3D3'})

    # judul
    bulan_str = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                 "Juli", "Agustus", "September", "Oktober", "November", "Desember"][bulan-1]
    ws.merge_range("A1:E1", "LAPORAN KEUANGAN KAS MASJID JAM'I AL FAIZIN", ttl_fmt)
    ws.merge_range("A2:E2", "LINGKUNGAN RT 010-RW 005 KEL. BENDUNGAN KEC. CILEGON", ttl_fmt)
    ws.merge_range("A3:E3", f"Periode Bulan {bulan_str} {tahun}", ttl_fmt)

    headers = ["tanggal", "KETERANGAN", "Debet", "Kredit", "Saldo"]
    for i, h in enumerate(headers):
        ws.write(5, i, h, hdr_fmt)

    # tulis data & merge tanggal yang sama
    start = 6
    i = 0
    data = df_export.values.tolist()
    while i < len(data):
        tgl = data[i][0]
        # cari sampai tanggal berubah
        j = i
        while j + 1 < len(data) and data[j + 1][0] == tgl:
            j += 1
        if i == j:
            ws.write(start, 0, tgl, date_fmt)
        else:
            ws.merge_range(start, 0, start + (j - i), 0, tgl, date_fmt)

        for k in range(i, j + 1):
            ws.write(start, 1, data[k][1], text_fmt)
            ws.write(start, 2, data[k][2], money_fmt)
            ws.write(start, 3, data[k][3], money_fmt)
            ws.write(start, 4, data[k][4], money_fmt)
            start += 1
        i = j + 1

    # baris total
    total_in = df["Pemasukan"].sum()
    total_out = df["Pengeluaran"].sum()
    saldo_akhir = df["Saldo"].iloc[-1] if not df.empty else 0
    ws.merge_range(start, 0, start, 1, "JUMLAH", sum_fmt)
    ws.write(start, 2, total_in, sum_money)
    ws.write(start, 3, total_out, sum_money)
    ws.write(start, 4, saldo_akhir, sum_money)

    # tanda tangan
    tgl_ttd = datetime.date(tahun, bulan, get_last_day_of_month(tahun, bulan))
    ttd_str = format_tanggal_indonesia(tgl_ttd)
    s_row = start + 3
    ws.write(s_row, 0, "Mengetahui,")
    ws.write(s_row, 3, ttd_str)
    ws.write(s_row + 1, 0, "Ketua DKM")
    ws.write(s_row + 1, 3, "Bendahara")
    ws.write(s_row + 4, 0, nama_ketua)
    ws.write(s_row + 4, 3, nama_bendahara)

    # ukuran kolom
    ws.set_column("A:A", 18)
    ws.set_column("B:B", 45)
    ws.set_column("C:E", 20)

    writer.close()
    return output.getvalue()

# ---------- Sidebar ----------
with st.sidebar:
    st.header("Pengaturan Laporan")

    existing_periods = get_available_periods()
    st.markdown("**Pilih Periode Tersimpan**")
    selected_period = st.selectbox("Periode Sebelumnya", ["(Buat Periode Baru)"] + existing_periods)

    now = datetime.date.today()
    bulan = st.selectbox("Bulan", range(1, 13), index=now.month - 1)
    tahun = st.selectbox("Tahun", range(now.year - 5, now.year + 1), index=5)
    periode = f"{tahun}-{bulan:02d}"

    nama_ketua = st.text_input("Nama Ketua DKM", "H Didi Rosyadi, ST")
    nama_bendahara = st.text_input("Nama Bendahara", "Sudiro")
    saldo_awal = st.number_input("Saldo Awal (Rp)", min_value=0, step=1000)

    if st.button("Reset Data Periode Ini"):
        conn.execute("DELETE FROM transaksi WHERE periode=?", (periode,))
        conn.commit()
        st.success("Data dihapus.")
        st.rerun()

# Tentukan periode aktif
if selected_period != "(Buat Periode Baru)":
    periode_aktif = selected_period
    tahun, bulan = map(int, selected_period.split("-"))
else:
    periode_aktif = periode

# ---------- Main Layout ----------
st.title("Laporan Keuangan Kas Masjid Jam'i Al Faizin")
st.caption("Antarmuka sederhana dengan riwayat periode laporan otomatis.")

st.markdown(f"### Periode Aktif: {calendar.month_name[bulan]} {tahun}")

col1, col2 = st.columns([2, 3])
with col1:
    st.subheader("Tambah Transaksi")
    tgl = st.date_input("Tanggal", value=datetime.date(tahun, bulan, 1),
                        min_value=datetime.date(tahun, bulan, 1),
                        max_value=datetime.date(tahun, bulan, get_last_day_of_month(tahun, bulan)))
    ket = st.text_area("Keterangan")
    jenis = st.radio("Jenis Transaksi", ["Pemasukan", "Pengeluaran"], horizontal=True)
    nominal = st.number_input("Nominal (Rp)", min_value=0, step=1000)
    if st.button("Tambah Transaksi", use_container_width=True):
        if ket.strip() == "" or nominal == 0:
            st.error("Isi keterangan dan nominal dengan benar.")
        else:
            insert_data(periode_aktif, tgl, ket, nominal if jenis == "Pemasukan" else 0,
                        nominal if jenis == "Pengeluaran" else 0)
            st.success("Transaksi ditambahkan.")
            st.rerun()

# ---------- Data ----------
df = load_data(periode_aktif)
init_row = pd.DataFrame({
    "id": [0],
    "periode": [periode_aktif],
    "tanggal": [datetime.date(tahun, bulan, 1)],
    "keterangan": [f"Saldo Awal {calendar.month_name[bulan]} {tahun}"],
    "pemasukan": [saldo_awal],
    "pengeluaran": [0.0]
})
df_all = pd.concat([init_row, df], ignore_index=True)
df_all["Saldo"] = df_all["pemasukan"].cumsum() - df_all["pengeluaran"].cumsum()

# ---------- Tabel ----------
st.markdown("### Daftar Transaksi")
if df_all.empty:
    st.info("Belum ada data transaksi untuk periode ini.")
else:
    df_display = df_all.rename(columns={
    "tanggal": "Tanggal",
    "keterangan": "Keterangan",
    "pemasukan": "Pemasukan",
    "pengeluaran": "Pengeluaran"
})

# ubah tampilan nominal di Streamlit biar pakai format Rp.
df_display["Pemasukan"] = df_display["Pemasukan"].apply(lambda x: f"Rp. {x:,.0f}".replace(",", "."))
df_display["Pengeluaran"] = df_display["Pengeluaran"].apply(lambda x: f"Rp. {x:,.0f}".replace(",", "."))
df_display["Saldo"] = df_display["Saldo"].apply(lambda x: f"Rp. {x:,.0f}".replace(",", "."))

st.data_editor(
    df_display[["Tanggal", "Keterangan", "Pemasukan", "Pengeluaran", "Saldo"]],
    hide_index=True,
    use_container_width=True,
    key="editor"
)


# ---------- Ringkasan ----------
total_in = df_all["pemasukan"].sum()
total_out = df_all["pengeluaran"].sum()
saldo = df_all["Saldo"].iloc[-1] if not df_all.empty else saldo_awal
st.markdown("### Ringkasan")
c1, c2, c3 = st.columns(3)
c1.metric("Total Pemasukan", f"Rp {total_in:,.0f}".replace(",", "."))
c2.metric("Total Pengeluaran", f"Rp {total_out:,.0f}".replace(",", "."))
c3.metric("Saldo Akhir", f"Rp {saldo:,.0f}".replace(",", "."))

# ---------- Excel ----------
st.markdown("### Unduh Laporan Excel")
excel_bytes = to_excel(df_all.rename(columns={
    "tanggal": "Tanggal",
    "keterangan": "Keterangan",
    "pemasukan": "Pemasukan",
    "pengeluaran": "Pengeluaran",
    "Saldo": "Saldo"
}), nama_ketua, nama_bendahara, bulan, tahun)
st.download_button(
    "💾 Download Laporan Excel",
    data=excel_bytes,
    file_name=f"Laporan_Keuangan_{calendar.month_name[bulan]}_{tahun}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
