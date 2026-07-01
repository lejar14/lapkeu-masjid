from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from jinja2 import Environment, FileSystemLoader
import sqlite3
import datetime
import calendar
import os
from io import BytesIO
import xlsxwriter

PREFIX = os.environ.get("PREFIX", "").rstrip("/")

app = FastAPI(root_path=PREFIX)
app.mount(f"{PREFIX}/static", StaticFiles(directory="static"), name="static")
_jinja = Environment(loader=FileSystemLoader("templates"), autoescape=True)
_jinja.globals["prefix"] = PREFIX

DB_PATH = "data/masjid.db"
os.makedirs("data", exist_ok=True)

BULAN = ["Januari","Februari","Maret","April","Mei","Juni",
         "Juli","Agustus","September","Oktober","November","Desember"]


def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS transaksi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            periode TEXT,
            tanggal TEXT,
            keterangan TEXT,
            pemasukan REAL DEFAULT 0,
            pengeluaran REAL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS periode_config (
            periode TEXT PRIMARY KEY,
            saldo_awal REAL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        );
    """)
    conn.commit()
    return conn


conn = get_conn()


def setting_get(key, default=""):
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row["value"] if row else default


def setting_set(key, value):
    conn.execute("INSERT OR REPLACE INTO settings (key,value) VALUES (?,?)", (key, value))
    conn.commit()


def saldo_awal_get(periode):
    row = conn.execute("SELECT saldo_awal FROM periode_config WHERE periode=?", (periode,)).fetchone()
    return float(row["saldo_awal"]) if row else 0.0


def saldo_awal_set(periode, nominal):
    conn.execute("INSERT OR REPLACE INTO periode_config (periode,saldo_awal) VALUES (?,?)", (periode, nominal))
    conn.commit()


def fmt_rp(val):
    if not val:
        return "Rp 0"
    return "Rp " + f"{int(val):,}".replace(",", ".")


def fmt_tgl(tgl_str):
    try:
        d = datetime.date.fromisoformat(str(tgl_str))
        return f"{d.day} {BULAN[d.month-1]} {d.year}"
    except Exception:
        return str(tgl_str)


def periode_label(periode):
    y, m = map(int, periode.split("-"))
    return f"{BULAN[m-1]} {y}"


def nav_periode(periode, delta):
    y, m = map(int, periode.split("-"))
    m += delta
    if m > 12:
        m, y = 1, y + 1
    elif m < 1:
        m, y = 12, y - 1
    return f"{y}-{m:02d}"


def compute(periode):
    saldo = saldo_awal_get(periode)
    rows = conn.execute(
        "SELECT * FROM transaksi WHERE periode=? ORDER BY tanggal ASC, id ASC", (periode,)
    ).fetchall()

    y, m = map(int, periode.split("-"))
    display = [{
        "id": None,
        "tanggal": f"{y}-{m:02d}-01",
        "keterangan": "Saldo Awal",
        "pemasukan": saldo,
        "pengeluaran": 0.0,
        "saldo": saldo,
        "is_awal": True,
    }]

    total_masuk = sum(r["pemasukan"] for r in rows)
    total_keluar = sum(r["pengeluaran"] for r in rows)

    for r in rows:
        saldo = saldo + r["pemasukan"] - r["pengeluaran"]
        display.append({
            "id": r["id"],
            "tanggal": r["tanggal"],
            "keterangan": r["keterangan"],
            "pemasukan": r["pemasukan"],
            "pengeluaran": r["pengeluaran"],
            "saldo": saldo,
            "is_awal": False,
        })

    return display, total_masuk, total_keluar, saldo


def current_periode():
    t = datetime.date.today()
    return f"{t.year}-{t.month:02d}"


_jinja.filters["fmt_rp"] = fmt_rp
_jinja.filters["fmt_tgl"] = fmt_tgl


def render(name: str, **ctx) -> HTMLResponse:
    return HTMLResponse(_jinja.get_template(name).render(**ctx))


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def index(request: Request, periode: str = None, edit: int = None, ok: str = None):
    if not periode:
        periode = current_periode()

    y, m = map(int, periode.split("-"))
    last_day = calendar.monthrange(y, m)[1]
    display, total_masuk, total_keluar, saldo_akhir = compute(periode)

    edit_row = None
    if edit:
        row = conn.execute("SELECT * FROM transaksi WHERE id=?", (edit,)).fetchone()
        if row:
            edit_row = dict(row)
            edit_row["jenis"] = "pemasukan" if row["pemasukan"] > 0 else "pengeluaran"
            edit_row["nominal"] = row["pemasukan"] if row["pemasukan"] > 0 else row["pengeluaran"]

    return render("index.html",
        periode=periode,
        periode_label=periode_label(periode),
        periode_prev=nav_periode(periode, -1),
        periode_next=nav_periode(periode, 1),
        display=display,
        total_masuk=total_masuk,
        total_keluar=total_keluar,
        saldo_akhir=saldo_akhir,
        saldo_awal=saldo_awal_get(periode),
        min_date=f"{y}-{m:02d}-01",
        max_date=f"{y}-{m:02d}-{last_day:02d}",
        nama_ketua=setting_get("nama_ketua", "H Didi Rosyadi, ST"),
        nama_bendahara=setting_get("nama_bendahara", "Sudiro"),
        ada_transaksi=len(display) > 1,
        edit_row=edit_row,
        flash=ok,
    )


@app.post("/tambah")
async def tambah(
    periode: str = Form(...),
    tanggal: str = Form(...),
    keterangan: str = Form(...),
    jenis: str = Form(...),
    nominal: float = Form(...),
):
    if keterangan.strip() and nominal > 0:
        masuk = nominal if jenis == "pemasukan" else 0.0
        keluar = nominal if jenis == "pengeluaran" else 0.0
        conn.execute(
            "INSERT INTO transaksi (periode,tanggal,keterangan,pemasukan,pengeluaran) VALUES (?,?,?,?,?)",
            (periode, tanggal, keterangan.strip(), masuk, keluar),
        )
        conn.commit()
    return RedirectResponse(f"{PREFIX}/?periode={periode}&ok=tambah", status_code=303)


@app.post("/hapus/{tid}")
async def hapus(tid: int, periode: str = Form(...)):
    conn.execute("DELETE FROM transaksi WHERE id=?", (tid,))
    conn.commit()
    return RedirectResponse(f"{PREFIX}/?periode={periode}&ok=hapus", status_code=303)


@app.post("/saldo-awal")
async def update_saldo_awal(
    periode: str = Form(...),
    saldo_awal: float = Form(0),
):
    saldo_awal_set(periode, saldo_awal)
    return RedirectResponse(f"{PREFIX}/?periode={periode}&ok=saldo", status_code=303)


@app.post("/edit/{tid}")
async def edit_transaksi(
    tid: int,
    periode: str = Form(...),
    tanggal: str = Form(...),
    keterangan: str = Form(...),
    jenis: str = Form(...),
    nominal: float = Form(...),
):
    if keterangan.strip() and nominal > 0:
        masuk = nominal if jenis == "pemasukan" else 0.0
        keluar = nominal if jenis == "pengeluaran" else 0.0
        conn.execute(
            "UPDATE transaksi SET tanggal=?, keterangan=?, pemasukan=?, pengeluaran=? WHERE id=?",
            (tanggal, keterangan.strip(), masuk, keluar, tid),
        )
        conn.commit()
    return RedirectResponse(f"{PREFIX}/?periode={periode}&ok=edit", status_code=303)


@app.post("/settings")
async def update_settings(
    periode: str = Form(...),
    nama_ketua: str = Form(""),
    nama_bendahara: str = Form(""),
):
    if nama_ketua.strip():
        setting_set("nama_ketua", nama_ketua.strip())
    if nama_bendahara.strip():
        setting_set("nama_bendahara", nama_bendahara.strip())
    return RedirectResponse(f"{PREFIX}/?periode={periode}&ok=nama", status_code=303)


@app.get("/export")
async def export(periode: str):
    y, m = map(int, periode.split("-"))
    display, _, _, _ = compute(periode)
    nama_ketua = setting_get("nama_ketua", "H Didi Rosyadi, ST")
    nama_bendahara = setting_get("nama_bendahara", "Sudiro")

    out = BytesIO()
    wb = xlsxwriter.Workbook(out)
    ws = wb.add_worksheet("Laporan")
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(4)
    ws.set_paper(9)

    ttl = wb.add_format({"bold": True, "font_size": 14, "align": "center"})
    hdr = wb.add_format({"bold": True, "border": 1, "align": "center", "valign": "vcenter", "bg_color": "#D3D3D3"})
    txt = wb.add_format({"border": 1, "align": "left", "text_wrap": True})
    tgl_fmt = wb.add_format({"border": 1, "align": "center"})
    rp = wb.add_format({"num_format": '"Rp. " #,##0', "border": 1, "align": "right"})
    tot = wb.add_format({"bold": True, "border": 1, "align": "center", "bg_color": "#D3D3D3"})
    tot_rp = wb.add_format({"bold": True, "num_format": '"Rp. " #,##0', "border": 1, "align": "right", "bg_color": "#D3D3D3"})

    ws.merge_range("A1:E1", "LAPORAN KEUANGAN KAS MASJID JAM'I AL FAIZIN", ttl)
    ws.merge_range("A2:E2", "LINGKUNGAN RT 010-RW 005 KEL. BENDUNGAN KEC. CILEGON", ttl)
    ws.merge_range("A3:E3", f"Periode Bulan {BULAN[m-1]} {y}", ttl)

    for i, h in enumerate(["Tanggal", "Keterangan", "Debet", "Kredit", "Saldo"]):
        ws.write(3, i, h, hdr)

    row = 4
    i = 0
    while i < len(display):
        tgl = display[i]["tanggal"]
        j = i
        while j + 1 < len(display) and display[j+1]["tanggal"] == tgl:
            j += 1
        tgl_str = fmt_tgl(tgl)
        if i == j:
            ws.write(row, 0, tgl_str, tgl_fmt)
        else:
            ws.merge_range(row, 0, row + (j - i), 0, tgl_str, tgl_fmt)
        for k in range(i, j + 1):
            d = display[k]
            ws.write(row, 1, d["keterangan"], txt)
            ws.write(row, 2, d["pemasukan"], rp)
            ws.write(row, 3, d["pengeluaran"], rp)
            ws.write(row, 4, d["saldo"], rp)
            row += 1
        i = j + 1

    total_masuk_all = sum(d["pemasukan"] for d in display)
    total_keluar_all = sum(d["pengeluaran"] for d in display)
    saldo_akhir = display[-1]["saldo"] if display else 0
    ws.merge_range(row, 0, row, 1, "JUMLAH", tot)
    ws.write(row, 2, total_masuk_all, tot_rp)
    ws.write(row, 3, total_keluar_all, tot_rp)
    ws.write(row, 4, saldo_akhir, tot_rp)

    last_day = calendar.monthrange(y, m)[1]
    ttd_str = fmt_tgl(f"{y}-{m:02d}-{last_day:02d}")
    s = row + 3
    ws.write(s, 0, "Mengetahui,")
    ws.write(s, 3, ttd_str)
    ws.write(s+1, 0, "Ketua DKM")
    ws.write(s+1, 3, "Bendahara")
    ws.write(s+4, 0, nama_ketua)
    ws.write(s+4, 3, nama_bendahara)

    ws.set_column("A:A", 18)
    ws.set_column("B:B", 45)
    ws.set_column("C:E", 20)
    wb.close()

    out.seek(0)
    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=Laporan_{BULAN[m-1]}_{y}.xlsx"},
    )
