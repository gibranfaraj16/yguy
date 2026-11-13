"""
=====================================================================
 main.py  —  SiCupang: Aplikasi Akuntansi (CASH BASIS)
 Streamlit (tampilan, 100% Python: tanpa HTML/CSS/div) + Supabase (HANYA database, tanpa Supabase Auth)
=====================================================================

CARA MENJALANKAN
  1. pip install -r requirements.txt
  2. Jalankan isi schema.sql di Supabase (menu SQL Editor).
  3. Buat akun: ubah username/password di buat_akun.sql, lalu jalankan di SQL Editor.
     (Tidak ada menu daftar; akun hanya dibuat dari sini.)
  4. Buat file  .streamlit/secrets.toml  (di folder project) isi:
         SUPABASE_URL = "https://xxxx.supabase.co"
         SUPABASE_SERVICE_KEY = "service-role-key-anda"
     service_role key hanya dipakai di server Streamlit (tidak dikirim ke browser).
     JANGAN dibagikan / di-upload ke GitHub.
  5. streamlit run main.py

LOGIN
  Akun & password ada di tabel `akun` (password disimpan sebagai hash bcrypt).
  Pengecekan dilakukan oleh fungsi database masuk_akun(). Karena tidak memakai
  Supabase Auth, pemisahan data antar akun dilakukan aplikasi lewat kolom user_id
  (lihat fungsi uid()).

LOGIKA AKUNTANSI (CASH BASIS)
  - Pendapatan dicatat saat kas DITERIMA, beban dicatat saat kas DIBAYAR.
  - Setiap transaksi = jurnal berpasangan: total Debit == total Kredit.
  - Pembelian ikan/pakan/obat/perlengkapan/lain-lain -> BEBAN (Debit beban, Kredit Kas)
  - Pembelian peralatan                              -> ASET  (Debit Peralatan, Kredit Kas)
  - Penjualan                                        -> Debit Kas, Kredit Pendapatan
  - Saldo awal: aset (Debit) - utang (Kredit) = Modal (Kredit).

ALUR PROGRAM
  main() -> belum login?  -> halaman_login()
         -> belum isi saldo awal? -> halaman_saldo_awal()
         -> sudah?  -> st.navigation (sidebar bergrup):
                       Ringkasan: Dashboard | Transaksi: Pembelian, Penjualan | Laporan: Akses Cepat
=====================================================================
"""

# ---------- IMPORT (memanggil library yang dipakai) ----------
import io                     # untuk membuat file Excel di memori (tanpa simpan ke disk)
import uuid                   # membuat kode unik untuk "bukti" jurnal
from datetime import date, timedelta   # tipe tanggal + selisih hari

import pandas as pd           # tabel data (DataFrame) + export Excel
import streamlit as st        # tampilan web
from openpyxl.utils import get_column_letter   # bantu mengatur lebar kolom Excel
from supabase import create_client             # koneksi ke Supabase

# Pengaturan halaman (judul tab browser, lebar penuh). Harus dipanggil paling awal.
st.set_page_config(page_title="SiCupang Akuntansi", page_icon="🐟", layout="wide",
                   initial_sidebar_state="expanded")


# =====================================================================
# 1. DAFTAR AKUN (Chart of Accounts)  —  boleh Anda tambah / ubah
# =====================================================================
# Format:  "kode": {"nama": ..., "tipe": ..., "normal": "D" atau "K"}
#   tipe   : Aset / Liabilitas / Ekuitas / Pendapatan / Beban
#   normal : sisi saldo normal akun. D = Debit, K = Kredit
# Jika menambah akun baru, cukup tambah baris di sini.
AKUN = {
    "1-101": {"nama": "Kas di Bank",           "tipe": "Aset",       "normal": "D"},
    "1-102": {"nama": "Piutang",               "tipe": "Aset",       "normal": "D"},
    "1-103": {"nama": "Persediaan Ikan",       "tipe": "Aset",       "normal": "D"},
    "1-104": {"nama": "Persediaan Pakan",      "tipe": "Aset",       "normal": "D"},
    "1-105": {"nama": "Persediaan Obat",       "tipe": "Aset",       "normal": "D"},
    "1-106": {"nama": "Persediaan Lainnya",    "tipe": "Aset",       "normal": "D"},
    "1-107": {"nama": "Perlengkapan",          "tipe": "Aset",       "normal": "D"},
    "1-108": {"nama": "Peralatan",             "tipe": "Aset",       "normal": "D"},
    "2-101": {"nama": "Utang Bank",            "tipe": "Liabilitas", "normal": "K"},
    "2-102": {"nama": "Utang Orang",           "tipe": "Liabilitas", "normal": "K"},
    "3-101": {"nama": "Modal",                 "tipe": "Ekuitas",    "normal": "K"},
    "4-101": {"nama": "Penjualan Ikan",        "tipe": "Pendapatan", "normal": "K"},
    "4-102": {"nama": "Penjualan Lainnya",     "tipe": "Pendapatan", "normal": "K"},
    "5-101": {"nama": "Beban Pembelian Ikan",  "tipe": "Beban",      "normal": "D"},
    "5-102": {"nama": "Beban Pakan",           "tipe": "Beban",      "normal": "D"},
    "5-103": {"nama": "Beban Obat",            "tipe": "Beban",      "normal": "D"},
    "5-104": {"nama": "Beban Perlengkapan",    "tipe": "Beban",      "normal": "D"},
    "5-105": {"nama": "Beban Lain-lain",       "tipe": "Beban",      "normal": "D"},
}
KAS = "1-101"   # kode akun kas, dipakai di semua transaksi

# Pilihan di dropdown halaman Pembelian: "tulisan di layar": "kode akun yang di-DEBIT"
PILIHAN_PEMBELIAN = {
    "Ikan":                           "5-101",
    "Pakan":                          "5-102",
    "Obat":                           "5-103",
    "Perlengkapan":                   "5-104",
    "Peralatan (dicatat sebagai aset)": "1-108",
    "Beban lain-lain (listrik, gaji, dll)": "5-105",
}
# Pilihan di dropdown halaman Penjualan: "tulisan": "kode akun yang di-KREDIT"
PILIHAN_PENJUALAN = {
    "Penjualan Ikan":    "4-101",
    "Penjualan Lainnya": "4-102",
}


# =====================================================================
# 2. FUNGSI BANTU UMUM
# =====================================================================
def rp(angka):
    """Ubah angka jadi teks rupiah: 1500000 -> 'Rp 1.500.000'."""
    angka = int(angka)
    # f"{x:,}" memberi koma ribuan (1,500,000), lalu koma kita ganti jadi titik
    teks = f"{abs(angka):,}".replace(",", ".")
    return f"-Rp {teks}" if angka < 0 else f"Rp {teks}"


@st.cache_resource          # dibuat sekali lalu dipakai ulang (tidak membuat koneksi baru tiap klik)
def _buat_client(url, key):
    return create_client(url, key)


def get_client():
    """Ambil koneksi Supabase (pakai service_role key dari secrets.toml)."""
    try:
        return _buat_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_SERVICE_KEY"])
    except Exception:
        st.error("Isi SUPABASE_URL dan SUPABASE_SERVICE_KEY di .streamlit/secrets.toml")
        st.stop()                      # hentikan program di sini


def uid():
    """ID akun yang sedang login. Dipakai di SEMUA query supaya data antar akun tidak tercampur."""
    return st.session_state.user["id"]


def ke_excel(sheets):
    """Buat file Excel (bytes). sheets = {"Nama Sheet": DataFrame, ...}"""
    buf = io.BytesIO()                                   # 'file' sementara di memori
    with pd.ExcelWriter(buf, engine="openpyxl") as w:    # 'with' = otomatis tutup file
        for nama, df in sheets.items():
            nama = nama[:31]                             # batas nama sheet Excel: 31 huruf
            df.to_excel(w, sheet_name=nama, index=False)
            ws = w.sheets[nama]
            # atur lebar tiap kolom sesuai isi
            for i, kol in enumerate(df.columns, start=1):
                isi = [len(str(kol))] + [len(str(x)) for x in df[kol].head(200)]
                ws.column_dimensions[get_column_letter(i)].width = min(max(isi) + 2, 50)
    return buf.getvalue()


def tombol_excel(label, sheets, nama_file, kunci):
    """Tombol unduh Excel. 'kunci' harus unik per tombol."""
    st.download_button(
        label, data=ke_excel(sheets), file_name=nama_file, key=kunci,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def format_uang(df, kolom):
    """Salinan DataFrame dengan kolom uang diformat 'Rp ...' (untuk tampilan saja)."""
    d = df.copy()
    for k in kolom:
        # lambda = fungsi mini. Kosong (NaN) ditampilkan kosong, selain itu diformat rupiah
        d[k] = d[k].map(lambda v: "" if pd.isna(v) else rp(v))
    return d


# =====================================================================
# 3. DATABASE: ambil & simpan jurnal
# =====================================================================
KOLOM_JURNAL = ["id", "bukti", "tanggal", "keterangan", "jenis", "kode_akun", "debit", "kredit"]


def ambil_semua(tabel):
    """Ambil semua baris tabel. Supabase membatasi 1000 baris per permintaan,
    jadi kita ambil bertahap (halaman demi halaman)."""
    hasil, mulai = [], 0
    while True:
        res = get_client().table(tabel).select("*").eq("user_id", uid()).order("id").range(mulai, mulai + 999).execute()
        hasil += res.data
        if len(res.data) < 1000:
            break
        mulai += 1000
    return hasil


def muat_jurnal():
    """Ambil seluruh jurnal user sebagai DataFrame. Disimpan sementara (cache)
    di session_state supaya tidak query berulang tiap klik."""
    if "jurnal_df" not in st.session_state:
        df = pd.DataFrame(ambil_semua("jurnal"), columns=KOLOM_JURNAL)
        df["tanggal"] = pd.to_datetime(df["tanggal"]).dt.date
        df[["debit", "kredit"]] = df[["debit", "kredit"]].astype("int64")
        df["akun"] = df["kode_akun"].map(lambda k: AKUN[k]["nama"])
        st.session_state.jurnal_df = df.sort_values(["tanggal", "id"]).reset_index(drop=True)
    return st.session_state.jurnal_df


def bersihkan_cache():
    """Hapus cache jurnal -> data diambil ulang dari database di pemuatan berikutnya."""
    st.session_state.pop("jurnal_df", None)


def baris_jurnal(tanggal, keterangan, jenis, entri):
    """Bentuk baris jurnal siap simpan.
    entri = daftar (kode_akun, debit, kredit). Wajib seimbang!"""
    total_d = sum(e[1] for e in entri)
    total_k = sum(e[2] for e in entri)
    if total_d != total_k:
        raise ValueError(f"Jurnal tidak seimbang: debit {total_d} != kredit {total_k}")
    bukti = str(uuid.uuid4())   # satu kode untuk semua baris dalam satu transaksi
    # 'for kode, d, k in entri' = membongkar tiap tuple (kode, debit, kredit)
    return [
        {"user_id": uid(), "bukti": bukti, "tanggal": tanggal.isoformat(), "keterangan": keterangan,
         "jenis": jenis, "kode_akun": kode, "debit": int(d), "kredit": int(k)}
        for kode, d, k in entri
    ]


def simpan_transaksi(tanggal, keterangan, jenis, entri):
    """Simpan satu transaksi (beberapa baris jurnal sekaligus) ke Supabase."""
    get_client().table("jurnal").insert(baris_jurnal(tanggal, keterangan, jenis, entri)).execute()
    bersihkan_cache()


# =====================================================================
# 4. PERHITUNGAN AKUNTANSI
# =====================================================================
def saldo_per_akun(df):
    """Hitung saldo tiap akun. Saldo dihitung sesuai sisi normalnya:
       akun normal Debit  -> saldo = debit - kredit
       akun normal Kredit -> saldo = kredit - debit
    Hasil: {"1-101": 5000000, ...}"""
    hasil = {}
    for kode, info in AKUN.items():
        d = int(df.loc[df["kode_akun"] == kode, "debit"].sum())
        k = int(df.loc[df["kode_akun"] == kode, "kredit"].sum())
        hasil[kode] = d - k if info["normal"] == "D" else k - d
    return hasil


def total_tipe(saldo, tipe):
    """Jumlahkan saldo semua akun bertipe tertentu (mis. 'Pendapatan')."""
    # sum(... for ... if ...) = jumlahkan hanya yang memenuhi syarat 'if'
    return sum(s for kode, s in saldo.items() if AKUN[kode]["tipe"] == tipe)


def filter_periode(df, awal=None, akhir=None):
    """Ambil baris jurnal pada rentang tanggal tertentu."""
    if awal is not None:
        df = df[df["tanggal"] >= awal]
    if akhir is not None:
        df = df[df["tanggal"] <= akhir]
    return df


def laporan_jurnal_umum(df):
    """Jurnal umum: urut tanggal, akun kredit dijorok ke kanan (aturan penulisan jurnal)."""
    j = pd.DataFrame({
        "Tanggal": df["tanggal"],
        "Bukti": df["bukti"].str[:8],
        "Keterangan": df["keterangan"],
        # kalau sisi kredit > 0, tambahkan spasi di depan nama akun
        "Akun": [("      " + a) if k > 0 else a for a, k in zip(df["akun"], df["kredit"])],
        "Debit": df["debit"],
        "Kredit": df["kredit"],
    })
    return j.reset_index(drop=True)


def laporan_buku_besar(df, kode):
    """Buku besar satu akun + saldo berjalan."""
    d = df[df["kode_akun"] == kode].sort_values(["tanggal", "id"])
    normal = AKUN[kode]["normal"]
    saldo, baris = 0, []
    for _, r in d.iterrows():        # iterrows = ulangi tiap baris tabel
        saldo += (r["debit"] - r["kredit"]) if normal == "D" else (r["kredit"] - r["debit"])
        baris.append({"Tanggal": r["tanggal"], "Keterangan": r["keterangan"],
                      "Debit": r["debit"], "Kredit": r["kredit"], "Saldo": saldo})
    return pd.DataFrame(baris, columns=["Tanggal", "Keterangan", "Debit", "Kredit", "Saldo"])


def laporan_neraca_saldo(df):
    """Neraca saldo (trial balance): saldo akhir tiap akun di kolom Debit ATAU Kredit.
    Total Debit harus sama dengan total Kredit."""
    baris = []
    for kode, info in AKUN.items():
        d = int(df.loc[df["kode_akun"] == kode, "debit"].sum())
        k = int(df.loc[df["kode_akun"] == kode, "kredit"].sum())
        if d == 0 and k == 0:
            continue                  # akun tanpa transaksi tidak ditampilkan
        net = d - k
        baris.append({"Kode": kode, "Akun": info["nama"],
                      "Debit": net if net > 0 else 0,
                      "Kredit": -net if net < 0 else 0})
    t = pd.DataFrame(baris, columns=["Kode", "Akun", "Debit", "Kredit"])
    total = pd.DataFrame([{"Kode": "", "Akun": "TOTAL",
                           "Debit": t["Debit"].sum(), "Kredit": t["Kredit"].sum()}])
    return pd.concat([t, total], ignore_index=True)


def laporan_laba_rugi(df):
    """Laporan laba rugi (income statement) dari akun Pendapatan & Beban."""
    saldo = saldo_per_akun(df)
    baris = []                          # tiap item: (keterangan, jumlah)

    def seksi(judul, tipe):             # fungsi di dalam fungsi (hanya dipakai di sini)
        baris.append((judul, None))
        for kode, info in AKUN.items():
            if info["tipe"] == tipe and saldo[kode] != 0:
                baris.append(("    " + info["nama"], saldo[kode]))
        total = total_tipe(saldo, tipe)
        baris.append((f"Total {judul.title()}", total))
        return total

    pendapatan = seksi("PENDAPATAN", "Pendapatan")
    beban = seksi("BEBAN", "Beban")
    baris.append(("LABA (RUGI) BERSIH", pendapatan - beban))
    return pd.DataFrame(baris, columns=["Keterangan", "Jumlah"])


# =====================================================================
# 5. LOGIN & SIDEBAR AKUN  (akun di tabel `akun`, tanpa Supabase Auth)
# =====================================================================
def judul_halaman(judul, deskripsi):
    """Kepala halaman yang seragam di semua menu: judul besar + keterangan kecil."""
    st.title(judul)
    st.caption(deskripsi)


def halaman_login():
    # Tiga kolom: kiri & kanan kosong, kartu login di tengah -> posisi di tengah layar
    _, tengah, _ = st.columns([1, 1.2, 1])
    with tengah:
        st.write("")
        st.write("")
        with st.container(border=True):
            st.header("🐟 SiCupang")
            st.caption("Aplikasi akuntansi sederhana (cash basis)")
            with st.form("form_masuk"):                   # form = kirim sekali saat tombol ditekan
                username = st.text_input("Username", placeholder="Masukkan username")
                password = st.text_input("Password", type="password", placeholder="Masukkan password")
                kirim = st.form_submit_button("Masuk", type="primary")
        if not kirim:
            return
        if not username.strip() or not password:
            st.error("Username dan password wajib diisi.")
            return
        try:
            # panggil fungsi database masuk_akun(); hasil: [] jika salah, [{id, nama}] jika benar
            res = get_client().rpc("masuk_akun", {"p_username": username.strip(),
                                                  "p_password": password}).execute()
        except Exception as e:
            st.error(f"Gagal terhubung ke database: {e}")
            return
        if not res.data:
            st.error("Username atau password salah.")
            return
        akun = res.data[0]
        st.session_state.user = {"id": akun["id"], "nama": akun["nama"]}
        bersihkan_cache()
        st.rerun()                                        # muat ulang halaman


def keluar():
    """Logout: kosongkan semua memori sesi."""
    st.session_state.clear()
    st.rerun()


def sidebar_akun():
    """Kartu akun di bagian bawah sidebar (di bawah menu): nama, info, tombol keluar."""
    with st.sidebar:
        st.divider()
        with st.container(border=True):
            st.markdown("**🐟 SiCupang**")
            st.caption(f":material/account_circle: {st.session_state.user['nama']}")
            st.caption("Pembukuan cash basis")
            if st.button("Keluar", icon=":material/logout:", key="btn_keluar"):
                keluar()


def saldo_awal_selesai():
    """Cek di database: apakah akun ini sudah menyelesaikan saldo awal?"""
    res = (get_client().table("profil").select("saldo_awal_selesai")
           .eq("user_id", uid()).execute())
    return bool(res.data and res.data[0]["saldo_awal_selesai"])


# =====================================================================
# 6. SALDO AWAL (wizard 7 langkah, pertama kali login)
# =====================================================================
TOTAL_LANGKAH = 7
OPSI = ["Tidak", "Ya"]

# Definisi tabel isian. Tiap kolom: (judul kolom, peran)
#   peran: pihak (Bank/Orang) | nama (teks) | nominal (rupiah) | umur (bulan)
# Ingin ubah judul kolom / tambah kolom? Ubah di sini.
TABEL = {
    "utang":        [("Pihak", "pihak"), ("Nama bank / orang", "nama"), ("Nominal (Rp)", "nominal")],
    "piutang":      [("Nama pihak", "nama"), ("Nominal (Rp)", "nominal")],
    "ikan":         [("Jenis ikan", "nama"), ("Harga (Rp)", "nominal")],
    "pakan":        [("Nama pakan", "nama"), ("Harga (Rp)", "nominal")],
    "obat":         [("Nama obat", "nama"), ("Harga (Rp)", "nominal")],
    "lainnya":      [("Nama barang", "nama"), ("Harga (Rp)", "nominal")],
    "peralatan":    [("Nama peralatan", "nama"), ("Harga (Rp)", "nominal"),
                     ("Perkiraan pemakaian (bulan)", "umur")],
    "perlengkapan": [("Nama perlengkapan", "nama"), ("Harga (Rp)", "nominal")],
}

def init_wizard():
    """Siapkan 'memori' wizard saat pertama dibuka."""
    if "sa" not in st.session_state:
        st.session_state.sa = {
            "langkah": 0,
            "kas_ada": "Tidak", "kas_nominal": 0,
            "utang_ada": "Tidak", "piutang_ada": "Tidak", "lain_ada": "Tidak",
            "tanggal": date.today(),
            "in": {},    # tabel awal yang ditampilkan editor
            "out": {},   # hasil editan terbaru dari editor
        }

def tanya(kunci, pertanyaan):
    """Pertanyaan Ya/Tidak. Jawaban disimpan di sa[kunci]. Return True jika 'Ya'."""
    sa = st.session_state.sa
    sa[kunci] = st.radio(pertanyaan, OPSI, index=OPSI.index(sa[kunci]),
                         horizontal=True, key=f"radio_{kunci}")
    return sa[kunci] == "Ya"


def editor(nama):
    """Tampilkan tabel isian yang bisa ditambah barisnya (klik baris kosong di bawah)."""
    sa = st.session_state.sa
    if nama not in sa["in"]:
        # tabel kosong dengan kolom sesuai TABEL (teks -> object, angka -> float)
        sa["in"][nama] = pd.DataFrame({
            kol: pd.Series(dtype="object" if peran in ("pihak", "nama") else "float64")
            for kol, peran in TABEL[nama]
        })
    config = {}                                       # pengaturan tampilan tiap kolom
    for kol, peran in TABEL[nama]:
        if peran == "pihak":
            config[kol] = st.column_config.SelectboxColumn(kol, options=["Bank", "Orang"])
        elif peran == "nominal":
            config[kol] = st.column_config.NumberColumn(kol, min_value=0, step=1000, format="%d")
        elif peran == "umur":
            config[kol] = st.column_config.NumberColumn(kol, min_value=0, step=1, format="%d")
    # num_rows="dynamic" = pengguna boleh menambah/menghapus baris
    sa["out"][nama] = st.data_editor(sa["in"][nama], num_rows="dynamic",
                                     column_config=config, key=f"editor_{nama}")


def bersih(nama):
    """Ambil isi tabel, buang baris kosong, cek kelengkapan.
    Return: list dict {'pihak','nama','nominal','umur'}. Raise ValueError jika ada yang salah."""
    df = st.session_state.sa["out"].get(nama)
    if df is None:
        return []
    hasil = []
    for no, (_, r) in enumerate(df.iterrows(), start=1):
        isi, kosong = {}, True
        for kol, peran in TABEL[nama]:
            v = r[kol]
            if peran in ("pihak", "nama"):
                v = "" if pd.isna(v) else str(v).strip()
            else:
                v = 0 if pd.isna(v) else int(v)
            if v:
                kosong = False
            isi[peran] = v
        if kosong:
            continue                                  # baris kosong dilewati
        for kol, peran in TABEL[nama]:
            if not isi[peran]:
                raise ValueError(f"Tabel {nama}, baris {no}: kolom '{kol}' belum diisi.")
        hasil.append(isi)
    return hasil


def kumpulkan():
    """Kumpulkan semua jawaban wizard menjadi satu dict data bersih."""
    sa = st.session_state.sa
    return {
        "kas": int(sa["kas_nominal"]) if sa["kas_ada"] == "Ya" else 0,
        "utang": bersih("utang") if sa["utang_ada"] == "Ya" else [],
        "piutang": bersih("piutang") if sa["piutang_ada"] == "Ya" else [],
        "ikan": bersih("ikan"),
        "pakan": bersih("pakan"),
        "obat": bersih("obat"),
        "lainnya": bersih("lainnya") if sa["lain_ada"] == "Ya" else [],
        "peralatan": bersih("peralatan"),
        "perlengkapan": bersih("perlengkapan"),
    }


def entri_saldo_awal(d):
    """Susun jurnal saldo awal:
       Debit  : kas, piutang, persediaan, perlengkapan, peralatan
       Kredit : utang bank, utang orang
       Modal  : selisihnya (aset - utang) supaya jurnal seimbang."""
    saldo = {}                                        # kode akun -> jumlah

    def tambah(kode, jumlah):
        if jumlah:
            saldo[kode] = saldo.get(kode, 0) + jumlah  # .get(x, 0) = ambil x, kalau tak ada pakai 0

    total = lambda daftar: sum(r["nominal"] for r in daftar)
    tambah("1-101", d["kas"])
    tambah("1-102", total(d["piutang"]))
    tambah("1-103", total(d["ikan"]))
    tambah("1-104", total(d["pakan"]))
    tambah("1-105", total(d["obat"]))
    tambah("1-106", total(d["lainnya"]))
    tambah("1-107", total(d["perlengkapan"]))
    tambah("1-108", total(d["peralatan"]))
    for r in d["utang"]:
        tambah("2-101" if r["pihak"] == "Bank" else "2-102", r["nominal"])

    entri = []
    for kode, jml in saldo.items():
        if AKUN[kode]["normal"] == "D":
            entri.append((kode, jml, 0))              # akun normal debit -> di sisi debit
        else:
            entri.append((kode, 0, jml))              # akun normal kredit -> di sisi kredit
    selisih = sum(e[1] for e in entri) - sum(e[2] for e in entri)
    if selisih > 0:
        entri.append(("3-101", 0, selisih))           # aset > utang: modal di kredit
    elif selisih < 0:
        entri.append(("3-101", -selisih, 0))          # utang > aset: modal negatif (debit)
    return entri


def item_saldo_awal(d):
    """Rincian saldo awal (jenis ikan, harga, umur pemakaian, dll) untuk tabel saldo_awal_item."""
    items = []
    if d["kas"]:
        items.append({"kelompok": "kas_bank", "pihak": None, "nama": "Kas di bank",
                      "nominal": d["kas"], "umur_bulan": None})
    for kelompok in ("utang", "piutang", "ikan", "pakan", "obat", "lainnya",
                     "peralatan", "perlengkapan"):
        for r in d[kelompok]:
            items.append({"kelompok": kelompok, "pihak": r.get("pihak"), "nama": r["nama"],
                          "nominal": r["nominal"], "umur_bulan": r.get("umur")})
    return items


def simpan_saldo_awal(d, tanggal):
    """Simpan saldo awal ke Supabase: jurnal + rincian + tanda 'selesai'."""
    sb = get_client()
    u = uid()
    entri = entri_saldo_awal(d)
    # hapus sisa percobaan sebelumnya milik akun ini (jika sempat gagal di tengah jalan)
    sb.table("jurnal").delete().eq("user_id", u).eq("jenis", "saldo_awal").execute()
    sb.table("saldo_awal_item").delete().eq("user_id", u).execute()
    if entri:
        sb.table("jurnal").insert(
            baris_jurnal(tanggal, "Saldo awal", "saldo_awal", entri)).execute()
    items = [{**it, "user_id": u} for it in item_saldo_awal(d)]   # tambahkan user_id ke tiap baris
    if items:
        sb.table("saldo_awal_item").insert(items).execute()
    # terakhir: tandai selesai (upsert = insert, atau update jika sudah ada)
    sb.table("profil").upsert({"user_id": u, "saldo_awal_selesai": True}).execute()
    bersihkan_cache()
    st.session_state.sa_selesai = True
    st.session_state.pop("sa", None)


def validasi_langkah(l):
    """Cek isian langkah ke-l. Return teks error, atau None jika aman."""
    sa = st.session_state.sa
    try:
        if l == 1 and sa["kas_ada"] == "Ya" and sa["kas_nominal"] <= 0:
            return "Nominal kas di bank harus lebih dari 0."
        if l == 2 and sa["utang_ada"] == "Ya" and not bersih("utang"):
            return "Isi minimal satu utang, atau pilih 'Tidak'."
        if l == 3 and sa["piutang_ada"] == "Ya" and not bersih("piutang"):
            return "Isi minimal satu piutang, atau pilih 'Tidak'."
        if l == 4:
            for t in ("ikan", "pakan", "obat"):
                bersih(t)
            if sa["lain_ada"] == "Ya" and not bersih("lainnya"):
                return "Isi persediaan lainnya, atau pilih 'Tidak'."
        if l == 5:
            bersih("peralatan")
        if l == 6:
            bersih("perlengkapan")
    except ValueError as e:                           # pesan dari bersih()
        return str(e)
    return None


def pindah(n):
    """Pindah ke langkah n. Hasil editan disalin supaya tidak hilang saat kembali."""
    sa = st.session_state.sa
    sa["in"] = {k: v.copy() for k, v in sa["out"].items()}
    sa["langkah"] = n
    st.rerun()


NAMA_LANGKAH = ["Mulai", "Kas di Bank", "Utang", "Piutang",
                "Persediaan", "Peralatan", "Perlengkapan"]


def halaman_saldo_awal():
    init_wizard()
    sa = st.session_state.sa
    l = sa["langkah"]

    # Kolom kiri & kanan kosong -> isi wizard berada di tengah, tidak melebar
    _, tengah, _ = st.columns([1, 3, 1])
    with tengah:
        st.title("Saldo Awal")
        st.progress((l + 1) / TOTAL_LANGKAH,
                    text=f"Langkah {l + 1} dari {TOTAL_LANGKAH}  ·  {NAMA_LANGKAH[l]}")

        with st.container(border=True):
            if l == 0:    # 1) pernyataan
                st.subheader("Anda harus memasukkan beberapa saldo untuk memulai❗")
                st.write("Semua isian berupa nominal uang, dipakai sebagai saldo awal pembukuan Anda.")

            elif l == 1:  # 2) kas di bank
                st.subheader("Kas di Bank")
                if tanya("kas_ada", "Apakah Anda memiliki kas di bank?"):
                    sa["kas_nominal"] = st.number_input(
                        "Nominal kas di bank (Rp)", min_value=0, step=100000,
                        value=int(sa["kas_nominal"]), key="num_kas")

            elif l == 2:  # 3) utang
                st.subheader("Utang")
                if tanya("utang_ada", "Apakah Anda memiliki utang?"):
                    st.caption("Pilih Bank atau Orang, isi nama dan nominal. Klik baris kosong di bawah tabel untuk menambah.")
                    editor("utang")

            elif l == 3:  # 4) piutang
                st.subheader("Piutang")
                if tanya("piutang_ada", "Apakah Anda memiliki piutang?"):
                    st.caption("Klik baris kosong di bawah tabel untuk menambah.")
                    editor("piutang")

            elif l == 4:  # 5) persediaan
                st.subheader("Persediaan")
                st.caption("Kosongkan tabel jika tidak ada. Klik baris kosong di bawah tabel untuk menambah.")
                ada_lain = tanya("lain_ada", "Apakah ada persediaan lain selain ikan, pakan, dan obat?")
                st.markdown("**Ikan**")
                editor("ikan")
                st.markdown("**Pakan**")
                editor("pakan")
                st.markdown("**Obat**")
                editor("obat")
                if ada_lain:
                    st.markdown("**Persediaan lainnya**")
                    editor("lainnya")

            elif l == 5:  # 6) peralatan
                st.subheader("Peralatan")
                st.caption("Perkiraan pemakaian = berapa bulan alat diperkirakan bisa dipakai. Kosongkan jika tidak ada.")
                editor("peralatan")

            elif l == 6:  # 7) perlengkapan + ringkasan
                st.subheader("Perlengkapan")
                st.caption("Kosongkan jika tidak ada.")
                editor("perlengkapan")
                sa["tanggal"] = st.date_input("Tanggal saldo awal", value=sa["tanggal"])
                try:                                   # ringkasan sementara
                    entri = entri_saldo_awal(kumpulkan())
                    aset = sum(e[1] for e in entri if e[0].startswith("1-"))
                    utang = sum(e[2] for e in entri if e[0].startswith("2-"))
                    modal = sum(e[2] - e[1] for e in entri if e[0] == "3-101")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Total Aset", rp(aset), border=True)
                    c2.metric("Total Utang", rp(utang), border=True)
                    c3.metric("Modal", rp(modal), border=True)
                except ValueError as e:
                    st.warning(str(e))

        # ---- tombol navigasi ----
        kiri, kanan = st.columns(2)
        kembali = kiri.button("Kembali", icon=":material/arrow_back:", disabled=(l == 0))
        teks = "Oke" if l == 0 else ("Simpan" if l == TOTAL_LANGKAH - 1 else "Lanjut")
        lanjut = kanan.button(teks, type="primary")

        if kembali:
            pindah(l - 1)
        if lanjut:
            err = validasi_langkah(l)
            if err:
                st.error(err)
            elif l == TOTAL_LANGKAH - 1:
                try:
                    simpan_saldo_awal(kumpulkan(), sa["tanggal"])
                except Exception as e:
                    st.error(f"Gagal menyimpan: {e}")
                else:
                    st.rerun()                         # masuk ke aplikasi utama
            else:
                pindah(l + 1)


# =====================================================================
# 7. HALAMAN DASHBOARD
# =====================================================================
PERIODE = ["Bulan ini", "30 hari", "Tahun ini", "Kustom"]


def pilih_periode(kunci, bawaan):
    """Bar filter periode: tombol cepat (Bulan ini / 30 hari / Tahun ini) atau Kustom.
    'kunci' harus unik per halaman. Return (tanggal_awal, tanggal_akhir)."""
    hari = date.today()
    with st.container(border=True):
        c1, c2, c3 = st.columns([2, 1.5, 1.5], vertical_alignment="bottom")
        pilih = c1.segmented_control("Periode", PERIODE, default=bawaan,
                                     key=f"{kunci}_periode") or bawaan
        if pilih == "Kustom":
            awal = c2.date_input("Dari tanggal", value=hari.replace(day=1), key=f"{kunci}_awal")
            akhir = c3.date_input("Sampai tanggal", value=hari, key=f"{kunci}_akhir")
        else:
            awal = {"Bulan ini": hari.replace(day=1),
                    "30 hari": hari - timedelta(days=29),
                    "Tahun ini": date(hari.year, 1, 1)}[pilih]
            akhir = hari
            c2.caption(f"Dari {awal.strftime('%d/%m/%Y')}")
            c3.caption(f"Sampai {akhir.strftime('%d/%m/%Y')}")
    if awal > akhir:
        st.error("Tanggal awal tidak boleh setelah tanggal akhir.")
        st.stop()
    return awal, akhir


def delta_rp(sekarang, sebelum):
    """Teks selisih untuk st.metric. Negatif -> diawali '-' (otomatis merah). Nol -> tanpa delta."""
    selisih = sekarang - sebelum
    return None if selisih == 0 else rp(selisih)


def data_grafik(df, awal, akhir):
    """Pendapatan vs Beban per hari (periode pendek) atau per bulan (periode panjang).
    Return DataFrame berindeks 'Periode' dengan kolom Pendapatan & Beban, atau None jika kosong."""
    d = filter_periode(df, awal, akhir)
    tipe = d["kode_akun"].map(lambda k: AKUN[k]["tipe"])
    d = d[tipe.isin(["Pendapatan", "Beban"])].copy()
    if d.empty:
        return None
    tipe = d["kode_akun"].map(lambda k: AKUN[k]["tipe"])
    d["Pendapatan"] = (d["kredit"] - d["debit"]).where(tipe == "Pendapatan", 0)
    d["Beban"] = (d["debit"] - d["kredit"]).where(tipe == "Beban", 0)
    format_waktu = "%Y-%m-%d" if (akhir - awal).days <= 62 else "%Y-%m"
    d["Periode"] = pd.to_datetime(d["tanggal"]).dt.strftime(format_waktu)
    return d.groupby("Periode")[["Pendapatan", "Beban"]].sum().sort_index()


def tabel_transaksi(df, jenis=None):
    """Daftar transaksi pembelian / penjualan (satu baris per transaksi), terbaru di atas.
    jenis=None -> gabungan keduanya (ada kolom 'Jenis')."""
    m_beli = (df["jenis"] == "pembelian") & (df["debit"] > 0)     # sisi beban/aset
    m_jual = (df["jenis"] == "penjualan") & (df["kredit"] > 0)    # sisi pendapatan
    if jenis == "pembelian":
        t = df[m_beli]
    elif jenis == "penjualan":
        t = df[m_jual]
    else:
        t = df[m_beli | m_jual]
    t = t.sort_values(["tanggal", "id"], ascending=False)
    tabel = pd.DataFrame({
        "Tanggal": t["tanggal"],
        "Jenis": t["jenis"].astype(str).str.capitalize(),
        "Kategori": t["kode_akun"].map(lambda k: AKUN[k]["nama"]),
        "Keterangan": t["keterangan"],
        "Nominal": t["debit"] + t["kredit"],
    })
    return tabel.drop(columns="Jenis") if jenis else tabel


def halaman_dashboard():
    judul_halaman("Dashboard", f"Ringkasan keuangan  ·  {st.session_state.user['nama']}")
    df = muat_jurnal()
    awal, akhir = pilih_periode("dash", "Bulan ini")

    # Periode pembanding = periode sebelumnya dengan jumlah hari yang sama
    panjang = (akhir - awal).days + 1
    akhir_lalu = awal - timedelta(days=1)
    awal_lalu = akhir_lalu - timedelta(days=panjang - 1)

    # Laba/pendapatan/beban = hanya dalam periode.  Kas/utang/piutang = saldo sampai tanggal akhir.
    s_now = saldo_per_akun(filter_periode(df, awal, akhir))
    s_lalu = saldo_per_akun(filter_periode(df, awal_lalu, akhir_lalu))
    s_akhir = saldo_per_akun(filter_periode(df, None, akhir))
    s_awal = saldo_per_akun(filter_periode(df, None, akhir_lalu))

    pendapatan, beban = total_tipe(s_now, "Pendapatan"), total_tipe(s_now, "Beban")
    pend_lalu, beban_lalu = total_tipe(s_lalu, "Pendapatan"), total_tipe(s_lalu, "Beban")
    laba, laba_lalu = pendapatan - beban, pend_lalu - beban_lalu
    kas, piutang = s_akhir[KAS], s_akhir["1-102"]
    utang_bank, utang_orang = s_akhir["2-101"], s_akhir["2-102"]

    # ---- Baris 1: 4 kartu angka utama ----
    k = st.columns(4)
    banding = "Dibanding periode sebelumnya (jumlah hari sama)"
    k[0].metric("Pendapatan", rp(pendapatan), delta_rp(pendapatan, pend_lalu),
                border=True, help=banding)
    k[1].metric("Beban", rp(beban), delta_rp(beban, beban_lalu), delta_color="inverse",
                border=True, help=banding)
    k[2].metric("Laba (Rugi)", rp(laba), delta_rp(laba, laba_lalu),
                border=True, help=banding)
    k[3].metric("Kas di Bank", rp(kas), delta_rp(kas, s_awal[KAS]),
                border=True, help="Perubahan dibanding saldo kas di awal periode")

    # ---- Baris 2: grafik (kiri, lebar) + utang/piutang (kanan) ----
    kiri, kanan = st.columns([2, 1], gap="medium")
    with kiri:
        with st.container(border=True):
            st.subheader("Pendapatan vs Beban")
            grafik = data_grafik(df, awal, akhir)
            if grafik is None:
                st.info("Belum ada pendapatan atau beban pada periode ini.")
            else:
                st.bar_chart(grafik, color=["#16a34a", "#dc2626"], stack=False)
    with kanan:
        with st.container(border=True):
            st.subheader("Utang / Piutang")
            st.metric("Piutang", rp(piutang))
            st.metric("Total Utang", rp(utang_bank + utang_orang))
            st.caption(f"Utang Bank: {rp(utang_bank)}  |  Utang Orang: {rp(utang_orang)}")
            st.divider()
            st.metric("Piutang - Utang", rp(piutang - utang_bank - utang_orang))

    # ---- Baris 3: transaksi terbaru ----
    with st.container(border=True):
        st.subheader("Transaksi Terbaru")
        terbaru = tabel_transaksi(df).head(8)
        if terbaru.empty:
            st.info("Belum ada transaksi. Mulai dari menu Pembelian atau Penjualan.")
        else:
            st.dataframe(format_uang(terbaru, ["Nominal"]), hide_index=True)


# =====================================================================
# 8. HALAMAN PEMBELIAN & PENJUALAN (satu fungsi untuk keduanya)
# =====================================================================
def halaman_transaksi(jenis, judul, pilihan):
    """jenis   : 'pembelian' atau 'penjualan'
       pilihan : dict {tulisan: kode akun}"""
    judul_halaman(judul, "Catat transaksi baru di kiri, riwayat terbaru ada di kanan."
                  if jenis == "pembelian" else
                  "Catat penjualan baru di kiri, riwayat terbaru ada di kanan.")

    # Wadah ringkasan dibuat dulu (supaya tampil paling atas),
    # tapi DIISI paling akhir supaya angkanya sudah memuat transaksi yang baru disimpan.
    ringkasan = st.container()

    kiri, kanan = st.columns([2, 3], gap="large")

    # ---------- Kolom kiri: form input ----------
    with kiri:
        with st.container(border=True):
            st.subheader(f"Input {judul}")
            with st.form(f"form_{jenis}", clear_on_submit=True):   # kosongkan form setelah submit
                c1, c2 = st.columns(2)
                tgl = c1.date_input("Tanggal", value=date.today())
                kategori = c2.selectbox("Kategori", list(pilihan.keys()))
                ket = st.text_input("Keterangan (nama barang / pembeli / penjual)")
                nominal = st.number_input("Nominal (Rp)", min_value=0, step=1000, value=0)
                kirim = st.form_submit_button("Submit", type="primary")

            if kirim:
                kode = pilihan[kategori]
                if nominal <= 0:
                    st.error("Nominal harus lebih dari 0.")
                elif not ket.strip():
                    st.error("Keterangan wajib diisi.")
                elif jenis == "pembelian" and nominal > saldo_per_akun(muat_jurnal())[KAS]:
                    # cash basis: tidak bisa bayar kalau kas tidak cukup
                    st.error("Kas tidak cukup untuk pembelian ini.")
                else:
                    if jenis == "pembelian":
                        # Debit akun beban/aset, Kredit Kas (uang keluar)
                        entri = [(kode, nominal, 0), (KAS, 0, nominal)]
                    else:
                        # Debit Kas (uang masuk), Kredit Pendapatan
                        entri = [(KAS, nominal, 0), (kode, 0, nominal)]
                    try:
                        simpan_transaksi(tgl, ket.strip(), jenis, entri)
                    except Exception as e:
                        st.error(f"Gagal menyimpan: {e}")
                    else:
                        st.success(f"Tersimpan: {kategori} {rp(nominal)}")

    # ---------- Kolom kanan: 10 transaksi terbaru ----------
    with kanan:
        with st.container(border=True):
            st.subheader(f"10 {judul} Terbaru")
            tabel = tabel_transaksi(muat_jurnal(), jenis)
            if tabel.empty:
                st.info("Belum ada data.")
            else:
                st.dataframe(format_uang(tabel.head(10), ["Nominal"]), hide_index=True)
                tombol_excel("⬇️ Unduh semua ke Excel", {judul: tabel},
                             f"{jenis}.xlsx", f"xl_{jenis}")

    # ---------- Ringkasan di atas halaman (diisi terakhir) ----------
    with ringkasan:
        df = muat_jurnal()
        tabel = tabel_transaksi(df, jenis)
        bulan_ini = tabel[tabel["Tanggal"] >= date.today().replace(day=1)]
        m = st.columns(3)
        m[0].metric(f"Total {judul} bulan ini", rp(bulan_ini["Nominal"].sum()), border=True)
        m[1].metric("Jumlah transaksi bulan ini", f"{len(bulan_ini)}", border=True)
        m[2].metric("Kas di Bank", rp(saldo_per_akun(df)[KAS]), border=True)


def halaman_pembelian():
    halaman_transaksi("pembelian", "Pembelian", PILIHAN_PEMBELIAN)


def halaman_penjualan():
    halaman_transaksi("penjualan", "Penjualan", PILIHAN_PENJUALAN)


# =====================================================================
# 9. HALAMAN AKSES CEPAT (laporan)
# =====================================================================
def halaman_akses_cepat():
    judul_halaman("Akses Cepat", "Laporan akuntansi: jurnal umum, buku besar, laba rugi, dan neraca saldo.")
    df = muat_jurnal()

    # Filter tanggal untuk semua laporan
    awal, akhir = pilih_periode("aq", "Tahun ini")
    df_p = filter_periode(df, awal, akhir)        # data dalam periode
    df_sd = filter_periode(df, None, akhir)       # data sampai tanggal akhir

    # Siapkan semua laporan (juga dipakai untuk Excel)
    jurnal = laporan_jurnal_umum(df_p)
    laba_rugi = laporan_laba_rugi(df_p)
    neraca = laporan_neraca_saldo(df_sd)
    buku_besar_semua = pd.concat(
        [laporan_buku_besar(df_sd, k).assign(Akun=f"{k} - {AKUN[k]['nama']}") for k in AKUN],
        ignore_index=True)

    # Ringkasan singkat periode
    laba = total_tipe(saldo_per_akun(df_p), "Pendapatan") - total_tipe(saldo_per_akun(df_p), "Beban")
    m = st.columns(3)
    m[0].metric("Laba (Rugi) periode", rp(laba), border=True)
    m[1].metric("Jumlah transaksi", f"{df_p['bukti'].nunique()}", border=True)
    m[2].metric("Kas di Bank", rp(saldo_per_akun(df_sd)[KAS]), border=True)

    # Unduh semua laporan
    with st.container(border=True):
        teks, tombol = st.columns([3, 1], vertical_alignment="center")
        teks.markdown("**Unduh laporan lengkap**")
        teks.caption("Jurnal umum, buku besar, laba rugi, dan neraca saldo dalam satu file Excel.")
        with tombol:
            tombol_excel("⬇️ Unduh SEMUA laporan ke Excel",
                         {"Jurnal Umum": jurnal, "Buku Besar": buku_besar_semua,
                          "Laba Rugi": laba_rugi, "Neraca Saldo": neraca},
                         "laporan_akuntansi.xlsx", "xl_semua")

    tab1, tab2, tab3, tab4 = st.tabs(["Jurnal Umum", "Buku Besar", "Laba Rugi", "Neraca Saldo"])

    with tab1:
        st.dataframe(format_uang(jurnal, ["Debit", "Kredit"]), hide_index=True)
        tombol_excel("⬇️ Excel", {"Jurnal Umum": jurnal}, "jurnal_umum.xlsx", "xl_jurnal")

    with tab2:
        # format_func = cara menampilkan pilihan (nilai aslinya tetap kode akun)
        kode = st.selectbox("Pilih akun", list(AKUN), key="bb_akun",
                            format_func=lambda k: f"{k} - {AKUN[k]['nama']}")
        bb = laporan_buku_besar(df_sd, kode)
        st.caption(f"Saldo normal: {'Debit' if AKUN[kode]['normal'] == 'D' else 'Kredit'}")
        st.dataframe(format_uang(bb, ["Debit", "Kredit", "Saldo"]), hide_index=True)
        tombol_excel("⬇️ Excel", {AKUN[kode]["nama"]: bb}, f"buku_besar_{kode}.xlsx", "xl_bb")

    with tab3:
        st.dataframe(format_uang(laba_rugi, ["Jumlah"]), hide_index=True)
        tombol_excel("⬇️ Excel", {"Laba Rugi": laba_rugi}, "laba_rugi.xlsx", "xl_lr")

    with tab4:
        st.dataframe(format_uang(neraca, ["Debit", "Kredit"]), hide_index=True)
        total = neraca.iloc[-1]                  # baris terakhir = TOTAL
        if total["Debit"] == total["Kredit"]:
            st.success("Seimbang ✅ (total Debit = total Kredit)")
        else:
            st.error("Tidak seimbang ❌ — periksa jurnal")
        tombol_excel("⬇️ Excel", {"Neraca Saldo": neraca}, "neraca_saldo.xlsx", "xl_ns")


# =====================================================================
# 10. PROGRAM UTAMA
# =====================================================================
def main():
    # 1) Belum login -> tampilkan halaman login (tanpa sidebar)
    if "user" not in st.session_state:
        halaman_login()
        return

    # 2) Sudah login, tapi saldo awal belum diisi -> wizard
    if "sa_selesai" not in st.session_state:
        st.session_state.sa_selesai = saldo_awal_selesai()
    if not st.session_state.sa_selesai:
        sidebar_akun()
        halaman_saldo_awal()
        return

    # 3) Aplikasi utama: menu sidebar dikelompokkan per bagian (st.navigation)
    halaman = st.navigation({
        "Ringkasan": [
            st.Page(halaman_dashboard, title="Dashboard", icon=":material/dashboard:",
                    url_path="dashboard", default=True),
        ],
        "Transaksi": [
            st.Page(halaman_pembelian, title="Pembelian", icon=":material/shopping_cart:",
                    url_path="pembelian"),
            st.Page(halaman_penjualan, title="Penjualan", icon=":material/sell:",
                    url_path="penjualan"),
        ],
        "Laporan": [
            st.Page(halaman_akses_cepat, title="Akses Cepat", icon=":material/receipt_long:",
                    url_path="akses-cepat"),
        ],
    })
    sidebar_akun()          # kartu akun muncul di bawah menu
    halaman.run()


# Streamlit menjalankan file ini sebagai "__main__", jadi main() otomatis dipanggil
if __name__ == "__main__":
    main()