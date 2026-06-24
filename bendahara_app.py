# =============================================================================
# APLIKASI KEUANGAN TERPADU MI AL MA'ARIF TEMPELSARI
# Sistem Keuangan Bendahara Sekolah + Keuangan Guru Kelas
# Python Streamlit + Supabase Cloud
# =============================================================================
#
# CARA MENJALANKAN LOKAL:
#   pip install streamlit supabase pandas plotly openpyxl
#   streamlit run bendahara_app.py
#
# SKema SQL Supabase (jalankan di SQL Editor jika belum ada):
# -----------------------------------------------------------------------
# CREATE TABLE tabel_transaksi_bendahara (
#     id BIGSERIAL PRIMARY KEY,
#     nama_bendahara TEXT NOT NULL,
#     pos_dana TEXT NOT NULL,
#     jenis_transaksi TEXT NOT NULL,
#     nominal NUMERIC NOT NULL DEFAULT 0,
#     tanggal DATE NOT NULL,
#     keterangan TEXT,
#     foto_nota TEXT,
#     created_at TIMESTAMPTZ DEFAULT NOW()
# );
#
# CREATE TABLE tabel_master_siswa (
#     id BIGSERIAL PRIMARY KEY,
#     no INTEGER,
#     nama_siswa TEXT NOT NULL,
#     nis TEXT,
#     nama_ibu TEXT,
#     alamat TEXT,
#     no_whatsapp TEXT,
#     biaya_mobil NUMERIC DEFAULT 0,
#     biaya_lks NUMERIC DEFAULT 0,
#     kelas TEXT NOT NULL,
#     created_at TIMESTAMPTZ DEFAULT NOW()
# );
#
# CREATE TABLE tabel_pembayaran_siswa (
#     id BIGSERIAL PRIMARY KEY,
#     siswa_id BIGINT REFERENCES tabel_master_siswa(id) ON DELETE CASCADE,
#     pos_tagihan TEXT NOT NULL,
#     nominal NUMERIC NOT NULL DEFAULT 0,
#     tanggal DATE NOT NULL,
#     kelas TEXT,
#     keterangan TEXT,
#     periode_bulan INTEGER,
#     periode_tahun INTEGER,
#     created_at TIMESTAMPTZ DEFAULT NOW()
# );
#
# CREATE TABLE tabel_target_manual (
#     id BIGSERIAL PRIMARY KEY,
#     kelas TEXT NOT NULL,
#     pos_tagihan TEXT NOT NULL,
#     nominal NUMERIC NOT NULL DEFAULT 0,
#     periode_tahun INTEGER NOT NULL,
#     keterangan TEXT,
#     created_at TIMESTAMPTZ DEFAULT NOW()
# );
#
# CREATE TABLE tabel_tabungan_siswa (
#     id BIGSERIAL PRIMARY KEY,
#     siswa_id BIGINT REFERENCES tabel_master_siswa(id) ON DELETE CASCADE,
#     jenis TEXT NOT NULL,
#     nominal NUMERIC NOT NULL DEFAULT 0,
#     tanggal DATE NOT NULL,
#     kelas TEXT,
#     keterangan TEXT,
#     created_at TIMESTAMPTZ DEFAULT NOW()
# );
#
# -- Storage bucket foto nota bendahara
# INSERT INTO storage.buckets (id, name, public)
# VALUES ('nota_transaksi', 'nota_transaksi', true) ON CONFLICT (id) DO NOTHING;
#
# -- Policy akses data (sesuaikan kebutuhan keamanan)
# ALTER TABLE tabel_transaksi_bendahara ENABLE ROW LEVEL SECURITY;
# ALTER TABLE tabel_master_siswa ENABLE ROW LEVEL SECURITY;
# ALTER TABLE tabel_pembayaran_siswa ENABLE ROW LEVEL SECURITY;
# ALTER TABLE tabel_target_manual ENABLE ROW LEVEL SECURITY;
# ALTER TABLE tabel_tabungan_siswa ENABLE ROW LEVEL SECURITY;
# CREATE POLICY "Akses transaksi" ON tabel_transaksi_bendahara FOR ALL USING (true);
# CREATE POLICY "Akses siswa" ON tabel_master_siswa FOR ALL USING (true);
# CREATE POLICY "Akses bayar siswa" ON tabel_pembayaran_siswa FOR ALL USING (true);
# CREATE POLICY "Akses target manual" ON tabel_target_manual FOR ALL USING (true);
# CREATE POLICY "Akses tabungan" ON tabel_tabungan_siswa FOR ALL USING (true);
# =============================================================================

import io
import uuid
import urllib.parse
import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import date, datetime
from supabase import create_client, Client

# =============================================================================
# KONFIGURASI SUPABASE (JANGAN DIUBAH)
# =============================================================================

SUPABASE_URL = "https://kwrpfcjyavybexkaunsq.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imt3cnBmY2p5YXZ5YmV4a2F1bnNxIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODIyMDAyMzUsImV4cCI6MjA5Nzc3NjIzNX0.T0lojeTGTt9Yv5daouRY2mA74OQsec6E61ZvGJVW7r4"
BUCKET_NOTA = "nota_transaksi"
TIPE_FILE_NOTA = ["png", "jpg", "jpeg"]

# =============================================================================
# WHATSAPP — MODE LINK MANUAL (TANPA GATEWAY / TANPA TOKEN)
# =============================================================================
# Fitur WA Gateway otomatis (Fonnte, WAH A, Twilio, dll.) SENGAJA DINONAKTIFKAN
# karena memerlukan token/API key berbayar dan konfigurasi server tambahan.
#
# WA_GATEWAY_AKTIF = False
# WA_GATEWAY_TOKEN = ""    # tidak dipakai
# WA_GATEWAY_URL = ""      # tidak dipakai
#
# Pengiriman WhatsApp via st.link_button → api.whatsapp.com (buka WA di HP/browser).
# =============================================================================
WA_GATEWAY_AKTIF = False

# =============================================================================
# KONFIGURASI PERAN, PASSWORD, DAN MENU
# =============================================================================

DAFTAR_PERAN = [
    "Kepala Sekolah & Komite",
    "Tata Usaha (TU)",
    "Bendahara",
    "Guru Kelas",
]

PASSWORD_PERAN = {
    "Kepala Sekolah & Komite": "kepsek123",
    "Tata Usaha (TU)": "tu123",
    "Bendahara": "bendahara123",
    "Guru Kelas": "guru123",
}

MENU_DASHBOARD = "🏫 Dashboard Utama"
MENU_INPUT = "📝 Input Transaksi Bendahara"
MENU_KELOLA = "✏️ Kelola Transaksi"
MENU_LAPORAN = "📊 Laporan Bulanan"
MENU_SISWA = "👨‍🎓 Kelola Data Siswa"
MENU_BAYAR = "💳 Input Pembayaran Siswa"
MENU_TUNGGAKAN = "📋 Cek Tunggakan Siswa"
MENU_TABUNGAN = "🏦 Tabungan Siswa"
MENU_TARGET = "⚙️ Target Tagihan Manual"

DAFTAR_BENDAHARA = ["Bu Heni", "Bu Ika", "Bu Yuli", "Bu Syifa", "Bu Minah", "Bu Izza"]
DAFTAR_POS_DANA = [
    "Infak Tahunan", "Ngaji Pagi", "Infak Jumat", "Mobil", "Buku LKS",
    "BOS", "Sumbangan Lainnya", "Kantin", "Syukuran", "PIP", "DANANTARA",
]
JENIS_TRANSAKSI = ["Pemasukan", "Pengeluaran"]

DAFTAR_KELAS = [
    "IA", "IB", "IIA", "IIB", "IIIA", "IIIB",
    "IVA", "IVB", "VA", "VB", "VIA", "VIB",
]

POS_TAGIHAN_SISWA = [
    "Infak Tahunan", "Infak Jumat", "Ngaji Pagi",
    "Mobil", "Buku LKS", "Syukuran", "Seragam",
]

KONFIG_TAGIHAN = {
    "Infak Tahunan": {"tipe": "flat_tahun", "nominal": 200_000},
    "Infak Jumat": {"tipe": "sukarela", "nominal": 0},
    "Ngaji Pagi": {"tipe": "flat_bulan", "nominal": 30_000},
    "Mobil": {"tipe": "kolom", "kolom": "biaya_mobil"},
    "Buku LKS": {"tipe": "kolom", "kolom": "biaya_lks"},
    "Syukuran": {"tipe": "manual"},
    "Seragam": {"tipe": "manual"},
}

KOLOM_TEMPLATE_SISWA = [
    "NO", "Nama Siswa", "NIS", "Nama Ibu", "Alamat",
    "No Whatsapp", "Biaya Mobil", "Biaya LKS",
]

NAMA_BULAN = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}


# =============================================================================
# FUNGSI BANTUAN UMUM
# =============================================================================

@st.cache_resource
def koneksi_supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)


def format_rupiah(angka: float) -> str:
    try:
        return f"Rp {float(angka):,.0f}".replace(",", ".")
    except (TypeError, ValueError):
        return "Rp 0"


def bersihkan_whatsapp(nomor) -> str:
    """Membersihkan nomor WA agar berawalan 62."""
    if pd.isna(nomor) or nomor is None:
        return ""
    teks = str(nomor).strip().replace(" ", "").replace("-", "").replace("+", "")
    if teks.startswith("0"):
        teks = "62" + teks[1:]
    elif not teks.startswith("62"):
        teks = "62" + teks
    return teks


def link_whatsapp(nomor: str, pesan: str) -> str:
    """
    Membuat link api.whatsapp.com dengan draf pesan terisi.
    Tidak memerlukan token — hanya membuka aplikasi WhatsApp di perangkat user.
    """
    nomor_bersih = bersihkan_whatsapp(nomor)
    if not nomor_bersih:
        return ""
    return (
        "https://api.whatsapp.com/send?phone="
        + nomor_bersih
        + "&text="
        + urllib.parse.quote(pesan)
    )


# def kirim_whatsapp_gateway(nomor, pesan):
#     """
#     [NONAKTIF] Contoh integrasi WA Gateway otomatis — memerlukan token.
#     Dikomentari sengaja. Gunakan tampilkan_tombol_whatsapp() sebagai gantinya.
#     """
#     if not WA_GATEWAY_AKTIF or not WA_GATEWAY_TOKEN:
#         return False
#     # requests.post(WA_GATEWAY_URL, headers={"Authorization": WA_GATEWAY_TOKEN}, ...)
#     return False


def pesan_nota_bayar_jawa(nama_siswa: str, pos_tagihan: str) -> str:
    """Draf pesan nota bayar — Bahasa Jawa Krama Ingil halus."""
    return (
        f"Matur nuwun sanget, iuran {pos_tagihan} kagem putra/putri panjenengan "
        f"{nama_siswa} sampun dipun tampi dening Guru Kelas. "
        f"Mugi dados berkah, amin."
    )


def pesan_tagihan_bulanan_jawa(nama_siswa: str, pos_tagihan: str, sisa: float) -> str:
    """Draf pesan tagihan bulanan — Bahasa Jawa Krama Ingil halus."""
    return (
        f"Nyuwun pirsa dhumateng Bapak/Ibu wali murid saking {nama_siswa}, "
        f"ngemutaken bilih rincian iuran sekolah {pos_tagihan} taksih wonten "
        f"kekirangan semanten ({format_rupiah(sisa)}). "
        f"Maturnuwun sanget dhumateng kawigatosanipun."
    )


def tampilkan_tombol_whatsapp(
    label: str,
    nomor: str,
    pesan: str,
    key: str | None = None,
) -> bool:
    """
    Menampilkan st.link_button ke api.whatsapp.com.
    Mengembalikan True jika tombol ditampilkan, False jika nomor kosong.
    """
    url = link_whatsapp(nomor, pesan)
    if not url:
        st.caption("⚠️ Nomor WhatsApp siswa belum diisi di data master.")
        return False
    st.link_button(label, url=url, use_container_width=True, key=key)
    return True


def muat_ulang_aplikasi():
    st.cache_data.clear()
    st.rerun()


def init_session_state():
    defaults = {
        "authenticated": False,
        "peran_aktif": None,
        "remember_me": False,
        "kelas_guru": "IA",
        "halaman_aktif": MENU_DASHBOARD,
        "flash_pesan": None,
        "wa_draft_link": None,
        "wa_draft_label": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def indeks_aman(daftar: list, nilai) -> int:
    try:
        return daftar.index(nilai)
    except ValueError:
        return 0


def menu_berdasarkan_peran(peran: str) -> list[str]:
    if peran == "Kepala Sekolah & Komite":
        return [MENU_DASHBOARD]
    if peran == "Tata Usaha (TU)":
        return [MENU_DASHBOARD, MENU_LAPORAN]
    if peran == "Bendahara":
        return [MENU_INPUT, MENU_KELOLA, MENU_LAPORAN, MENU_SISWA, MENU_DASHBOARD]
    if peran == "Guru Kelas":
        return [
            MENU_BAYAR, MENU_TUNGGAKAN, MENU_TABUNGAN,
            MENU_TARGET, MENU_SISWA,
        ]
    return [MENU_DASHBOARD]


def cek_akses_halaman(peran: str, halaman: str) -> bool:
    return halaman in menu_berdasarkan_peran(peran)


def proses_login(peran: str, password: str, ingat: bool) -> bool:
    if password != PASSWORD_PERAN.get(peran, ""):
        return False
    st.session_state.authenticated = True
    st.session_state.peran_aktif = peran
    st.session_state.remember_me = ingat
    menu = menu_berdasarkan_peran(peran)
    st.session_state.halaman_aktif = menu[0]
    return True


def proses_logout():
    st.session_state.authenticated = False
    st.session_state.peran_aktif = None
    st.session_state.remember_me = False
    st.session_state.halaman_aktif = MENU_DASHBOARD
    st.session_state.wa_draft_link = None
    st.session_state.wa_draft_label = None
    muat_ulang_aplikasi()


def pastikan_bucket_storage():
    try:
        sb = koneksi_supabase()
        nama = [b.get("name", getattr(b, "name", "")) for b in sb.storage.list_buckets()]
        if BUCKET_NOTA not in nama:
            sb.storage.create_bucket(BUCKET_NOTA, options={"public": True})
    except Exception:
        pass


def unggah_foto_nota(file_bytes, nama_file, tipe_file) -> str | None:
    try:
        sb = koneksi_supabase()
        pastikan_bucket_storage()
        ext = nama_file.rsplit(".", 1)[-1].lower()
        path = f"nota/{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:8]}.{ext}"
        sb.storage.from_(BUCKET_NOTA).upload(
            path=path, file=file_bytes,
            file_options={"content-type": tipe_file or "image/jpeg", "upsert": "false"},
        )
        return sb.storage.from_(BUCKET_NOTA).get_public_url(path)
    except Exception as e:
        st.error(f"Gagal unggah foto: {e}")
        return None


# =============================================================================
# FUNGSI DATABASE — BENDAHARA
# =============================================================================

def ambil_transaksi_bendahara() -> pd.DataFrame:
    try:
        sb = koneksi_supabase()
        res = sb.table("tabel_transaksi_bendahara").select("*").order("tanggal", desc=True).execute()
        if not res.data:
            return pd.DataFrame()
        df = pd.DataFrame(res.data)
        df["nominal"] = pd.to_numeric(df["nominal"], errors="coerce").fillna(0)
        df["tanggal"] = pd.to_datetime(df["tanggal"], errors="coerce")
        if "foto_nota" not in df.columns:
            df["foto_nota"] = None
        return df
    except Exception as e:
        st.error(f"Gagal ambil transaksi bendahara: {e}")
        return pd.DataFrame()


def simpan_transaksi_bendahara(**data) -> bool:
    try:
        koneksi_supabase().table("tabel_transaksi_bendahara").insert(data).execute()
        return True
    except Exception as e:
        st.error(f"Gagal simpan transaksi: {e}")
        return False


def update_transaksi_bendahara(transaksi_id, **data) -> bool:
    try:
        koneksi_supabase().table("tabel_transaksi_bendahara").update(data).eq("id", transaksi_id).execute()
        return True
    except Exception as e:
        st.error(f"Gagal update transaksi: {e}")
        return False


def hapus_transaksi_bendahara(transaksi_id) -> bool:
    try:
        koneksi_supabase().table("tabel_transaksi_bendahara").delete().eq("id", transaksi_id).execute()
        return True
    except Exception as e:
        st.error(f"Gagal hapus transaksi: {e}")
        return False


def ringkasan_bendahara(df: pd.DataFrame) -> dict:
    if df.empty:
        return {"pemasukan": 0.0, "pengeluaran": 0.0, "saldo": 0.0}
    masuk = df.loc[df["jenis_transaksi"] == "Pemasukan", "nominal"].sum()
    keluar = df.loc[df["jenis_transaksi"] == "Pengeluaran", "nominal"].sum()
    return {"pemasukan": float(masuk), "pengeluaran": float(keluar), "saldo": float(masuk - keluar)}


# =============================================================================
# FUNGSI DATABASE — SISWA & TAGIHAN GURU
# =============================================================================

def ambil_siswa(kelas: str | None = None) -> pd.DataFrame:
    try:
        sb = koneksi_supabase()
        q = sb.table("tabel_master_siswa").select("*").order("no", desc=False)
        if kelas:
            q = q.eq("kelas", kelas)
        res = q.execute()
        if not res.data:
            return pd.DataFrame()
        df = pd.DataFrame(res.data)
        for kol in ["biaya_mobil", "biaya_lks"]:
            df[kol] = pd.to_numeric(df.get(kol, 0), errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"Gagal ambil data siswa: {e}")
        return pd.DataFrame()


def hapus_siswa_per_kelas(kelas: str) -> bool:
    try:
        koneksi_supabase().table("tabel_master_siswa").delete().eq("kelas", kelas).execute()
        return True
    except Exception as e:
        st.error(f"Gagal hapus data siswa kelas {kelas}: {e}")
        return False


def insert_siswa_bulk(records: list[dict]) -> bool:
    try:
        if records:
            koneksi_supabase().table("tabel_master_siswa").insert(records).execute()
        return True
    except Exception as e:
        st.error(f"Gagal simpan data siswa: {e}")
        return False


def buat_template_excel_siswa() -> bytes:
    df = pd.DataFrame(columns=KOLOM_TEMPLATE_SISWA)
    df.loc[0] = [1, "Contoh Nama Siswa", "12345", "Contoh Nama Ibu", "Alamat lengkap", "6281234567890", 50000, 75000]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Template Siswa", index=False)
    return buf.getvalue()


def proses_upload_excel_siswa(file_bytes, kelas: str) -> tuple[bool, str]:
    try:
        df = pd.read_excel(io.BytesIO(file_bytes))
        df.columns = [str(c).strip() for c in df.columns]
        wajib = {"Nama Siswa"}
        if not wajib.issubset(set(df.columns)):
            return False, "Kolom 'Nama Siswa' wajib ada di Excel."

        records = []
        for _, row in df.iterrows():
            nama = str(row.get("Nama Siswa", "")).strip()
            if not nama or nama.lower() == "contoh nama siswa":
                continue
            records.append({
                "no": int(row.get("NO", 0) or 0),
                "nama_siswa": nama,
                "nis": str(row.get("NIS", "") or ""),
                "nama_ibu": str(row.get("Nama Ibu", "") or ""),
                "alamat": str(row.get("Alamat", "") or ""),
                "no_whatsapp": bersihkan_whatsapp(row.get("No Whatsapp", "")),
                "biaya_mobil": float(row.get("Biaya Mobil", 0) or 0),
                "biaya_lks": float(row.get("Biaya LKS", 0) or 0),
                "kelas": kelas,
            })

        if not records:
            return False, "Tidak ada data siswa valid di file Excel."

        if not hapus_siswa_per_kelas(kelas):
            return False, "Gagal mengganti data lama kelas tersebut."
        if not insert_siswa_bulk(records):
            return False, "Gagal menyimpan data siswa baru."

        return True, f"Berhasil mengimpor {len(records)} siswa kelas {kelas}."
    except Exception as e:
        return False, f"Error baca Excel: {e}"


def ambil_pembayaran_siswa(kelas: str | None = None) -> pd.DataFrame:
    try:
        sb = koneksi_supabase()
        q = sb.table("tabel_pembayaran_siswa").select("*").order("tanggal", desc=True)
        if kelas:
            q = q.eq("kelas", kelas)
        res = q.execute()
        if not res.data:
            return pd.DataFrame()
        df = pd.DataFrame(res.data)
        df["nominal"] = pd.to_numeric(df["nominal"], errors="coerce").fillna(0)
        df["tanggal"] = pd.to_datetime(df["tanggal"], errors="coerce")
        return df
    except Exception as e:
        st.error(f"Gagal ambil pembayaran siswa: {e}")
        return pd.DataFrame()


def simpan_pembayaran_siswa(**data) -> bool:
    try:
        koneksi_supabase().table("tabel_pembayaran_siswa").insert(data).execute()
        return True
    except Exception as e:
        st.error(f"Gagal simpan pembayaran: {e}")
        return False


def ambil_target_manual(kelas: str, tahun: int) -> pd.DataFrame:
    try:
        res = (
            koneksi_supabase()
            .table("tabel_target_manual")
            .select("*")
            .eq("kelas", kelas)
            .eq("periode_tahun", tahun)
            .execute()
        )
        return pd.DataFrame(res.data or [])
    except Exception:
        return pd.DataFrame()


def simpan_target_manual(kelas, pos, nominal, tahun, ket="") -> bool:
    try:
        sb = koneksi_supabase()
        existing = (
            sb.table("tabel_target_manual")
            .select("id")
            .eq("kelas", kelas)
            .eq("pos_tagihan", pos)
            .eq("periode_tahun", tahun)
            .execute()
        )
        payload = {
            "kelas": kelas, "pos_tagihan": pos,
            "nominal": nominal, "periode_tahun": tahun, "keterangan": ket,
        }
        if existing.data:
            sb.table("tabel_target_manual").update(payload).eq("id", existing.data[0]["id"]).execute()
        else:
            sb.table("tabel_target_manual").insert(payload).execute()
        return True
    except Exception as e:
        st.error(f"Gagal simpan target manual: {e}")
        return False


def ambil_tabungan(kelas: str | None = None) -> pd.DataFrame:
    try:
        q = koneksi_supabase().table("tabel_tabungan_siswa").select("*").order("tanggal", desc=True)
        if kelas:
            q = q.eq("kelas", kelas)
        res = q.execute()
        if not res.data:
            return pd.DataFrame()
        df = pd.DataFrame(res.data)
        df["nominal"] = pd.to_numeric(df["nominal"], errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"Gagal ambil tabungan: {e}")
        return pd.DataFrame()


def simpan_tabungan(**data) -> bool:
    try:
        koneksi_supabase().table("tabel_tabungan_siswa").insert(data).execute()
        return True
    except Exception as e:
        st.error(f"Gagal simpan tabungan: {e}")
        return False


def hitung_target_siswa(siswa: pd.Series, pos: str, target_manual: pd.DataFrame, bulan: int, tahun: int) -> float:
    cfg = KONFIG_TAGIHAN.get(pos, {})
    tipe = cfg.get("tipe")
    if tipe == "flat_tahun":
        return float(cfg["nominal"])
    if tipe == "flat_bulan":
        return float(cfg["nominal"])
    if tipe == "sukarela":
        return 0.0
    if tipe == "kolom":
        return float(siswa.get(cfg["kolom"], 0) or 0)
    if tipe == "manual":
        if target_manual.empty:
            return 0.0
        baris = target_manual[target_manual["pos_tagihan"] == pos]
        if baris.empty:
            return 0.0
        return float(baris.iloc[0]["nominal"])
    return 0.0


def total_bayar_siswa_pos(df_bayar: pd.DataFrame, siswa_id: int, pos: str, bulan: int, tahun: int) -> float:
    if df_bayar.empty:
        return 0.0
    mask = (
        (df_bayar["siswa_id"] == siswa_id)
        & (df_bayar["pos_tagihan"] == pos)
        & (df_bayar["periode_bulan"] == bulan)
        & (df_bayar["periode_tahun"] == tahun)
    )
    return float(df_bayar.loc[mask, "nominal"].sum())


def status_pelunasan(persen: float, target: float, bayar: float) -> str:
    if target <= 0:
        return "Sukarela" if bayar > 0 else "Bebas"
    if persen >= 100:
        return "✅ Lunas"
    if bayar <= 0:
        return "❌ Belum Bayar"
    return f"🟡 Cicil ({persen:.0f}%)"


def buat_rekap_tunggakan(kelas: str, bulan: int, tahun: int) -> pd.DataFrame:
    df_siswa = ambil_siswa(kelas)
    if df_siswa.empty:
        return pd.DataFrame()

    df_bayar = ambil_pembayaran_siswa(kelas)
    target_manual = ambil_target_manual(kelas, tahun)
    baris = []

    for _, siswa in df_siswa.iterrows():
        for pos in POS_TAGIHAN_SISWA:
            target = hitung_target_siswa(siswa, pos, target_manual, bulan, tahun)
            bayar = total_bayar_siswa_pos(df_bayar, siswa["id"], pos, bulan, tahun)
            if target <= 0 and pos == "Infak Jumat":
                persen = 100.0 if bayar > 0 else 0.0
            elif target <= 0:
                continue
            else:
                persen = min(100.0, (bayar / target) * 100) if target else 0.0
            sisa = max(0.0, target - bayar)
            baris.append({
                "ID Siswa": siswa["id"],
                "Nama Siswa": siswa["nama_siswa"],
                "Pos Tagihan": pos,
                "Target": target,
                "Terbayar": bayar,
                "Sisa": sisa,
                "Persentase": round(persen, 1),
                "Status": status_pelunasan(persen, target, bayar),
                "No Whatsapp": siswa.get("no_whatsapp", ""),
            })
    return pd.DataFrame(baris)


def metrik_pelunasan_sekolah(bulan: int, tahun: int) -> dict:
    total_target = 0.0
    total_bayar = 0.0
    per_kelas = []

    for kelas in DAFTAR_KELAS:
        rekap = buat_rekap_tunggakan(kelas, bulan, tahun)
        if rekap.empty:
            continue
        t_target = rekap["Target"].sum()
        t_bayar = rekap["Terbayar"].sum()
        total_target += t_target
        total_bayar += t_bayar
        persen = (t_bayar / t_target * 100) if t_target else 0
        per_kelas.append({"Kelas": kelas, "Target": t_target, "Terbayar": t_bayar, "Persentase": round(persen, 1)})

    persen_sekolah = (total_bayar / total_target * 100) if total_target else 0
    return {
        "total_target": total_target,
        "total_bayar": total_bayar,
        "persen_sekolah": round(persen_sekolah, 1),
        "per_kelas": pd.DataFrame(per_kelas),
    }


# =============================================================================
# LAPORAN BULANAN BENDAHARA — PER POS DANA
# =============================================================================

def filter_bulan_tahun(df, bulan, tahun):
    if df.empty:
        return df.copy()
    m = (df["tanggal"].dt.month == bulan) & (df["tanggal"].dt.year == tahun)
    return df.loc[m].copy()


def laporan_per_pos_dana(df_bulan: pd.DataFrame) -> dict:
    hasil = {}
    if df_bulan.empty:
        return hasil
    total_masuk = df_bulan.loc[df_bulan["jenis_transaksi"] == "Pemasukan", "nominal"].sum()
    for pos in DAFTAR_POS_DANA:
        df_pos = df_bulan[df_bulan["pos_dana"] == pos].copy()
        if df_pos.empty:
            continue
        masuk = df_pos.loc[df_pos["jenis_transaksi"] == "Pemasukan", "nominal"].sum()
        keluar = df_pos.loc[df_pos["jenis_transaksi"] == "Pengeluaran", "nominal"].sum()
        persen = (masuk / total_masuk * 100) if total_masuk else 0
        hasil[pos] = {
            "detail": df_pos.sort_values("tanggal"),
            "pemasukan": float(masuk),
            "pengeluaran": float(keluar),
            "saldo": float(masuk - keluar),
            "persen": round(persen, 1),
        }
    return hasil


def buat_excel_laporan_pos(bulan, tahun, ringkasan, laporan_pos: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame([
            ["Laporan Keuangan Bendahara MI Al Ma'arif Tempelsari"],
            ["Periode", f"{NAMA_BULAN[bulan]} {tahun}"],
            ["Total Pemasukan", ringkasan["pemasukan"]],
            ["Total Pengeluaran", ringkasan["pengeluaran"]],
            ["Saldo", ringkasan["saldo"]],
        ]).to_excel(w, sheet_name="Ringkasan", index=False, header=False)

        for pos, data in laporan_pos.items():
            sheet = pos[:31]
            df_out = data["detail"].copy()
            df_out["tanggal"] = df_out["tanggal"].dt.strftime("%d/%m/%Y")
            df_out = df_out[[
                "tanggal", "nama_bendahara", "jenis_transaksi",
                "nominal", "keterangan",
            ]]
            df_out.columns = ["Tanggal", "Bendahara", "Jenis", "Nominal", "Keterangan"]
            df_out.to_excel(w, sheet_name=sheet, index=False)
            start = len(df_out) + 2
            pd.DataFrame([
                ["Subtotal Pemasukan", data["pemasukan"]],
                ["Subtotal Pengeluaran", data["pengeluaran"]],
                ["Saldo Pos", data["saldo"]],
                ["Persentase dr Total Pemasukan", f"{data['persen']}%"],
            ]).to_excel(w, sheet_name=sheet, index=False, header=False, startrow=start)
    return buf.getvalue()


# =============================================================================
# INISIALISASI APLIKASI
# =============================================================================

init_session_state()
pastikan_bucket_storage()

st.set_page_config(
    page_title="Keuangan Terpadu MI Al Ma'arif Tempelsari",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.judul-aplikasi { text-align:center; color:#1a5276; font-size:1.9rem; font-weight:bold; }
.sub-judul { text-align:center; color:#566573; margin-bottom:1rem; }
.metric-box { background:linear-gradient(135deg,#1a5276,#2980b9); color:white;
              padding:1rem; border-radius:10px; text-align:center; }
.login-box { background:#f8f9fa; padding:1rem; border-radius:8px; border:1px solid #ddd; }
.tombol-hapus button { background:#e74c3c !important; color:white !important; }
</style>
""", unsafe_allow_html=True)


# =============================================================================
# SIDEBAR — LOGIN, PERAN, NAVIGASI
# =============================================================================

with st.sidebar:
    st.image("https://img.icons8.com/color/96/school-building.png", width=72)
    st.title("🏫 MI Tempelsari")
    st.caption("Aplikasi Keuangan Terpadu")
    st.markdown("---")

    if not st.session_state.authenticated:
        st.markdown("##### 🔐 Login")
        with st.container():
            peran_login = st.selectbox("Peran", DAFTAR_PERAN, key="login_peran")
            password = st.text_input("Password", type="password", key="login_password")
            ingat_saya = st.checkbox("Ingat Saya di Perangkat Ini", key="login_ingat")

            if st.button("🔓 Masuk", use_container_width=True, type="primary"):
                if proses_login(peran_login, password, ingat_saya):
                    st.session_state.flash_pesan = f"Selamat datang, {peran_login}!"
                    muat_ulang_aplikasi()
                else:
                    st.error("Password salah. Silakan coba lagi.")

        st.info("Password default:\n- kepsek123\n- tu123\n- bendahara123\n- guru123")
        st.stop()

    # Sudah login
    peran = st.session_state.peran_aktif
    st.success(f"Login: **{peran}**")
    if st.session_state.remember_me:
        st.caption("✅ Ingat saya aktif")

    if peran == "Guru Kelas":
        st.session_state.kelas_guru = st.selectbox(
            "Kelas Anda", DAFTAR_KELAS,
            index=indeks_aman(DAFTAR_KELAS, st.session_state.kelas_guru),
            key="select_kelas_guru",
        )

    if st.button("🚪 Logout", use_container_width=True):
        proses_logout()

    st.markdown("---")
    menu_tersedia = menu_berdasarkan_peran(peran)
    if st.session_state.halaman_aktif not in menu_tersedia:
        st.session_state.halaman_aktif = menu_tersedia[0]

    halaman = st.radio(
        "Menu", menu_tersedia,
        index=menu_tersedia.index(st.session_state.halaman_aktif),
        label_visibility="collapsed",
    )
    st.session_state.halaman_aktif = halaman

    if st.button("🔄 Muat Ulang Data", use_container_width=True):
        muat_ulang_aplikasi()


# =============================================================================
# HEADER
# =============================================================================

st.markdown('<p class="judul-aplikasi">🏫 Aplikasi Keuangan Terpadu</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-judul">MI Al Ma\'arif Tempelsari — Bendahara & Guru Kelas</p>', unsafe_allow_html=True)

if st.session_state.flash_pesan:
    st.success(st.session_state.flash_pesan)
    st.session_state.flash_pesan = None

peran = st.session_state.peran_aktif
if not cek_akses_halaman(peran, halaman):
    st.error("⛔ Anda tidak memiliki akses ke halaman ini.")
    st.stop()


# =============================================================================
# HALAMAN: DASHBOARD UTAMA
# =============================================================================

if halaman == MENU_DASHBOARD:
    st.header("🏫 Dashboard Utama")
    bulan_ini = date.today().month
    tahun_ini = date.today().year

    df_trx = ambil_transaksi_bendahara()
    ringkas_b = ringkasan_bendahara(df_trx)
    metrik_siswa = metrik_pelunasan_sekolah(bulan_ini, tahun_ini)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Saldo Bendahara", format_rupiah(ringkas_b["saldo"]))
    k2.metric("Pemasukan Total", format_rupiah(ringkas_b["pemasukan"]))
    k3.metric("Pengeluaran Total", format_rupiah(ringkas_b["pengeluaran"]))
    k4.metric(
        "Pelunasan Iuran Siswa",
        f"{metrik_siswa['persen_sekolah']}%",
        help="Persentase total dana masuk vs target tagihan seluruh kelas",
    )

    st.markdown("---")
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Keuangan Bendahara per Pos Dana")
        if df_trx.empty:
            st.info("Belum ada transaksi bendahara.")
        else:
            df_m = df_trx[df_trx["jenis_transaksi"] == "Pemasukan"].groupby("pos_dana")["nominal"].sum().reset_index()
            if not df_m.empty:
                fig = px.bar(df_m, x="pos_dana", y="nominal", title="Pemasukan per Pos Dana", color="nominal",
                             color_continuous_scale="Blues")
                fig.update_layout(xaxis_tickangle=-45, showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.subheader("Capaian Pelunasan per Kelas")
        df_kelas = metrik_siswa["per_kelas"]
        if df_kelas.empty:
            st.info("Belum ada data tagihan siswa.")
        else:
            fig2 = px.bar(
                df_kelas, x="Kelas", y="Persentase", text="Persentase",
                title=f"Pelunasan Iuran — {NAMA_BULAN[bulan_ini]} {tahun_ini}",
                color="Persentase", color_continuous_scale="Greens",
            )
            fig2.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            fig2.update_layout(yaxis_range=[0, 110])
            st.plotly_chart(fig2, use_container_width=True)

    if not df_kelas.empty:
        st.subheader("Detail Capaian per Kelas")
        tampil = df_kelas.copy()
        tampil["Target"] = tampil["Target"].apply(format_rupiah)
        tampil["Terbayar"] = tampil["Terbayar"].apply(format_rupiah)
        tampil["Persentase"] = tampil["Persentase"].astype(str) + "%"
        st.dataframe(tampil, use_container_width=True, hide_index=True)


# =============================================================================
# HALAMAN: INPUT TRANSAKSI BENDAHARA
# =============================================================================

elif halaman == MENU_INPUT:
    st.header("📝 Input Transaksi Bendahara")

    with st.form("form_input_bendahara", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            nama_b = st.selectbox("Nama Bendahara *", DAFTAR_BENDAHARA)
            jenis = st.selectbox("Jenis Transaksi *", JENIS_TRANSAKSI)
            tanggal = st.date_input("Tanggal *", value=date.today())
        with c2:
            pos = st.selectbox("Pos Dana *", DAFTAR_POS_DANA)
            nominal = st.number_input("Nominal (Rp) *", min_value=0, step=1000, format="%d")
            ket = st.text_input("Keterangan")

        foto = st.file_uploader("Unggah Bukti Nota / Kuitansi (Opsional)", type=TIPE_FILE_NOTA)
        if foto:
            st.image(foto, width=260)

        if st.form_submit_button("💾 Simpan Transaksi", type="primary", use_container_width=True):
            if nominal <= 0:
                st.warning("Nominal harus lebih dari 0.")
            else:
                url_foto = None
                if foto:
                    url_foto = unggah_foto_nota(foto.getvalue(), foto.name, foto.type)
                    if url_foto is None:
                        st.stop()
                ok = simpan_transaksi_bendahara(
                    nama_bendahara=nama_b, pos_dana=pos, jenis_transaksi=jenis,
                    nominal=float(nominal), tanggal=tanggal.isoformat(),
                    keterangan=ket, foto_nota=url_foto,
                )
                if ok:
                    st.success("Transaksi berhasil disimpan!")
                    st.balloons()


# =============================================================================
# HALAMAN: KELOLA TRANSAKSI BENDAHARA
# =============================================================================

elif halaman == MENU_KELOLA:
    st.header("✏️ Kelola Transaksi Bendahara")
    df = ambil_transaksi_bendahara()

    if df.empty:
        st.info("Belum ada transaksi.")
    else:
        df_t = df.copy()
        df_t["tanggal"] = df_t["tanggal"].dt.strftime("%d/%m/%Y")
        df_t["nominal"] = df_t["nominal"].apply(format_rupiah)
        st.dataframe(df_t[["id", "tanggal", "nama_bendahara", "pos_dana", "jenis_transaksi", "nominal", "keterangan"]],
                     use_container_width=True, hide_index=True)

        ids = df["id"].tolist()
        id_pilih = st.selectbox("Pilih transaksi", ids, format_func=lambda i: f"ID {i} | {df.loc[df.id==i,'tanggal'].iloc[0].strftime('%d/%m/%Y')} | {format_rupiah(df.loc[df.id==i,'nominal'].iloc[0])}")
        baris = df[df["id"] == id_pilih].iloc[0]

        with st.form("form_edit"):
            c1, c2 = st.columns(2)
            with c1:
                e_b = st.selectbox("Bendahara", DAFTAR_BENDAHARA, index=indeks_aman(DAFTAR_BENDAHARA, baris["nama_bendahara"]))
                e_j = st.selectbox("Jenis", JENIS_TRANSAKSI, index=indeks_aman(JENIS_TRANSAKSI, baris["jenis_transaksi"]))
                e_t = st.date_input("Tanggal", value=baris["tanggal"].date())
            with c2:
                e_p = st.selectbox("Pos Dana", DAFTAR_POS_DANA, index=indeks_aman(DAFTAR_POS_DANA, baris["pos_dana"]))
                e_n = st.number_input("Nominal", min_value=0, step=1000, value=int(baris["nominal"]))
                e_k = st.text_input("Keterangan", value=str(baris.get("keterangan") or ""))

            foto_baru = st.file_uploader("Ganti foto nota (opsional)", type=TIPE_FILE_NOTA)
            if st.form_submit_button("💾 Simpan Perubahan", type="primary"):
                foto_final = baris.get("foto_nota")
                if foto_baru:
                    foto_final = unggah_foto_nota(foto_baru.getvalue(), foto_baru.name, foto_baru.type)
                if update_transaksi_bendahara(id_pilih,
                    nama_bendahara=e_b, pos_dana=e_p, jenis_transaksi=e_j,
                    nominal=float(e_n), tanggal=e_t.isoformat(), keterangan=e_k, foto_nota=foto_final):
                    st.session_state.flash_pesan = "Transaksi diperbarui."
                    muat_ulang_aplikasi()

        st.markdown("---")
        yakin = st.checkbox("Yakin hapus transaksi ini")
        st.markdown('<div class="tombol-hapus">', unsafe_allow_html=True)
        if st.button("🗑️ Hapus Transaksi", disabled=not yakin):
            if hapus_transaksi_bendahara(id_pilih):
                st.session_state.flash_pesan = "Transaksi dihapus."
                muat_ulang_aplikasi()
        st.markdown("</div>", unsafe_allow_html=True)


# =============================================================================
# HALAMAN: LAPORAN BULANAN (PER POS DANA)
# =============================================================================

elif halaman == MENU_LAPORAN:
    st.header("📊 Laporan Bulanan Bendahara")
    c1, c2 = st.columns(2)
    with c1:
        bln = st.selectbox("Bulan", list(NAMA_BULAN.keys()), format_func=lambda x: NAMA_BULAN[x], index=date.today().month - 1)
    with c2:
        thn = st.selectbox("Tahun", list(range(2024, date.today().year + 2)), index=list(range(2024, date.today().year + 2)).index(date.today().year))

    df_all = ambil_transaksi_bendahara()
    df_bln = filter_bulan_tahun(df_all, bln, thn)
    ringkas = ringkasan_bendahara(df_bln)
    lap_pos = laporan_per_pos_dana(df_bln)

    st.markdown("##### 📥 Cetak / Unduh Laporan")
    if df_bln.empty:
        st.warning("Tidak ada data pada periode ini.")
    else:
        excel_bytes = buat_excel_laporan_pos(bln, thn, ringkas, lap_pos)
        st.download_button(
            "📊 Unduh Excel per Pos Dana",
            data=excel_bytes,
            file_name=f"Laporan_{NAMA_BULAN[bln]}_{thn}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.markdown("---")
    m1, m2, m3 = st.columns(3)
    m1.metric("Pemasukan", format_rupiah(ringkas["pemasukan"]))
    m2.metric("Pengeluaran", format_rupiah(ringkas["pengeluaran"]))
    m3.metric("Saldo", format_rupiah(ringkas["saldo"]))

    if not lap_pos:
        st.info("Tidak ada transaksi per pos dana.")
    else:
        for pos, data in lap_pos.items():
            with st.expander(f"📁 {pos} — Masuk: {format_rupiah(data['pemasukan'])} | Keluar: {format_rupiah(data['pengeluaran'])} | {data['persen']}% dr total pemasukan"):
                d = data["detail"].copy()
                d["tanggal"] = d["tanggal"].dt.strftime("%d/%m/%Y")
                d["nominal"] = d["nominal"].apply(format_rupiah)
                st.dataframe(
                    d[["tanggal", "nama_bendahara", "jenis_transaksi", "nominal", "keterangan"]].rename(columns={
                        "tanggal": "Tanggal", "nama_bendahara": "Bendahara",
                        "jenis_transaksi": "Jenis", "nominal": "Nominal", "keterangan": "Keterangan",
                    }),
                    use_container_width=True, hide_index=True,
                )
                st.caption(f"Subtotal pemasukan pos ini: **{format_rupiah(data['pemasukan'])}** | Persentase: **{data['persen']}%**")


# =============================================================================
# HALAMAN: KELOLA DATA SISWA
# =============================================================================

elif halaman == MENU_SISWA:
    st.header("👨‍🎓 Kelola Data Siswa")
    kelas_upload = st.session_state.kelas_guru if peran == "Guru Kelas" else st.selectbox("Pilih Kelas", DAFTAR_KELAS, key="kelas_upload_siswa")

    st.markdown("##### 📥 Template Excel")
    st.download_button(
        "📥 Unduh Template Excel",
        data=buat_template_excel_siswa(),
        file_name="Template_Data_Siswa_MI_Tempelsari.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("Kolom: NO | Nama Siswa | NIS | Nama Ibu | Alamat | No Whatsapp | Biaya Mobil | Biaya LKS")

    file_up = st.file_uploader("Unggah Excel siswa yang sudah diisi", type=["xlsx", "xls"])
    if file_up and st.button("⬆️ Impor ke Database", type="primary"):
        ok, msg = proses_upload_excel_siswa(file_up.getvalue(), kelas_upload)
        if ok:
            st.success(msg)
        else:
            st.error(msg)

    st.markdown("---")
    st.subheader(f"Data Siswa Kelas {kelas_upload}")
    df_s = ambil_siswa(kelas_upload)
    if df_s.empty:
        st.info("Belum ada data siswa untuk kelas ini.")
    else:
        tampil = df_s.copy()
        st.dataframe(tampil, use_container_width=True, hide_index=True)
        st.caption(f"Total: {len(tampil)} siswa")


# =============================================================================
# HALAMAN: INPUT PEMBAYARAN SISWA (GURU KELAS)
# =============================================================================

elif halaman == MENU_BAYAR:
    st.header("💳 Input Pembayaran / Cicilan Siswa")
    st.caption(
        "Setelah simpan pembayaran, gunakan tombol WhatsApp untuk mengirim nota "
        "ke orang tua (pesan Jawa Krama terisi otomatis — tanpa token gateway)."
    )

    # Tampilkan tombol WA setelah pembayaran berhasil disimpan (link api.whatsapp.com)
    if st.session_state.get("wa_draft_link"):
        st.success("✅ Pembayaran terakhir berhasil disimpan!")
        st.link_button(
            st.session_state.get("wa_draft_label", "📱 Kirim Nota via WhatsApp"),
            url=st.session_state["wa_draft_link"],
            use_container_width=True,
            key="wa_link_setelah_bayar",
        )
        st.caption("Klik tombol di atas → WhatsApp terbuka → pesan Jawa Krama sudah terisi → tekan Kirim.")
        if st.button("✖ Tutup notifikasi WhatsApp", key="tutup_wa_draft"):
            st.session_state.wa_draft_link = None
            st.session_state.wa_draft_label = None
            st.rerun()

    kelas = st.session_state.kelas_guru
    bulan = st.selectbox("Periode Bulan", list(NAMA_BULAN.keys()), format_func=lambda x: NAMA_BULAN[x], index=date.today().month - 1)
    tahun = st.selectbox("Periode Tahun", list(range(2024, date.today().year + 2)), index=list(range(2024, date.today().year + 2)).index(date.today().year))

    df_siswa = ambil_siswa(kelas)
    if df_siswa.empty:
        st.warning(f"Belum ada siswa kelas {kelas}. Impor data di menu Kelola Data Siswa.")
    else:
        map_siswa = {f"{r['nama_siswa']} (NIS: {r.get('nis','-')})": r for _, r in df_siswa.iterrows()}
        pilih_siswa = st.selectbox("Nama Siswa", list(map_siswa.keys()))
        siswa = map_siswa[pilih_siswa]
        pos = st.selectbox("Pos Tagihan", POS_TAGIHAN_SISWA)

        target_manual = ambil_target_manual(kelas, tahun)
        target = hitung_target_siswa(siswa, pos, target_manual, bulan, tahun)
        df_bayar = ambil_pembayaran_siswa(kelas)
        sudah_bayar = total_bayar_siswa_pos(df_bayar, siswa["id"], pos, bulan, tahun)
        sisa = max(0, target - sudah_bayar) if target > 0 else 0

        st.info(f"Target: **{format_rupiah(target)}** | Sudah bayar: **{format_rupiah(sudah_bayar)}** | Sisa: **{format_rupiah(sisa)}**")

        with st.form("form_bayar_siswa"):
            nominal = st.number_input("Nominal Pembayaran (Rp)", min_value=0, step=1000, format="%d")
            tgl = st.date_input("Tanggal Bayar", value=date.today())
            ket = st.text_input("Keterangan", placeholder="Cicilan ke-1, lunas, dll.")
            if st.form_submit_button("💾 Simpan Pembayaran", type="primary"):
                if nominal <= 0:
                    st.warning("Nominal harus lebih dari 0.")
                else:
                    berhasil_bayar = simpan_pembayaran_siswa(
                        siswa_id=int(siswa["id"]), pos_tagihan=pos, nominal=float(nominal),
                        tanggal=tgl.isoformat(), kelas=kelas, keterangan=ket,
                        periode_bulan=bulan, periode_tahun=tahun,
                    )
                    if berhasil_bayar:
                        wa = bersihkan_whatsapp(siswa.get("no_whatsapp", ""))
                        if wa:
                            pesan = pesan_nota_bayar_jawa(siswa["nama_siswa"], pos)
                            st.session_state.wa_draft_link = link_whatsapp(wa, pesan)
                            st.session_state.wa_draft_label = (
                                f"📱 Kirim Nota — {siswa['nama_siswa']} ({pos})"
                            )
                        else:
                            st.session_state.wa_draft_link = None
                            st.session_state.wa_draft_label = None
                        st.session_state.flash_pesan = "Pembayaran tersimpan!"
                        muat_ulang_aplikasi()


# =============================================================================
# HALAMAN: CEK TUNGGAKAN SISWA
# =============================================================================

elif halaman == MENU_TUNGGAKAN:
    st.header("📋 Cek Tunggakan & Pelunasan Siswa")
    kelas = st.session_state.kelas_guru
    bulan = st.selectbox("Bulan", list(NAMA_BULAN.keys()), format_func=lambda x: NAMA_BULAN[x], index=date.today().month - 1, key="tgg_bln")
    tahun = st.selectbox("Tahun", list(range(2024, date.today().year + 2)), index=list(range(2024, date.today().year + 2)).index(date.today().year), key="tgg_thn")

    rekap = buat_rekap_tunggakan(kelas, bulan, tahun)
    if rekap.empty:
        st.info("Belum ada data.")
    else:
        filter_status = st.multiselect(
            "Filter Status",
            ["✅ Lunas", "🟡 Cicil", "❌ Belum Bayar"],
            default=["🟡 Cicil", "❌ Belum Bayar"],
        )
        df_t = rekap.copy()
        if filter_status:
            mask = pd.Series(False, index=df_t.index)
            if "✅ Lunas" in filter_status:
                mask |= df_t["Status"].str.startswith("✅")
            if "🟡 Cicil" in filter_status:
                mask |= df_t["Status"].str.contains("Cicil", na=False)
            if "❌ Belum Bayar" in filter_status:
                mask |= df_t["Status"].str.startswith("❌")
            df_t = df_t[mask]

        # Visualisasi persentase per siswa (rata-rata semua pos)
        avg_siswa = rekap.groupby("Nama Siswa")["Persentase"].mean().reset_index()
        fig = px.bar(avg_siswa, x="Nama Siswa", y="Persentase", color="Persentase",
                     title=f"Rata-rata Pelunasan Siswa Kelas {kelas}", color_continuous_scale="RdYlGn")
        fig.update_layout(xaxis_tickangle=-45, yaxis_range=[0, 105])
        st.plotly_chart(fig, use_container_width=True)

        tampil = df_t.copy()
        tampil["Target"] = tampil["Target"].apply(format_rupiah)
        tampil["Terbayar"] = tampil["Terbayar"].apply(format_rupiah)
        tampil["Sisa"] = tampil["Sisa"].apply(format_rupiah)
        tampil["Persentase"] = tampil["Persentase"].astype(str) + "%"
        st.dataframe(tampil.drop(columns=["ID Siswa", "No Whatsapp"]), use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("📱 Tagihan Bulanan via WhatsApp (Belum Lunas)")
        st.caption(
            "Tombol di bawah membuka api.whatsapp.com dengan draf pesan Jawa Krama Ingil. "
            "Tidak memerlukan token gateway — guru tinggal klik dan kirim."
        )
        belum_lunas = rekap[(rekap["Sisa"] > 0) & (rekap["Target"] > 0)]
        if belum_lunas.empty:
            st.success("Semua tagihan wajib sudah lunas!")
        else:
            for _, r in belum_lunas.iterrows():
                pesan = pesan_tagihan_bulanan_jawa(r["Nama Siswa"], r["Pos Tagihan"], r["Sisa"])
                tampilkan_tombol_whatsapp(
                    f"📱 WA Tagihan — {r['Nama Siswa']} · {r['Pos Tagihan']} · sisa {format_rupiah(r['Sisa'])}",
                    r["No Whatsapp"],
                    pesan,
                    key=f"wa_{r['ID Siswa']}_{r['Pos Tagihan']}",
                )


# =============================================================================
# HALAMAN: TABUNGAN SISWA
# =============================================================================

elif halaman == MENU_TABUNGAN:
    st.header("🏦 Tabungan Siswa")
    st.caption("Catatan tabungan pribadi siswa — terpisah dari uang sekolah.")
    kelas = st.session_state.kelas_guru
    df_siswa = ambil_siswa(kelas)

    if df_siswa.empty:
        st.warning("Belum ada siswa di kelas ini.")
    else:
        map_s = {r["nama_siswa"]: r for _, r in df_siswa.iterrows()}
        nama = st.selectbox("Nama Siswa", list(map_s.keys()))
        siswa = map_s[nama]

        with st.form("form_tabungan"):
            jenis = st.selectbox("Jenis", ["Setor", "Tarik"])
            nominal = st.number_input("Nominal", min_value=0, step=1000)
            tgl = st.date_input("Tanggal")
            ket = st.text_input("Keterangan")
            if st.form_submit_button("💾 Simpan", type="primary"):
                nilai = float(nominal) if jenis == "Setor" else -float(nominal)
                if simpan_tabungan(
                    siswa_id=int(siswa["id"]), jenis=jenis, nominal=nilai,
                    tanggal=tgl.isoformat(), kelas=kelas, keterangan=ket,
                ):
                    st.success("Tabungan tersimpan.")
                    muat_ulang_aplikasi()

        df_tab = ambil_tabungan(kelas)
        if not df_tab.empty:
            saldo = df_tab[df_tab["siswa_id"] == siswa["id"]]["nominal"].sum()
            st.metric(f"Saldo Tabungan {nama}", format_rupiah(saldo))
            st.dataframe(df_tab[df_tab["siswa_id"] == siswa["id"]], use_container_width=True, hide_index=True)


# =============================================================================
# HALAMAN: TARGET TAGIHAN MANUAL (SYUKURAN & SERAGAM)
# =============================================================================

elif halaman == MENU_TARGET:
    st.header("⚙️ Target Tagihan Manual")
    st.info("Atur nominal target **Syukuran** dan **Seragam** per kelas per tahun.")
    kelas = st.session_state.kelas_guru
    tahun = st.selectbox("Tahun Ajaran", list(range(2024, date.today().year + 2)), index=list(range(2024, date.today().year + 2)).index(date.today().year))

    for pos in ["Syukuran", "Seragam"]:
        with st.form(f"target_{pos}"):
            st.subheader(pos)
            target_df = ambil_target_manual(kelas, tahun)
            lama = 0.0
            if not target_df.empty:
                baris = target_df[target_df["pos_tagihan"] == pos]
                if not baris.empty:
                    lama = float(baris.iloc[0]["nominal"])
            nominal = st.number_input(f"Nominal Target {pos} per Siswa (Rp)", min_value=0, step=1000, value=int(lama))
            ket = st.text_input("Keterangan", key=f"ket_{pos}")
            if st.form_submit_button(f"Simpan Target {pos}"):
                if simpan_target_manual(kelas, pos, float(nominal), tahun, ket):
                    st.success(f"Target {pos} disimpan.")
                    muat_ulang_aplikasi()


# =============================================================================
# FOOTER + PANDUAN DEPLOY ONLINE
# =============================================================================

st.markdown("---")
st.caption("© 2026 MI Al Ma'arif Tempelsari — Aplikasi Keuangan Terpadu | Streamlit + Supabase Cloud")

# =============================================================================
# PANDUAN DEPLOY GRATIS KE STREAMLIT COMMUNITY CLOUD (share.streamlit.io)
# =============================================================================
# Agar aplikasi bisa diakses dari HP kapan saja & di mana saja:
#
# 1. Buat akun GitHub gratis di https://github.com
# 2. Buat repository baru, upload file:
#    - bendahara_app.py
#    - requirements.txt  (isi: streamlit, supabase, pandas, plotly, openpyxl)
# 3. Buka https://share.streamlit.io dan login dengan GitHub
# 4. Klik "New app" → pilih repository → Main file: bendahara_app.py
# 5. Klik Deploy — tunggu beberapa menit
# 6. Aplikasi live di URL: https://nama-app.streamlit.app
# 7. Buka URL tersebut dari HP/laptop guru, bendahara, kepsek — data otomatis
#    sinkron karena tersimpan di Supabase cloud
#
# Catatan keamanan: untuk produksi, pindahkan SUPABASE_KEY ke Streamlit Secrets
# (menu Settings → Secrets) dan ganti password default peran dengan yang kuat.
# =============================================================================
