import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Otomatisasi HPP & Laporan Keuangan",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

# Mengambil password dari Streamlit Secrets (atau fallback default 'admin123')
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "admin123")
def apply_login_styles():
    st.markdown("""
        <style>
        header, footer, [data-testid="stToolbar"] {visibility: hidden;}
        .block-container {
            padding: 0rem !important;
            max-width: 100% !important;
        }
        [data-testid="stAppViewContainer"] {
            background: linear-gradient(90deg, #0b3c91 0%, #1d63dc 50%, #eef2f8 50%, #f8fafd 100%);
            min-height: 100vh;
        }
        .left-hero {
            height: 100vh;
            display: flex;
            flex-direction: column;
            justify-content: center;
            padding: 4rem 5rem;
            color: white;
            position: relative;
            overflow: hidden;
            background: radial-gradient(circle at 20% 30%, rgba(255,255,255,0.12) 0%, transparent 60%),
                        linear-gradient(135deg, #0a3685 0%, #1b5fd4 60%, #3a82f6 100%);
        }
        .left-hero h1 {
            font-size: 2.8rem; font-weight: 700; line-height: 1.25;
            margin-bottom: 1rem; z-index: 2; color: #ffffff;
        }
        .left-hero p {
            font-size: 1.1rem; opacity: 0.85; z-index: 2; max-width: 460px;
        }
        .glass-cube-1 {
            position: absolute; top: 12%; right: 15%; width: 180px; height: 180px;
            background: rgba(255, 255, 255, 0.08); border: 1px solid rgba(255, 255, 255, 0.25);
            border-radius: 16px; transform: rotate(15deg); backdrop-filter: blur(6px);
        }
        .glass-cube-2 {
            position: absolute; bottom: 14%; left: 10%; width: 220px; height: 220px;
            background: rgba(255, 255, 255, 0.06); border: 1px solid rgba(255, 255, 255, 0.2);
            border-radius: 20px; transform: rotate(-12deg); backdrop-filter: blur(4px);
        }
        div[data-testid="column"]:nth-of-type(2) > div {
            display: flex; flex-direction: column; justify-content: center;
            min-height: 100vh; padding: 2rem 3.5rem;
        }
        .glass-card-header {
            background: rgba(255, 255, 255, 0.68);
            backdrop-filter: blur(18px); -webkit-backdrop-filter: blur(18px);
            border: 1px solid rgba(255, 255, 255, 0.9);
            border-radius: 24px 24px 0 0; border-bottom: none;
            padding: 2.5rem 2.5rem 1rem 2.5rem;
            box-shadow: 0 20px 40px rgba(15, 45, 95, 0.08); text-align: center;
        }
        .lock-badge {
            width: 68px; height: 68px; margin: 0 auto 1.2rem auto;
            background: linear-gradient(135deg, #e0ebff 0%, #c2d7ff 100%);
            border-radius: 18px; display: flex; align-items: center;
            justify-content: center; font-size: 2rem;
            box-shadow: 0 8px 20px rgba(29, 99, 220, 0.18);
        }
        .glass-card-header h2 {
            color: #0f172a; font-size: 1.65rem; font-weight: 700; margin-bottom: 0.4rem;
        }
        .glass-card-header p {
            color: #475569; font-size: 0.95rem; margin-bottom: 0;
        }
        [data-testid="stForm"] {
            background: rgba(255, 255, 255, 0.68);
            backdrop-filter: blur(18px); -webkit-backdrop-filter: blur(18px);
            border: 1px solid rgba(255, 255, 255, 0.9);
            border-top: none; border-radius: 0 0 24px 24px;
            padding: 1rem 2.5rem 2.5rem 2.5rem !important;
            box-shadow: 0 20px 40px rgba(15, 45, 95, 0.08);
        }
        [data-testid="stFormSubmitButton"] button {
            width: 100%;
            background: linear-gradient(90deg, #1354c6 0%, #256ef1 100%) !important;
            color: white !important; border: none !important;
            border-radius: 12px !important; padding: 0.65rem 1rem !important;
            font-weight: 600 !important; font-size: 1rem !important;
            margin-top: 0.5rem; box-shadow: 0 6px 16px rgba(29, 99, 220, 0.3);
        }
        </style>
    """, unsafe_allow_html=True)
def render_login_page():
    apply_login_styles()
    col_left, col_right = st.columns([1.15, 1], gap="small")

    with col_left:
        st.markdown("""
            <div class="left-hero">
                <div class="glass-cube-1"></div>
                <div class="glass-cube-2"></div>
                <h1>Sistem Otomatisasi HPP & Laporan Keuangan</h1>
                <p>Platform terintegrasi untuk pengolahan General Ledger, kalkulasi harga pokok penjualan, dan pelaporan akurat.</p>
            </div>
        """, unsafe_allow_html=True)

    with col_right:
        _, center_box, _ = st.columns([0.08, 1.84, 0.08])
        with center_box:
            st.markdown("""
                <div class="glass-card-header">
                    <div class="lock-badge">🔒</div>
                    <h2>Akses Terbatas - Aplikasi HPP</h2>
                    <p>Silakan masukkan kata sandi untuk mengakses sistem laporan HPP</p>
                </div>
            """, unsafe_allow_html=True)

            with st.form("login_form"):
                password = st.text_input("Masukkan Password Aplikasi:", type="password", placeholder="••••••••")
                submitted = st.form_submit_button("Masuk")
                if submitted:
                    if password == APP_PASSWORD:
                        st.session_state.authenticated = True
                        st.rerun()
                    else:
                        st.error("Password salah. Silakan coba lagi.")

if not st.session_state.authenticated:
    render_login_page()
    st.stop()
st.markdown("""
    <style>
    .block-container { padding: 2rem 3rem !important; max-width: 1200px !important; }
    </style>
""", unsafe_allow_html=True)

col_title, col_logout = st.columns([5, 1])
with col_title:
    st.title("📊 Otomatisasi Worksheet HPP & Rekapitulasi GL")
    st.caption("Unggah file General Ledger (Excel/CSV) untuk menghasilkan worksheet konsolidasi dan ringkasan pivot otomatis.")
with col_logout:
    st.write("")
    if st.button("🚪 Keluar", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

st.divider()

KATEGORI_HPP = [
    "1. Pembelian Bahan Baku / Barang Dagang",
    "2. Biaya Angkut Pembelian (Freight In)",
    "3. Retur & Potongan Pembelian",
    "4. Biaya Tenaga Kerja Langsung (BTKL)",
    "5. Biaya Overhead Pabrik (BOP)",
    "6. Penyesuaian Persediaan (Selisih Stok)",
    "7. Beban Pokok Lainnya / Subkontrak"
]
def parse_uploaded_gl(uploaded_file, kategori_label):
    fname = uploaded_file.name.lower()
    if fname.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").copy()

    # Deteksi otomatis kolom utama
    col_map = {}
    for c in df.columns:
        cl = c.lower()
        if any(k in cl for k in ["date", "tanggal", "tgl"]) and "date" not in col_map:
            col_map["date"] = c
        elif any(k in cl for k in ["account", "akun", "kode"]) and "account" not in col_map:
            col_map["account"] = c
        elif any(k in cl for k in ["journal", "jurnal", "ref", "no. bukti", "nomor"]) and "ref" not in col_map:
            col_map["ref"] = c
        elif any(k in cl for k in ["partner", "vendor", "supplier", "rekanan", "kontak"]) and "partner" not in col_map:
            col_map["partner"] = c
        elif any(k in cl for k in ["label", "desc", "keterangan", "uraian", "memo", "name"]) and "desc" not in col_map:
            col_map["desc"] = c
        elif any(k in cl for k in ["debit", "debet"]) and "debit" not in col_map:
            col_map["debit"] = c
        elif any(k in cl for k in ["credit", "kredit"]) and "credit" not in col_map:
            col_map["credit"] = c
        elif any(k in cl for k in ["balance", "saldo", "amount", "nilai", "jumlah"]) and "amount" not in col_map:
            col_map["amount"] = c

    out = pd.DataFrame()
    out["Kategori HPP"] = [kategori_label] * len(df)
    out["Tanggal"] = pd.to_datetime(df[col_map["date"]], errors="coerce") if "date" in col_map else pd.NaT
    out["Akun"] = df[col_map["account"]].astype(str).str.strip() if "account" in col_map else "-"
    out["No. Referensi"] = df[col_map["ref"]].astype(str).str.strip() if "ref" in col_map else "-"
    out["Rekanan / Vendor"] = df[col_map["partner"]].astype(str).str.strip() if "partner" in col_map else "-"
    out["Keterangan"] = df[col_map["desc"]].astype(str).str.strip() if "desc" in col_map else "-"

    # Konversi angka bersih
    def to_num(series):
        return pd.to_numeric(
            series.astype(str).str.replace(",", "", regex=False).str.replace(r"[^\d.\-]", "", regex=True),
            errors="coerce"
        ).fillna(0.0)

    debit = to_num(df[col_map["debit"]]) if "debit" in col_map else pd.Series(0.0, index=df.index)
    credit = to_num(df[col_map["credit"]]) if "credit" in col_map else pd.Series(0.0, index=df.index)

    if "debit" in col_map or "credit" in col_map:
        out["Debit"] = debit
        out["Kredit"] = credit
        out["Net HPP"] = debit - credit
    elif "amount" in col_map:
        amt = to_num(df[col_map["amount"]])
        out["Debit"] = amt.clip(lower=0)
        out["Kredit"] = (-amt).clip(lower=0)
        out["Net HPP"] = amt
    else:
        out["Debit"] = 0.0
        out["Kredit"] = 0.0
        out["Net HPP"] = 0.0

    # Hapus baris total/saldo awal yang tidak memiliki transaksi valid
    mask_invalid = out["Akun"].str.lower().isin(["nan", "none", "total", "subtotal", "initial balance", "saldo awal"])
    out = out[~mask_invalid].copy()
    out = out[(out["Debit"] != 0) | (out["Kredit"] != 0) | (out["Tanggal"].notna())].copy()
    out["Bulan"] = out["Tanggal"].dt.strftime("%Y-%m").fillna("Tanpa Tanggal")

    return out
st.subheader("1️⃣ Unggah File General Ledger (7 Kelompok Akun HPP)")

tab_satuan, tab_massal = st.tabs(["📂 Unggah Per Kategori Akun (7 Slot)", "⚡ Unggah Sekaligus (Multi-File)"])

uploaded_map = {}

with tab_satuan:
    cols = st.columns(2)
    for idx, kat in enumerate(KATEGORI_HPP):
        with cols[idx % 2]:
            f = st.file_uploader(kat, type=["xlsx", "xls", "csv"], key=f"gl_slot_{idx}")
            if f is not None:
                uploaded_map[kat] = [f]

with tab_massal:
    multi_files = st.file_uploader(
        "Pilih hingga 7 file GL sekaligus (akan dikelompokkan berdasarkan urutan atau nama file)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="gl_multi"
    )
    if multi_files:
        for idx, f in enumerate(multi_files):
            kat_target = KATEGORI_HPP[idx % len(KATEGORI_HPP)]
            uploaded_map.setdefault(kat_target, []).append(f)
def build_hpp_excel(df_all):
    wb = Workbook()
    ACC_FORMAT = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

    header_fill = PatternFill(start_color="0B3C91", end_color="0B3C91", fill_type="solid")
    sub_fill = PatternFill(start_color="DCE6F7", end_color="DCE6F7", fill_type="solid")
    total_fill = PatternFill(start_color="E8EEF5", end_color="E8EEF5", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, color="0F172A", size=11)
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    # --- SHEET 1: KONSOLIDASI GL ---
    ws_gl = wb.active
    ws_gl.title = "Konsolidasi_GL"
    cols_gl = ["Kategori HPP", "Tanggal", "Bulan", "Akun", "No. Referensi", "Rekanan / Vendor", "Keterangan", "Debit", "Kredit", "Net HPP"]
    ws_gl.append(cols_gl)

    for col_idx in range(1, len(cols_gl) + 1):
        c = ws_gl.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")

    for _, r in df_all.iterrows():
        tgl_val = r["Tanggal"].strftime("%Y-%m-%d") if pd.notna(r["Tanggal"]) else ""
        ws_gl.append([
            r["Kategori HPP"], tgl_val, r["Bulan"], r["Akun"],
            r["No. Referensi"], r["Rekanan / Vendor"], r["Keterangan"],
            float(r["Debit"]), float(r["Kredit"]), float(r["Net HPP"])
        ])

    for row in ws_gl.iter_rows(min_row=2):
        for cell in row:
            cell.border = thin_border
            if cell.column in [8, 9, 10] and cell.value is not None:
                cell.number_format = ACC_FORMAT

    # --- SHEET 2: RINGKASAN PIVOT HPP ---
    ws_piv = wb.create_sheet(title="Ringkasan_Pivot_HPP")
    ws_piv["A1"] = "RINGKASAN KALKULASI HARGA POKOK PENJUALAN (HPP)"
    ws_piv["A1"].font = Font(bold=True, size=14, color="0B3C91")

    # Tabel 1: Pivot Kategori vs Bulan
    piv_bulanan = pd.pivot_table(
        df_all, values="Net HPP", index="Kategori HPP", columns="Bulan",
        aggfunc="sum", fill_value=0.0, margins=True, margins_name="Total HPP"
    ).reset_index()

    start_r1 = 3
    for c_idx, col_name in enumerate(piv_bulanan.columns, start=1):
        cell = ws_piv.cell(row=start_r1, column=c_idx, value=str(col_name))
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for r_offset, (_, row_data) in enumerate(piv_bulanan.iterrows(), start=1):
        curr_r = start_r1 + r_offset
        is_total = (str(row_data.iloc[0]) == "Total HPP")
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws_piv.cell(row=curr_r, column=c_idx, value=float(val) if c_idx >= 2 else str(val))
            cell.border = thin_border
            if is_total:
                cell.fill = total_fill
                cell.font = bold_font

    # Tabel 2: Rincian Per Akun
    hdr_p2 = start_r1 + len(piv_bulanan) + 3
    ws_piv.cell(row=hdr_p2 - 1, column=1, value="RINCIAN NET HPP PER KATEGORI & AKUN").font = bold_font

    piv_akun = df_all.groupby(["Kategori HPP", "Akun"], as_index=False)[["Debit", "Kredit", "Net HPP"]].sum()
    for c_idx, col_name in enumerate(piv_akun.columns, start=1):
        cell = ws_piv.cell(row=hdr_p2, column=c_idx, value=str(col_name))
        cell.fill = sub_fill
        cell.font = bold_font
        cell.border = thin_border

    for r_offset, (_, row_data) in enumerate(piv_akun.iterrows(), start=1):
        curr_r = hdr_p2 + r_offset
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws_piv.cell(row=curr_r, column=c_idx, value=float(val) if c_idx >= 3 else str(val))
            cell.border = thin_border

    # Format angka akuntansi dengan cell.row (bebas AttributeError)
    for row in ws_piv.iter_rows(min_row=4):
        for cell in row:
            if cell.column >= 2 and cell.value is not None and not (cell.column == 2 and cell.row >= hdr_p2):
                if isinstance(cell.value, (int, float)):
                    cell.number_format = ACC_FORMAT

    # Auto-fit lebar kolom kedua sheet
    for sheet in [ws_gl, ws_piv]:
        for col in sheet.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            sheet.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 5, 42)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
st.subheader("2️⃣ Proses & Unduh Hasil Worksheet HPP")

if uploaded_map:
    if st.button("🚀 Proses & Buat Worksheet HPP", type="primary", use_container_width=True):
        with st.spinner("Sedang menggabungkan dan menghitung rekapitulasi 7 akun GL..."):
            frames = []
            for kat, files in uploaded_map.items():
                for f in files:
                    f.seek(0)
                    frames.append(parse_uploaded_gl(f, kat))

            if frames:
                df_combined = pd.concat(frames, ignore_index=True)
                total_debit = df_combined["Debit"].sum()
                total_kredit = df_combined["Kredit"].sum()
                total_hpp = df_combined["Net HPP"].sum()

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Jumlah File GL", f"{sum(len(v) for v in uploaded_map.values())} File")
                m2.metric("Total Debit", f"Rp {total_debit:,.2f}")
                m3.metric("Total Kredit", f"Rp {total_kredit:,.2f}")
                m4.metric("Total Net HPP", f"Rp {total_hpp:,.2f}")

                with st.expander("🔍 Lihat Pratinjau Data Konsolidasi (20 Baris Pertama)", expanded=True):
                    st.dataframe(df_combined.head(20), use_container_width=True)

                excel_bytes = build_hpp_excel(df_combined)
                st.success("✅ Worksheet HPP berhasil dibuat dan siap diunduh!")
                st.download_button(
                    label="📥 Unduh Worksheet_HPP_Terkonsolidasi.xlsx",
                    data=excel_bytes,
                    file_name="Worksheet_HPP_Terkonsolidasi.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
else:
    st.info("Silakan unggah minimal 1 file General Ledger pada kolom di atas untuk memulai pemrosesan.")
