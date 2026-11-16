import streamlit as st
import pandas as pd
import numpy as np
import io
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

APP_PASSWORD = st.secrets.get("APP_PASSWORD", "admin123")
def apply_login_styles():
    st.markdown("""
        <style>
        header, footer, [data-testid="stToolbar"], [data-testid="stDecoration"] {
            display: none !important;
        }
        .block-container {
            padding: 0rem !important;
            max-width: 100% !important;
        }
        [data-testid="stAppViewContainer"] {
            background: linear-gradient(90deg, #0b3c91 0%, #1d63dc 50%, #eef2f8 50%, #f4f7fc 100%);
            min-height: 100vh;
            overflow-x: hidden;
        }
        [data-testid="stHorizontalBlock"] {
            gap: 0rem !important;
            min-height: 100vh !important;
            align-items: stretch !important;
        }
        .left-hero {
            height: 100vh;
            display: flex;
            flex-direction: column;
            justify-content: center;
            padding: 3rem 4.5rem;
            color: white;
            position: relative;
            overflow: hidden;
            background: radial-gradient(circle at 30% 25%, #2b75f5 0%, #0d429e 65%, #072966 100%);
        }
        .left-hero-svg {
            position: absolute;
            inset: 0;
            width: 100%;
            height: 100%;
            z-index: 1;
            opacity: 0.92;
        }
        .left-hero-content {
            position: relative;
            z-index: 2;
            margin-top: 2rem;
        }
        .left-hero h1 {
            font-size: 2.75rem;
            font-weight: 700;
            line-height: 1.22;
            margin-bottom: 1rem;
            color: #ffffff;
            text-shadow: 0 4px 18px rgba(0, 15, 50, 0.35);
        }
        .left-hero p {
            font-size: 1.05rem;
            color: rgba(255, 255, 255, 0.88);
            max-width: 440px;
            line-height: 1.55;
        }
        [data-testid="stHorizontalBlock"] > [data-testid="column"]:nth-child(2) {
            position: relative;
            display: flex !important;
            flex-direction: column !important;
            justify-content: center !important;
            align-items: center !important;
            min-height: 100vh !important;
            background: radial-gradient(circle at 80% 18%, rgba(163, 198, 255, 0.45) 0%, transparent 22%),
                        radial-gradient(circle at 22% 78%, rgba(130, 177, 255, 0.5) 0%, transparent 25%),
                        radial-gradient(circle at 82% 85%, rgba(190, 214, 255, 0.4) 0%, transparent 18%),
                        #eff3fa;
            padding: 2rem !important;
        }
        .orb-1, .orb-2, .orb-3 {
            position: fixed;
            border-radius: 50%;
            background: radial-gradient(circle at 30% 30%, rgba(255,255,255,0.95), rgba(145,185,255,0.55) 60%, rgba(85,140,245,0.35));
            box-shadow: inset -6px -6px 16px rgba(45, 95, 195, 0.25), 0 12px 28px rgba(30, 75, 165, 0.15);
            pointer-events: none;
            z-index: 1;
        }
        .orb-1 { width: 95px; height: 95px; top: 14%; right: 8%; }
        .orb-2 { width: 125px; height: 125px; bottom: 13%; left: 53%; }
        .orb-3 { width: 52px; height: 52px; top: 20%; left: 56%; }

        .login-wrapper {
            position: relative;
            z-index: 3;
            width: 100%;
            max-width: 440px;
            margin: 0 auto;
        }
        .glass-card-top {
            background: rgba(255, 255, 255, 0.58);
            backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            border: 1.5px solid rgba(255, 255, 255, 0.92);
            border-bottom: none;
            border-radius: 26px 26px 0 0;
            padding: 2.5rem 2.2rem 1.2rem 2.2rem;
            text-align: center;
            box-shadow: 0 -10px 35px rgba(20, 55, 120, 0.06);
        }
        .lock-badge {
            width: 74px;
            height: 74px;
            margin: 0 auto 1.25rem auto;
            background: linear-gradient(145deg, #ffffff 0%, #dbe8ff 100%);
            border: 1px solid rgba(255, 255, 255, 0.9);
            border-radius: 22px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 2.2rem;
            box-shadow: 0 12px 26px rgba(25, 85, 195, 0.22);
        }
        .glass-card-top h2 {
            color: #0f172a;
            font-size: 1.6rem;
            font-weight: 700;
            margin-bottom: 0.45rem;
        }
        .glass-card-top p {
            color: #475569;
            font-size: 0.92rem;
            line-height: 1.45;
            margin: 0;
        }
        div[data-testid="stForm"] {
            position: relative;
            z-index: 3;
            width: 100%;
            max-width: 440px;
            margin: 0 auto !important;
            background: rgba(255, 255, 255, 0.58) !important;
            backdrop-filter: blur(20px) !important;
            -webkit-backdrop-filter: blur(20px) !important;
            border: 1.5px solid rgba(255, 255, 255, 0.92) !important;
            border-top: none !important;
            border-radius: 0 0 26px 26px !important;
            padding: 0.5rem 2.2rem 2.5rem 2.2rem !important;
            box-shadow: 0 22px 45px rgba(20, 55, 120, 0.12) !important;
        }
        div[data-testid="stFormSubmitButton"],
        div[data-testid="stFormSubmitButton"] > button {
            width: 100% !important;
            display: block !important;
        }
        div[data-testid="stFormSubmitButton"] > button {
            background: linear-gradient(90deg, #1252c4 0%, #256ef1 100%) !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 999px !important;
            padding: 0.7rem 1.5rem !important;
            font-size: 1rem !important;
            font-weight: 600 !important;
            margin-top: 0.75rem !important;
            box-shadow: 0 8px 20px rgba(25, 92, 215, 0.35) !important;
            transition: all 0.2s ease !important;
        }
        div[data-testid="stFormSubmitButton"] > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 10px 24px rgba(25, 92, 215, 0.48) !important;
        }
        </style>
    """, unsafe_allow_html=True)
def render_login_page():
    apply_login_styles()
    col_left, col_right = st.columns([1.05, 1], gap="small")

    with col_left:
        st.markdown("""
            <div class="left-hero">
                <svg class="left-hero-svg" viewBox="0 0 800 900" fill="none" xmlns="http://w3.org">
                    <polygon points="360,120 560,235 360,350 160,235" fill="rgba(255,255,255,0.14)" stroke="rgba(255,255,255,0.35)" stroke-width="2"/>
                    <polygon points="160,235 360,350 360,580 160,465" fill="rgba(255,255,255,0.07)" stroke="rgba(255,255,255,0.28)" stroke-width="2"/>
                    <polygon points="560,235 360,350 360,580 560,465" fill="rgba(255,255,255,0.11)" stroke="rgba(255,255,255,0.32)" stroke-width="2"/>
                    <polygon points="620,70 740,140 620,210 500,140" fill="rgba(255,255,255,0.12)" stroke="rgba(255,255,255,0.3)" stroke-width="1.5"/>
                    <polygon points="500,140 620,210 620,350 500,280" fill="rgba(255,255,255,0.06)" stroke="rgba(255,255,255,0.22)" stroke-width="1.5"/>
                    <polygon points="740,140 620,210 620,350 740,280" fill="rgba(255,255,255,0.09)" stroke="rgba(255,255,255,0.25)" stroke-width="1.5"/>
                    <polygon points="560,510 740,615 560,720 380,615" fill="rgba(255,255,255,0.12)" stroke="rgba(255,255,255,0.3)" stroke-width="1.5"/>
                    <polygon points="380,615 560,720 560,880 380,775" fill="rgba(255,255,255,0.06)" stroke="rgba(255,255,255,0.25)" stroke-width="1.5"/>
                    <polygon points="740,615 560,720 560,880 740,775" fill="rgba(255,255,255,0.10)" stroke="rgba(255,255,255,0.28)" stroke-width="1.5"/>
                    <rect x="430" y="730" width="16" height="55" rx="3" fill="rgba(160,210,255,0.45)"/>
                    <rect x="456" y="715" width="16" height="82" rx="3" fill="rgba(180,220,255,0.55)"/>
                    <rect x="482" y="700" width="16" height="110" rx="3" fill="rgba(210,235,255,0.65)"/>
                    <rect x="508" y="685" width="16" height="138" rx="3" fill="rgba(255,255,255,0.75)"/>
                </svg>
                <div class="left-hero-content">
                    <h1>Sistem Otomatisasi HPP & Laporan Keuangan</h1>
                    <p>Platform terintegrasi untuk pengolahan General Ledger, kalkulasi harga pokok penjualan, dan pelaporan akurat.</p>
                </div>
            </div>
        """, unsafe_allow_html=True)

    with col_right:
        st.markdown("""
            <div class="orb-1"></div>
            <div class="orb-2"></div>
            <div class="orb-3"></div>
            <div class="login-wrapper">
                <div class="glass-card-top">
                    <div class="lock-badge">🔒</div>
                    <h2>Akses Terbatas - Aplikasi HPP</h2>
                    <p>Silakan masukkan kata sandi untuk mengakses<br>sistem laporan HPP</p>
                </div>
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
def apply_dashboard_styles():
    st.markdown("""
        <style>
        header, footer, [data-testid="stToolbar"], [data-testid="stDecoration"] {
            display: none !important;
        }
        [data-testid="stAppViewContainer"] {
            background: radial-gradient(circle at 12% 10%, rgba(160, 198, 255, 0.35) 0%, transparent 32%),
                        radial-gradient(circle at 88% 85%, rgba(135, 180, 255, 0.32) 0%, transparent 35%),
                        linear-gradient(135deg, #eaf0fa 0%, #f5f8fe 50%, #e6eefb 100%);
            min-height: 100vh;
        }
        .block-container {
            padding: 2rem 3rem 4rem 3rem !important;
            max-width: 1240px !important;
        }
        .dash-hero {
            background: linear-gradient(125deg, #0b3c91 0%, #1d63dc 65%, #3882f6 100%);
            border-radius: 22px;
            padding: 1.8rem 2.2rem;
            color: white;
            box-shadow: 0 16px 36px rgba(13, 62, 152, 0.22);
            border: 1px solid rgba(255, 255, 255, 0.35);
            margin-bottom: 1.5rem;
        }
        .dash-hero h1 {
            font-size: 1.85rem;
            font-weight: 700;
            margin: 0 0 0.35rem 0;
            color: #ffffff;
        }
        .dash-hero p {
            font-size: 0.96rem;
            margin: 0;
            color: rgba(255, 255, 255, 0.88);
        }
        .section-glass-title {
            background: rgba(255, 255, 255, 0.68);
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            border: 1px solid rgba(255, 255, 255, 0.9);
            border-radius: 16px;
            padding: 0.95rem 1.4rem;
            margin: 1.2rem 0 1rem 0;
            box-shadow: 0 8px 24px rgba(15, 45, 95, 0.06);
            color: #0f172a;
            font-weight: 700;
            font-size: 1.15rem;
        }
        [data-testid="stFileUploader"] {
            background: rgba(255, 255, 255, 0.65) !important;
            backdrop-filter: blur(16px) !important;
            -webkit-backdrop-filter: blur(16px) !important;
            border: 1px solid rgba(255, 255, 255, 0.92) !important;
            border-radius: 18px !important;
            padding: 1rem 1.2rem !important;
            box-shadow: 0 10px 26px rgba(20, 55, 120, 0.06) !important;
            margin-bottom: 0.85rem !important;
        }
        [data-testid="stFileUploader"] section {
            background: rgba(240, 246, 255, 0.75) !important;
            border: 1.5px dashed #7aa7f7 !important;
            border-radius: 12px !important;
        }
        [data-testid="stMetric"] {
            background: rgba(255, 255, 255, 0.72) !important;
            backdrop-filter: blur(16px) !important;
            border: 1px solid rgba(255, 255, 255, 0.95) !important;
            border-radius: 18px !important;
            padding: 1.1rem 1.3rem !important;
            box-shadow: 0 10px 25px rgba(20, 55, 120, 0.07) !important;
        }
        .stButton > button, .stDownloadButton > button {
            border-radius: 999px !important;
            font-weight: 600 !important;
            padding: 0.65rem 1.4rem !important;
            box-shadow: 0 8px 20px rgba(25, 92, 215, 0.25) !important;
            transition: all 0.2s ease !important;
        }
        .stButton > button[kind="primary"], .stDownloadButton > button {
            background: linear-gradient(90deg, #1252c4 0%, #256ef1 100%) !important;
            color: #ffffff !important;
            border: none !important;
        }
        </style>
    """, unsafe_allow_html=True)

apply_dashboard_styles()

col_head, col_out = st.columns([5.5, 1])
with col_head:
    st.markdown("""
        <div class="dash-hero">
            <h1>📊 Otomatisasi Worksheet HPP & Rekapitulasi GL</h1>
            <p>Unggah file Master HPP, Klasifikasi, Trial Balance, dan General Ledger bulan berjalan untuk pemrosesan otomatis.</p>
        </div>
    """, unsafe_allow_html=True)
with col_out:
    st.write("")
    st.write("")
    if st.button("🚪 Keluar", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()
def parse_excel_generic(uploaded_file, source_label):
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").copy()

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
    out["Sumber File"] = [source_label] * len(df)
    out["Tanggal"] = pd.to_datetime(df[col_map["date"]], errors="coerce") if "date" in col_map else pd.NaT
    out["Akun"] = df[col_map["account"]].astype(str).str.strip() if "account" in col_map else "-"
    out["No. Referensi"] = df[col_map["ref"]].astype(str).str.strip() if "ref" in col_map else "-"
    out["Rekanan / Vendor"] = df[col_map["partner"]].astype(str).str.strip() if "partner" in col_map else "-"
    out["Keterangan"] = df[col_map["desc"]].astype(str).str.strip() if "desc" in col_map else "-"

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

    mask_invalid = out["Akun"].str.lower().isin(["nan", "none", "total", "subtotal", "initial balance", "saldo awal"])
    out = out[~mask_invalid].copy()
    out = out[(out["Debit"] != 0) | (out["Kredit"] != 0) | (out["Tanggal"].notna())].copy()
    out["Bulan"] = out["Tanggal"].dt.strftime("%Y-%m").fillna("Tanpa Tanggal")

    return out
st.markdown('<div class="section-glass-title">📂 Unggah Dokumen Sumber HPP & General Ledger</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader("1️⃣ Upload File Master HPP (s.d. September)", type=["xlsx", "xlsm"])
    klasifikasi_file = st.file_uploader("2️⃣ Upload File Klasifikasi (Pembantu)", type=["xlsx"])
with col2:
    tb_file = st.file_uploader("3️⃣ Upload Trial Balance (TB) Bulan Berjalan", type=["xlsx"])
    gl_files = st.file_uploader("4️⃣ Upload Semua File GL Bulan Berjalan (Bisa pilih banyak)", type=["xlsx"], accept_multiple_files=True)
def build_hpp_excel(df_all, df_tb=None, df_klas=None):
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

    ws_gl = wb.active
    ws_gl.title = "Konsolidasi_GL"
    cols_gl = ["Sumber File", "Tanggal", "Bulan", "Akun", "No. Referensi", "Rekanan / Vendor", "Keterangan", "Debit", "Kredit", "Net HPP"]
    ws_gl.append(cols_gl)

    for col_idx in range(1, len(cols_gl) + 1):
        c = ws_gl.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")

    for _, r in df_all.iterrows():
        tgl_val = r["Tanggal"].strftime("%Y-%m-%d") if pd.notna(r["Tanggal"]) else ""
        ws_gl.append([
            r["Sumber File"], tgl_val, r["Bulan"], r["Akun"],
            r["No. Referensi"], r["Rekanan / Vendor"], r["Keterangan"],
            float(r["Debit"]), float(r["Kredit"]), float(r["Net HPP"])
        ])

    for row in ws_gl.iter_rows(min_row=2):
        for cell in row:
            cell.border = thin_border
            if cell.column in (8, 9, 10) and cell.value is not None:
                cell.number_format = ACC_FORMAT

    ws_piv = wb.create_sheet(title="Ringkasan_Pivot_HPP")
    ws_piv["A1"] = "RINGKASAN KALKULASI HARGA POKOK PENJUALAN (HPP)"
    ws_piv["A1"].font = Font(bold=True, size=14, color="0B3C91")

    piv_bulanan = pd.pivot_table(
        df_all, values="Net HPP", index="Sumber File", columns="Bulan",
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
        is_total = (str(row_data.iloc) == "Total HPP")
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws_piv.cell(row=curr_r, column=c_idx, value=float(val) if c_idx >= 2 else str(val))
            cell.border = thin_border
            if is_total:
                cell.fill = total_fill
                cell.font = bold_font

    hdr_p2 = start_r1 + len(piv_bulanan) + 3
    ws_piv.cell(row=hdr_p2 - 1, column=1, value="RINCIAN NET HPP PER SUMBER FILE & AKUN").font = bold_font

    piv_akun = df_all.groupby(["Sumber File", "Akun"], as_index=False)[["Debit", "Kredit", "Net HPP"]].sum()
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

    for row in ws_piv.iter_rows(min_row=4):
        for cell in row:
            if cell.column >= 2 and cell.value is not None and not (cell.column == 2 and cell.row >= hdr_p2):
                if isinstance(cell.value, (int, float)):
                    cell.number_format = ACC_FORMAT

    for sheet in [ws_gl, ws_piv]:
        for col in sheet.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            sheet.column_dimensions[get_column_letter(col.column)].width = min(max_len + 5, 42)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
st.markdown('<div class="section-glass-title">⚙️ Proses & Unduh Hasil Worksheet HPP</div>', unsafe_allow_html=True)

has_any_file = any([master_file, klasifikasi_file, tb_file]) or (gl_files and len(gl_files) > 0)

if has_any_file:
    if st.button("🚀 Proses & Buat Worksheet HPP", type="primary", use_container_width=True):
        with st.spinner("Sedang memproses Master HPP, Klasifikasi, TB, dan seluruh file GL..."):
            frames = []
            total_uploaded = 0

            if master_file is not None:
                master_file.seek(0)
                frames.append(parse_excel_generic(master_file, f"Master HPP ({master_file.name})"))
                total_uploaded += 1

            if tb_file is not None:
                tb_file.seek(0)
                frames.append(parse_excel_generic(tb_file, f"Trial Balance ({tb_file.name})"))
                total_uploaded += 1

            if gl_files:
                for f in gl_files:
                    f.seek(0)
                    frames.append(parse_excel_generic(f, f"GL - {f.name}"))
                    total_uploaded += 1

            df_klas = pd.read_excel(klasifikasi_file) if klasifikasi_file is not None else None
            if klasifikasi_file is not None:
                total_uploaded += 1

            if frames:
                df_combined = pd.concat(frames, ignore_index=True)
                total_debit = df_combined["Debit"].sum()
                total_kredit = df_combined["Kredit"].sum()
                total_hpp = df_combined["Net HPP"].sum()

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Total Dokumen Diproses", f"{total_uploaded} File")
                m2.metric("Total Debit", f"Rp {total_debit:,.2f}")
                m3.metric("Total Kredit", f"Rp {total_kredit:,.2f}")
                m4.metric("Total Net HPP", f"Rp {total_hpp:,.2f}")

                with st.expander("🔍 Lihat Pratinjau Data Konsolidasi (20 Baris Pertama)", expanded=True):
                    st.dataframe(df_combined.head(20), use_container_width=True)

                excel_bytes = build_hpp_excel(df_combined, df_klas=df_klas)
                st.success("✅ Worksheet HPP berhasil dibuat dan siap diunduh!")
                st.download_button(
                    label="📥 Unduh Worksheet_HPP_Terkonsolidasi.xlsx",
                    data=excel_bytes,
                    file_name="Worksheet_HPP_Terkonsolidasi.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.warning("Unggah minimal 1 file transaksi (Master HPP, TB, atau GL) untuk dikonsolidasi.")
else:
    st.info("Silakan unggah file pada keempat kolom di atas untuk memulai pemrosesan.")
