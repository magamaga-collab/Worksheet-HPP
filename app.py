import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import ILLEGAL_CHARACTERS_RE
from io import BytesIO
import re

st.set_page_config(page_title="Otomatisasi HPP & Rekonsiliasi GL-TB", layout="wide")

# --- SISTEM KEAMANAN PASSWORD ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        st.title("🔒 Akses Terbatas - Aplikasi HPP")
        pwd = st.text_input("Masukkan Password Aplikasi:", type="password")
        if st.button("Masuk"):
            if "APP_PASSWORD" in st.secrets and pwd == st.secrets["APP_PASSWORD"]:
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("❌ Password salah atau belum diatur di Streamlit Secrets.")
        st.stop()

check_password()
# --------------------------------

st.title("📊 Otomatisasi HPP, Checking GL vs TB & Pivot Otomatis")
st.markdown("Unggah **Master HPP (s.d. Bulan Sebelumnya)**, **File Klasifikasi**, **Trial Balance (TB) Bulan Berjalan**, dan **File-file GL** untuk memproses laporan secara otomatis.")

ACC_FORMAT = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

def sanitize_excel_str(val):
    """Menghapus karakter kontrol ilegal agar XML Excel tidak corrupt"""
    if pd.isna(val) or val is None:
        return ""
    return ILLEGAL_CHARACTERS_RE.sub("", str(val)).strip()

def clean_account_name(val):
    """Membersihkan nama akun dan menghapus akhiran jumlah baris seperti (331)"""
    s = sanitize_excel_str(val)
    return re.sub(r'\s*\(\d+\)$', '', s).strip()

def split_account_code_desc(acc_str):
    """Memisahkan 5 digit nomor akun di awal string dan deskripsi akunnya"""
    s = clean_account_name(acc_str)
    if not s:
        return "", ""
    match = re.match(r'^(\d{5})\s*[-–]?\s*(.*)$', s)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    match_any = re.match(r'^(\d+)\s*[-–]?\s*(.*)$', s)
    if match_any:
        return match_any.group(1).strip(), match_any.group(2).strip()
    return "", s

def extract_account_code(acc_str):
    code, _ = split_account_code_desc(acc_str)
    return code
def load_klasifikasi(file):
    """Membaca file klasifikasi untuk mapping & koordinat sel referensi di Excel"""
    df_raw = pd.read_excel(file, header=None)
    
    header_row_idx = 0
    for idx, row in df_raw.iterrows():
        row_vals = [str(x).strip().lower() for x in row.values if pd.notna(x)]
        if 'nama project' in row_vals or 'account' in row_vals:
            header_row_idx = idx
            break
            
    df_k = pd.read_excel(file, header=header_row_idx)
    df_k.columns = [str(c).strip() for c in df_k.columns]
    
    proj_map = {}
    pct_map = {}
    if 'Nama Project' in df_k.columns:
        for _, r in df_k.dropna(subset=['Nama Project']).iterrows():
            p_name = sanitize_excel_str(r['Nama Project'])
            k_proj = r.get('Klasifikasi Project New', r.get('Klasifikasi Project', ''))
            pct_val = r.get('Percentage 2026', r.get('Percentage', 0))
            try:
                pct_val = float(pct_val) if pd.notna(pct_val) else 0.0
            except:
                pct_val = 0.0
            proj_map[p_name] = sanitize_excel_str(k_proj)
            pct_map[p_name] = pct_val

    acc_map = {}
    if 'Account' in df_k.columns and 'Klasifikasi LK' in df_k.columns:
        for _, r in df_k.dropna(subset=['Account']).iterrows():
            acc_full = sanitize_excel_str(r['Account'])
            acc_code = extract_account_code(acc_full)
            k_lk = sanitize_excel_str(r['Klasifikasi LK'])
            acc_map[acc_full] = k_lk
            if acc_code:
                acc_map[acc_code] = k_lk
                
    return proj_map, pct_map, acc_map, df_raw

def parse_gl_files(gl_files):
    """Membaca banyak file GL, membuang baris summary akun/total, dan memisahkan kode & deskripsi akun"""
    all_tx = []
    gl_summary_list = []
    
    for f in gl_files:
        df = pd.read_excel(f)
        df.columns = [str(c).strip() for c in df.columns]
        
        rename_dict = {}
        for col in df.columns:
            if col.lower() in ['detail project', 'detil project']:
                rename_dict[col] = 'detil project'
        df = df.rename(columns=rename_dict)
        
        for num_col in ['Debit', 'Credit', 'Amount in Currency']:
            if num_col in df.columns:
                df[num_col] = pd.to_numeric(df[num_col], errors='coerce').fillna(0.0)
                
        if 'Account' in df.columns:
            df['Account'] = df['Account'].ffill()
            
        parsed_dates = pd.to_datetime(df['Date'], errors='coerce')
        df_valid = df[parsed_dates.notna()].copy()
        
        cleaned_accs = [clean_account_name(x) for x in df_valid['Account'].tolist()]
        split_pairs = [split_account_code_desc(x) for x in cleaned_accs]
        
        df_valid['Account'] = cleaned_accs
        df_valid['Account_Code'] = [pair for pair in split_pairs]
        df_valid['GL_Account_Desc'] = [pair for pair in split_pairs]
        
        df_valid = df_valid[df_valid['Account_Code'] != ""].copy()
        df_valid['Tags_Calc'] = df_valid['Debit'] - df_valid['Credit']
        
        all_tx.append(df_valid)
        
        if not df_valid.empty:
            grp = df_valid.groupby(['Account_Code', 'GL_Account_Desc'], as_index=False).agg(
                GL_Debit=('Debit', 'sum'),
                GL_Credit=('Credit', 'sum'),
                GL_Net_Mutasi=('Tags_Calc', 'sum')
            )
            gl_summary_list.append(grp)
            
    df_gl_all = pd.concat(all_tx, ignore_index=True) if all_tx else pd.DataFrame()
    df_gl_sum = pd.concat(gl_summary_list, ignore_index=True) if gl_summary_list else pd.DataFrame()
    if not df_gl_sum.empty:
        df_gl_sum = df_gl_sum.groupby(['Account_Code', 'GL_Account_Desc'], as_index=False).sum()
        
    return df_gl_all, df_gl_sum

def parse_tb_file(tb_file):
    """Membaca Kolom A dari file TB yang diawali 5 digit kode akun dan mengambil Debit & Credit Bulan Berjalan"""
    df_raw = pd.read_excel(tb_file, header=None)
    
    bb_col_idx = None
    for r_idx in range(min(5, len(df_raw))):
        for c_idx in range(1, len(df_raw.columns)):
            val = df_raw.iat[r_idx, c_idx]
            if pd.notna(val):
                val_low = str(val).lower()
                if any(k in val_low for k in ['bulan berjalan', 'jan', 'feb', 'mar', 'apr', 'mei', 'may', 'jun', 'jul', 'agu', 'aug', 'sep', 'okt', 'oct', 'nov', 'des', 'dec']):
                    bb_col_idx = c_idx
                    break
        if bb_col_idx is not None:
            break
            
    if bb_col_idx is None:
        bb_col_idx = 3
        
    tb_records = []
    for r_idx in range(len(df_raw)):
        acc_val = df_raw.iat[r_idx, 0]
        if pd.isna(acc_val):
            continue
        acc_str = str(acc_val).strip()
        
        match_5digit = re.match(r'^(\d{5})\s*[-–]?\s*(.*)$', acc_str)
        if not match_5digit:
            continue
            
        acc_code = match_5digit.group(1).strip()
        acc_desc = match_5digit.group(2).strip()
        
        debit_raw = df_raw.iat[r_idx, bb_col_idx] if len(df_raw.columns) > bb_col_idx else 0.0
        credit_raw = df_raw.iat[r_idx, bb_col_idx + 1] if len(df_raw.columns) > bb_col_idx + 1 else 0.0
        
        debit_bb = pd.to_numeric(debit_raw, errors='coerce')
        credit_bb = pd.to_numeric(credit_raw, errors='coerce')
        debit_bb = 0.0 if pd.isna(debit_bb) else float(debit_bb)
        credit_bb = 0.0 if pd.isna(credit_bb) else float(credit_bb)
        
        tb_records.append({
            'Account_Code': acc_code,
            'TB_Account_Desc': acc_desc,
            'TB_Debit_Berjalan': debit_bb,
            'TB_Credit_Berjalan': credit_bb,
            'TB_Net_Mutasi': debit_bb - credit_bb
        })
        
    return pd.DataFrame(tb_records)
col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader("1️⃣ Upload File Master HPP (s.d. September)", type=["xlsx", "xlsm"])
    klasifikasi_file = st.file_uploader("2️⃣ Upload File Klasifikasi (Pembantu)", type=["xlsx"])
with col2:
    tb_file = st.file_uploader("3️⃣ Upload Trial Balance (TB) Bulan Berjalan", type=["xlsx"])
    gl_files = st.file_uploader("4️⃣ Upload Semua File GL Bulan Berjalan (Bisa pilih banyak)", type=["xlsx"], accept_multiple_files=True)

if master_file and klasifikasi_file and tb_file and gl_files:
    if st.button("🚀 Proses & Rekonsiliasi Data", type="primary"):
        with st.spinner("Sedang memproses file Master, menyertakan Ref_Klasifikasi, dan menghidupkan seluruh rumus..."):
            proj_map, pct_map, acc_map, df_klas_raw = load_klasifikasi(klasifikasi_file)
            df_gl_all, df_gl_sum = parse_gl_files(gl_files)
            df_tb = parse_tb_file(tb_file)
            
            df_merge = pd.merge(df_gl_sum, df_tb, on='Account_Code', how='left')
            for col in ['TB_Debit_Berjalan', 'TB_Credit_Berjalan', 'TB_Net_Mutasi']:
                df_merge[col] = df_merge[col].fillna(0.0)
            df_merge['TB_Account_Desc'] = df_merge['TB_Account_Desc'].fillna('(Tidak ditemukan di TB)')
            
            vertical_rows = []
            for _, r in df_merge.iterrows():
                acc_code = r['Account_Code']
                gl_desc = r['GL_Account_Desc']
                tb_desc = r['TB_Account_Desc']
                
                gl_deb = round(float(r['GL_Debit']), 2)
                gl_cred = round(float(r['GL_Credit']), 2)
                gl_net = round(float(r['GL_Net_Mutasi']), 2)
                
                tb_deb = round(float(r['TB_Debit_Berjalan']), 2)
                tb_cred = round(float(r['TB_Credit_Berjalan']), 2)
                tb_net = round(float(r['TB_Net_Mutasi']), 2)
                
                diff_deb = round(gl_deb - tb_deb, 2)
                diff_cred = round(gl_cred - tb_cred, 2)
                diff_net = round(gl_net - tb_net, 2)
                
                status_str = 'BALANCE' if (abs(diff_deb) < 1.0 and abs(diff_cred) < 1.0 and abs(diff_net) < 1.0) else 'SELISIH'
                
                vertical_rows.append({
                    'Sumber': 'GL',
                    'Nomor Akun': acc_code,
                    'Nama Akun': gl_desc,
                    'Debet': gl_deb,
                    'Kredit': gl_cred,
                    'Net Mutasi': gl_net
                })
                vertical_rows.append({
                    'Sumber': 'TB',
                    'Nomor Akun': acc_code,
                    'Nama Akun': tb_desc,
                    'Debet': tb_deb,
                    'Kredit': tb_cred,
                    'Net Mutasi': tb_net
                })
                vertical_rows.append({
                    'Sumber': 'Selisih (GL - TB)',
                    'Nomor Akun': acc_code,
                    'Nama Akun': status_str,
                    'Debet': diff_deb,
                    'Kredit': diff_cred,
                    'Net Mutasi': diff_net
                })
                
            ordered_cols = ['Sumber', 'Nomor Akun', 'Nama Akun', 'Debet', 'Kredit', 'Net Mutasi']
            df_check = pd.DataFrame(vertical_rows, columns=ordered_cols)
            
            st.subheader("🔍 Hasil Checking Mutasi GL vs TB")
            st.dataframe(df_check, use_container_width=True, hide_index=True)
            master_bytes = master_file.getvalue()
            wb_val = openpyxl.load_workbook(BytesIO(master_bytes), data_only=True)
            ws_val = wb_val[wb_val.sheetnames]
            
            wb = openpyxl.load_workbook(BytesIO(master_bytes))
            main_sheet_name = wb.sheetnames
            ws_main = wb[main_sheet_name]
            
            # 1. Buat / Perbarui Sheet Referensi 'Ref_Klasifikasi' agar semua rumus INDEX-MATCH hidup di dalam file ini
            if 'Ref_Klasifikasi' in wb.sheetnames:
                del wb['Ref_Klasifikasi']
            ws_ref = wb.create_sheet(title='Ref_Klasifikasi')
            
            for r_i in range(len(df_klas_raw)):
                row_vals = [sanitize_excel_str(v) if isinstance(v, str) else (None if pd.isna(v) else v) for v in df_klas_raw.iloc[r_i].tolist()]
                ws_ref.append(row_vals)
                
            # Deteksi koordinat kolom pada sheet Ref_Klasifikasi untuk rumus INDEX-MATCH
            ref_proj_col, ref_kproj_col, ref_pct_col = "B", "C", "D"
            ref_acc_col, ref_klk_col = "F", "G"
            ref_start_row = 2
            ref_end_row = max(ws_ref.max_row, 2)
            
            for r_i in range(1, min(10, ws_ref.max_row + 1)):
                for c_i in range(1, ws_ref.max_column + 1):
                    v = str(ws_ref.cell(row=r_i, column=c_i).value or '').strip().lower()
                    col_ltr = get_column_letter(c_i)
                    if v == 'nama project':
                        ref_proj_col = col_ltr
                        ref_start_row = r_i + 1
                    elif v in ['klasifikasi project new', 'klasifikasi project']:
                        ref_kproj_col = col_ltr
                    elif v in ['percentage 2026', 'percentage']:
                        ref_pct_col = col_ltr
                    elif v == 'account':
                        ref_acc_col = col_ltr
                    elif v == 'klasifikasi lk':
                        ref_klk_col = col_ltr

            # Format persentase di Ref_Klasifikasi
            pct_col_num = openpyxl.utils.column_index_from_string(ref_pct_col)
            for r_i in range(ref_start_row, ref_end_row + 1):
                c_pct = ws_ref.cell(row=r_i, column=pct_col_num)
                if isinstance(c_pct.value, (int, float)):
                    c_pct.number_format = '0.00%'

            # 2. Baca struktur kolom Worksheet Utama
            headers = [cell.value for cell in ws_main]
            col_idx_map = {str(h).strip(): idx + 1 for idx, h in enumerate(headers) if h is not None}
            
            last_row = ws_main.max_row
            while last_row > 1 and ws_main.cell(row=last_row, column=1).value is None and ws_main.cell(row=last_row, column=3).value is None:
                last_row -= 1
                
            no_urut_col = col_idx_map.get('no urut', 25)
            last_no_urut = 0
            for r in range(last_row, 1, -1):
                val = ws_val.cell(row=r, column=no_urut_col).value
                if isinstance(val, (int, float)):
                    last_no_urut = int(val)
                    break
                    
            # Ambil huruf kolom dinamis dari Worksheet Utama untuk menyusun rumus yang presisi
            c_acc = get_column_letter(col_idx_map.get('Account', 3))
            c_proj = get_column_letter(col_idx_map.get('Project', 5))
            c_deb = get_column_letter(col_idx_map.get('Debit', 12))
            c_cred = get_column_letter(col_idx_map.get('Credit', 13))
            c_tags = get_column_letter(col_idx_map.get('Tags', 14))
            c_pct26 = get_column_letter(col_idx_map.get('Percentage 2026', 22))
            c_hpp26 = get_column_letter(col_idx_map.get('HPP 2026', 23))
            c_sisa26 = get_column_letter(col_idx_map.get('Sisa Persediaan 2026', 24))
            c_klk = get_column_letter(col_idx_map.get('Klasifikasi LK', 26))
            c_kpn = get_column_letter(col_idx_map.get('Klasifikasi Project New', 28))
            
            current_row = last_row + 1
            current_no = last_no_urut + 1
            new_rows_for_pivot = []
            
            acc_cols_list = ['Amount in Currency', 'Debit', 'Credit', 'Tags', 'HPP', 'Sisa Persediaan', 'HPP 2025', 'Sisa Persediaan 25', 'HPP 2026', 'Sisa Persediaan 2026']
            pct_cols_list = ['Percentage', 'Percentage 2025', 'Percentage 2026']

            # 3. Append Baris Baru dengan Rumus Hidup (INDEX-MATCH & Perhitungan HPP)
            for _, tx in df_gl_all.iterrows():
                acc_full = clean_account_name(tx.get('Account', ''))
                acc_code = str(tx.get('Account_Code', '')).strip()
                proj_name = sanitize_excel_str(tx.get('Project', ''))
                
                pct_2026 = pct_map.get(proj_name, 0.0)
                klas_lk = acc_map.get(acc_full, acc_map.get(acc_code, ""))
                klas_proj_new = proj_map.get(proj_name, "")
                
                tx_date = pd.to_datetime(tx.get('Date'), errors='coerce')
                date_val = tx_date.to_pydatetime() if pd.notna(tx_date) else tx.get('Date')
                
                # Rumus Excel Aktif yang merujuk ke sheet Ref_Klasifikasi
                f_pct26 = f'=IFERROR(INDEX(Ref_Klasifikasi!${ref_pct_col}${ref_start_row}:${ref_pct_col}${ref_end_row}, MATCH({c_proj}{current_row}, Ref_Klasifikasi!${ref_proj_col}${ref_start_row}:${ref_proj_col}${ref_end_row}, 0)), 0)'
                f_klk = f'=IFERROR(INDEX(Ref_Klasifikasi!${ref_klk_col}${ref_start_row}:${ref_klk_col}${ref_end_row}, MATCH({c_acc}{current_row}, Ref_Klasifikasi!${ref_acc_col}${ref_start_row}:${ref_acc_col}${ref_end_row}, 0)), "")'
                f_kpn = f'=IFERROR(INDEX(Ref_Klasifikasi!${ref_kproj_col}${ref_start_row}:${ref_kproj_col}${ref_end_row}, MATCH({c_proj}{current_row}, Ref_Klasifikasi!${ref_proj_col}${ref_start_row}:${ref_proj_col}${ref_end_row}, 0)), "")'
                
                row_data = {
                    'Date': date_val,
                    'Number': sanitize_excel_str(tx.get('Number', '')),
                    'Account': acc_full,
                    'Branch': sanitize_excel_str(tx.get('Branch', '')),
                    'Project': proj_name,
                    'Analytic Account': sanitize_excel_str(tx.get('Analytic Account', '')),
                    'Budget Code': sanitize_excel_str(tx.get('Budget Code', '')),
                    'Partner': sanitize_excel_str(tx.get('Partner', '')),
                    'Label': sanitize_excel_str(tx.get('Label', '')),
                    'detil project': sanitize_excel_str(tx.get('detil project', '')),
                    'Amount in Currency': float(tx.get('Amount in Currency', 0.0)),
                    'Debit': float(tx.get('Debit', 0.0)),
                    'Credit': float(tx.get('Credit', 0.0)),
                    'Tags': f"={c_deb}{current_row}-{c_cred}{current_row}",
                    'Percentage 2026': f_pct26,
                    'HPP 2026': f"={c_tags}{current_row}*{c_pct26}{current_row}",
                    'Sisa Persediaan 2026': f"={c_tags}{current_row}-{c_hpp26}{current_row}",
                    'no urut': current_no,
                    'Klasifikasi LK': f_klk,
                    'Nama Project New': f"={c_proj}{current_row}",
                    'Klasifikasi Project New': f_kpn
                }
                
                for col_name, c_idx in col_idx_map.items():
                    if col_name in row_data:
                        cell = ws_main.cell(row=current_row, column=c_idx, value=row_data[col_name])
                        if col_name == 'Date' and pd.notna(tx_date):
                            cell.number_format = 'YYYY-MM-DD'
                        elif col_name in acc_cols_list:
                            cell.number_format = ACC_FORMAT
                        elif col_name in pct_cols_list:
                            cell.number_format = '0.00%'
                            
                tags_val = float(tx.get('Debit', 0.0)) - float(tx.get('Credit', 0.0))
                hpp_val = tags_val * pct_2026
                sisa_val = tags_val - hpp_val
                new_rows_for_pivot.append({
                    'Account': acc_full,
                    'Klasifikasi LK': klas_lk if klas_lk else '(blank)',
                    'Klasifikasi Project New': klas_proj_new if klas_proj_new else '(blank)',
                    'HPP 2026': hpp_val,
                    'Sisa Persediaan 2026': sisa_val
                })
                
                current_row += 1
                current_no += 1
                
            final_data_row = current_row - 1
            # Terapkan format Accounting pada seluruh baris historis di Worksheet Utama
            for col_name, c_idx in col_idx_map.items():
                if col_name in acc_cols_list:
                    for r in range(2, last_row + 1):
                        c_cell = ws_main.cell(row=r, column=c_idx)
                        if isinstance(c_cell.value, (int, float)) or (isinstance(c_cell.value, str) and c_cell.value.startswith('=')):
                            c_cell.number_format = ACC_FORMAT
                elif col_name in pct_cols_list:
                    for r in range(2, last_row + 1):
                        c_cell = ws_main.cell(row=r, column=c_idx)
                        if isinstance(c_cell.value, (int, float)) or (isinstance(c_cell.value, str) and c_cell.value.startswith('=')):
                            c_cell.number_format = '0.00%'

            # Kumpulkan kategori unik tahun 2026 untuk membangun struktur baris Pivot
            hist_2026_records = []
            date_col = col_idx_map.get('Date', 1)
            acc_col = col_idx_map.get('Account', 3)
            hpp26_col = col_idx_map.get('HPP 2026', 23)
            sisa26_col = col_idx_map.get('Sisa Persediaan 2026', 24)
            klk_col = col_idx_map.get('Klasifikasi LK', 26)
            kpn_col = col_idx_map.get('Klasifikasi Project New', 28)
            
            for r in range(2, last_row + 1):
                d_val = ws_val.cell(row=r, column=date_col).value
                d_parsed = pd.to_datetime(d_val, errors='coerce')
                h_val = ws_val.cell(row=r, column=hpp26_col).value
                s_val = ws_val.cell(row=r, column=sisa26_col).value
                
                is_2026 = (pd.notna(d_parsed) and d_parsed.year == 2026)
                if is_2026 or (isinstance(h_val, (int, float)) and h_val != 0) or (isinstance(s_val, (int, float)) and s_val != 0):
                    hist_2026_records.append({
                        'Account': sanitize_excel_str(ws_val.cell(row=r, column=acc_col).value or ''),
                        'Klasifikasi LK': sanitize_excel_str(ws_val.cell(row=r, column=klk_col).value or '') or '(blank)',
                        'Klasifikasi Project New': sanitize_excel_str(ws_val.cell(row=r, column=kpn_col).value or '') or '(blank)',
                        'HPP 2026': float(h_val) if isinstance(h_val, (int, float)) else 0.0,
                        'Sisa Persediaan 2026': float(s_val) if isinstance(s_val, (int, float)) else 0.0
                    })
                    
            wb_val.close()
            df_pivot_src = pd.DataFrame(hist_2026_records + new_rows_for_pivot)
            
            # 4. Buat Sheet Checking_TB_vs_GL (Format Vertikal + Format Accounting)
            if 'Checking_TB_vs_GL' in wb.sheetnames:
                del wb['Checking_TB_vs_GL']
            ws_chk = wb.create_sheet(title='Checking_TB_vs_GL')
            
            ws_chk.append(ordered_cols)
            for cell in ws_chk:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                
            for _, r_chk in df_check.iterrows():
                row_idx = ws_chk.max_row + 1
                ws_chk.append(list(r_chk.values))
                if r_chk['Sumber'] == 'Selisih (GL - TB)':
                    fill_color = "E2EFDA" if r_chk['Nama Akun'] == 'BALANCE' else "FCE4D6"
                    for cell in ws_chk[row_idx]:
                        cell.font = Font(bold=True)
                        cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                        
            for row in ws_chk.iter_rows(min_row=2):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = ACC_FORMAT
                        
            # 5. Buat Sheet Pivot_HPP_2026 dengan Rumus Aktif SUMIFS & SUM
            if 'Pivot_HPP_2026' in wb.sheetnames:
                del wb['Pivot_HPP_2026']
            ws_piv = wb.create_sheet(title='Pivot_HPP_2026')
            
            ms = f"'{main_sheet_name}'"
            rng_hpp = f"{ms}!${c_hpp26}$2:${c_hpp26}${final_data_row}"
            rng_sisa = f"{ms}!${c_sisa26}$2:${c_sisa26}${final_data_row}"
            rng_klk = f"{ms}!${c_klk}$2:${c_klk}${final_data_row}"
            rng_kpn = f"{ms}!${c_kpn}$2:${c_kpn}${final_data_row}"
            rng_acc = f"{ms}!${c_acc}$2:${c_acc}${final_data_row}"
            
            # --- TABEL PIVOT 1: Klasifikasi LK (Menggunakan Rumus SUMIFS) ---
            ws_piv.append(["RINGKASAN BERDASARKAN KLASIFIKASI LK (2026)"])
            ws_piv["A1"].font = Font(bold=True, size=12)
            ws_piv.append(["Row Labels", "Sum of HPP 2026", "Sum of Sisa Persediaan 2026"])
            for cell in ws_piv:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                
            p1_start_row = 3
            unique_klk = sorted(df_pivot_src['Klasifikasi LK'].unique().tolist())
            for klk_val in unique_klk:
                r_curr = ws_piv.max_row + 1
                if klk_val == '(blank)':
                    f_hpp = f'=SUMIFS({rng_hpp}, {rng_klk}, "")'
                    f_sisa = f'=SUMIFS({rng_sisa}, {rng_klk}, "")'
                else:
                    f_hpp = f'=SUMIFS({rng_hpp}, {rng_klk}, A{r_curr})'
                    f_sisa = f'=SUMIFS({rng_sisa}, {rng_klk}, A{r_curr})'
                ws_piv.append([klk_val, f_hpp, f_sisa])
                
            p1_end_row = ws_piv.max_row
            gt1_row = p1_end_row + 1
            ws_piv.append([
                "Grand Total",
                f"=SUM(B{p1_start_row}:B{p1_end_row})",
                f"=SUM(C{p1_start_row}:C{p1_end_row})"
            ])
            for cell in ws_piv[gt1_row]:
                cell.font = Font(bold=True)
                
            ws_piv.append([])
            ws_piv.append([])
            
            # --- TABEL PIVOT 2: Rincian Klasifikasi Project New & Account (Menggunakan Rumus SUMIFS) ---
            start_p2 = ws_piv.max_row + 1
            ws_piv.append(["RINCIAN BERDASARKAN KLASIFIKASI PROJECT NEW & ACCOUNT (2026)"])
            ws_piv.cell(row=start_p2, column=1).font = Font(bold=True, size=12)
            
            hdr_p2 = start_p2 + 1
            ws_piv.append(["Klasifikasi Project New", "Account", "Sum of HPP 2026", "Sum of Sisa Persediaan 2026"])
            for cell in ws_piv[hdr_p2]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                
            piv2_grp = df_pivot_src.groupby(['Klasifikasi Project New', 'Account'], as_index=False)[['HPP 2026', 'Sisa Persediaan 2026']].sum()
            subtotal_rows = []
            
            for proj_cat, group in piv2_grp.groupby('Klasifikasi Project New'):
                r_parent = ws_piv.max_row + 1
                subtotal_rows.append(r_parent)
                first_child = r_parent + 1
                last_child = r_parent + len(group)
                
                # Baris Induk (Subtotal per Klasifikasi Project New menggunakan SUM dari baris anaknya)
                ws_piv.append([
                    proj_cat,
                    "",
                    f"=SUM(C{first_child}:C{last_child})",
                    f"=SUM(D{first_child}:D{last_child})"
                ])
                for cell in ws_piv[r_parent]:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                    
                # Baris Anak (Rincian per Account menggunakan SUMIFS 2 kriteria)
                for _, item in group.iterrows():
                    r_child = ws_piv.max_row + 1
                    acc_val = item['Account']
                    crit_proj = '""' if proj_cat == '(blank)' else f'$A${r_parent}'
                    crit_acc = '""' if acc_val == '' else f'B{r_child}'
                    
                    f_hpp_c = f'=SUMIFS({rng_hpp}, {rng_kpn}, {crit_proj}, {rng_acc}, {crit_acc})'
                    f_sisa_c = f'=SUMIFS({rng_sisa}, {rng_kpn}, {crit_proj}, {rng_acc}, {crit_acc})'
                    ws_piv.append(["", acc_val, f_hpp_c, f_sisa_c])
                    
            gt2_row = ws_piv.max_row + 1
            if subtotal_rows:
                sum_c_expr = "+".join([f"C{r}" for r in subtotal_rows])
                sum_d_expr = "+".join([f"D{r}" for r in subtotal_rows])
                ws_piv.append(["Grand Total", "", f"={sum_c_expr}", f"={sum_d_expr}"])
            else:
                ws_piv.append(["Grand Total", "", 0, 0])
            for cell in ws_piv[gt2_row]:
                cell.font = Font(bold=True)
                
            # Terapkan Format Accounting ke seluruh angka & rumus di Sheet Pivot_HPP_2026
            for row in ws_piv.iter_rows(min_row=3):
                for cell in row:
                    if cell.column >= 2 and cell.value is not None and not (cell.column == 2 and row[0].row > hdr_p2):
                        cell.number_format = ACC_FORMAT
                        
            # Bersihkan objek Pivot Table bawaan agar tidak memicu pesan recovery di Excel
            for sheet in wb.worksheets:
                sheet._pivots = []
                
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success("✅ Proses selesai! Seluruh rumus aktif, sheet Ref_Klasifikasi, dan format Accounting telah diterapkan.")
            st.download_button(
                label="📥 Download File Excel HPP Lengkap (Rumus Aktif & Accounting)",
                data=output,
                file_name="HPP_Update_Oktober_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
