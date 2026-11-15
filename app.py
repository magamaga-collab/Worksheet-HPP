import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill
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

def clean_account_name(val):
    """Membersihkan nama akun dan menghapus akhiran jumlah baris seperti (331)"""
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    return re.sub(r'\s*\(\d+\)$', '', s).strip()

def split_account_code_desc(acc_str):
    """Memisahkan 5 digit nomor akun di awal string dan deskripsi akunnya"""
    if pd.isna(acc_str) or acc_str is None:
        return "", ""
    s = clean_account_name(acc_str)
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
    """Membaca file klasifikasi untuk mapping Project dan Account"""
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
            p_name = str(r['Nama Project']).strip()
            k_proj = r.get('Klasifikasi Project New', r.get('Klasifikasi Project', ''))
            pct_val = r.get('Percentage 2026', r.get('Percentage', 0))
            try:
                pct_val = float(pct_val) if pd.notna(pct_val) else 0.0
            except:
                pct_val = 0.0
            proj_map[p_name] = str(k_proj).strip() if pd.notna(k_proj) else ""
            pct_map[p_name] = pct_val

    acc_map = {}
    if 'Account' in df_k.columns and 'Klasifikasi LK' in df_k.columns:
        for _, r in df_k.dropna(subset=['Account']).iterrows():
            acc_full = str(r['Account']).strip()
            acc_code = extract_account_code(acc_full)
            k_lk = str(r['Klasifikasi LK']).strip() if pd.notna(r['Klasifikasi LK']) else ""
            acc_map[acc_full] = k_lk
            if acc_code:
                acc_map[acc_code] = k_lk
                
    return proj_map, pct_map, acc_map
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
                
        is_summary_row = df['Date'].isna() & df['Number'].isna()
        df_valid = df[~is_summary_row].copy()
        
        if 'Account' in df.columns:
            df['Account_Filled'] = df['Account'].ffill()
            df_valid['Account'] = df.loc[df_valid.index, 'Account_Filled']
            
        cleaned_accs = [clean_account_name(x) for x in df_valid['Account'].tolist()]
        split_pairs = [split_account_code_desc(x) for x in cleaned_accs]
        
        df_valid['Account'] = cleaned_accs
        df_valid['Account_Code'] = [p[0] for p in split_pairs]
        df_valid['Account_Desc'] = [p[1] for p in split_pairs]
        df_valid['Tags_Calc'] = df_valid['Debit'] - df_valid['Credit']
        
        all_tx.append(df_valid)
        
        if not df_valid.empty:
            grp = df_valid.groupby(['Account_Code', 'Account_Desc'], as_index=False).agg(
                GL_Debit=('Debit', 'sum'),
                GL_Credit=('Credit', 'sum'),
                GL_Net_Mutasi=('Tags_Calc', 'sum')
            )
            gl_summary_list.append(grp)
            
    df_gl_all = pd.concat(all_tx, ignore_index=True) if all_tx else pd.DataFrame()
    df_gl_sum = pd.concat(gl_summary_list, ignore_index=True) if gl_summary_list else pd.DataFrame()
    if not df_gl_sum.empty:
        df_gl_sum = df_gl_sum.groupby(['Account_Code', 'Account_Desc'], as_index=False).sum()
        
    return df_gl_all, df_gl_sum

def parse_tb_file(tb_file):
    """Membaca Kolom A (index 0) dari file TB yang diawali 5 digit kode akun dan mengambil Debit & Credit Bulan Berjalan"""
    df_raw = pd.read_excel(tb_file, header=None)
    
    # Cari kolom 'Bulan Berjalan' atau nama bulan di baris header atas (misal Oct 2026)
    bb_col_idx = None
    for r_idx in range(min(6, len(df_raw))):
        for c_idx, val in enumerate(df_raw.iloc[r_idx]):
            if pd.notna(val):
                val_low = str(val).lower()
                if any(k in val_low for k in ['bulan berjalan', 'jan', 'feb', 'mar', 'apr', 'mei', 'may', 'jun', 'jul', 'agu', 'aug', 'sep', 'okt', 'oct', 'nov', 'des', 'dec']):
                    if c_idx >= 1:
                        bb_col_idx = c_idx
                        break
        if bb_col_idx is not None:
            break
            
    # Default standar TB: Kolom A (0) = Akun, Kolom B-C (1-2) = Initial, Kolom D-E (3-4) = Bulan Berjalan
    if bb_col_idx is None:
        bb_col_idx = 3
        
    tb_records = []
    for _, row in df_raw.iterrows():
        acc_val = row.iloc[0]
        if pd.isna(acc_val):
            continue
        acc_str = str(acc_val).strip()
        
        # Deteksi baris yang diawali dengan 5 digit kode akun di Kolom A
        match_5digit = re.match(r'^(\d{5})\s*[-–]?\s*(.*)$', acc_str)
        if not match_5digit:
            continue
            
        acc_code = match_5digit.group(1).strip()
        acc_desc = match_5digit.group(2).strip()
        
        debit_bb = pd.to_numeric(row.iloc[bb_col_idx], errors='coerce') if len(row) > bb_col_idx else 0.0
        credit_bb = pd.to_numeric(row.iloc[bb_col_idx + 1], errors='coerce') if len(row) > bb_col_idx + 1 else 0.0
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
        with st.spinner("Sedang memproses file Master, GL, TB, dan membuat Pivot..."):
            proj_map, pct_map, acc_map = load_klasifikasi(klasifikasi_file)
            df_gl_all, df_gl_sum = parse_gl_files(gl_files)
            df_tb = parse_tb_file(tb_file)
            
            df_check = pd.merge(df_gl_sum, df_tb, on='Account_Code', how='left')
            for col in ['TB_Debit_Berjalan', 'TB_Credit_Berjalan', 'TB_Net_Mutasi']:
                df_check[col] = df_check[col].fillna(0.0)
            df_check['TB_Account_Desc'] = df_check['TB_Account_Desc'].fillna('(Tidak ditemukan di TB)')
                
            df_check['Selisih_Debit'] = (df_check['GL_Debit'] - df_check['TB_Debit_Berjalan']).round(2)
            df_check['Selisih_Credit'] = (df_check['GL_Credit'] - df_check['TB_Credit_Berjalan']).round(2)
            df_check['Selisih_Net'] = (df_check['GL_Net_Mutasi'] - df_check['TB_Net_Mutasi']).round(2)
            
            is_balanced = (
                (df_check['Selisih_Debit'].abs() < 1.0) &
                (df_check['Selisih_Credit'].abs() < 1.0) &
                (df_check['Selisih_Net'].abs() < 1.0)
            )
            df_check['Status'] = np.where(is_balanced, 'BALANCE', 'SELISIH')
            
            ordered_cols = [
                'Account_Code', 'Account_Desc', 'TB_Account_Desc',
                'GL_Debit', 'TB_Debit_Berjalan', 'Selisih_Debit',
                'GL_Credit', 'TB_Credit_Berjalan', 'Selisih_Credit',
                'GL_Net_Mutasi', 'TB_Net_Mutasi', 'Selisih_Net',
                'Status'
            ]
            df_check = df_check[ordered_cols]
            
            st.subheader("🔍 Hasil Checking Mutasi GL vs TB (Rincian Debit & Credit)")
            st.dataframe(df_check, use_container_width=True)
            master_bytes = master_file.getvalue()
            wb_val = openpyxl.load_workbook(BytesIO(master_bytes), data_only=True)
            ws_val = wb_val[wb_val.sheetnames[0]]
            
            wb = openpyxl.load_workbook(BytesIO(master_bytes))
            main_sheet_name = wb.sheetnames[0]
            ws_main = wb[main_sheet_name]
            
            headers = [cell.value for cell in ws_main[1]]
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
                    
            current_row = last_row + 1
            current_no = last_no_urut + 1
            new_rows_for_pivot = []
            
            for _, tx in df_gl_all.iterrows():
                acc_full = str(tx.get('Account', '')).strip()
                acc_code = str(tx.get('Account_Code', '')).strip()
                proj_name = str(tx.get('Project', '')).strip() if pd.notna(tx.get('Project')) else ""
                
                pct_2026 = pct_map.get(proj_name, 0.0)
                klas_lk = acc_map.get(acc_full, acc_map.get(acc_code, ""))
                klas_proj_new = proj_map.get(proj_name, "")
                
                tx_date = pd.to_datetime(tx.get('Date'), errors='coerce')
                date_val = tx_date.to_pydatetime() if pd.notna(tx_date) else tx.get('Date')
                
                row_data = {
                    'Date': date_val,
                    'Number': tx.get('Number', ''),
                    'Account': acc_full,
                    'Branch': tx.get('Branch', ''),
                    'Project': proj_name,
                    'Analytic Account': tx.get('Analytic Account', ''),
                    'Budget Code': tx.get('Budget Code', ''),
                    'Partner': tx.get('Partner', ''),
                    'Label': tx.get('Label', ''),
                    'detil project': tx.get('detil project', ''),
                    'Amount in Currency': float(tx.get('Amount in Currency', 0.0)),
                    'Debit': float(tx.get('Debit', 0.0)),
                    'Credit': float(tx.get('Credit', 0.0)),
                    'Tags': f"=L{current_row}-M{current_row}",
                    'Percentage 2026': pct_2026,
                    'HPP 2026': f"=N{current_row}*V{current_row}",
                    'Sisa Persediaan 2026': f"=N{current_row}-W{current_row}",
                    'no urut': current_no,
                    'Klasifikasi LK': klas_lk,
                    'Nama Project New': proj_name,
                    'Klasifikasi Project New': klas_proj_new
                }
                
                for col_name, c_idx in col_idx_map.items():
                    if col_name in row_data:
                        cell = ws_main.cell(row=current_row, column=c_idx, value=row_data[col_name])
                        if col_name == 'Date' and pd.notna(tx_date):
                            cell.number_format = 'YYYY-MM-DD'
                            
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
                        'Account': str(ws_val.cell(row=r, column=acc_col).value or ''),
                        'Klasifikasi LK': str(ws_val.cell(row=r, column=klk_col).value or '(blank)'),
                        'Klasifikasi Project New': str(ws_val.cell(row=r, column=kpn_col).value or '(blank)'),
                        'HPP 2026': float(h_val) if isinstance(h_val, (int, float)) else 0.0,
                        'Sisa Persediaan 2026': float(s_val) if isinstance(s_val, (int, float)) else 0.0
                    })
                    
            wb_val.close()
            df_pivot_src = pd.DataFrame(hist_2026_records + new_rows_for_pivot)
            
            if 'Checking_TB_vs_GL' in wb.sheetnames:
                del wb['Checking_TB_vs_GL']
            ws_chk = wb.create_sheet(title='Checking_TB_vs_GL')
            
            ws_chk.append(ordered_cols)
            for cell in ws_chk[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                
            for _, r_chk in df_check.iterrows():
                ws_chk.append(list(r_chk.values))
                
            if 'Pivot_HPP_2026' in wb.sheetnames:
                del wb['Pivot_HPP_2026']
            ws_piv = wb.create_sheet(title='Pivot_HPP_2026')
            
            ws_piv.append(["RINGKASAN BERDASARKAN KLASIFIKASI LK (2026)"])
            ws_piv["A1"].font = Font(bold=True, size=12)
            ws_piv.append(["Row Labels", "Sum of HPP 2026", "Sum of Sisa Persediaan 2026"])
            for cell in ws_piv[2]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                
            piv1 = df_pivot_src.groupby('Klasifikasi LK', as_index=False)[['HPP 2026', 'Sisa Persediaan 2026']].sum()
            for _, r_p1 in piv1.iterrows():
                ws_piv.append([r_p1['Klasifikasi LK'], r_p1['HPP 2026'], r_p1['Sisa Persediaan 2026']])
            ws_piv.append(["Grand Total", piv1['HPP 2026'].sum(), piv1['Sisa Persediaan 2026'].sum()])
            for cell in ws_piv[ws_piv.max_row]:
                cell.font = Font(bold=True)
                
            ws_piv.append([])
            ws_piv.append([])
            
            start_p2 = ws_piv.max_row + 1
            ws_piv.append(["RINCIAN BERDASARKAN KLASIFIKASI PROJECT NEW & ACCOUNT (2026)"])
            ws_piv.cell(row=start_p2, column=1).font = Font(bold=True, size=12)
            
            hdr_p2 = start_p2 + 1
            ws_piv.append(["Row Labels (Project / Account)", "Sum of HPP 2026", "Sum of Sisa Persediaan 2026"])
            for cell in ws_piv[hdr_p2]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                
            piv2_grp = df_pivot_src.groupby(['Klasifikasi Project New', 'Account'], as_index=False)[['HPP 2026', 'Sisa Persediaan 2026']].sum()
            for proj_cat, group in piv2_grp.groupby('Klasifikasi Project New'):
                r_idx = ws_piv.max_row + 1
                ws_piv.append([proj_cat, group['HPP 2026'].sum(), group['Sisa Persediaan 2026'].sum()])
                for cell in ws_piv[r_idx]:
                    cell.font = Font(bold=True)
                for _, item in group.iterrows():
                    ws_piv.append([f"    {item['Account']}", item['HPP 2026'], item['Sisa Persediaan 2026']])
                    
            gt_row = ws_piv.max_row + 1
            ws_piv.append(["Grand Total", piv2_grp['HPP 2026'].sum(), piv2_grp['Sisa Persediaan 2026'].sum()])
            for cell in ws_piv[gt_row]:
                cell.font = Font(bold=True)
                
            for sheet in [ws_chk, ws_piv]:
                for row in sheet.iter_rows(min_row=2):
                    for cell in row:
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = '#,##0.00'
                            
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success("✅ Proses selesai! Seluruh data Oktober telah ditambahkan dan direkonsiliasi.")
            st.download_button(
                label="📥 Download File Excel HPP Lengkap (Update Oktober)",
                data=output,
                file_name="HPP_Update_Oktober_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
