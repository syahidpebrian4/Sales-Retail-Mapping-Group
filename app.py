import streamlit as st
import pandas as pd
import base64
import os
import io

# ================= CONFIG & UI STYLE =================
st.set_page_config(page_title="Sales Report Per Mapping Group", layout="wide", initial_sidebar_state="expanded")

def get_base64_image(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return None

# --- CSS CUSTOM ---
logo_b64 = get_base64_image("lotte_logo.png")
st.markdown(f"""
    <style>
        .custom-header {{
            position: fixed; top: 0; left: 0; width: 100%; height: 90px;
            background-color: white; display: flex; align-items: center;
            padding: 0 30px; border-bottom: 3px solid #eeeeee; z-index: 999999;
        }}
        .header-logo {{ height: 55px; margin-right: 25px; }}
        .header-title {{
            font-size: 36px; font-weight: 900;
            font-family: 'Arial Black', sans-serif; color: #333333; margin: 0;
        }}
        [data-testid="stSidebar"] {{
            background-color: #FF0000 !important; margin-top: 90px !important;
            min-width: 320px !important;
        }}
        [data-testid="stSidebar"] .stMarkdown p, 
        [data-testid="stSidebar"] label {{
            color: white !important; font-weight: bold !important;
        }}
        .main .block-container {{ padding-top: 130px !important; }}
        header {{ visibility: hidden; }}
    </style>
    <div class="custom-header">
        <img src="data:image/png;base64,{logo_b64 if logo_b64 else ''}" class="header-logo">
        <h1 class="header-title">Retail Report Mapping Group</h1>
    </div>
""", unsafe_allow_html=True)

# ================= DATA MAPPING STORE =================
STORE_MAP = {
    "6001": "Pasar Rebo", "6002": "Sidoarjo", "6003": "Kelapa Gading", 
    "6004": "Meruya", "6005": "Bandung", "6006": "Ciputat",
    "6007": "Alam Sutera", "6008": "Cibitung", "6009": "Denpasar", 
    "6010": "Medan", "6011": "Semarang", "6013": "Makasar", 
    "6014": "Palembang", "6015": "Pekanbaru", "6016": "Yogyakarta", 
    "6017": "Banjarmasin", "6018": "Bekasi", "6019": "Solo", 
    "6020": "Balikpapan", "6021": "Jatake", "6022": "Serang", 
    "6023": "Cikarang", "6024": "Cirebon", "6026": "Bogor", 
    "6027": "Tasikmalaya", "6028": "Mastrip", "6029": "Batam", 
    "6030": "Pakansari", "6031": "Lampung", "6032": "Samarinda", 
    "6033": "Manado", "6034": "Kerawang", "6036": "Cimahi", 
    "6037": "Mataram", "6038": "Tegal", "6039": "Serpong"
}

# ================= LOGIKA DATA =================

def load_and_clean_data(uploaded_file):
    df_raw = pd.read_excel(uploaded_file, header=None)
    rows = []
    for i in range(len(df_raw)):
        current_row = df_raw.iloc[i].values.tolist()
        if pd.notnull(current_row[0]) and str(current_row[0]).strip().isdigit():
            rows.append(current_row)
        elif pd.isnull(current_row[0]) and len(current_row) > 4 and pd.notnull(current_row[4]):
            if rows:
                rows[-1][4] = f"{rows[-1][4]} {current_row[4]}".strip()
    
    df = pd.DataFrame(rows)
    df = df[[0, 2, 4, 5, 6, 9, 10, 13, 14]]
    df.columns = ['Str_cd', 'Group', 'Item', 'D_TY', 'D_LY', 'M_TY', 'M_LY', 'Y_TY', 'Y_LY']
    
    df['Str_cd'] = df['Str_cd'].astype(int).astype(str)
    df['Group'] = df['Group'].astype(str).str.strip().str.upper()
    df['Item'] = df['Item'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
    
    for col in ['D_TY', 'D_LY', 'M_TY', 'M_LY', 'Y_TY', 'Y_LY']:
        df[col] = df[col].astype(str).str.replace(',', '').str.replace(' ', '')
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df

def to_excel_with_style(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Bongkar index supaya Store Code & Store Name jadi kolom normal di Excel
        df_export = df.reset_index()
        df_export.to_excel(writer, sheet_name='Sales Report', header=False, startrow=2, index=False)
        
        workbook  = writer.book
        worksheet = writer.sheets['Sales Report']
        
        # Border Format (PENTING BIAR ADA GARISNYA)
        header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D3D3D3', 'border': 1})
        num_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1, 'align': 'right'})
        pct_fmt = workbook.add_format({'num_format': '0.0%', 'border': 1, 'align': 'right'})
        txt_fmt = workbook.add_format({'border': 1, 'align': 'left'}) # Format garis untuk nama toko

        # Merge Header
        worksheet.merge_range('A1:A2', 'Store Code', header_fmt)
        worksheet.merge_range('B1:B2', 'Store Name', header_fmt)
        worksheet.set_column('A:A', 12, txt_fmt) # Kasih border ke kolom A
        worksheet.set_column('B:B', 25, txt_fmt) # Kasih border ke kolom B

        current_col = 2
        categories = []
        for cat, sub in df.columns:
            if cat not in categories: categories.append(cat)
        
        for cat in categories:
            sub_cols = [c[1] for c in df.columns if c[0] == cat]
            count = len(sub_cols)
            if count > 1:
                worksheet.merge_range(0, current_col, 0, current_col + count - 1, str(cat), header_fmt)
            else:
                worksheet.write(0, current_col, str(cat), header_fmt)
            
            for i, sub in enumerate(sub_cols):
                col_idx = current_col + i
                worksheet.write(1, col_idx, str(sub), header_fmt)
                if "YEAR" in str(sub).upper(): worksheet.set_column(col_idx, col_idx, 15, num_fmt)
                else: worksheet.set_column(col_idx, col_idx, 12, pct_fmt)
            current_col += count
            
    return output.getvalue()

# ================= UI & FILTER =================

uploaded_file = st.file_uploader("📂 UPLOAD EXCEL", type=["xlsx"])

if uploaded_file:
    df = load_and_clean_data(uploaded_file)
    
    with st.sidebar:
        st.markdown("---")
        all_items = sorted(df['Item'].unique())
        all_stores = sorted(df['Str_cd'].unique(), key=int)
        selected_stores = st.multiselect("SELECT STORES", all_stores, default=all_stores)
        selected_groups = st.multiselect("SELECT MAPPING GROUP", ['SMALL', 'MEDIUM', 'BIG'], default=['SMALL', 'MEDIUM', 'BIG'])
        selected_item = st.selectbox("SELECT ITEM", all_items, index=0)
        period = st.selectbox("SELECT PERIOD", ["Daily", "MTD", "YTD"])
        st.markdown("---")

    p_map = {"Daily": "D", "MTD": "M", "YTD": "Y"}
    suffix = p_map[period]
    
    final_rows = []
    for store in selected_stores:
        df_match = df[(df['Str_cd'] == store) & (df['Item'] == selected_item)]
        if df_match.empty: continue
            
        # Store Data
        row = {
            'Store Code': int(store),
            'Store Name': STORE_MAP.get(store, "Unknown")
        }
        
        df_total = df_match[df_match['Group'].isin(['SMALL','MEDIUM','BIG'])]
        col_ty, col_ly = f'{suffix}_TY', f'{suffix}_LY'
        t_ty, t_ly = df_total[col_ty].sum(), df_total[col_ly].sum()
        
        row[('TOTAL SALES', 'THIS YEAR')] = t_ty
        row[('TOTAL SALES', 'LAST YEAR')] = t_ly
        row[('TOTAL SALES', 'GROWTH (%)')] = ((t_ty - t_ly)/t_ly) if t_ly != 0 else 0
        
        for g in ['SMALL', 'MEDIUM', 'BIG']:
            if g in selected_groups:
                df_g = df_match[df_match['Group'] == g]
                g_ty, g_ly = df_g[col_ty].sum(), df_g[col_ly].sum()
                row[(g, 'THIS YEAR')] = g_ty
                row[(g, 'LAST YEAR')] = g_ly
                row[(g, 'GROWTH (%)')] = ((g_ty - g_ly)/g_ly) if g_ly != 0 else 0
                row[(g, 'CONT (%)')] = (g_ty / t_ty) if t_ty != 0 else 0
        
        final_rows.append(row)

    if final_rows:
        res_df = pd.DataFrame(final_rows)
        # --- SOLUSI GARIS TABEL ---
        # Jadikan kodenya dan namanya sebagai index gabungan. 
        # Streamlit otomatis kasih garis kolom untuk semua index!
        res_df = res_df.set_index(['Store Code', 'Store Name']).sort_index()
        
        st.markdown(f"### 📋 {selected_item} Report ({period})")
        
        fmt = {col: ("{:,.0f}" if "YEAR" in col[1].upper() else "{:.1%}") for col in res_df.columns}
        
        # Streamlit Styling
        st.dataframe(
            res_df.style.format(fmt).applymap(
                lambda x: 'color: red' if isinstance(x, (int, float)) and x < 0 else None
            ), use_container_width=True, height=550
        )
        
        excel_bin = to_excel_with_style(res_df)
        st.download_button(label="📥 DOWNLOAD EXCEL", data=excel_bin, file_name=f"Report.xlsx", use_container_width=True)
