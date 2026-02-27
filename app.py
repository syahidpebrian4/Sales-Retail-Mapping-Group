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

# --- CSS CUSTOM: LOTTE STYLE ---
logo_b64 = get_base64_image("lotte_logo.png")
st.markdown(f"""
    <style>
        /* Header Putih */
        .custom-header {{
            position: fixed;
            top: 0; left: 0; width: 100%; height: 90px;
            background-color: white;
            display: flex; align-items: center;
            padding: 0 30px; border-bottom: 3px solid #eeeeee;
            z-index: 999999;
        }}
        .header-logo {{ height: 55px; margin-right: 25px; }}
        .header-title {{
            font-size: 36px; font-weight: 900;
            font-family: 'Arial Black', sans-serif; color: #333333; margin: 0;
        }}
        
        /* Sidebar Merah Lotte */
        [data-testid="stSidebar"] {{
            background-color: #FF0000 !important;
            margin-top: 90px !important;
            min-width: 320px !important;
        }}
        [data-testid="stSidebar"] .stMarkdown p, 
        [data-testid="stSidebar"] label {{
            color: white !important; font-weight: bold !important;
            font-size: 1.1rem !important;
        }}
        
        /* Tombol & Input Sidebar */
        div[data-testid="stSelectbox"] > label {{ color: white !important; }}
        
        /* Main Container Adjustment */
        .main .block-container {{ padding-top: 130px !important; }}
        header {{ visibility: hidden; }}
    </style>
    <div class="custom-header">
        <img src="data:image/png;base64,{logo_b64 if logo_b64 else ''}" class="header-logo">
        <h1 class="header-title">Sales Report Mapping Group</h1>
    </div>
""", unsafe_allow_html=True)

# ================= LOGIKA DATA =================

def load_and_clean_data(uploaded_file):
    # Baca mentah
    df_raw = pd.read_excel(uploaded_file, header=None)
    
    # Logika menggabungkan baris yang pecah (seperti /VisitCusts.)
    rows = []
    for i in range(len(df_raw)):
        current_row = df_raw.iloc[i].values.tolist()
        
        # Jika kolom 0 berisi angka (Store ID)
        if pd.notnull(current_row[0]) and str(current_row[0]).strip().isdigit():
            rows.append(current_row)
        # Jika kolom 0 kosong tapi kolom 4 ada isi (kasus Net sale tumpah)
        elif pd.isnull(current_row[0]) and pd.notnull(current_row[4]):
            if rows:
                # Gabungkan teks ke baris sebelumnya di kolom Item (index 4)
                rows[-1][4] = f"{rows[-1][4]} {current_row[4]}".strip()
    
    df = pd.DataFrame(rows)
    
    # Ambil kolom spesifik sesuai struktur file: 0, 2, 4, 5, 6, 9, 10, 13, 14
    df = df[[0, 2, 4, 5, 6, 9, 10, 13, 14]]
    df.columns = ['Str_cd', 'Group', 'Item', 'D_TY', 'D_LY', 'M_TY', 'M_LY', 'Y_TY', 'Y_LY']
    
    # Cleaning Teks
    df['Str_cd'] = df['Str_cd'].astype(int).astype(str)
    df['Group'] = df['Group'].astype(str).str.strip().str.upper()
    df['Item'] = df['Item'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
    
    # Cleaning Angka (Hapus koma jika ada)
    for col in df.columns[3:]:
        df[col] = df[col].astype(str).str.replace(',', '').str.replace(' ', '')
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
    return df

# ================= UI & FILTER =================

uploaded_file = st.file_uploader("📂 UPLOAD EXCEL", type=["xlsx"])

if uploaded_file:
    df = load_and_clean_data(uploaded_file)
    
    with st.sidebar:
        st.markdown("---")
        # Item Filter
        all_items = sorted(df['Item'].unique())
        # Default cari yang mengandung "Net sale"
        def_idx = 0
        for i, v in enumerate(all_items):
            if "NET SALE" in v.upper():
                def_idx = i
                break
        
        # Store Filter
        all_stores = sorted(df['Str_cd'].unique(), key=int)
        selected_stores = st.multiselect("📍 SELECT STORES", all_stores, default=all_stores)
        
        # Group Filter
        selected_groups = st.multiselect("SELECT MAPPING GROUP", ['SMALL', 'MEDIUM', 'BIG'], default=['SMALL', 'MEDIUM', 'BIG'])

        selected_item = st.selectbox("SELECT ITEM", all_items, index=def_idx)
        
        # Period Filter
        period = st.selectbox("SELECT PERIOD", ["Daily", "MTD", "YTD"])
        st.markdown("---")

    # --- PROCESSING SUMMARY ---
    suffix = 'D' if period == "Daily" else period
    final_rows = []
    
    for store in selected_stores:
        df_match = df[(df['Str_cd'] == store) & (df['Item'] == selected_item)]
        if df_match.empty: continue
            
        res = {'Store': store}
        
        # TOTAL BASE (S+M+B)
        df_total = df_match[df_match['Group'].isin(['SMALL','MEDIUM','BIG'])]
        t_ty = df_total[f'{suffix}_TY'].sum()
        t_ly = df_total[f'{suffix}_LY'].sum()
        
        res[('TOTAL SALES', 'THIS YEAR')] = t_ty
        res[('TOTAL SALES', 'LAST YEAR')] = t_ly
        res[('TOTAL SALES', 'GROWTH (%)')] = ((t_ty - t_ly)/t_ly*100) if t_ly != 0 else 0
        
        # PER GROUP
        for g in ['SMALL', 'MEDIUM', 'BIG']:
            df_g = df_match[df_match['Group'] == g]
            g_ty = df_g[f'{suffix}_TY'].sum()
            g_ly = df_g[f'{suffix}_LY'].sum()
            
            if g in selected_groups:
                res[(g, 'THIS YEAR')] = g_ty
                res[(g, 'LAST YEAR')] = g_ly
                res[(g, 'GROWTH (%)')] = ((g_ty - g_ly)/g_ly*100) if g_ly != 0 else 0
                res[(g, 'CONT (%)')] = (g_ty / t_ty * 100) if t_ty != 0 else 0
        
        final_rows.append(res)

    # --- DISPLAY TABLE ---
    if final_rows:
        res_df = pd.DataFrame(final_rows).set_index('Store')
        res_df.columns = pd.MultiIndex.from_tuples(res_df.columns)
        
        st.markdown(f"### 📋 {selected_item} Report ({period})")
        
        # Format angka ribuan & %
        format_dict = {}
        for col in res_df.columns:
            if "YEAR" in col[1]: format_dict[col] = "{:,.0f}"
            else: format_dict[col] = "{:.1f}%"
            
        st.dataframe(
            res_df.style.format(format_dict).applymap(
                lambda x: 'color: red' if isinstance(x, (int, float)) and x < 0 else None
            ),
            use_container_width=True,
            height=500
        )
        
        # Download Button
        csv = res_df.to_csv().encode('utf-8')
        st.download_button("📥 DOWNLOAD REPORT CSV", csv, "sales_report.xlsx", use_container_width=True)
    else:
        st.info("No data found for selected filters.")
