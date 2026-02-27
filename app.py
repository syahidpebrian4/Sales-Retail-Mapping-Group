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
    "6001": "Pasar Rebo", "6003": "Kelapa Gading", "6006": "Ciputat",
    "6007": "Alam Sutera", "6010": "Medan", "6014": "Palembang",
    "6015": "Pekanbaru", "6021": "Jatake", "6022": "Serang",
    "6029": "Batam", "6031": "Lampung", "6039": "Serpong",
    "6004": "Meruya", "6005": "Bandung", "6008": "Cibitung",
    "6018": "Bekasi", "6023": "Cikarang", "6024": "Cirebon",
    "6026": "Bogor", "6027": "Tasikmalaya", "6030": "Pakansari",
    "6034": "Kerawang", "6036": "Cimahi", "6038": "Tegal",
    "6002": "Sidoarjo", "6009": "Denpasar", "6011": "Semarang",
    "6013": "Makasar", "6016": "Yogyakarta", "6017": "Banjarmasin",
    "6019": "Solo", "6020": "Balikpapan", "6028": "Mastrip",
    "6032": "Samarinda", "6033": "Manado", "6037": "Mataram"
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
        # Tulis data mulai baris ke-3 (index 2) tanpa header default pandas
        df.to_excel(writer, sheet_name='Sales Report', startrow=2, header=False)
        
        workbook  = writer.book
        worksheet = writer.sheets['Sales Report']
        
        # --- FORMATS ---
        header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D3D3D3', 'border': 1})
        num_fmt = workbook.add_format({'num_format': '#,##0;[Red]▼#,##0', 'border': 1, 'align': 'right'})
        pct_fmt = workbook.add_format({'num_format': '0.0%;[Red]▼0.0%', 'border': 1, 'align': 'right'})
        text_fmt = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})
        code_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})

        # --- FIX BORDER & COLUMN WIDTH ---
        worksheet.set_column(0, 0, 12, code_fmt) # Store Code
        worksheet.set_column(1, 1, 25, text_fmt) # Store Name (Fixed border)

        # Merge Header Manual
        worksheet.merge_range('A1:A2', 'Store Code', header_fmt)
        worksheet.merge_range('B1:B2', 'Store Name', header_fmt)

        current_col = 2
        categories = df.columns.get_level_values(0).unique()
        
        for cat in categories:
            if cat == 'Store Name': continue
            
            sub_cols = df[cat].columns
            sub_cols_count = len(sub_cols)
            
            if sub_cols_count > 1:
                worksheet.merge_range(0, current_col, 0, current_col + sub_cols_count - 1, str(cat), header_fmt)
            else:
                worksheet.write(0, current_col, str(cat), header_fmt)
            
            for i, met in enumerate(sub_cols):
                col_idx = current_col + i
                worksheet.write(1, col_idx, str(met), header_fmt)
                
                # Apply numeric formats to data cells in this column
                if "YEAR" in str(met).upper():
                    worksheet.set_column(col_idx, col_idx, 15, num_fmt)
                else:
                    worksheet.set_column(col_idx, col_idx, 12, pct_fmt)
            
            current_col += sub_cols_count
            
        worksheet.hide_gridlines(2)
        
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
            
        res = {'Store Name': STORE_MAP.get(store, "Unknown")}
        store_code_val = store
        
        df_total = df_match[df_match['Group'].isin(['SMALL','MEDIUM','BIG'])]
        col_ty = f'{suffix}_TY'
        col_ly = f'{suffix}_LY'
        
        t_ty = df_total[col_ty].sum() if col_ty in df_total.columns else 0
        t_ly = df_total[col_ly].sum() if col_ly in df_total.columns else 0
        
        res[('TOTAL SALES', 'THIS YEAR')] = t_ty
        res[('TOTAL SALES', 'LAST YEAR')] = t_ly
        res[('TOTAL SALES', 'GROWTH (%)')] = ((t_ty - t_ly)/t_ly) if t_ly != 0 else 0
        
        for g in ['SMALL', 'MEDIUM', 'BIG']:
            if g in selected_groups:
                df_g = df_match[df_match['Group'] == g]
                g_ty = df_g[col_ty].sum() if col_ty in df_g.columns else 0
                g_ly = df_g[col_ly].sum() if col_ly in df_g.columns else 0
                
                res[(g, 'THIS YEAR')] = g_ty
                res[(g, 'LAST YEAR')] = g_ly
                res[(g, 'GROWTH (%)')] = ((g_ty - g_ly)/g_ly) if g_ly != 0 else 0
                res[(g, 'CONT (%)')] = (g_ty / t_ty) if t_ty != 0 else 0
        
        res['_str_code'] = store_code_val
        final_rows.append(res)

    if final_rows:
        res_df = pd.DataFrame(final_rows)
        res_df['_str_code'] = pd.to_numeric(res_df['_str_code'])
        res_df = res_df.sort_values('_str_code').set_index('_str_code')
        res_df.index.name = 'Store Code'
        
        # Merge Header Streamlit
        new_columns = []
        for col in res_df.columns:
            if col == 'Store Name': new_columns.append(('Store Name', ''))
            elif isinstance(col, tuple): new_columns.append(col)
        res_df.columns = pd.MultiIndex.from_tuples(new_columns)
        
        st.markdown(f"### 📋 {selected_item} Report ({period})")
        
        # Styling Streamlit View
        format_dict = {}
        for col in res_df.columns:
            if col[1] == "": continue
            if "YEAR" in col[1].upper(): format_dict[col] = "{:,.0f}"
            else: format_dict[col] = "{:.1f}%"
        
        # Konversi ke Persen untuk tampilan Streamlit (dikali 100 karena format_dict {:.1f}%)
        display_df = res_df.copy()
        for col in display_df.columns:
            if "%" in col[1]:
                display_df[col] = display_df[col] * 100

        st.dataframe(
            display_df.style.format(format_dict).applymap(
                lambda x: 'color: red' if isinstance(x, (int, float)) and x < 0 else None
            ), use_container_width=True, height=500
        )
        
        # Download Button
        excel_bin = to_excel_with_style(res_df)
        st.download_button(
            label="📥 DOWNLOAD EXCEL", data=excel_bin,
            file_name=f"{selected_item}_{period}.xlsx", use_container_width=True
        )
    else:
        st.info("No data found for selected filters.")
