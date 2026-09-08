import streamlit as st
import pandas as pd
import pypdf
import re
import io

st.set_page_config(page_title="Validasi Dokumen Ekspor/Impor", layout="wide")

st.title("📋 App Validasi Dokumen: Invoice, Packing List & CEISA")
st.write("Unggah file Excel, CSV, atau PDF untuk ketiga dokumen, lalu tekan tombol **Jalankan Validasi**.")

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. File Invoice")
    file_inv = st.file_uploader("Upload Invoice", type=["xlsx", "xls", "csv", "pdf"], key="inv")

with col2:
    st.subheader("2. File Packing List")
    file_pl = st.file_uploader("Upload Packing List", type=["xlsx", "xls", "csv", "pdf"], key="pl")

with col3:
    st.subheader("3. File CEISA")
    file_ceisa = st.file_uploader("Upload Data CEISA", type=["xlsx", "xls", "csv", "pdf"], key="ceisa")

def clean_num(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ['nan', 'none', '-']:
        return 0.0
    
    if '.' in val_str and ',' in val_str:
        if val_str.rfind(',') > val_str.rfind('.'):
            val_str = val_str.replace('.', '').replace(',', '.')
        else:
            val_str = val_str.replace(',', '')
    elif ',' in val_str:
        parts = val_str.split(',')
        if len(parts) == 2 and len(parts[1]) in [1, 2]:
            val_str = val_str.replace(',', '.')
        else:
            val_str = val_str.replace(',', '')
            
    cleaned = re.sub(r'[^\d\.]', '', val_str)
    try:
        return float(cleaned)
    except ValueError:
        return 0.0

def extract_numbers_from_pdf(pdf_file):
    reader = pypdf.PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    
    def find_val(pattern):
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return clean_num(match.group(1))
        return 0.0

    qty = find_val(r"(?:qty|quantity|jumlah|total\s*qty)\s*[:=]?\s*([\d\.,]+)")
    gw = find_val(r"(?:gross\s*weight|gw|berat\s*kotor|net\s*weight|total\s*gw)\s*[:=]?\s*([\d\.,]+)")
    amt = find_val(r"(?:fob|cif|total\s*amount|amount|nilai\s*pabean|total\s*harga|price)\s*[:=]?\s*([\d\.,]+)")
    return qty, gw, amt, text

def read_excel_smart(uploaded_file):
    bytes_data = uploaded_file.read()
    uploaded_file.seek(0)
    
    try:
        return pd.read_excel(io.BytesIO(bytes_data), engine='calamine')
    except Exception:
        pass

    try:
        return pd.read_excel(io.BytesIO(bytes_data))
    except Exception:
        pass

    try:
        dfs = pd.read_html(io.BytesIO(bytes_data))
        if dfs:
            return dfs[0]
    except Exception:
        pass

    for sep in [',', ';', '\t', '|']:
        try:
            return pd.read_csv(io.BytesIO(bytes_data), sep=sep)
        except Exception:
            pass

    try:
        return bytes_data.decode('utf-8', errors='ignore')
    except Exception:
        pass

    raise ValueError("File tidak dapat dibaca.")

def clean_and_extract_df(df):
    if df is None or df.empty:
        return df, 0.0, 0.0, 0.0

    qty_keys = ['qty', 'quantity', 'jumlah', 'jml', 'jumlah barang', 'kuantitas', 'jumlah_satuan', 'jumlah_kemasan']
    gw_keys = ['gross weight', 'gross_weight', 'gw', 'berat kotor', 'bruto', 'gross', 'berat_kotor', 'berat_bruto', 'net weight', 'net_weight']
    amt_keys = ['fob', 'cif', 'amount', 'total amount', 'amount us $', 'total harga', 'nilai pabean', 'nilai', 'price', 'total price', 'nilai_incoterm', 'harga']

    cols_lower = [str(c).strip().lower() for c in df.columns]

    has_qty = any(any(k == c or k in c for k in qty_keys) for c in cols_lower)
    has_gw = any(any(k == c or k in c for k in gw_keys) for c in cols_lower)
    has_amt = any(any(k == c or k in c for k in amt_keys) for c in cols_lower)

    if not (has_qty or has_gw or has_amt):
        header_row_idx = None
        for idx, row in df.iterrows():
            row_str_list = [str(val).lower() for val in row.values if pd.notna(val)]
            row_combined = " ".join(row_str_list)
            if any(k in row_combined for k in qty_keys + gw_keys + amt_keys):
                header_row_idx = idx
                break

        if header_row_idx is not None:
            df.columns = [str(c).strip().lower() for c in df.iloc[header_row_idx].values]
            df = df.iloc[header_row_idx + 1:].reset_index(drop=True)
    else:
        df.columns = cols_lower

    df_clean = df.copy()
    row_text_summary = df_clean.apply(lambda row: " ".join([str(v).lower() for v in row.values if pd.notna(v)]), axis=1)
    df_clean = df_clean[~row_text_summary.str.contains(r'sub\s*total|subtotal|grand\s*total|^total', regex=True)].reset_index(drop=True)

    qty_col = next((c for c in df_clean.columns if any(k == str(c) or k in str(c) for k in qty_keys)), None)
    gw_col = next((c for c in df_clean.columns if any(k == str(c) or k in str(c) for k in gw_keys)), None)
    amt_col = next((c for c in df_clean.columns if any(k == str(c) or k in str(c) for k in amt_keys)), None)

    total_qty = float(df_clean[qty_col].apply(clean_num).sum()) if qty_col else 0.0
    total_gw = float(df_clean[gw_col].apply(clean_num).sum()) if gw_col else 0.0
    total_amt = float(df_clean[amt_col].apply(clean_num).sum()) if amt_col else 0.0

    return df_clean, total_qty, total_gw, total_amt

def load_data(uploaded_file):
    if uploaded_file is None:
        return None, 0.0, 0.0, 0.0
    
    file_name = uploaded_file.name.lower()
    
    if file_name.endswith('.pdf'):
        qty, gw, amt, raw_text = extract_numbers_from_pdf(uploaded_file)
        return "pdf", qty, gw, amt, raw_text
    
    parsed = read_excel_smart(uploaded_file)
    
    if isinstance(parsed, pd.DataFrame):
        df, qty, gw, amt = clean_and_extract_df(parsed)
        return "table", df, qty, gw, amt
    else:
        text = str(parsed)
        def find_val(pattern):
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return clean_num(match.group(1))
            return 0.0

        qty = find_val(r"(?:qty|quantity|jumlah)\s*[:=]?\s*([\d\.,]+)")
        gw = find_val(r"(?:gross\s*weight|gw|berat\s*kotor)\s*[:=]?\s*([\d\.,]+)")
        amt = find_val(r"(?:fob|cif|total\s*amount|amount|nilai)\s*[:=]?\s*([\d\.,]+)")
        return "text", None, qty, gw, amt

def check_item_level_mismatches(df_inv, df_pl, df_ceisa):
    """Pemeriksaan detail presisi per seri barang (mencakup Nilai FOB dan Berat sekaligus)."""
    mismatches = []
    if df_ceisa is None:
        return mismatches

    code_keys = ['product code', 'kode barang', 'kode_barang', 'item code', 'part number']
    fob_keys = ['fob', 'amount us $', 'amount', 'total amount', 'nilai pabean']
    gw_keys = ['gross weight', 'gross_weight', 'gw', 'berat kotor', 'bruto', 'gross', 'berat_kotor', 'berat_bruto', 'net weight', 'net_weight']
    qty_keys = ['qty', 'quantity', 'jumlah']

    # Sort CEISA berdasarkan nomor seri
    ceisa_seri_col = next((c for c in df_ceisa.columns if 'seri' in str(c).lower()), None)
    df_ceisa_sorted = df_ceisa.copy()
    if ceisa_seri_col:
        df_ceisa_sorted['seri_num'] = pd.to_numeric(df_ceisa_sorted[ceisa_seri_col], errors='coerce')
        df_ceisa_sorted = df_ceisa_sorted.sort_values(by='seri_num').drop(columns=['seri_num']).reset_index(drop=True)

    # Menyiapkan kolom Invoice
    df_inv_clean = df_inv.dropna(how='all').copy() if df_inv is not None else None
    inv_code_col = next((c for c in df_inv_clean.columns if any(k in str(c).lower() for k in code_keys)), None) if df_inv_clean is not None else None
    inv_item_col = next((c for c in df_inv_clean.columns if 'item code' in str(c).lower()), None) if df_inv_clean is not None else None
    inv_fob_col = next((c for c in df_inv_clean.columns if any(k in str(c).lower() for k in fob_keys)), None) if df_inv_clean is not None else None
    inv_qty_col = next((c for c in df_inv_clean.columns if any(k in str(c).lower() for k in qty_keys)), None) if df_inv_clean is not None else None

    # Menyiapkan kolom Packing List
    df_pl_clean = df_pl.dropna(how='all').copy() if df_pl is not None else None
    pl_code_col = next((c for c in df_pl_clean.columns if any(k in str(c).lower() for k in code_keys)), None) if df_pl_clean is not None else None
    pl_gw_col = next((c for c in df_pl_clean.columns if any(k in str(c).lower() for k in gw_keys)), None) if df_pl_clean is not None else None
    pl_qty_col = next((c for c in df_pl_clean.columns if any(k in str(c).lower() for k in qty_keys)), None) if df_pl_clean is not None else None

    # Kolom CEISA
    ceisa_code_col = next((c for c in df_ceisa_sorted.columns if any(k in str(c).lower() for k in code_keys)), None)
    ceisa_fob_col = next((c for c in df_ceisa_sorted.columns if any(k in str(c).lower() for k in fob_keys)), None)
    ceisa_gw_col = next((c for c in df_ceisa_sorted.columns if any(k in str(c).lower() for k in gw_keys)), None)
    ceisa_qty_col = next((c for c in df_ceisa_sorted.columns if any(k in str(c).lower() for k in qty_keys)), None)

    # Pemetaan dari Invoice: (Product Code, Qty) -> (FOB Amount, Item Code)
    inv_bridge = {}
    if df_inv_clean is not None and inv_code_col and inv_fob_col:
        for _, r in df_inv_clean.iterrows():
            p_code = str(r[inv_code_col]).strip().lower()
            q_val = str(r[inv_qty_col]).strip() if inv_qty_col else ""
            f_val = clean_num(r[inv_fob_col])
            i_code = str(r[inv_item_col]).strip().lower() if inv_item_col else ""
            if p_code and p_code != 'nan' and f_val > 0:
                inv_bridge[(p_code, q_val)] = (f_val, i_code)
                inv_bridge[p_code] = (f_val, i_code)

    # Pemetaan dari Packing List: (Item Code / Product Code, Qty) -> Weight
    pl_map = {}
    if df_pl_clean is not None and pl_code_col and pl_gw_col:
        for _, r in df_pl_clean.iterrows():
            c_val = str(r[pl_code_col]).strip().lower()
            q_val = str(r[pl_qty_col]).strip() if pl_qty_col else ""
            w_val = clean_num(r[pl_gw_col])
            if c_val and c_val != 'nan' and w_val > 0:
                pl_map[(c_val, q_val)] = w_val
                pl_map[c_val] = w_val

    for idx, row_ceisa in df_ceisa_sorted.iterrows():
        seri = row_ceisa[ceisa_seri_col] if ceisa_seri_col else idx + 1
        code = str(row_ceisa[ceisa_code_col]).strip() if ceisa_code_col else "-"
        q_str = str(row_ceisa[ceisa_qty_col]).strip() if ceisa_qty_col else ""
        ceisa_fob = clean_num(row_ceisa[ceisa_fob_col]) if ceisa_fob_col else 0.0
        ceisa_gw = clean_num(row_ceisa[ceisa_gw_col]) if ceisa_gw_col else 0.0
        uraian = str(row_ceisa.get('uraian', row_ceisa.get('description', '-')))

        # Pencarian acuan FOB Invoice & Item Code
        inv_fob, item_code = None, None
        if (code.lower(), q_str) in inv_bridge:
            inv_fob, item_code = inv_bridge[(code.lower(), q_str)]
        elif code.lower() in inv_bridge:
            inv_fob, item_code = inv_bridge[code.lower()]

        # Pencarian acuan Berat Packing List
        pl_gw = None
        if item_code and (item_code, q_str) in pl_map:
            pl_gw = pl_map[(item_code, q_str)]
        elif item_code and item_code in pl_map:
            pl_gw = pl_map[item_code]
        elif (code.lower(), q_str) in pl_map:
            pl_gw = pl_map[(code.lower(), q_str)]
        elif code.lower() in pl_map:
            pl_gw = pl_map[code.lower()]

        fob_diff = abs(ceisa_fob - inv_fob) if inv_fob is not None else 0.0
        gw_diff = abs(ceisa_gw - pl_gw) if pl_gw is not None else 0.0

        if fob_diff > 0.01 or gw_diff > 0.01:
            reks = []
            if fob_diff > 0.01 and inv_fob is not None:
                reks.append(f"Revisi nilai FOB pada Seri {seri} di CEISA menjadi ${inv_fob:,.2f}")
            if gw_diff > 0.01 and pl_gw is not None:
                reks.append(f"Revisi nilai Berat pada Seri {seri} di CEISA menjadi {pl_gw:,.2f} KG")

            mismatches.append({
                "Seri CEISA": seri,
                "Kode Barang": code.upper(),
                "Uraian Barang": uraian,
                "Nilai Invoice (USD)": f"${inv_fob:,.2f}" if inv_fob is not None else "-",
                "Nilai CEISA (USD)": f"${ceisa_fob:,.2f}",
                "Selisih FOB (USD)": f"${fob_diff:,.2f}" if fob_diff > 0.01 else "Cocok",
                "Berat PL (KG)": f"{pl_gw:,.2f} KG" if pl_gw is not None else "-",
                "Berat CEISA (KG)": f"{ceisa_gw:,.2f} KG",
                "Selisih Berat (KG)": f"{gw_diff:,.2f} KG" if gw_diff > 0.01 else "Cocok",
                "Rekomendasi Revisi": " & ".join(reks)
            })

    return mismatches

if st.button("🚀 Jalankan Validasi", type="primary"):
    if file_inv and file_pl and file_ceisa:
        try:
            res_inv = load_data(file_inv)
            res_pl = load_data(file_pl)
            res_ceisa = load_data(file_ceisa)

            type_inv, data_inv = res_inv[0], res_inv[1] if res_inv[0] == "table" else None
            qty_inv, gw_inv, amt_inv = res_inv[1 if res_inv[0] != "table" else 2], res_inv[2 if res_inv[0] != "table" else 3], res_inv[3 if res_inv[0] != "table" else 4]

            type_pl, data_pl = res_pl[0], res_pl[1] if res_pl[0] == "table" else None
            qty_pl, gw_pl, amt_pl = res_pl[1 if res_pl[0] != "table" else 2], res_pl[2 if res_pl[0] != "table" else 3], res_pl[3 if res_pl[0] != "table" else 4]

            type_ceisa, data_ceisa = res_ceisa[0], res_ceisa[1] if res_ceisa[0] == "table" else None
            qty_ceisa, gw_ceisa, amt_ceisa = res_ceisa[1 if res_ceisa[0] != "table" else 2], res_ceisa[2 if res_ceisa[0] != "table" else 3], res_ceisa[3 if res_ceisa[0] != "table" else 4]

            st.markdown("---")
            st.header("🔍 Hasil Pemeriksaan Validasi")

            def is_match(v1, v2, v3=None):
                if v3 is None:
                    return abs(v1 - v2) < 0.05
                return abs(v1 - v2) < 0.05 and abs(v2 - v3) < 0.05

            validasi_data = [
                {
                    "Parameter": "Total Kuantitas (Qty)",
                    "Invoice": f"{qty_inv:,.2f}",
                    "Packing List": f"{qty_pl:,.2f}",
                    "CEISA": f"{qty_ceisa:,.2f}",
                    "Status": "✅ MATCH" if is_match(qty_inv, qty_pl, qty_ceisa) else "❌ MISMATCH"
                },
                {
                    "Parameter": "Total Gross / Net Weight (KG)",
                    "Invoice": f"{gw_inv:,.2f}" if gw_inv > 0 else "-",
                    "Packing List": f"{gw_pl:,.2f}",
                    "CEISA": f"{gw_ceisa:,.2f}",
                    "Status": "✅ MATCH" if is_match(gw_pl, gw_ceisa) else "❌ MISMATCH"
                },
                {
                    "Parameter": "Total Amount (FOB / Nilai Pabean)",
                    "Invoice": f"{amt_inv:,.2f}",
                    "Packing List": f"{amt_pl:,.2f}" if amt_pl > 0 else "-",
                    "CEISA": f"{amt_ceisa:,.2f}",
                    "Status": "✅ MATCH" if is_match(amt_inv, amt_ceisa) else "❌ MISMATCH"
                }
            ]

            df_res = pd.DataFrame(validasi_data)
            st.dataframe(df_res, use_container_width=True)

            is_all_valid = all(d["Status"] == "✅ MATCH" for d in validasi_data)
            if is_all_valid:
                st.success("Semua data cocok! Dokumen siap diproses.")
            else:
                st.error("Ditemukan ketidakcocokan data. Periksa detail revisi di bawah ini!")

                if type_inv == "table" and type_ceisa == "table":
                    item_errors = check_item_level_mismatches(data_inv, data_pl, data_ceisa)
                    if item_errors:
                        st.markdown("### ⚠️ Detail Dokumen & Nomor Seri Yang Harus Direvisi di CEISA")
                        df_err = pd.DataFrame(item_errors)
                        st.dataframe(df_err, use_container_width=True)
                    else:
                        st.info("Penyebab selisih diduga akibat perbedaan total gabungan atau entri yang belum lengkap di CEISA.")

            with st.expander("Lihat Detail Data Uploaded"):
                col_a, col_b, col_c = st.columns(3)
                
                with col_a:
                    st.subheader("Detail Invoice")
                    if type_inv == "table":
                        st.dataframe(data_inv)
                    else:
                        st.text("Konten dibaca sebagai teks/PDF")

                with col_b:
                    st.subheader("Detail Packing List")
                    if type_pl == "table":
                        st.dataframe(data_pl)
                    else:
                        st.text("Konten dibaca sebagai teks/PDF")

                with col_c:
                    st.subheader("Detail CEISA")
                    if type_ceisa == "table":
                        st.dataframe(data_ceisa)
                    else:
                        st.text("Konten dibaca sebagai teks/PDF")

        except Exception as e:
            st.error(f"Terjadi kesalahan saat membaca file: {e}")
    else:
        st.warning("Mohon unggah ketiga file terlebih dahulu!")
