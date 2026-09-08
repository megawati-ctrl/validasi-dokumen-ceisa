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

def extract_numbers_from_pdf(pdf_file):
    reader = pypdf.PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    
    def find_val(pattern):
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            val_str = match.group(1).replace(',', '')
            try:
                return float(val_str)
            except ValueError:
                return 0.0
        return 0.0

    qty = find_val(r"(?:qty|quantity|jumlah)\s*[:=]?\s*([\d\.,]+)")
    gw = find_val(r"(?:gross\s*weight|gw|berat\s*kotor)\s*[:=]?\s*([\d\.,]+)")
    return qty, gw, text

def read_excel_smart(uploaded_file):
    """Membaca berbagai jenis format file Excel menggunakan berbagai engine secara bergantian"""
    bytes_data = uploaded_file.read()
    uploaded_file.seek(0)
    
    # 1. Coba Pakai Engine Calamine (Paling ampuh membaca file rusak/lama)
    try:
        return pd.read_excel(io.BytesIO(bytes_data), engine='calamine')
    except Exception:
        pass

    # 2. Coba Pakai Engine xlrd/openpyxl standar
    try:
        return pd.read_excel(io.BytesIO(bytes_data))
    except Exception:
        pass

    # 3. Coba Baca sebagai HTML Table
    try:
        dfs = pd.read_html(io.BytesIO(bytes_data))
        if dfs:
            return dfs[0]
    except Exception:
        pass

    # 4. Coba Baca sebagai CSV dengan berbagai pemisah (koma, titik koma, tab)
    for sep in [',', ';', '\t', '|']:
        try:
            return pd.read_csv(io.BytesIO(bytes_data), sep=sep)
        except Exception:
            pass

    # 5. Jika gagal total, baca isi file sebagai teks untuk dicari angkanya
    try:
        text_content = bytes_data.decode('utf-8', errors='ignore')
        return text_content
    except Exception:
        pass

    raise ValueError("File korup atau formatnya tidak didukung.")

def load_data(uploaded_file):
    if uploaded_file is None:
        return None, 0.0, 0.0
    
    file_name = uploaded_file.name.lower()
    
    if file_name.endswith('.pdf'):
        qty, gw, raw_text = extract_numbers_from_pdf(uploaded_file)
        return "pdf", qty, gw, raw_text
    
    parsed = read_excel_smart(uploaded_file)
    
    if isinstance(parsed, pd.DataFrame):
        df = parsed
        df.columns = df.columns.astype(str).str.strip().str.lower()
        
        qty = df['qty'].sum() if 'qty' in df.columns else 0.0
        gw = df['gross_weight'].sum() if 'gross_weight' in df.columns else 0.0
        return "table", df, qty, gw
    else:
        # Jika berupa teks mentah
        text = parsed
        def find_val(pattern):
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                val_str = match.group(1).replace(',', '')
                try:
                    return float(val_str)
                except ValueError:
                    return 0.0
            return 0.0

        qty = find_val(r"(?:qty|quantity|jumlah)\s*[:=]?\s*([\d\.,]+)")
        gw = find_val(r"(?:gross\s*weight|gw|berat\s*kotor)\s*[:=]?\s*([\d\.,]+)")
        return "text", None, qty, gw

if st.button("🚀 Jalankan Validasi", type="primary"):
    if file_inv and file_pl and file_ceisa:
        try:
            res_inv = load_data(file_inv)
            res_pl = load_data(file_pl)
            res_ceisa = load_data(file_ceisa)

            type_inv, data_inv, qty_inv, gw_inv = (res_inv[0], None, res_inv[1], res_inv[2]) if res_inv[0] in ["pdf", "text"] else (res_inv[0], res_inv[1], res_inv[2], res_inv[3])
            type_pl, data_pl, qty_pl, gw_pl = (res_pl[0], None, res_pl[1], res_pl[2]) if res_pl[0] in ["pdf", "text"] else (res_pl[0], res_pl[1], res_pl[2], res_pl[3])
            type_ceisa, data_ceisa, qty_ceisa, gw_ceisa = (res_ceisa[0], None, res_ceisa[1], res_ceisa[2]) if res_ceisa[0] in ["pdf", "text"] else (res_ceisa[0], res_ceisa[1], res_ceisa[2], res_ceisa[3])

            st.markdown("---")
            st.header("🔍 Hasil Pemeriksaan Validasi")

            validasi_data = [
                {
                    "Parameter": "Total Kuantitas (Qty)",
                    "Invoice": qty_inv,
                    "Packing List": qty_pl,
                    "CEISA": qty_ceisa,
                    "Status": "✅ MATCH" if (qty_inv == qty_pl == qty_ceisa) else "❌ MISMATCH"
                },
                {
                    "Parameter": "Total Gross Weight (KG)",
                    "Invoice": gw_inv if gw_inv > 0 else "-",
                    "Packing List": gw_pl,
                    "CEISA": gw_ceisa,
                    "Status": "✅ MATCH" if (gw_pl == gw_ceisa) else "❌ MISMATCH"
                }
            ]

            df_res = pd.DataFrame(validasi_data)
            st.dataframe(df_res, use_container_width=True)

            is_all_valid = all(d["Status"] == "✅ MATCH" for d in validasi_data)
            if is_all_valid:
                st.success("Semua data cocok! Dokumen siap diproses.")
            else:
                st.error("Ditemukan ketidakcocokan data. Periksa kembali entri dokumen!")

            with st.expander("Lihat Detail Data Uploaded"):
                col_a, col_b, col_c = st.columns(3)
                with col_a:
                    st.subheader("Detail Invoice")
                    st.dataframe(data_inv) if type_inv == "table" else st.text("Konten berhasil dibaca sebagai teks")
                with col_b:
                    st.subheader("Detail Packing List")
                    st.dataframe(data_pl) if type_pl == "table" else st.text("Konten berhasil dibaca sebagai teks")
                with col_c:
                    st.subheader("Detail CEISA")
                    st.dataframe(data_ceisa) if type_ceisa == "table" else st.text("Konten berhasil dibaca sebagai teks")

        except Exception as e:
            st.error(f"Terjadi kesalahan saat membaca file: {e}")
    else:
        st.warning("Mohon unggah ketiga file terlebih dahulu!")
