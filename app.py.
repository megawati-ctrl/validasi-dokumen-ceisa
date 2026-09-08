import streamlit as st
import pandas as pd

st.set_page_config(page_title="Validasi Dokumen Ekspor/Impor", layout="wide")

st.title("📋 App Validasi Dokumen: Invoice, Packing List & CEISA")
st.write("Unggah file Excel/CSV untuk ketiga dokumen, lalu tekan tombol **Jalankan Validasi**.")

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. File Invoice")
    file_inv = st.file_uploader("Upload Invoice", type=["xlsx", "xls", "csv"], key="inv")

with col2:
    st.subheader("2. File Packing List")
    file_pl = st.file_uploader("Upload Packing List", type=["xlsx", "xls", "csv"], key="pl")

with col3:
    st.subheader("3. File CEISA")
    file_ceisa = st.file_uploader("Upload Data CEISA", type=["xlsx", "xls", "csv"], key="ceisa")

def load_data(uploaded_file):
    if uploaded_file is not None:
        if uploaded_file.name.endswith('.csv'):
            return pd.read_csv(uploaded_file)
        else:
            return pd.read_excel(uploaded_file)
    return None

if st.button("🚀 Jalankan Validasi", type="primary"):
    if file_inv and file_pl and file_ceisa:
        try:
            df_inv = load_data(file_inv)
            df_pl = load_data(file_pl)
            df_ceisa = load_data(file_ceisa)

            # Normalisasi nama kolom menjadi huruf kecil & tanpa spasi berlebih
            for df in [df_inv, df_pl, df_ceisa]:
                df.columns = df.columns.str.strip().str.lower()

            st.markdown("---")
            st.header("🔍 Hasil Pemeriksaan Validasi")

            # 1. Ringkasan Total Kuantitas / Berat
            # Sesuaikan nama kolom default jika berbeda di file Excel kamu
            qty_inv = df_inv['qty'].sum() if 'qty' in df_inv.columns else 0
            qty_pl = df_pl['qty'].sum() if 'qty' in df_pl.columns else 0
            qty_ceisa = df_ceisa['qty'].sum() if 'qty' in df_ceisa.columns else 0

            gw_pl = df_pl['gross_weight'].sum() if 'gross_weight' in df_pl.columns else 0
            gw_ceisa = df_ceisa['gross_weight'].sum() if 'gross_weight' in df_ceisa.columns else 0

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
                    "Invoice": "-",
                    "Packing List": gw_pl,
                    "CEISA": gw_ceisa,
                    "Status": "✅ MATCH" if (gw_pl == gw_ceisa) else "❌ MISMATCH"
                }
            ]

            df_res = pd.DataFrame(validasi_data)
            st.dataframe(df_res, use_container_width=True)

            # Status Kesimpulan
            is_all_valid = all(d["Status"] == "✅ MATCH" for d in validasi_data)
            if is_all_valid:
                st.success("Semua data cocok! Dokumen siap diproses.")
            else:
                st.error("Ditemukan ketidakcocokan data. Periksa kembali entri dokumen!")

            # Detail Tampilan Data
            with st.expander("Lihat Data Mentah (Raw Data)"):
                st.subheader("Data Invoice")
                st.dataframe(df_inv)
                st.subheader("Data Packing List")
                st.dataframe(df_pl)
                st.subheader("Data CEISA")
                st.dataframe(df_ceisa)

        except Exception as e:
            st.error(f"Terjadi kesalahan saat membaca file: {e}")
            st.info("Pastikan file kamu memiliki nama kolom seperti: `qty`, `gross_weight`, dll.")
    else:
        st.warning("Mohon unggah ketiga file terlebih dahulu!")
