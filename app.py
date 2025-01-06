import pandas as pd
import streamlit as st

st.title("Kertas Kerja Anggota PAR")
st.write("""File ini berisikan Delinquency.xlsx dan DbSimpanan.xlsx""")

# Upload file Delinquency dan DbSimpanan
uploaded_delinquency = st.file_uploader("Upload Delinquency.xlsx", type="xlsx")
uploaded_dbsimpanan = st.file_uploader("Upload DbSimpanan.xlsx", type="xlsx")

if uploaded_delinquency and uploaded_dbsimpanan:
    try:
        # Load Delinquency dan skip baris 1-3
        delinquency_df = pd.read_excel(uploaded_delinquency, skiprows=3)
        dbsimpanan_df = pd.read_excel(uploaded_dbsimpanan, skiprows=1)

        # Rename kolom DbSimpanan untuk konsistensi
        dbsimpanan_df = dbsimpanan_df.rename(columns={"Center ID": "Ctr ID"})

        # Proses Data
        kk_anggota_df = delinquency_df[[
            "Client ID", "Loan No", "Client Name", "Ctr ID", "Total Balance", "Arreas Due"
        ]].copy()

        # Tambahkan kolom "No."
        kk_anggota_df.insert(0, "No.", range(1, len(kk_anggota_df) + 1))

        # Lakukan VLOOKUP untuk mendapatkan Officer Name berdasarkan Ctr ID
        kk_anggota_df = kk_anggota_df.merge(
            dbsimpanan_df[["Ctr ID", "Officer Name"]],
            on="Ctr ID",
            how="left"
        )

        # Tambahkan kolom kosong
        kk_anggota_df["Ditemui/ Tidak Ditemukan"] = ""
        kk_anggota_df["KETERANGAN (Kelemahan)"] = ""

        # Pilih kolom sesuai kebutuhan output
        kk_anggota_df = kk_anggota_df[[
            "No.", "Client ID", "Loan No", "Client Name", "Ctr ID", "Officer Name",
            "Total Balance", "Arreas Due", "Ditemui/ Tidak Ditemukan", "KETERANGAN (Kelemahan)"
        ]]

        # Tambahkan filter untuk Center
        selected_center = st.selectbox(
            "Pilih Center (Ctr ID)", 
            options=["Semua"] + sorted(kk_anggota_df["Ctr ID"].unique().tolist())
        )
        
        if selected_center != "Semua":
            kk_anggota_df = kk_anggota_df[kk_anggota_df["Ctr ID"] == selected_center]

        # Tambahkan filter untuk Top N berdasarkan Total Balance atau Arreas Due
        filter_column = st.selectbox("Filter Berdasarkan", ["Total Balance", "Arreas Due"])
        top_n = st.slider("Tampilkan Top N Data", min_value=1, max_value=len(kk_anggota_df), value=10)
        
        # Urutkan data berdasarkan kolom yang dipilih dan ambil Top N
        kk_anggota_df = kk_anggota_df.sort_values(by=filter_column, ascending=False).head(top_n)

        # Tampilkan hasil pada Streamlit
        st.success("Data berhasil diproses dengan filter!")
        st.dataframe(kk_anggota_df)

        # Fungsi untuk convert DataFrame ke Excel
        @st.cache_data
        def convert_df_to_excel(df):
            from io import BytesIO
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="KK Anggota PAR")
            processed_data = output.getvalue()
            return processed_data

        # Buat file Excel untuk diunduh
        excel_data = convert_df_to_excel(kk_anggota_df)
        st.download_button(
            label="Unduh KK Anggota PAR",
            data=excel_data,
            file_name="KK Anggota PAR.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
else:
    st.warning("Silakan unggah kedua file untuk melanjutkan.")
