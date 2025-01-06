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

        # Load DbSimpanan dan skip baris 1
        dbsimpanan_df = pd.read_excel(uploaded_dbsimpanan, skiprows=1)

        # Rename kolom DbSimpanan untuk konsistensi
        dbsimpanan_df = dbsimpanan_df.rename(columns={"Center ID": "Ctr ID"})

        # Pastikan kolom 'Ctr ID' memiliki tipe data yang sama (string)
        delinquency_df['Ctr ID'] = delinquency_df['Ctr ID'].astype(str)
        dbsimpanan_df['Ctr ID'] = dbsimpanan_df['Ctr ID'].astype(str)

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

        # Tambahkan filter untuk memilih beberapa Center
        unique_centers = sorted(kk_anggota_df["Ctr ID"].unique().tolist())
        selected_centers = st.multiselect(
            "Pilih Center (Ctr ID)", 
            options=unique_centers,
            default=unique_centers  # Default semua Center terpilih
        )
        
        # Filter data berdasarkan Center yang dipilih
        if selected_centers:
            kk_anggota_df = kk_anggota_df[kk_anggota_df["Ctr ID"].isin(selected_centers)]

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
