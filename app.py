import streamlit as st
import pandas as pd
import numpy as np
from sklearn.cluster import KMeans
from sklearn.model_selection import train_test_split
from sklearn.naive_bayes import GaussianNB
from sklearn.metrics import accuracy_score, classification_report, confusion_matrix
import plotly.express as px
import io
import joblib
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Pipeline K-Means + NB Lengkap", layout="wide", page_icon="🔬")

st.title("🔬 Pipeline K-Means & Gaussian Naive Bayes - Versi Lengkap")

with st.sidebar:
    st.header("Pengaturan")
    st.caption("Versi 1.4 - Excel Mirip Contoh")

uploaded_file = st.file_uploader("Upload File Excel", type=["xlsx"])

if uploaded_file is not None:
    df_raw = pd.read_excel(uploaded_file)
    st.success(f"File dimuat: {len(df_raw)} baris")

    # ====================== DETEKSI KOLOM ======================
    col_map = {}
    for col in df_raw.columns:
        c = str(col).lower().strip()
        if any(x in c for x in ['nomor', 'soal', 'no']):
            col_map['Nomor Soal'] = col
        elif any(x in c for x in ['persentase', 'persen', 'p (%)']):
            col_map['Persentase'] = col
        elif any(x in c for x in ['waktu', 'time', 'detik']):
            col_map['Waktu'] = col
        elif any(x in c for x in ['benar', 'correct', 'jumlah benar']):
            col_map['Jumlah Benar'] = col
        elif any(x in c for x in ['siswa', 'total', 'jumlah siswa']):
            col_map['Jumlah Siswa'] = col

    # Ambil kolom yang tersedia
    needed = ['Nomor Soal', 'Persentase', 'Waktu']
    for key in needed:
        if key not in col_map:
            st.error(f"Kolom '{key}' tidak ditemukan. Kolom yang ada: {list(df_raw.columns)}")
            st.stop()

    df = df_raw[[col_map['Nomor Soal'], col_map['Persentase'], col_map['Waktu']]].copy()
    df = df.dropna().reset_index(drop=True)
    df = df.rename(columns={
        col_map['Nomor Soal']: 'Soal',
        col_map['Persentase']: 'Persentase',
        col_map['Waktu']: 'Waktu (detik)'
    })

    X = df[['Persentase', 'Waktu (detik)']]
    X_scaled = (X - X.min()) / (X.max() - X.min())

    # ====================== K-MEANS ======================
    # Centroid awal sesuai contoh kamu
    init_centroids_scaled = np.array([[1.0, 0.0], [0.85, 0.5], [0.6, 0.9]])  # dinormalisasi kasar

    kmeans = KMeans(n_clusters=3, init=init_centroids_scaled, n_init=1, max_iter=10, random_state=42)
    df['Cluster'] = kmeans.fit_predict(X_scaled)
    cluster_map = {0: 'Mudah', 1: 'Sedang', 2: 'Sulit'}
    df['Label'] = df['Cluster'].map(cluster_map)

    # ====================== NAIVE BAYES ======================
    X_nb = df[['Persentase', 'Waktu (detik)']]
    y_nb = df['Label']

    X_train, X_test, y_train, y_test = train_test_split(X_nb, y_nb, test_size=0.25, 
                                                        stratify=y_nb, random_state=42)

    model = GaussianNB()
    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)
    acc = accuracy_score(y_test, y_pred)

    # ====================== GENERATE EXCEL LENGKAP ======================
    if st.button("🚀 Generate Excel Report Lengkap (Mirip Contoh)", type="primary", use_container_width=True):
        with st.spinner("Sedang membuat Excel lengkap dengan banyak sheet..."):
            output = io.BytesIO()

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Sheet Data
                df.to_excel(writer, sheet_name='Data', index=False)

                # Sheet Hasil K-Means
                df[['Soal', 'Persentase', 'Waktu (detik)', 'Label']].to_excel(
                    writer, sheet_name='Hasil K-Means', index=False)

                # Sheet Evaluasi NB
                report = pd.DataFrame(classification_report(y_test, y_pred, output_dict=True)).round(4).transpose()
                report.to_excel(writer, sheet_name='Evaluasi NB')

                # Sheet Stratified Split (sederhana)
                split_info = pd.DataFrame({
                    'Keterangan': ['Total Soal', 'Train (75%)', 'Test (25%)'],
                    'Jumlah': [len(df), len(X_train), len(X_test)]
                })
                split_info.to_excel(writer, sheet_name='Stratified Split', index=False)

            output.seek(0)
            wb = load_workbook(output)

            # ====================== STYLING WARNA ======================
            ws = wb['Hasil K-Means']
            hijau = PatternFill(start_color="C6EFCE", fill_type="solid")
            kuning = PatternFill(start_color="FFEB9C", fill_type="solid")
            merah = PatternFill(start_color="FFC7CE", fill_type="solid")

            for row in range(2, ws.max_row + 1):
                label_cell = ws.cell(row=row, column=4)
                if label_cell.value == "Mudah":
                    label_cell.fill = hijau
                elif label_cell.value == "Sedang":
                    label_cell.fill = kuning
                elif label_cell.value == "Sulit":
                    label_cell.fill = merah

            # Simpan final
            final_buf = io.BytesIO()
            wb.save(final_buf)
            final_buf.seek(0)

            st.download_button(
                label="📥 Download Excel Lengkap",
                data=final_buf,
                file_name=f"Hasil_KMeans_NB_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            # Download Model
            model_buf = io.BytesIO()
            joblib.dump(model, model_buf)
            model_buf.seek(0)
            st.download_button(
                label="📥 Download Model Naive Bayes",
                data=model_buf,
                file_name="model_nb.pkl",
                use_container_width=True
            )

            st.success("✅ Excel berhasil dibuat!")

    # Tampilan di Streamlit
    tab1, tab2, tab3 = st.tabs(["Hasil K-Means", "Naive Bayes", "Visualisasi"])
    with tab1:
        st.dataframe(df, use_container_width=True)
    with tab2:
        st.metric("Akurasi", f"{acc:.2%}")
        st.dataframe(pd.DataFrame(classification_report(y_test, y_pred, output_dict=True)).round(4).transpose())
    with tab3:
        fig = px.scatter(df, x='Persentase', y='Waktu (detik)', color='Label', title="Visualisasi K-Means")
        st.plotly_chart(fig, use_container_width=True)

else:
    st.info("Silakan upload file Excel berisi data soal matematika.")

st.caption("Pipeline sesuai flowchart • Excel dengan warna + multiple sheet")
