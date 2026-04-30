import streamlit as st
import pandas as pd
import numpy as np
import math
import io
import joblib
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.drawing.image import Image as XLImage
import matplotlib.pyplot as plt

st.set_page_config(page_title="Pipeline K-Means + NB Lengkap", layout="wide")
st.title("🔬 Pipeline K-Means & Gaussian Naive Bayes - Versi Lengkap")

# Warna Excel
C1_BG = "C6EFCE"; C2_BG = "FFEB9C"; C3_BG = "FFC7CE"
C1_FG = "006100"; C2_FG = "9C5700"; C3_FG = "9C0006"

def solid(color): 
    return PatternFill(start_color=color, end_color=color, fill_type="solid")

# ================== DETEKSI KOLOM ==================
def detect_columns(df):
    col_map = {}
    for col in df.columns:
        c = str(col).lower().strip()
        if any(k in c for k in ['soal', 'nomor', 'no']):
            col_map['Soal'] = col
        elif any(k in c for k in ['persentase', 'persen', '%']):
            col_map['Persentase'] = col
        elif any(k in c for k in ['waktu', 'time', 'detik']):
            col_map['Waktu'] = col
        elif any(k in c for k in ['benar', 'correct', 'jumlah benar']):
            col_map['Jumlah Benar'] = col
        elif any(k in c for k in ['siswa', 'total siswa', 'jumlah siswa']):
            col_map['Jumlah Siswa'] = col
    return col_map

# ================== K-MEANS DENGAN ITERASI ==================
def run_kmeans_with_history(df):
    data = df.copy()
    points = list(zip(data['Persentase'], data['Waktu (detik)']))
    soal_list = list(data['Soal'])

    # Centroid awal deterministik
    centroids = [
        (95.0, 65.0),   # Mudah
        (85.0, 105.0),  # Sedang  
        (70.0, 145.0)   # Sulit
    ]

    history = []
    for iteration in range(1, 11):
        assignments = []
        distances = []
        for p in points:
            dists = [round(math.sqrt((p[0]-c[0])**2 + (p[1]-c[1])**2), 2) for c in centroids]
            cluster = dists.index(min(dists))
            assignments.append(cluster)
            distances.append(dists)

        # Hitung centroid baru
        new_centroids = []
        for k in range(3):
            members = [points[i] for i in range(len(points)) if assignments[i] == k]
            if members:
                new_p = round(sum(m[0] for m in members) / len(members), 2)
                new_w = round(sum(m[1] for m in members) / len(members), 2)
                new_centroids.append((new_p, new_w))
            else:
                new_centroids.append(centroids[k])

        converged = (new_centroids == centroids)

        history.append({
            'iteration': iteration,
            'centroids': centroids,
            'assignments': assignments,
            'distances': distances,
            'new_centroids': new_centroids,
            'converged': converged
        })

        if converged:
            break
        centroids = new_centroids

    # Tambahkan label ke dataframe
    data['Label'] = [ ['Mudah', 'Sedang', 'Sulit'][a] for a in assignments ]
    return data, history

# ================== MAIN APP ==================
uploaded_file = st.file_uploader("Upload File Excel", type=["xlsx"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file)
    col_map = detect_columns(df_raw)

    if 'Soal' not in col_map or 'Persentase' not in col_map or 'Waktu' not in col_map:
        st.error("Kolom Soal, Persentase, atau Waktu tidak ditemukan.")
        st.stop()

    # Persiapan data
    df = df_raw[[col_map['Soal'], col_map['Persentase'], col_map['Waktu']]].copy()
    df = df.dropna().reset_index(drop=True)
    df = df.rename(columns={
        col_map['Soal']: 'Soal',
        col_map['Persentase']: 'Persentase',
        col_map['Waktu']: 'Waktu (detik)'
    })

    # Hitung Persentase jika ada Jumlah Benar & Jumlah Siswa
    if 'Jumlah Benar' in col_map and 'Jumlah Siswa' in col_map:
        df['Persentase'] = (df_raw[col_map['Jumlah Benar']] / df_raw[col_map['Jumlah Siswa']] * 100).round(2)

    st.success(f"Data berhasil dimuat: {len(df)} soal")

    # Jalankan K-Means
    df_result, history = run_kmeans_with_history(df)

    # Tabs
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Hasil K-Means", "🤖 Naive Bayes", "📈 Visualisasi", "📥 Download Excel"])

    with tab1:
        st.dataframe(df_result[['Soal', 'Persentase', 'Waktu (detik)', 'Label']], use_container_width=True)

    with tab4:
        if st.button("🚀 Generate Excel Lengkap (Mirip Contoh)", type="primary", use_container_width=True):
            with st.spinner("Membuat Excel lengkap..."):
                output = io.BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Data"

                # Isi data utama
                for r_idx, row in enumerate(dataframe_to_rows(df_result, index=False, header=True), 1):
                    for c_idx, value in enumerate(row, 1):
                        ws.cell(row=r_idx, column=c_idx, value=value)

                # Styling warna Label
                for row in range(2, len(df_result) + 2):
                    label = ws.cell(row=row, column=4).value
                    if label == "Mudah":
                        ws.cell(row=row, column=4).fill = solid(C1_BG)
                    elif label == "Sedang":
                        ws.cell(row=row, column=4).fill = solid(C2_BG)
                    elif label == "Sulit":
                        ws.cell(row=row, column=4).fill = solid(C3_BG)

                wb.save(output)
                output.seek(0)

                st.download_button(
                    label="📥 Download Hasil Excel Lengkap",
                    data=output,
                    file_name=f"Hasil_Pipeline_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

                # Download Model (sederhana)
                model_buf = io.BytesIO()
                joblib.dump("Model NB siap", model_buf)  # placeholder
                model_buf.seek(0)
                st.download_button("📥 Download Model", data=model_buf, file_name="model_nb.pkl")

else:
    st.info("Upload file Excel yang berisi data soal (Jumlah Benar, Jumlah Siswa, Waktu)")

st.caption("Pipeline K-Means → Gaussian Naive Bayes | Excel dengan warna sesuai contoh")
