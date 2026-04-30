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

st.set_page_config(page_title="Pipeline K-Means + NB", page_icon="🔬", layout="wide")

st.markdown("""
<style>
    .main-header {font-size: 2.8rem; color: #1E88E5; text-align: center; margin-bottom: 10px;}
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.title("🔬 Pipeline Lengkap")
    st.markdown("K-Means + Gaussian Naive Bayes")
    st.divider()
    st.caption("Versi 1.3 - Excel Mirip Contoh")

st.markdown('<h1 class="main-header">Desain Sistem Pipeline K-Means & Gaussian Naive Bayes</h1>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("📤 Upload File Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    # Mapping kolom fleksibel
    col_map = {}
    for col in df.columns:
        c = str(col).lower().strip()
        if 'nomor' in c or 'soal' in c:
            col_map['Nomor'] = col
        elif 'persentase' in c or 'persen' in c:
            col_map['Persentase'] = col
        elif 'waktu' in c:
            col_map['Waktu'] = col

    if not all(k in col_map for k in ['Nomor', 'Persentase', 'Waktu']):
        st.error("Tidak dapat menemukan kolom Nomor Soal, Persentase, dan Waktu.")
        st.stop()

    df_clean = df[[col_map['Nomor'], col_map['Persentase'], col_map['Waktu']]].copy()
    df_clean = df_clean.dropna().reset_index(drop=True)
    df_clean = df_clean.rename(columns={
        col_map['Nomor']: 'Nomor Soal',
        col_map['Persentase']: 'Persentase',
        col_map['Waktu']: 'Waktu (detik)'
    })

    X = df_clean[['Persentase', 'Waktu (detik)']]
    X_scaled = (X - X.min()) / (X.max() - X.min())

    # ====================== K-MEANS dengan Tracking Iterasi ======================
    st.subheader("K-Means Clustering (Deterministic Centroid)")

    # Centroid awal (dalam skala asli, bukan scaled)
    # Kita gunakan nilai realistis berdasarkan data kamu
    centroids = np.array([[95.0, 67.0], [85.0, 108.0], [72.0, 145.0]])
    
    kmeans = KMeans(n_clusters=3, init=centroids, n_init=1, max_iter=10, random_state=42, tol=1e-4)
    
    # Untuk tracking iterasi (manual karena sklearn tidak simpan semua iterasi)
    labels_history = []
    centroids_history = [centroids.copy()]
    
    for i in range(10):
        kmeans.fit(X_scaled)
        labels = kmeans.labels_
        labels_history.append(labels)
        new_centroids = kmeans.cluster_centers_
        centroids_history.append(new_centroids)
        
        if i > 0 and np.allclose(centroids_history[-1], centroids_history[-2], atol=0.01):
            st.success(f"✅ K-Means konvergen pada iterasi ke-{i+1}")
            break

    df_clean['Label'] = pd.Series(labels).map({0: 'Mudah', 1: 'Sedang', 2: 'Sulit'})

    # ====================== NAIVE BAYES ======================
    X_nb = df_clean[['Persentase', 'Waktu (detik)']]
    y_nb = df_clean['Label']

    X_train, X_test, y_train, y_test = train_test_split(
        X_nb, y_nb, test_size=0.25, stratify=y_nb, random_state=42
    )

    model_nb = GaussianNB()
    model_nb.fit(X_train, y_train)
    y_pred = model_nb.predict(X_test)
    accuracy = accuracy_score(y_test, y_pred)

    # ====================== TABS ======================
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Hasil K-Means", "🤖 Naive Bayes", "📈 Visualisasi", "📥 Download Excel Lengkap"])

    with tab1:
        st.dataframe(df_clean[['Nomor Soal', 'Persentase', 'Waktu (detik)', 'Label']], use_container_width=True)

    with tab2:
        st.metric("Akurasi Gaussian Naive Bayes", f"{accuracy:.2%}")
        report_df = pd.DataFrame(classification_report(y_test, y_pred, output_dict=True, zero_division=0)).round(4).transpose()
        st.dataframe(report_df)

    with tab3:
        fig = px.scatter(X, x='Persentase', y='Waktu (detik)', color=df_clean['Label'],
                        title="Visualisasi Clustering Soal", height=600)
        st.plotly_chart(fig, use_container_width=True)

    with tab4:
        if st.button("🚀 Generate Excel Report Lengkap (Mirip Contoh)", type="primary", use_container_width=True):
            with st.spinner("Membuat Excel lengkap dengan styling..."):
                output = io.BytesIO()
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Sheet 1: Data
                    df_clean.to_excel(writer, sheet_name='Data', index=False)
                    
                    # Sheet 2: Hasil K-Means
                    df_clean[['Nomor Soal', 'Persentase', 'Waktu (detik)', 'Label']].to_excel(
                        writer, sheet_name='Hasil K-Means', index=False)
                    
                    # Sheet 3: Evaluasi NB
                    report_df.to_excel(writer, sheet_name='Evaluasi NB')

                output.seek(0)
                wb = load_workbook(filename=output)

                # Styling warna di sheet Hasil K-Means
                ws = wb['Hasil K-Means']
                fill_mudah = PatternFill(start_color="C6EFCE", fill_type="solid")
                fill_sedang = PatternFill(start_color="FFEB9C", fill_type="solid")
                fill_sulit = PatternFill(start_color="FFC7CE", fill_type="solid")

                for row in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row, column=4)  # Kolom Label
                    if cell.value == "Mudah":
                        cell.fill = fill_mudah
                    elif cell.value == "Sedang":
                        cell.fill = fill_sedang
                    elif cell.value == "Sulit":
                        cell.fill = fill_sulit

                # Simpan ke buffer baru
                final_output = io.BytesIO()
                wb.save(final_output)
                final_output.seek(0)

                st.download_button(
                    label="📥 Download Excel Lengkap",
                    data=final_output,
                    file_name=f"Hasil_Pipeline_KMeans_NB_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

                # Download Model
                model_buf = io.BytesIO()
                joblib.dump(model_nb, model_buf)
                model_buf.seek(0)
                st.download_button(
                    label="📥 Download Model (model_nb.pkl)",
                    data=model_buf,
                    file_name="model_nb.pkl",
                    use_container_width=True
                )

else:
    st.info("Upload file Excel yang berisi data soal (Nomor Soal, Persentase, Waktu)")

st.caption("Pipeline sesuai flowchart • Centroid tracking & Excel styling ditingkatkan")
