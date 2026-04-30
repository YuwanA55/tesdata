import streamlit as st
import pandas as pd
import numpy as np
from sklearn.cluster import KMeans
from sklearn.model_selection import train_test_split
from sklearn.naive_bayes import GaussianNB
from sklearn.metrics import accuracy_score, classification_report, confusion_matrix
import plotly.express as px
import plotly.figure_factory as ff
import io
import joblib
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ========================= CONFIG =========================
st.set_page_config(page_title="Pipeline K-Means + NB", page_icon="🔬", layout="wide")

st.markdown("""
<style>
    .main-header {font-size: 2.8rem; color: #1E88E5; text-align: center;}
    .success {color: #4CAF50;}
</style>
""", unsafe_allow_html=True)

# ========================= SIDEBAR =========================
with st.sidebar:
    st.title("🔬 Pipeline ML")
    st.markdown("**K-Means + Gaussian Naive Bayes**")
    st.divider()
    st.caption("Versi 1.2 - Kolom Fleksibel")

# ========================= HEADER =========================
st.markdown('<h1 class="main-header">Desain Sistem Pipeline K-Means & Gaussian Naive Bayes</h1>', unsafe_allow_html=True)
st.caption("Mengikuti flowchart yang kamu buat")

# ========================= FILE UPLOAD =========================
uploaded_file = st.file_uploader("📤 Upload File Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    st.success(f"✅ File berhasil dimuat! Jumlah baris: **{len(df)}**")

    # ====================== MAPPING KOLOM OTOMATIS ======================
    col_mapping = {}
    
    for col in df.columns:
        col_lower = str(col).lower().strip()
        if 'nomor' in col_lower or 'soal' in col_lower:
            col_mapping['Nomor soal'] = col
        elif 'persentase' in col_lower or 'persen' in col_lower:
            col_mapping['Persentase'] = col
        elif 'waktu' in col_lower or 'time' in col_lower:
            col_mapping['Waktu'] = col

    # Validasi
    if not all(key in col_mapping for key in ['Nomor soal', 'Persentase', 'Waktu']):
        st.error("❌ Tidak dapat menemukan kolom yang diperlukan.")
        st.write("Kolom yang terdeteksi:", list(df.columns))
        st.stop()

    st.info(f"Kolom yang digunakan → Nomor: **{col_mapping['Nomor soal']}** | Persentase: **{col_mapping['Persentase']}** | Waktu: **{col_mapping['Waktu']}**")

    # ====================== PREPROCESSING ======================
    df_clean = df[[col_mapping['Nomor soal'], 
                   col_mapping['Persentase'], 
                   col_mapping['Waktu']]].copy()
    
    df_clean = df_clean.dropna()
    df_clean = df_clean.rename(columns={
        col_mapping['Nomor soal']: 'Nomor soal',
        col_mapping['Persentase']: 'Persentase',
        col_mapping['Waktu']: 'Waktu'
    })

    X = df_clean[['Persentase', 'Waktu']]
    X_scaled = (X - X.min()) / (X.max() - X.min())

    # ====================== K-MEANS ======================
    init_centroids = np.array([[0.85, 0.15], [0.55, 0.45], [0.20, 0.80]])
    kmeans = KMeans(n_clusters=3, init=init_centroids, n_init=1, max_iter=300, random_state=42)
    df_clean['Cluster'] = kmeans.fit_predict(X_scaled)
    
    cluster_map = {0: 'Mudah', 1: 'Sedang', 2: 'Sulit'}
    df_clean['Label'] = df_clean['Cluster'].map(cluster_map)

    # ====================== NAIVE BAYES ======================
    X_nb = df_clean[['Persentase', 'Waktu']]
    y_nb = df_clean['Label']
    
    X_train, X_test, y_train, y_test = train_test_split(X_nb, y_nb, test_size=0.25, 
                                                        stratify=y_nb, random_state=42)
    
    model_nb = GaussianNB()
    model_nb.fit(X_train, y_train)
    y_pred = model_nb.predict(X_test)
    accuracy = accuracy_score(y_test, y_pred)

    # ====================== TABS ======================
    tab1, tab2, tab3, tab4 = st.tabs(["📊 K-Means", "🤖 Naive Bayes", "📈 Visualisasi", "🔮 Prediksi & Download"])

    with tab1:
        st.subheader("Hasil K-Means Clustering")
        display_cols = ['Nomor soal', 'Persentase', 'Waktu', 'Label']
        st.dataframe(df_clean[display_cols], use_container_width=True)

    with tab2:
        st.metric("Akurasi Model", f"{accuracy:.2%}")
        report_df = pd.DataFrame(classification_report(y_test, y_pred, output_dict=True, zero_division=0)).transpose().round(4)
        st.dataframe(report_df)

    with tab3:
        fig = px.scatter(X_scaled, x='Persentase', y='Waktu', color=df_clean['Label'],
                        title="Visualisasi Hasil Clustering", height=600)
        st.plotly_chart(fig, use_container_width=True)

    with tab4:
        st.subheader("🔮 Prediksi Manual")
        col1, col2 = st.columns(2)
        with col1:
            persen = st.number_input("Persentase", min_value=0.0, max_value=100.0, value=75.0)
        with col2:
            waktu = st.number_input("Waktu (detik)", min_value=0, max_value=300, value=45)
        
        if st.button("Prediksi Tingkat Kesulitan", type="primary"):
            input_data = pd.DataFrame([[persen, waktu]], columns=['Persentase', 'Waktu'])
            prediksi = model_nb.predict(input_data)[0]
            st.success(f"**Prediksi: Soal ini termasuk kategori {prediksi}**")

        st.divider()
        st.subheader("📥 Download Laporan Excel + Model")

        if st.button("🚀 Generate Laporan Excel dengan Warna", type="primary", use_container_width=True):
            with st.spinner("Membuat laporan dengan styling warna..."):
                output = io.BytesIO()
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_clean.to_excel(writer, sheet_name='Data_Lengkap', index=False)
                    df_clean[['Nomor soal', 'Persentase', 'Waktu', 'Label']].to_excel(
                        writer, sheet_name='Hasil_KMeans', index=False)
                    report_df.to_excel(writer, sheet_name='Evaluasi_NB')

                output.seek(0)
                
                # Styling warna
                wb = load_workbook(filename=output)
                ws = wb['Hasil_KMeans']
                
                fill_mudah = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Hijau
                fill_sedang = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid") # Kuning
                fill_sulit = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")   # Merah
                
                for row in range(2, ws.max_row + 1):
                    label_cell = ws.cell(row=row, column=4)  # Kolom Label
                    if label_cell.value == "Mudah":
                        label_cell.fill = fill_mudah
                    elif label_cell.value == "Sedang":
                        label_cell.fill = fill_sedang
                    elif label_cell.value == "Sulit":
                        label_cell.fill = fill_sulit
                
                styled_output = io.BytesIO()
                wb.save(styled_output)
                styled_output.seek(0)

                st.download_button(
                    label="📥 Download Laporan Excel (Diwarnai)",
                    data=styled_output,
                    file_name=f"Laporan_Pipeline_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
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
                    mime="application/octet-stream",
                    use_container_width=True
                )

                st.success("✅ Laporan berhasil dibuat dengan styling warna!")

else:
    st.info("📌 Silakan upload file Excel kamu.")
    st.markdown("**Kolom yang didukung:** `Nomor soal`, `Persentase`, `Waktu` (nama kolom tidak harus persis)")

st.caption("Pipeline sesuai flowchart • Versi kolom fleksibel")
