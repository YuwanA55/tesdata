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
from openpyxl.styles import PatternFill, Font

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
    st.caption("Versi 1.1 - Improved")

# ========================= HEADER =========================
st.markdown('<h1 class="main-header">Desain Sistem Pipeline K-Means & Gaussian Naive Bayes</h1>', unsafe_allow_html=True)

# ========================= FILE UPLOAD =========================
uploaded_file = st.file_uploader("📤 Upload File Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    required_cols = ['Nomor soal', 'Persentase (%)', 'Waktu (detik)']
    if not all(col in df.columns for col in required_cols):
        st.error(f"Kolom harus ada: {required_cols}")
        st.stop()

    df_clean = df.dropna(subset=['Persentase (%)', 'Waktu (detik)']).copy()
    X = df_clean[['Persentase (%)', 'Waktu (detik)']]
    X_scaled = (X - X.min()) / (X.max() - X.min())

    # ====================== K-MEANS ======================
    init_centroids = np.array([[0.85, 0.15], [0.55, 0.45], [0.20, 0.80]])
    kmeans = KMeans(n_clusters=3, init=init_centroids, n_init=1, max_iter=300, random_state=42)
    df_clean['Cluster'] = kmeans.fit_predict(X_scaled)
    cluster_map = {0: 'Mudah', 1: 'Sedang', 2: 'Sulit'}
    df_clean['Label'] = df_clean['Cluster'].map(cluster_map)

    # ====================== NAIVE BAYES ======================
    X_nb = df_clean[['Persentase (%)', 'Waktu (detik)']]
    y_nb = df_clean['Label']
    
    X_train, X_test, y_train, y_test = train_test_split(X_nb, y_nb, test_size=0.25, 
                                                        stratify=y_nb, random_state=42)
    
    model_nb = GaussianNB()
    model_nb.fit(X_train, y_train)
    y_pred = model_nb.predict(X_test)
    accuracy = accuracy_score(y_test, y_pred)

    # ====================== TABS ======================
    tab1, tab2, tab3, tab4 = st.tabs(["📊 K-Means", "🤖 Naive Bayes", "📈 Visualisasi", "🔮 Prediksi Manual & Download"])

    with tab1:
        st.subheader("Hasil K-Means Clustering")
        st.dataframe(df_clean[['Nomor soal', 'Persentase (%)', 'Waktu (detik)', 'Label']], use_container_width=True)

    with tab2:
        st.metric("Akurasi Model Gaussian Naive Bayes", f"{accuracy:.2%}")
        report_df = pd.DataFrame(classification_report(y_test, y_pred, output_dict=True)).transpose().round(4)
        st.dataframe(report_df)

    with tab3:
        fig = px.scatter(X_scaled, x='Persentase (%)', y='Waktu (detik)', color=df_clean['Label'],
                        title="Visualisasi Cluster Soal", height=600)
        st.plotly_chart(fig, use_container_width=True)

    with tab4:
        st.subheader("🔮 Prediksi Manual")
        col1, col2 = st.columns(2)
        with col1:
            persen = st.number_input("Persentase (%)", min_value=0.0, max_value=100.0, value=75.0)
        with col2:
            waktu = st.number_input("Waktu (detik)", min_value=0, max_value=300, value=45)
        
        if st.button("Prediksi Label Soal", type="primary"):
            input_data = pd.DataFrame([[persen, waktu]], columns=['Persentase (%)', 'Waktu (detik)'])
            prediksi = model_nb.predict(input_data)[0]
            st.success(f"**Soal ini diprediksi: {prediksi}**")

        st.divider()
        st.subheader("📥 Download Laporan & Model")

        if st.button("🚀 Generate Laporan Excel Lengkap + Model", type="primary", use_container_width=True):
            with st.spinner("Membuat file Excel dengan styling..."):
                output = io.BytesIO()
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_clean.to_excel(writer, sheet_name='Data_Lengkap', index=False)
                    df_clean[['Nomor soal','Persentase (%)','Waktu (detik)','Label']].to_excel(
                        writer, sheet_name='Hasil_KMeans', index=False)
                    report_df.to_excel(writer, sheet_name='Evaluasi_NB', index=True)
                    
                    summary = pd.DataFrame({
                        'Metrik': ['Total Soal', 'Akurasi', 'Tanggal'],
                        'Nilai': [len(df_clean), f"{accuracy:.2%}", datetime.now().strftime("%Y-%m-%d %H:%M")]
                    })
                    summary.to_excel(writer, sheet_name='Ringkasan', index=False)

                output.seek(0)
                
                # Styling Excel (Warna per Label)
                wb = load_workbook(filename=output)
                ws = wb['Hasil_KMeans']
                
                # Warna fill
                fill_mudah = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                fill_sedang = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
                fill_sulit = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                
                for row in ws.iter_rows(min_row=2, min_col=4, max_col=4):  # Kolom Label
                    for cell in row:
                        if cell.value == "Mudah":
                            cell.fill = fill_mudah
                        elif cell.value == "Sedang":
                            cell.fill = fill_sedang
                        elif cell.value == "Sulit":
                            cell.fill = fill_sulit
                
                # Simpan kembali ke buffer
                styled_output = io.BytesIO()
                wb.save(styled_output)
                styled_output.seek(0)

                st.download_button(
                    label="📥 Download Laporan Excel (Sudah diwarnai)",
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
                    label="📥 Download Model Naive Bayes (model_nb.pkl)",
                    data=model_buf,
                    file_name="model_nb.pkl",
                    mime="application/octet-stream",
                    use_container_width=True
                )

                st.success("✅ Laporan Excel dengan styling warna berhasil dibuat!")

else:
    st.info("Silakan upload file Excel yang berisi kolom: **Nomor soal**, **Persentase (%)**, dan **Waktu (detik)**")

st.caption("Pipeline sesuai flowchart yang kamu buat • Prediksi Manual + Excel Styling + Dark Theme Ready")
