import streamlit as st
import pandas as pd
import numpy as np
from sklearn.cluster import KMeans
from sklearn.model_selection import train_test_split
from sklearn.naive_bayes import GaussianNB
from sklearn.metrics import (
    accuracy_score, 
    classification_report, 
    confusion_matrix
)
import plotly.express as px
import plotly.figure_factory as ff
import io
import joblib
from datetime import datetime

# ========================= CONFIGURATION =========================
st.set_page_config(
    page_title="Pipeline K-Means + Gaussian NB",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS untuk tampilan lebih profesional
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E88E5;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .subheader {
        color: #424242;
        font-weight: 500;
    }
</style>
""", unsafe_allow_html=True)

# ========================= SIDEBAR =========================
with st.sidebar:
    st.image("https://via.placeholder.com/150x150.png?text=Pipeline", width=150)
    st.title("Tentang Aplikasi")
    st.markdown("""
    **Pipeline Machine Learning**  
    - K-Means Clustering (Deterministic Centroid)  
    - Gaussian Naive Bayes Classifier  
    - Evaluasi Lengkap + Visualisasi  
    """)
    
    st.divider()
    st.caption("Versi 1.0 | Dibuat dengan ❤️ menggunakan Streamlit")

# ========================= MAIN TITLE =========================
st.markdown('<h1 class="main-header">🔬 Desain Sistem Pipeline K-Means & Gaussian Naive Bayes</h1>', unsafe_allow_html=True)
st.markdown("**Mengikuti flowchart dan desain sistem yang telah dibuat**")

# ========================= FILE UPLOADER =========================
uploaded_file = st.file_uploader(
    "📤 Upload File Excel (.xlsx)", 
    type=["xlsx"],
    help="File harus memiliki kolom: Nomor soal, Persentase (%), Waktu (detik)"
)

if uploaded_file is not None:
    try:
        # Baca data
        df = pd.read_excel(uploaded_file)
        
        st.success(f"✅ File berhasil dimuat! Jumlah soal: **{len(df)}** baris")
        
        # Validasi kolom
        required_cols = ['Nomor soal', 'Persentase (%)', 'Waktu (detik)']
        if not all(col in df.columns for col in required_cols):
            st.error(f"❌ Kolom yang dibutuhkan: {required_cols}")
            st.stop()
        
        # Preprocessing
        df_clean = df.dropna(subset=['Persentase (%)', 'Waktu (detik)']).copy()
        X = df_clean[['Persentase (%)', 'Waktu (detik)']]
        
        # Normalisasi Min-Max
        X_scaled = (X - X.min()) / (X.max() - X.min())
        
        # Tab layout
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Preprocessing & K-Means", 
                                         "🤖 Gaussian Naive Bayes", 
                                         "📈 Visualisasi", 
                                         "📥 Hasil & Download"])

        with tab1:
            st.subheader("1. Preprocessing")
            st.dataframe(df_clean.head(), use_container_width=True)
            
            st.subheader("2. K-Means Clustering")
            st.markdown("**Centroid Awal Deterministik:** C1=Mudah, C2=Sedang, C3=Sulit")
            
            # Centroid awal
            init_centroids = np.array([
                [0.85, 0.15],   # Mudah
                [0.55, 0.45],   # Sedang
                [0.20, 0.80]    # Sulit
            ])
            
            kmeans = KMeans(n_clusters=3, init=init_centroids, n_init=1, 
                           max_iter=300, random_state=42)
            df_clean['Cluster'] = kmeans.fit_predict(X_scaled)
            
            cluster_map = {0: 'Mudah', 1: 'Sedang', 2: 'Sulit'}
            df_clean['Label'] = df_clean['Cluster'].map(cluster_map)
            
            col1, col2 = st.columns(2)
            with col1:
                st.write("**Distribusi Label**")
                st.bar_chart(df_clean['Label'].value_counts())
            with col2:
                st.write("**Jumlah per Label**")
                st.dataframe(df_clean['Label'].value_counts().reset_index())

        with tab2:
            st.subheader("Gaussian Naive Bayes")
            
            X_nb = df_clean[['Persentase (%)', 'Waktu (detik)']]
            y_nb = df_clean['Label']
            
            # Stratified Split 75:25
            X_train, X_test, y_train, y_test = train_test_split(
                X_nb, y_nb, test_size=0.25, stratify=y_nb, random_state=42
            )
            
            # Training
            model_nb = GaussianNB()
            model_nb.fit(X_train, y_train)
            
            # Prediksi & Evaluasi
            y_pred = model_nb.predict(X_test)
            accuracy = accuracy_score(y_test, y_pred)
            
            st.metric(label="**Akurasi Model**", value=f"{accuracy:.2%}", delta=None)
            
            report = classification_report(y_test, y_pred, output_dict=True, zero_division=0)
            report_df = pd.DataFrame(report).transpose().round(4)
            
            col_a, col_b = st.columns(2)
            with col_a:
                st.subheader("Classification Report")
                st.dataframe(report_df)
            with col_b:
                st.subheader("Confusion Matrix")
                cm = confusion_matrix(y_test, y_pred, labels=['Mudah', 'Sedang', 'Sulit'])
                fig_cm = ff.create_annotated_heatmap(
                    cm, x=['Mudah','Sedang','Sulit'], y=['Mudah','Sedang','Sulit'],
                    colorscale='Blues', showscale=True
                )
                st.plotly_chart(fig_cm, use_container_width=True)

        with tab3:
            st.subheader("Visualisasi Cluster")
            fig = px.scatter(
                X_scaled, x='Persentase (%)', y='Waktu (detik)',
                color=df_clean['Label'],
                title="Visualisasi Hasil K-Means Clustering",
                labels={'color': 'Label Kesulitan'},
                hover_data={'index': df_clean.index}
            )
            fig.update_layout(height=600)
            st.plotly_chart(fig, use_container_width=True)

        with tab4:
            st.subheader("Generate Laporan Lengkap")
            
            if st.button("🚀 Generate & Download Laporan Excel + Model", type="primary", use_container_width=True):
                with st.spinner("Sedang membuat laporan lengkap..."):
                    output = io.BytesIO()
                    
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_clean.to_excel(writer, sheet_name='Data_Lengkap', index=False)
                        df_clean[['Nomor soal', 'Persentase (%)', 'Waktu (detik)', 'Label']].to_excel(
                            writer, sheet_name='Hasil_KMeans', index=False)
                        report_df.to_excel(writer, sheet_name='Evaluasi_NB')
                        
                        # Ringkasan
                        summary = pd.DataFrame({
                            'Metrik': ['Total Soal', 'Akurasi Model', 'Tanggal Proses', 'Versi Pipeline'],
                            'Nilai': [len(df_clean), f"{accuracy:.2%}", 
                                     datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "1.0"]
                        })
                        summary.to_excel(writer, sheet_name='Ringkasan', index=False)
                    
                    output.seek(0)
                    
                    # Download Excel
                    st.download_button(
                        label="📥 Download Laporan Excel Lengkap",
                        data=output,
                        file_name=f"Laporan_Pipeline_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    # Download Model
                    model_buffer = io.BytesIO()
                    joblib.dump(model_nb, model_buffer)
                    model_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Download Model (model_nb.pkl)",
                        data=model_buffer,
                        file_name="model_nb.pkl",
                        mime="application/octet-stream",
                        use_container_width=True
                    )
                    
                    st.success("✅ Laporan dan model berhasil dibuat!")

    except Exception as e:
        st.error(f"Terjadi kesalahan: {str(e)}")
        st.info("Pastikan file Excel Anda memiliki kolom yang benar dan tidak rusak.")

else:
    st.info("👆 Silakan upload file Excel untuk memulai pipeline.")
    st.markdown("""
    **Kolom yang diperlukan:**
    - `Nomor soal`
    - `Persentase (%)`
    - `Waktu (detik)`
    """)

st.divider()
st.caption("Pipeline sesuai flowchart: Input → Preprocessing → K-Means (Centroid Deterministik) → Gaussian Naive Bayes → Output Excel + Model + Grafik")
