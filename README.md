# Pipeline K-Means + Gaussian Naive Bayes

Aplikasi web untuk mengklasifikasikan tingkat kesulitan soal menggunakan **K-Means Clustering** dan **Gaussian Naive Bayes**.

### Fitur
- Upload file Excel
- Preprocessing otomatis
- K-Means dengan centroid awal deterministik (Mudah, Sedang, Sulit)
- Gaussian Naive Bayes dengan stratified split 75:25
- Evaluasi lengkap (Akurasi, F1-Score, Confusion Matrix)
- Visualisasi interaktif (Plotly)
- Download laporan Excel multi-sheet + model `.pkl`

### Cara Menjalankan Lokal
```bash
pip install -r requirements.txt
streamlit run app.py
