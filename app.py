import streamlit as st
import pandas as pd
import numpy as np
from sklearn.cluster import KMeans
from sklearn.model_selection import train_test_split
from sklearn.naive_bayes import GaussianNB
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import accuracy_score, classification_report, confusion_matrix
import plotly.express as px
import io, joblib
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.chart import BarChart, Reference

# ================= CONFIG =================
st.set_page_config(page_title="Pipeline Skripsi ML", layout="wide")
st.title("🎓 Pipeline K-Means + Machine Learning")
st.caption("Versi Skripsi (Lengkap + Evaluasi Model)")

uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)
    st.success(f"Data: {len(df)} baris")

    # ================= AUTO DETECT =================
    col_map = {}
    for col in df.columns:
        c = col.lower()
        if 'soal' in c:
            col_map['Nomor soal'] = col
        elif 'persen' in c:
            col_map['Persentase'] = col
        elif 'waktu' in c:
            col_map['Waktu'] = col

    if len(col_map) < 3:
        st.error("Kolom tidak lengkap")
        st.stop()

    # ================= PREPROCESS =================
    df_clean = df[[col_map['Nomor soal'],
                   col_map['Persentase'],
                   col_map['Waktu']]].dropna()

    df_clean.columns = ['Nomor soal', 'Persentase', 'Waktu']

    X = df_clean[['Persentase', 'Waktu']]
    X_scaled = (X - X.min()) / (X.max() - X.min())

    # ================= KMEANS AUTO =================
    kmeans = KMeans(n_clusters=3, init='k-means++', random_state=42)
    df_clean['Cluster'] = kmeans.fit_predict(X_scaled)

    # Mapping otomatis
    cluster_mean = df_clean.groupby('Cluster')['Persentase'].mean().sort_values()

    mapping = {
        cluster_mean.index[0]: 'Sulit',
        cluster_mean.index[1]: 'Sedang',
        cluster_mean.index[2]: 'Mudah'
    }

    df_clean['Label'] = df_clean['Cluster'].map(mapping)

    # ================= SPLIT =================
    X_nb = df_clean[['Persentase', 'Waktu']]
    y_nb = df_clean['Label']

    X_train, X_test, y_train, y_test = train_test_split(
        X_nb, y_nb, test_size=0.25, stratify=y_nb, random_state=42
    )

    # ================= MODEL =================
    nb = GaussianNB()
    dt = DecisionTreeClassifier(random_state=42)

    nb.fit(X_train, y_train)
    dt.fit(X_train, y_train)

    pred_nb = nb.predict(X_test)
    pred_dt = dt.predict(X_test)

    # ================= METRICS =================
    def get_metrics(y_true, y_pred):
        report = classification_report(y_true, y_pred, output_dict=True, zero_division=0)
        return {
            "Accuracy": accuracy_score(y_true, y_pred),
            "Precision": report['weighted avg']['precision'],
            "Recall": report['weighted avg']['recall'],
            "F1-Score": report['weighted avg']['f1-score']
        }

    metrics_nb = get_metrics(y_test, pred_nb)
    metrics_dt = get_metrics(y_test, pred_dt)

    # ================= CONFUSION =================
    cm_nb = pd.DataFrame(
        confusion_matrix(y_test, pred_nb, labels=nb.classes_),
        index=nb.classes_,
        columns=nb.classes_
    )

    # ================= TABS =================
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "K-Means",
        "Perbandingan Model",
        "📊 Evaluasi Model",
        "Visualisasi",
        "Prediksi & Export"
    ])

    # ===== KMEANS =====
    with tab1:
        st.dataframe(df_clean, use_container_width=True)

    # ===== PERBANDINGAN =====
    with tab2:
        comp_df = pd.DataFrame({
            "Model": ["Naive Bayes", "Decision Tree"],
            "Accuracy": [metrics_nb["Accuracy"], metrics_dt["Accuracy"]]
        })
        st.dataframe(comp_df)
        st.bar_chart(comp_df.set_index("Model"))

    # ===== EVALUASI =====
    with tab3:
        st.subheader("Evaluasi Model")

        eval_df = pd.DataFrame({
            "Metric": ["Accuracy", "Precision", "Recall", "F1-Score"],
            "Naive Bayes": list(metrics_nb.values()),
            "Decision Tree": list(metrics_dt.values())
        })

        st.dataframe(eval_df, use_container_width=True)

        st.subheader("Grafik Evaluasi")
        st.bar_chart(eval_df.set_index("Metric"))

        st.subheader("Confusion Matrix (Naive Bayes)")
        st.dataframe(cm_nb)

    # ===== VISUAL =====
    with tab4:
        fig = px.scatter(
            X_scaled,
            x='Persentase',
            y='Waktu',
            color=df_clean['Label'],
            title="Clustering"
        )
        st.plotly_chart(fig, use_container_width=True)

    # ===== PREDIKSI =====
    with tab5:

        st.subheader("Prediksi Manual")

        col1, col2, col3 = st.columns(3)

        with col1:
            jml = st.number_input("Jumlah Siswa", 1, 100, 10)
        with col2:
            benar = st.number_input("Jumlah Benar", 0, 100, 7)
        with col3:
            waktu = st.number_input("Waktu", 0, 300, 45)

        if st.button("Prediksi"):
            if benar > jml:
                st.error("Input tidak valid")
            else:
                persen = (benar / jml) * 100
                input_df = pd.DataFrame([[persen, waktu]],
                                        columns=['Persentase', 'Waktu'])

                hasil = nb.predict(input_df)[0]

                st.info(f"Persentase: {persen:.2f}%")
                st.success(f"Hasil: {hasil}")

        st.divider()

        # ===== EXPORT EXCEL =====
        if st.button("Generate Excel Lengkap"):

            output = io.BytesIO()

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Input')
                df_clean.to_excel(writer, sheet_name='KMeans')

                comp_df.to_excel(writer, sheet_name='Perbandingan', index=False)
                eval_df.to_excel(writer, sheet_name='Evaluasi', index=False)
                cm_nb.to_excel(writer, sheet_name='Confusion')

            output.seek(0)
            wb = load_workbook(output)

            # ===== WARNA =====
            ws = wb['KMeans']
            fill = {
                "Mudah": PatternFill(start_color="C6EFCE", fill_type="solid"),
                "Sedang": PatternFill(start_color="FFEB9C", fill_type="solid"),
                "Sulit": PatternFill(start_color="FFC7CE", fill_type="solid"),
            }

            for r in range(2, ws.max_row + 1):
                val = ws.cell(row=r, column=5).value
                if val in fill:
                    ws.cell(row=r, column=5).fill = fill[val]

            # ===== GRAFIK EXCEL =====
            ws_chart = wb['Perbandingan']
            chart = BarChart()
            data = Reference(ws_chart, min_col=2, min_row=1, max_row=3)
            cats = Reference(ws_chart, min_col=1, min_row=2, max_row=3)

            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            chart.title = "Perbandingan Model"

            ws_chart.add_chart(chart, "E5")

            # ===== AUTO WIDTH =====
            for s in wb.sheetnames:
                ws = wb[s]
                for col in ws.columns:
                    max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                    ws.column_dimensions[col[0].column_letter].width = max_len + 2

            final = io.BytesIO()
            wb.save(final)
            final.seek(0)

            st.download_button(
                "Download Excel",
                final,
                file_name=f"skripsi_{datetime.now().strftime('%H%M')}.xlsx"
            )

            # MODEL
            buf = io.BytesIO()
            joblib.dump(nb, buf)
            buf.seek(0)

            st.download_button(
                "Download Model NB",
                buf,
                file_name="model_nb.pkl"
            )

else:
    st.info("Upload file Excel dulu")
