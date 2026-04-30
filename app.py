import streamlit as st
import pandas as pd
import numpy as np
from sklearn.cluster import KMeans
from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.naive_bayes import GaussianNB
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import accuracy_score, classification_report, confusion_matrix
import plotly.express as px

st.set_page_config(layout="wide")
st.title("🎓 Pipeline ML (Validasi Lengkap)")

uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)
    st.success(f"Jumlah data: {len(df)}")

    # ================= DETEKSI KOLOM =================
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

    # ================= KMEANS =================
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
    X_model = df_clean[['Persentase', 'Waktu']]
    y_model = df_clean['Label']

    X_train, X_test, y_train, y_test = train_test_split(
        X_model, y_model, test_size=0.3, stratify=y_model, random_state=42
    )

    # ================= MODEL =================
    nb = GaussianNB()
    dt = DecisionTreeClassifier(random_state=42)

    nb.fit(X_train, y_train)
    dt.fit(X_train, y_train)

    pred_train_nb = nb.predict(X_train)
    pred_test_nb = nb.predict(X_test)

    pred_train_dt = dt.predict(X_train)
    pred_test_dt = dt.predict(X_test)

    # ================= METRICS =================
    def metrics(y_true, y_pred):
        report = classification_report(y_true, y_pred, output_dict=True, zero_division=0)
        return {
            "Accuracy": accuracy_score(y_true, y_pred),
            "Precision": report['weighted avg']['precision'],
            "Recall": report['weighted avg']['recall'],
            "F1": report['weighted avg']['f1-score']
        }

    nb_train = metrics(y_train, pred_train_nb)
    nb_test = metrics(y_test, pred_test_nb)

    dt_train = metrics(y_train, pred_train_dt)
    dt_test = metrics(y_test, pred_test_dt)

    # ================= CROSS VALIDATION =================
    cv_nb = cross_val_score(nb, X_model, y_model, cv=5).mean()
    cv_dt = cross_val_score(dt, X_model, y_model, cv=5).mean()

    # ================= DETEKSI OVERFITTING =================
    def status(train, test):
        if train - test > 0.1:
            return "⚠ Overfitting"
        elif test > 0.95:
            return "⚠ Terlalu sempurna (cek data)"
        else:
            return "Normal"

    status_nb = status(nb_train["Accuracy"], nb_test["Accuracy"])
    status_dt = status(dt_train["Accuracy"], dt_test["Accuracy"])

    # ================= CONFUSION =================
    cm_nb = pd.DataFrame(confusion_matrix(y_test, pred_test_nb),
                         index=nb.classes_, columns=nb.classes_)

    # ================= NAVBAR =================
    tab1, tab2, tab3, tab4 = st.tabs([
        "K-Means",
        "Perbandingan",
        "📊 Evaluasi Model",
        "Visualisasi"
    ])

    # ===== KMEANS =====
    with tab1:
        st.dataframe(df_clean, use_container_width=True)

    # ===== PERBANDINGAN =====
    with tab2:
        comp = pd.DataFrame({
            "Model": ["Naive Bayes", "Decision Tree"],
            "Test Accuracy": [nb_test["Accuracy"], dt_test["Accuracy"]],
            "Cross Val": [cv_nb, cv_dt],
            "Status": [status_nb, status_dt]
        })
        st.dataframe(comp)
        st.bar_chart(comp.set_index("Model")[["Test Accuracy", "Cross Val"]])

    # ===== EVALUASI =====
    with tab3:
        st.subheader("Naive Bayes")
        st.write("Train vs Test")
        st.write(nb_train, nb_test)

        st.subheader("Decision Tree")
        st.write(dt_train, dt_test)

        st.subheader("Confusion Matrix")
        st.dataframe(cm_nb)

    # ===== VISUAL =====
    with tab4:
        fig = px.scatter(
            X_scaled,
            x='Persentase',
            y='Waktu',
            color=df_clean['Label']
        )
        st.plotly_chart(fig, use_container_width=True)

else:
    st.info("Upload file Excel terlebih dahulu")
