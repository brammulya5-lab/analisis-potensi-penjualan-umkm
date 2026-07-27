import io
from pathlib import Path

import joblib
import pandas as pd
import plotly.express as px
import streamlit as st
import xgboost as xgb
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# config halaman biar full width dan sidebar kebuka
st.set_page_config(
    page_title="Analisis Potensi Penjualan UMKM",
    layout="wide",
    initial_sidebar_state="expanded"
)

# custom styling css biar tampilan dashboard & sidebar makin rapi dan simetris
st.markdown("""
    <style>
    .stApp { background-color: #0b1329; color: #f1f5f9; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    
    .kartu-utama {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        border: 1px solid #334155;
        padding: 28px;
        border-radius: 16px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.3);
        margin-bottom: 24px;
    }
    
    div[data-testid="stMetric"] {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        border: 1px solid #334155;
        padding: 14px 16px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    div[data-testid="stMetric"] label { color: #94a3b8 !important; font-weight: 600; text-transform: uppercase; font-size: 0.7rem; letter-spacing: 0.03em; }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] { color: #f8fafc !important; font-size: 1.4rem; font-weight: 700; }
    
    .stButton > button {
        border-radius: 10px; font-weight: 600;
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%);
        color: white; border: none; padding: 0.65rem 1.4rem;
        box-shadow: 0 4px 6px -1px rgba(2, 132, 199, 0.3);
        transition: all 0.3s ease;
        width: 100%;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #0369a1 100%, #075985 100%);
    }
    
    section[data-testid="stSidebar"] {
        background-color: #070a14;
        border-right: 1px solid #1e293b;
        padding-top: 1rem;
    }
    
    .kotak-judul-sidebar {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        border: 1px solid #334155;
        padding: 16px;
        border-radius: 10px;
        margin-bottom: 20px;
        text-align: center;
    }
    .kotak-judul-sidebar h3 {
        color: #f8fafc;
        font-size: 1.1rem;
        margin: 0 0 4px 0;
        font-weight: 700;
    }
    .kotak-judul-sidebar p {
        color: #94a3b8;
        font-size: 0.7rem;
        margin: 0;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }

    section[data-testid="stSidebar"] .stFileUploader section { pointer-events: none; }
    section[data-testid="stSidebar"] .stFileUploader button { pointer-events: auto; }
    div[data-testid="stFileUploader"] section { pointer-events: none; }
    div[data-testid="stFileUploader"] button { pointer-events: auto; }

    /* Styling Menu Navigasi Sidebar - Diperkuat agar posisinya benar-benar pas di tengah */
    section[data-testid="stSidebar"] .stRadio > label { display: none; }
    section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] { gap: 8px; }
    section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {
        background-color: #162032;
        padding: 10px 14px;
        border-radius: 8px;
        border: 1px solid #283548;
        transition: all 0.2s ease-in-out;
        width: 100%;
        display: flex !important;
        align-items: center !important;
        cursor: pointer;
    }
    section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover {
        background-color: #1e293b;
        border-color: #0284c7;
    }
    section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label div {
        display: flex !important;
        align-items: center !important;
    }
    section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label p {
        margin: 0 !important;
        padding: 0 !important;
        font-size: 0.9rem;
        font-weight: 500;
        color: #f1f5f9;
        line-height: normal !important;
    }
    
    hr { border-color: #334155; margin: 1.5rem 0; }
    </style>
""", unsafe_allow_html=True)

# inisialisasi path file model & encoder
MODEL_PATH = "models/model_terbaik.json"
ENCODER_PATH = "models/label_encoder.pkl"

# fitur yang dipake model machine learning kita
MODEL_FEATURES = ["Frekuensi", "Rata_Qty", "Rata_Harga"]

# kamus mapping rekomendasi stok
RECOMMENDATION_MAPPING = {
    "Potensi Tinggi": "Tingkatkan Stok",
    "Potensi Sedang": "Pertahankan Stok",
    "Potensi Rendah": "Kurangi Stok",
}

# palet warna chart plotly
CHART_COLOR_MAP = {
    "Potensi Tinggi": "#10b981",
    "Potensi Sedang": "#f59e0b",
    "Potensi Rendah": "#ef4444",
}

# load model xgboost & label encoder
@st.cache_resource
def load_trained_model():
    model = xgb.XGBClassifier()
    model.load_model(MODEL_PATH)
    encoder = joblib.load(ENCODER_PATH)
    return model, encoder

# auto deteksi nama kolom excel jaga-jaga kalau beda penulisan
def auto_detect_columns(raw_df):
    clean_df = raw_df.loc[:, ~raw_df.columns.duplicated()].copy()
    col_mapping = {}
    mapped_check = set()
    
    for raw_col in clean_df.columns:
        low_col = str(raw_col).strip().lower()
        
        if "tanggal" not in mapped_check and any(kw in low_col for kw in ["tanggal", "tgl", "date", "faktur", "waktu", "hari"]):
            col_mapping[raw_col] = "Tanggal Faktur"
            mapped_check.add("tanggal")
        elif "barang" not in mapped_check and any(kw in low_col for kw in ["barang", "produk", "item", "nama", "product", "menu"]):
            col_mapping[raw_col] = "Nama Barang"
            mapped_check.add("barang")
        elif "qty" not in mapped_check and any(kw in low_col for kw in ["qty", "jumlah", "terjual", "quantity", "jml", "count", "banyak", "pcs", "volume"]):
            col_mapping[raw_col] = "Qty"
            mapped_check.add("qty")
        elif "harga" not in mapped_check and any(kw in low_col for kw in ["harga", "price", "satuan", "jual", "cost", "nilai"]):
            col_mapping[raw_col] = "Harga Satuan"
            mapped_check.add("harga")
            
    clean_df = clean_df.rename(columns=col_mapping)
    return clean_df

# fungsi utama buat pembersihan data (handle duplikat & missing value)
@st.cache_data
def clean_transaction_data(input_df):
    processed_df = auto_detect_columns(input_df)
    required_cols = ["Tanggal Faktur", "Nama Barang", "Qty", "Harga Satuan"]
    
    for col in required_cols:
        if col not in processed_df.columns:
            raise ValueError(f"Kolom wajib '{col}' tidak ditemukan. Cek lagi header excelnya ya.")
            
    processed_df = processed_df.loc[:, ~processed_df.columns.duplicated()].copy()
            
    duplicates_qty = int(processed_df.duplicated().sum())
    if duplicates_qty > 0:
        processed_df = processed_df.drop_duplicates()
        
    processed_df["Nama Barang"] = processed_df["Nama Barang"].astype(str).str.strip()
    processed_df["Tanggal Faktur"] = pd.to_datetime(processed_df["Tanggal Faktur"], errors="coerce")
    processed_df["Qty"] = pd.to_numeric(processed_df["Qty"], errors="coerce")
    processed_df["Harga Satuan"] = pd.to_numeric(processed_df["Harga Satuan"], errors="coerce")
    
    missing_qty = int(processed_df.isnull().sum().sum())
    processed_df = processed_df.dropna(subset=required_cols)
    processed_df = processed_df[(processed_df["Qty"] > 0) & (processed_df["Harga Satuan"] > 0)]
    
    summary_info = {
        "duplikat_dihapus": duplicates_qty,
        "missing_value": missing_qty,
        "total_bersih": len(processed_df)
    }
    return processed_df, summary_info

# agregasi data per produk buat dapetin fitur frekuensi & rata-rata harga
@st.cache_data
def extract_product_features(clean_df):
    features_df = clean_df.groupby("Nama Barang").agg(
        Total_Qty=("Qty", "sum"),
        Rata_Qty=("Qty", "mean"),
        Frekuensi=("Qty", "count"),
        Rata_Harga=("Harga Satuan", "mean")
    ).reset_index()
    return features_df

# helper buat prediksi model klasifikasi
def run_model_prediction(model, encoder, df):
    return pd.Series(
        encoder.inverse_transform(model.predict(df[MODEL_FEATURES]).astype(int)), 
        index=df.index
    )

# sidebar navigasi menu aplikasi dengan sedikit sentuhan ikon natural
st.sidebar.markdown("""
    <div class="kotak-judul-sidebar">
        <h3>📌 Tahapan Analisis</h3>
        <p>Sistem Klasifikasi Stok</p>
    </div>
""", unsafe_allow_html=True)

nav = st.sidebar.radio(
    "Navigasi:",
    [
        "1. Upload Data", 
        "2. Pembersihan Data", 
        "3. Pembentukan Fitur", 
        "4. Prediksi & Klasifikasi", 
        "5. Dashboard & Visualisasi",
        "6. Unduh Laporan"
    ],
    label_visibility="collapsed"
)

# inisialisasi session state biar data gak ke-reset pas pindah menu
for state_name in ["data_mentah", "data_bersih", "info_bersih", "data_fitur", "data_hasil", "nama_file_aktif"]:
    if state_name not in st.session_state:
        st.session_state[state_name] = None


# MENU 1: UPLOAD DATA 
if nav == "1. Upload Data":
    st.markdown("""
        <div class="kartu-utama">
            <h1 style="color: #f8fafc; font-size: 1.8rem; font-weight: 800; margin-bottom: 0.5rem; line-height: 1.3;">Sistem Pendukung Keputusan Analisis Potensi Penjualan Produk UMKM</h1>
            <p style="color: #94a3b8; font-size: 1rem; margin-bottom: 0; line-height: 1.5;">Platform berbasis sistem cerdas yang dirancang untuk membantu pelaku usaha dalam mengelompokkan tingkat potensi produk dan mengelola persediaan stok secara objektif.</p>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
        <div class="kartu-utama">
            <h3 style="color: #f8fafc; margin-top: 0; font-size: 1.2rem;">📁 Unggah Laporan Data Penjualan</h3>
            <p style="color: #94a3b8; margin-bottom: 1.0rem; font-size: 0.92rem;">Silahkan masukkan dokumen laporan data penjualan berformat <b>Excel (.xlsx)</b>. Sistem secara otomatis mendeteksi nama kolom dari berbagai format data UMKM tanpa perlu pengaturan manual.</p>
    """, unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("Pilih file Excel laporan data penjualan (.xlsx)", type=["xlsx"], key="uploader_excel")
    
    if uploaded_file is not None:
        try:
            if st.session_state["nama_file_aktif"] != uploaded_file.name:
                raw_df = pd.read_excel(uploaded_file)
                raw_df.columns = raw_df.columns.str.strip()
                
                detected_df = auto_detect_columns(raw_df)
                st.session_state["data_mentah"] = detected_df
                st.session_state["nama_file_aktif"] = uploaded_file.name
        except Exception as e:
            st.error(f"Gagal membaca file excel: {e}")

    st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state["data_mentah"] is not None:
        df_preview = st.session_state["data_mentah"]
        
        item_col = df_preview["Nama Barang"] if "Nama Barang" in df_preview.columns else df_preview.iloc[:, 1]
        if isinstance(item_col, pd.DataFrame):
            item_col = item_col.iloc[:, 0]
        total_unique_items = int(item_col.nunique())
        
        qty_col = df_preview["Qty"] if "Qty" in df_preview.columns else df_preview.iloc[:, 2]
        if isinstance(qty_col, pd.DataFrame):
            qty_col = qty_col.iloc[:, 0]
        sum_qty_all = int(pd.to_numeric(qty_col, errors='coerce').sum())
        
        date_col = df_preview["Tanggal Faktur"] if "Tanggal Faktur" in df_preview.columns else df_preview.iloc[:, 0]
        if isinstance(date_col, pd.DataFrame):
            date_col = date_col.iloc[:, 0]
        valid_dates = pd.to_datetime(date_col, errors="coerce").dropna()
        date_range_str = f"{valid_dates.min():%d/%m/%Y} - {valid_dates.max():%d/%m/%Y}" if not valid_dates.empty else "Tidak terdeteksi"
        
        st.success(f"File '{st.session_state['nama_file_aktif']}' berhasil di upload!")
            
        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("TOTAL BARIS", f"{len(df_preview):,}")
        col2.metric("JUMLAH PRODUK", f"{total_unique_items:,}")
        col3.metric("TOTAL QTY", f"{sum_qty_all:,}")
        col4.metric("PERIODE", date_range_str)
        
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("Tampilkan Pratinjau 10 Baris Pertama Data"):
            st.dataframe(df_preview.head(10), use_container_width=True, hide_index=True)
    else:
        st.info("Silahkan upload file laporan data penjualan terlebih dahulu.")


# MENU 2: PEMBERSIHAN DATA 
elif nav == "2. Pembersihan Data":
    st.markdown("""
        <div class="kartu-utama">
            <h3 style="color: #f8fafc; margin-top: 0; font-size: 1.2rem;">🧹 Pembersihan Data (Data Cleaning)</h3>
            <p style="color: #94a3b8; margin-bottom: 1.2rem; font-size: 0.92rem;"><b>Fungsi Menu Ini:</b> Memastikan kualitas dataset dengan menyaring anomali, menghapus data ganda (duplikat), merapihkan spasi penulisan teks, serta menyeleksi baris transaksi yang bernilai valid.</p>
    """, unsafe_allow_html=True)
    
    if st.session_state["data_mentah"] is not None:
        if st.button("✨ Jalankan Proses Pembersihan"):
            clean_df, info_dict = clean_transaction_data(st.session_state["data_mentah"])
            st.session_state["data_bersih"] = clean_df
            st.session_state["info_bersih"] = info_dict
            st.success("Pembersihan data berhasil dilakukan.")
            
        st.markdown("</div>", unsafe_allow_html=True)
        
        if st.session_state["data_bersih"] is not None:
            info_tampil = st.session_state["info_bersih"]
            st.markdown("<br>", unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)
            col1.metric("Data Duplikat", f"{info_tampil['duplikat_dihapus']} baris")
            col2.metric("Data Kosong", f"{info_tampil['missing_value']} anomali")
            col3.metric("Data Bersih", f"{info_tampil['total_bersih']} baris")
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            st.info(f"Sistem mendeteksi adanya {info_tampil['duplikat_dihapus']} baris data ganda (duplikat) dan {info_tampil['missing_value']} data kosong/anomali yang kemudian otomatis dibersihkan oleh sistem, sehingga menghasilkan total {info_tampil['total_bersih']} baris data bersih yang siap digunakan ke tahap ekstraksi fitur.")
            
            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander("Tampilkan Sampel Data yang Sudah Bersih"):
                st.dataframe(st.session_state["data_bersih"].head(10), use_container_width=True, hide_index=True)
    else:
        st.markdown("</div>", unsafe_allow_html=True)
        st.warning("Belum ada data. Selesaikan Menu 1 terlebih dahulu.")


# MENU 3: PEMBENTUKAN FITUR 
elif nav == "3. Pembentukan Fitur":
    st.markdown("""
        <div class="kartu-utama">
            <h3 style="color: #f8fafc; margin-top: 0; font-size: 1.2rem;">⚙️ Pembentukan Fitur (Feature Engineering)</h3>
            <p style="color: #94a3b8; margin-bottom: 1rem; font-size: 0.92rem;"><b>Fungsi Menu Ini:</b> Merangkum semua laporan data penjualan yang tadinya terpisah menjadi satu baris ringkasan khusus untuk setiap nama barang. Berikut penjelasan kolom pada tabel hasil:</p>
            <ul style="color: #94a3b8; font-size: 0.9rem; margin-bottom: 0; line-height: 1.6;">
                <li><b>Nama Barang:</b> Nama produk yang dijual.</li>
                <li><b>Total Qty:</b> Total keseluruhan jumlah barang yang berhasil terjual.</li>
                <li><b>Rata-rata Qty:</b> Jumlah rata-rata barang yang biasanya keluar dalam sekali transaksi oleh pembeli.</li>
                <li><b>Frekuensi Transaksi:</b> Seberapa sering suatu produk dibeli oleh pelanggan.</li>
                <li><b>Rata-rata Harga:</b> Kisaran harga jual rata-rata dari produk tersebut.</li>
            </ul>
    """, unsafe_allow_html=True)
    
    if st.session_state["data_bersih"] is not None:
        if st.button("⚡ Mulai Ekstraksi Fitur Produk"):
            feat_df = extract_product_features(st.session_state["data_bersih"])
            st.session_state["data_fitur"] = feat_df
            st.success("Ekstraksi fitur per produk selesai.")
            
        st.markdown("</div>", unsafe_allow_html=True)
        
        if st.session_state["data_fitur"] is not None:
            df_tampil_fitur = st.session_state["data_fitur"]
            st.markdown("<br>", unsafe_allow_html=True)
            st.metric("TOTAL JENIS PRODUK", f"{len(df_tampil_fitur)} Produk")
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.write("Tabel Ringkasan Fitur Produk:")
            st.dataframe(df_tampil_fitur, use_container_width=True, hide_index=True)
    else:
        st.markdown("</div>", unsafe_allow_html=True)
        st.warning("Selesaikan Menu 2 (Pembersihan Data) terlebih dahulu.")


# HALAMAN 4: PREDIKSI & KLASIFIKASI 
elif nav == "4. Prediksi & Klasifikasi":
    st.markdown("""
        <div class="kartu-utama">
            <h3 style="color: #f8fafc; margin-top: 0; font-size: 1.2rem;">🔍 Prediksi & Klasifikasi Potensi Produk</h3>
            <p style="color: #94a3b8; margin-bottom: 0; font-size: 0.92rem;"><b>Fungsi Menu Ini:</b> Menerapkan model kecerdasan buatan untuk mengklasifikasikan status produk ke dalam kategori <b>Potensi Tinggi, Sedang, atau Rendah</b> serta memberikan arahan rekomendasi stok secara objektif.</p>
        </div>
    """, unsafe_allow_html=True)
    
    if st.session_state["data_fitur"] is not None:
        try:
            model_obj, encoder_obj = load_trained_model()
            df_temp = st.session_state["data_fitur"].copy()
            
            df_temp["Kategori Potensi"] = run_model_prediction(model_obj, encoder_obj, df_temp)
            df_temp["Rekomendasi Stok"] = df_temp["Kategori Potensi"].map(RECOMMENDATION_MAPPING)
            st.session_state["data_hasil"] = df_temp
            
            st.success("Proses prediksi klasifikasi berhasil dijalankan.")
            st.markdown("<br>", unsafe_allow_html=True)
            
            c1, c2 = st.columns(2)
            with c1:
                keyword = st.text_input("Cari Berdasarkan Nama Produk", placeholder="Ketik nama produk...")
            with c2:
                selected_cat = st.multiselect("Filter Kategori Potensi", options=["Potensi Tinggi", "Potensi Sedang", "Potensi Rendah"])
            
            filtered_result = df_temp.copy()
            if keyword:
                filtered_result = filtered_result[filtered_result["Nama Barang"].str.contains(keyword, case=False, na=False)]
            if selected_cat:
                filtered_result = filtered_result[filtered_result["Kategori Potensi"].isin(selected_cat)]
                
            st.markdown("<br>", unsafe_allow_html=True)
            st.dataframe(filtered_result, use_container_width=True, hide_index=True)
        except Exception as err_msg:
            st.error(f"Terjadi kendala saat eksekusi model ML: {err_msg}")
    else:
        st.warning("Selesaikan Menu 3 (Pembentukan Fitur) terlebih dahulu.")


# HALAMAN 5: DASHBOARD & VISUALISASI 
elif nav == "5. Dashboard & Visualisasi":
    st.markdown("""
        <div class="kartu-utama">
            <h3 style="color: #f8fafc; margin-top: 0; font-size: 1.2rem;">📊 Dashboard & Visualisasi Grafik</h3>
            <p style="color: #94a3b8; margin-bottom: 0; font-size: 0.92rem;"><b>Fungsi Menu Ini:</b> Menyajikan infografis analitik interaktif beserta narasi kesimpulan otomatis guna mempermudah manajemen dalam mengambil keputusan bisnis.</p>
        </div>
    """, unsafe_allow_html=True)
    
    if st.session_state["data_hasil"] is not None:
        df_dash = st.session_state["data_hasil"]
        
        count_high = int((df_dash["Kategori Potensi"] == "Potensi Tinggi").sum())
        count_med = int((df_dash["Kategori Potensi"] == "Potensi Sedang").sum())
        count_low = int((df_dash["Kategori Potensi"] == "Potensi Rendah").sum())
        total_items_all = len(df_dash)
        
        col1, col2, col3 = st.columns(3)
        col1.metric("POTENSI TINGGI", f"{count_high} Produk")
        col2.metric("POTENSI SEDANG", f"{count_med} Produk")
        col3.metric("POTENSI RENDAH", f"{count_low} Produk")
        
        st.markdown("---")
        
        pie_df = df_dash["Kategori Potensi"].value_counts().reset_index()
        pie_df.columns = ["Kategori", "Jumlah Produk"]
        
        tab_v1, tab_v2, tab_v3 = st.tabs(["Proporsi Kategori", "10 Produk Terbesar", "Tren Historis Bulanan"])
        
        with tab_v1:
            fig_pie = px.pie(pie_df, names="Kategori", values="Jumlah Produk", 
                             title="Distribusi Klasifikasi Potensi Penjualan Produk",
                             color="Kategori", color_discrete_map=CHART_COLOR_MAP, hole=0.45)
            st.plotly_chart(fig_pie, use_container_width=True)
            
            p_high = (count_high / total_items_all) * 100 if total_items_all > 0 else 0
            p_med = (count_med / total_items_all) * 100 if total_items_all > 0 else 0
            p_low = (count_low / total_items_all) * 100 if total_items_all > 0 else 0
            
            st.info(f"""
                Ringkasan Analisis Kategori Potensi Penjualan:
                * Dari total {total_items_all} jenis produk yang dianalisis oleh sistem, berikut rincian pembagian status dan langkah yang sebaiknya diambil:
                  - Potensi Tinggi ({count_high} produk / {p_high:.1f}%): Produk paling laris dan menjadi andalan utama. Tindakan: Tingkatkan Stok.
                  - Potensi Sedang ({count_med} produk / {p_med:.1f}%): Produk dengan tingkat penjualan stabil. Tindakan: Pertahankan Stok.
                  - Potensi Rendah ({count_low} produk / {p_low:.1f}%): Produk kurang diminati. Tindakan: Kurangi Stok.
            """)
            
        with tab_v2:
            top_10 = df_dash.sort_values(by="Total_Qty", ascending=False).head(10)
            fig_bar = px.bar(top_10, x="Total_Qty", y="Nama Barang", orientation='h', 
                             color="Kategori Potensi", color_discrete_map=CHART_COLOR_MAP,
                             title="10 Produk dengan Akumulasi Penjualan (Qty) Tertinggi")
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_bar, use_container_width=True)
            
            top_product_name = top_10.iloc[0]["Nama Barang"] if not top_10.empty else "-"
            top_product_qty = top_10.iloc[0]["Total_Qty"] if not top_10.empty else 0
            
            st.info(f"""
                Analisis & Kesimpulan Produk Terbesar:
                * Volume penjualan fisik terbanyak dipegang oleh "{top_product_name}" dengan total pencapaian {top_product_qty:,.0f} unit.
                * Sebagian besar produk di jajaran teratas masuk dalam kategori Potensi Tinggi.
            """)
            
        with tab_v3:
            df_trend = st.session_state["data_bersih"].copy()
            df_trend["Bulan"] = df_trend["Tanggal Faktur"].dt.to_period("M").dt.to_timestamp()
            monthly_rekap = df_trend.groupby("Bulan").agg(Total_Qty=("Qty", "sum")).reset_index().sort_values("Bulan")
            
            fig_line = px.line(monthly_rekap, x="Bulan", y="Total_Qty", markers=True,
                               title="Grafik Tren Historis Kuantitas Penjualan Bulanan")
            st.plotly_chart(fig_line, use_container_width=True)
            
            total_months = len(monthly_rekap)
            total_qty_sum = df_trend["Qty"].sum()
            
            st.info(f"""
                Analisis & Kesimpulan Tren Bulanan:
                * Grafik memperlihatkan fluktuasi pergerakan volume penjualan selama rentang {total_months} bulan historis, dengan akumulasi total barang keluar mencapai {total_qty_sum:,.0f} unit.
            """)
    else:
        st.warning("Jalankan Menu 4 (Prediksi & Klasifikasi) terlebih dahulu.")


# HALAMAN 6: UNDUH LAPORAN 
elif nav == "6. Unduh Laporan":
    st.markdown("""
        <div class="kartu-utama">
            <h3 style="color: #f8fafc; margin-top: 0; font-size: 1.2rem;">📥 Unduh Laporan Hasil Analisis</h3>
            <p style="color: #94a3b8; margin-bottom: 1.5rem; font-size: 0.92rem;"><b>Fungsi Menu Ini:</b> Mengekspor seluruh rekapitulasi klasifikasi potensi produk beserta rekomendasi stok ke dalam dokumen format Excel (.xlsx) yang telah diformat secara profesional agar siap dilampirkan dalam laporan resmi.</p>
    """, unsafe_allow_html=True)
    
    if st.session_state["data_hasil"] is not None:
        tabel_export = st.session_state["data_hasil"]
        
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            tabel_export.to_excel(writer, index=False, sheet_name="Hasil Klasifikasi")
            
            wb_file = writer.book
            ws_file = writer.sheets["Hasil Klasifikasi"]
            
            header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            box_border = Border(
                left=Side(style='thin', color='CBD5E1'),
                right=Side(style='thin', color='CBD5E1'),
                top=Side(style='thin', color='CBD5E1'),
                bottom=Side(style='thin', color='CBD5E1')
            )
            
            for c_idx in range(1, len(tabel_export.columns) + 1):
                cell_h = ws_file.cell(row=1, column=c_idx)
                cell_h.fill = header_fill
                cell_h.font = header_font
                cell_h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell_h.border = box_border
            
            for r_idx in range(2, len(tabel_export) + 2):
                for c_idx in range(1, len(tabel_export.columns) + 1):
                    cell_d = ws_file.cell(row=r_idx, column=c_idx)
                    cell_d.font = Font(name="Calibri", size=11)
                    cell_d.border = box_border
                    
                    col_key = tabel_export.columns[c_idx - 1]
                    
                    if col_key == "Rata_Harga":
                        cell_d.number_format = '#,##0.00'
                        cell_d.alignment = Alignment(horizontal="right", vertical="center")
                    elif col_key == "Rata_Qty":
                        # Dibatasi 2 angka di belakang koma agar rapi dan profesional
                        cell_d.number_format = '0.00'
                        cell_d.alignment = Alignment(horizontal="right", vertical="center")
                    elif c_idx == 1:
                        cell_d.alignment = Alignment(horizontal="left", vertical="center")
                    elif c_idx in [6, 7]:
                        cell_d.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell_d.alignment = Alignment(horizontal="right", vertical="center")
            
            for column_cells in ws_file.columns:
                max_width_val = 0
                col_letter_str = get_column_letter(column_cells[0].column)
                for cell_item in column_cells:
                    try:
                        if cell_item.value:
                            max_width_val = max(max_width_val, len(str(cell_item.value)))
                    except:
                        pass
                ws_file.column_dimensions[col_letter_str].width = max(max_width_val + 5, 16)
                
        excel_buffer.seek(0)
        
        st.success("File laporan Excel Anda sudah siap untuk diunduh.")
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            label="Unduh Dokumen Laporan Rekomendasi (.xlsx)",
            data=excel_buffer,
            file_name="Laporan_Analisis_Potensi_Penjualan_UMKM.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.warning("Belum terdapat data hasil analisis. Harap selesaikan proses hingga tahap 4 terlebih dahulu.")
    
    st.markdown("</div>", unsafe_allow_html=True)