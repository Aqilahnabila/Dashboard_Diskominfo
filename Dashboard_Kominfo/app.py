# -*- coding: utf-8 -*-
"""
Dashboard Analisis Data CCTV Lalu Lintas
=========================================
"""

import os
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
import tempfile
import subprocess
import shutil
import base64
from pathlib import Path

APP_DIR = os.path.dirname(os.path.abspath(__file__))

import parser as dp  # data parsing & cleaning
import report as rpt

st.set_page_config(
    page_title="Dashboard Analisis CCTV",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------------------------------------------------------
# Styling dashboard — palet biru Kominfo
# ----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    :root {
        --kominfo-navy: #063B68;
        --kominfo-blue: #0874B9;
        --kominfo-mid: #258FCB;
        --kominfo-pale: #EAF4FB;
        --surface: #FFFFFF;
        --page-bg: #F4F7FA;
        --text-main: #263746;
        --text-muted: #687B8D;
        --line: #D9E3EB;
    }

    .stApp { background: var(--page-bg); color: var(--text-main); }
    .main > div { padding-top: 1.45rem; padding-bottom: 2.5rem; }
    .block-container { max-width: 100%; padding-left: 2.3rem; padding-right: 2.3rem; }

    /* Header utama: tegas, tetapi tidak berlebihan */
    h1 {
        color: var(--kominfo-navy) !important;
        font-size: 2.15rem !important;
        font-weight: 750 !important;
        letter-spacing: -.035em;
        line-height: 1.2 !important;
        padding-bottom: .65rem;
        border-bottom: 3px solid #D7EAF6;
        margin-bottom: .45rem !important;
    }
    h2, h3 { color: var(--kominfo-navy) !important; font-weight: 700 !important; letter-spacing: -.018em; }
    h2 { font-size: 1.55rem !important; }
    h3 { font-size: 1.25rem !important; }
    [data-testid="stCaptionContainer"] p { color: var(--text-muted); line-height: 1.55; font-size: .9rem; }

    /* Sidebar dengan bidang biru muda yang lebih berkarakter */
    section[data-testid="stSidebar"] {
        background: #EAF1F7;
        border-right: 1px solid #D2DFE9;
    }
    section[data-testid="stSidebar"] > div { padding-top: 1.35rem; }
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        color: var(--kominfo-navy) !important;
        font-weight: 750 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    section[data-testid="stSidebar"] label { color: #40586D; }

    /* KPI: lebih menonjol dan proporsional */
    div[data-testid="stMetric"] {
        background: var(--surface);
        border: 1px solid #D7E2EB;
        border-top: 4px solid var(--kominfo-blue);
        border-radius: 11px;
        padding: 1.15rem 1rem 1rem;
        box-shadow: 0 2px 7px rgba(16, 54, 83, .06);
        min-height: 118px;
        transition: box-shadow .18s ease, transform .18s ease;
    }
    div[data-testid="stMetric"]:hover {
        box-shadow: 0 4px 12px rgba(16, 54, 83, .10);
        transform: translateY(-1px);
    }
    [data-testid="stMetricLabel"] { display:flex !important; justify-content:center !important; width:100% !important; }
    [data-testid="stMetricLabel"] p {
        color: #5C7183 !important; font-size: .96rem !important;
        font-weight: 700 !important; text-align:center !important; width:100% !important;
    }
    div[data-testid="stMetricValue"] {
        color: var(--kominfo-navy) !important; width:100%;
        justify-content:center; text-align:center;
        font-size: 1.85rem !important; font-weight: 750 !important;
    }
    div[data-testid="stMetricDelta"] {
        color: #25805B !important; width:100%; justify-content:center;
        text-align:center; font-size:.86rem !important;
    }
    div[data-testid="stHorizontalBlock"] > div:nth-child(5) div[data-testid="stMetricValue"] {
        white-space:normal !important; overflow-wrap:anywhere !important;
        line-height:1.2 !important; font-size:clamp(1.05rem, 1.6vw, 1.5rem) !important;
    }
    div[data-testid="stHorizontalBlock"] > div:nth-child(5) div[data-testid="stMetricDelta"] svg { display:none !important; }

    /* Navigasi tab: area aktif jelas, tetap sederhana */
    div[data-baseweb="tab-list"] {
        gap: .35rem; border-bottom: 1px solid #D5E0E9;
        padding: .35rem .25rem 0; background: transparent;
    }
    button[data-baseweb="tab"] {
        color: #536A7D !important; font-weight: 650 !important;
        padding: .8rem 1.05rem !important; border-radius: 7px 7px 0 0;
    }
    button[data-baseweb="tab"]:hover { color: var(--kominfo-blue) !important; background: #E8F2F9 !important; }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: var(--kominfo-navy) !important; background: #E5F1F9 !important;
        font-weight: 750 !important;
    }
    div[data-baseweb="tab-highlight"] { background-color: var(--kominfo-blue) !important; height: 3px !important; }

    /* Input dan filter */
    div[data-baseweb="select"] > div,
    div[data-testid="stTextInput"] input,
    div[data-testid="stDateInput"] input {
        border-color: #C9D8E4 !important; border-radius: 8px !important;
        background: #FFFFFF !important;
    }
    div[data-baseweb="select"] > div:focus-within,
    div[data-testid="stTextInput"] input:focus {
        border-color: var(--kominfo-blue) !important;
        box-shadow: 0 0 0 1px var(--kominfo-blue) !important;
    }
    input[type="checkbox"] { accent-color: var(--kominfo-blue); }
    div[data-testid="stSlider"] [data-baseweb="slider"] div[role="slider"] {
        background-color: var(--kominfo-blue) !important; border-color: var(--kominfo-blue) !important;
    }

    /* Tombol */
    .stButton > button, .stDownloadButton > button {
        border-radius: 8px; font-weight: 650;
        border: 1px solid var(--kominfo-blue);
        padding: .48rem 1rem;
        transition: background-color .15s ease, border-color .15s ease, box-shadow .15s ease;
    }
    .stButton > button[kind="primary"] {
        background: var(--kominfo-blue) !important; color: #FFFFFF !important;
    }
    .stButton > button[kind="primary"]:hover {
        background: var(--kominfo-navy) !important; border-color: var(--kominfo-navy) !important;
        box-shadow: 0 2px 7px rgba(6, 59, 104, .18);
    }
    .stDownloadButton > button { color: var(--kominfo-blue) !important; background:#FFFFFF !important; }
    .stDownloadButton > button:hover { background:var(--kominfo-pale) !important; }

    /* Grafik dan tabel: bidang putih terpisah dari latar halaman */
    div[data-testid="stVegaLiteChart"], div[data-testid="stArrowVegaLiteChart"] {
        background: #FFFFFF; border: 1px solid #E0E8EF;
        border-radius: 10px; padding: .45rem .55rem .2rem;
        box-shadow: 0 1px 4px rgba(16, 54, 83, .035);
    }
    div[data-testid="stDataFrame"], div[data-testid="stTable"] {
        border: 1px solid var(--line); border-radius: 9px; overflow:hidden;
        background:#FFFFFF;
    }

    /* Pesan, uploader, dan pemisah */
    div[data-testid="stAlert"] { border-radius:9px; border:1px solid var(--line); }
    div[data-testid="stFileUploader"] section {
        background:#FFFFFF; border:1px dashed #9DBBD0; border-radius:9px;
    }
    hr { border-color:#D8E2EA !important; }

    .badge-aktif {background:#E7F3EC;color:#267247;padding:2px 10px;border-radius:5px;font-size:.8rem;}
    .badge-nonaktif {background:#F9ECEB;color:#A23C35;padding:2px 10px;border-radius:5px;font-size:.8rem;}
    .badge-parsial {background:#F8F1E2;color:#87651E;padding:2px 10px;border-radius:5px;font-size:.8rem;}

    @media (max-width: 900px) {
        .block-container { padding-left:1rem; padding-right:1rem; }
        h1 { font-size:1.65rem !important; }
        div[data-testid="stMetric"] { padding:.8rem .55rem; min-height:100px; }
        div[data-testid="stMetricValue"] { font-size:1.35rem !important; }
    }

    /* Penyempurnaan visual resmi Diskominfo — CSS saja */
    .block-container {
        max-width: 1500px;
        padding-top: 1.8rem;
        padding-bottom: 3rem;
    }
    h1 {
        font-size: 2.25rem !important;
        letter-spacing: -0.025em !important;
        border-bottom: 3px solid #D6E5F0;
        padding-bottom: .8rem;
        margin-bottom: .65rem !important;
    }
    h2 { font-size: 1.45rem !important; margin-top: .45rem !important; }
    h3 { font-size: 1.12rem !important; margin-top: .45rem !important; }
    [data-testid="stCaptionContainer"] p {
        color: #627487 !important; font-size: .9rem !important;
    }
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #EAF2F8 0%, #F3F7FA 100%);
        border-right: 1px solid #D5E1EA;
    }
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {
        padding-bottom: .35rem; border-bottom: 1px solid #D4E1EA;
    }
    div[data-testid="stMetric"] {
        border: 1px solid #D8E3EC !important;
        border-top: 4px solid #0874B9 !important;
        border-radius: 12px !important;
        background: #FFFFFF !important;
        box-shadow: 0 3px 12px rgba(18,55,82,.055) !important;
        padding: 1.05rem .9rem .95rem !important;
        min-height: 126px;
    }
    div[data-testid="stMetric"]:hover {
        border-color: #B8CDDC !important;
        box-shadow: 0 6px 16px rgba(18,55,82,.09) !important;
        transform: translateY(-2px);
    }
    [data-testid="stMetricLabel"] p {
        color: #607487 !important; font-size: .84rem !important;
    }
    div[data-testid="stMetricValue"] {
        color: #063B68 !important; font-size: 1.8rem !important;
        font-weight: 750 !important;
    }
    div[data-baseweb="tab-list"] {
        gap: .45rem; padding: .45rem .35rem 0;
        border-bottom: 1px solid #D7E2EB;
    }
    button[data-baseweb="tab"] {
        border-radius: 8px 8px 0 0 !important;
        padding: .78rem 1.05rem !important;
        color: #526A7D !important; font-size: .91rem !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #063B68 !important; background: #E5F1F9 !important;
        box-shadow: inset 0 -3px 0 #0874B9;
    }
    div[data-testid="stVegaLiteChart"],
    div[data-testid="stArrowVegaLiteChart"] {
        border: 1px solid #DFE7EE !important;
        border-radius: 12px !important; background: #FFFFFF !important;
        padding: .65rem .7rem .35rem !important;
        box-shadow: 0 2px 8px rgba(18,55,82,.04) !important;
    }
    div[data-testid="stDataFrame"], div[data-testid="stTable"] {
        border: 1px solid #DCE5EC !important;
        border-radius: 10px !important; overflow: hidden;
        background: #FFFFFF !important;
    }
    .stButton > button, .stDownloadButton > button {
        min-height: 2.45rem; border-radius: 8px !important;
        font-weight: 650 !important;
    }
    .stButton > button[kind="primary"] {
        background: #0874B9 !important; border-color: #0874B9 !important;
    }
    .stButton > button[kind="primary"]:hover {
        background: #063B68 !important; border-color: #063B68 !important;
    }
    div[data-testid="stAlert"] { border-radius: 10px !important; padding: .85rem 1rem !important; }
    div[data-testid="stFileUploader"] section {
        border: 1.5px dashed #9BB8CC !important;
        border-radius: 10px !important; background: rgba(255,255,255,.85) !important;
    }
    hr { margin: 1.35rem 0 !important; border-color: #DCE5EC !important; }
    @media (max-width: 900px) {
        .block-container { padding-left: 1rem; padding-right: 1rem; }
        h1 { font-size: 1.7rem !important; }
        div[data-testid="stMetric"] { min-height: 105px; }
    }

    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Dashboard Analisis Data Lalu Lintas")
st.caption(
    "Unggah data rekap CCTV (Excel/CSV) — data akan otomatis dibersihkan, dianalisis, "
    "dan ditampilkan sebagai dashboard lengkap serta siap diunduh sebagai laporan."
)

# Sidebar
with st.sidebar:

    # Logo Kota Madiun dan Diskominfo
    logo_kota_path = Path(__file__).parent / "assets" / "madiun.png"
    logo_kominfo_path = Path(__file__).parent / "assets" / "kominfo.png"
    
    logo_col1, logo_col2 = st.columns([1, 1], gap="small", vertical_alignment="center")

    with logo_col1:
        st.image(str(logo_kota_path), width=80)

    with logo_col2:
        st.image(str(logo_kominfo_path), width=165)

    st.markdown("<br>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# Sidebar: upload data
# ----------------------------------------------------------------------------
with st.sidebar:
    st.header("Input Data CCTV")

    uploaded_files = st.file_uploader(
        "Unggah file data CCTV",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        help=(
            "Anda dapat mengunggah satu atau beberapa file sekaligus. "
            "Format yang didukung: Excel (.xlsx/.xls) dan CSV."
        ),
    )

    st.markdown("---")

    use_sample = st.checkbox(
        "Gunakan data contoh",
        value=False,
        help="Gunakan data bawaan aplikasi untuk mencoba dashboard.",
    )

    # ------------------------------------------------------------------------
    # Daftar file yang diinput
    # ------------------------------------------------------------------------
    if uploaded_files:
        st.markdown("#### File Terinput")

        for file in uploaded_files:
            st.markdown(f"- `{file.name}`")

    elif use_sample:
        st.markdown("#### File Terinput")
        st.markdown("- `sample_data.xlsx`")


# ----------------------------------------------------------------------------
# Validasi input
# ----------------------------------------------------------------------------
if not uploaded_files and not use_sample:
    st.info(
        "**Silakan unggah data CCTV terlebih dahulu.**\n\n"
        "Anda dapat mengunggah satu atau beberapa file Excel/CSV "
        "melalui panel di sebelah kiri."
    )
    st.stop()


# ----------------------------------------------------------------------------
# Load & parse data
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner="Memproses data...")
def _load_multiple(files_data):
    """
    Membaca satu atau beberapa file Excel/CSV menggunakan parser yang sama,
    kemudian menggabungkan seluruh hasil cleaning menjadi satu DataFrame.
    """

    import io

    hasil = []

    for file_name, file_bytes in files_data:

        fobj = io.BytesIO(file_bytes)
        fobj.name = file_name

        # Setiap file diproses menggunakan parser yang sama
        df_file = dp.load_and_clean(fobj)

        if df_file is not None and not df_file.empty:
            hasil.append(df_file)

    # Tidak ada data valid
    if not hasil:
        return pd.DataFrame()

    # Gabungkan seluruh file
    df_gabungan = pd.concat(hasil, ignore_index=True)

    # Hilangkan baris duplikat
    df_gabungan = df_gabungan.drop_duplicates().reset_index(drop=True)

    return df_gabungan


# ----------------------------------------------------------------------------
# Proses data
# ----------------------------------------------------------------------------
try:

    if uploaded_files:

        # Simpan isi seluruh file dalam bentuk bytes
        files_data = tuple((file.name, file.getvalue()) for file in uploaded_files)

        # Proses seluruh file
        df_raw = _load_multiple(files_data)

        # Nama sumber data
        if len(uploaded_files) == 1:
            sumber_data = uploaded_files[0].name
        else:
            sumber_data = f"{len(uploaded_files)} file data CCTV"

    else:

        # --------------------------------------------------------------------
        # Gunakan data contoh
        # --------------------------------------------------------------------
        sample_path = os.path.join(APP_DIR, "sample_data.xlsx")

        with open(sample_path, "rb") as f:
            sample_bytes = f.read()

        df_raw = _load_multiple((("sample_data.xlsx", sample_bytes),))

        sumber_data = "data contoh (bawaan aplikasi)"


# ----------------------------------------------------------------------------
# Error handling
# ----------------------------------------------------------------------------
except dp.DataFormatError as e:

    st.error(f"❌ **Format data tidak sesuai:** {e}")
    st.stop()

except FileNotFoundError:

    st.error(
        "❌ **File data contoh tidak ditemukan.** "
        "Silakan unggah file data Anda sendiri."
    )
    st.stop()

except Exception as e:

    st.error(f"❌ **Terjadi kesalahan saat membaca data:** {e}")
    st.stop()


# ----------------------------------------------------------------------------
# Validasi hasil akhir
# ----------------------------------------------------------------------------
if df_raw.empty:

    st.error(
        "❌ File berhasil dibaca, tetapi tidak ada baris data " "yang valid ditemukan."
    )
    st.stop()


# ----------------------------------------------------------------------------
# Ringkasan data berhasil diproses
# ----------------------------------------------------------------------------
jumlah_lokasi = df_raw["Lokasi"].nunique()
jumlah_hari = df_raw["Tanggal"].nunique()
jumlah_baris = len(df_raw)

# Format angka Indonesia
jumlah_hari_fmt = f"{jumlah_hari:,}".replace(",", ".")
jumlah_baris_fmt = f"{jumlah_baris:,}".replace(",", ".")


if uploaded_files and len(uploaded_files) > 1:

    st.success(
        f"**Data berhasil diproses dan digabungkan** — "
        f"**{jumlah_lokasi} lokasi**, "
        f"**{jumlah_hari_fmt} hari**, "
        f"**{jumlah_baris_fmt} baris data**."
    )

else:

    st.success(
        f"**Data berhasil diproses** — "
        f"**{jumlah_lokasi} lokasi**, "
        f"**{jumlah_hari_fmt} hari**, "
        f"**{jumlah_baris_fmt} baris data**."
    )


# ----------------------------------------------------------------------------
# Sidebar: filter lanjutan (setelah data ada)
# ----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("---")
    st.header("Filter")

    all_lokasi = sorted(df_raw["Lokasi"].unique().tolist())
    lokasi_terpilih = st.multiselect(
        "Pilih lokasi (kosongkan = semua lokasi)", all_lokasi, default=[]
    )
    if not lokasi_terpilih:
        lokasi_terpilih = all_lokasi

    tgl_min, tgl_max = df_raw["Tanggal"].min().date(), df_raw["Tanggal"].max().date()
    if tgl_min == tgl_max:
        rentang_tanggal = (tgl_min, tgl_max)
        st.caption(f"Tanggal data: {tgl_min.strftime('%d-%m-%Y')} (hanya 1 hari)")
    else:
        rentang_tanggal = st.slider(
            "Rentang tanggal",
            min_value=tgl_min,
            max_value=tgl_max,
            value=(tgl_min, tgl_max),
            format="DD-MM-YYYY",
        )

    periode = st.selectbox(
        "Periode analisis",
        ["Harian", "Mingguan", "Bulanan"],
        index=0,
        help="Menentukan bagaimana data digabungkan pada grafik tren & tabel: "
        "per hari, per minggu, atau per bulan.",
    )

# ----------------------------------------------------------------------------
# Terapkan filter
# ----------------------------------------------------------------------------
mask = (
    df_raw["Lokasi"].isin(lokasi_terpilih)
    & (df_raw["Tanggal"].dt.date >= rentang_tanggal[0])
    & (df_raw["Tanggal"].dt.date <= rentang_tanggal[1])
)
df = df_raw.loc[mask].copy()

if df.empty:
    st.warning(
        "Tidak ada data pada kombinasi filter yang dipilih. Silakan ubah filter di sidebar."
    )
    st.stop()

dq = dp.data_quality_report(df)
agg = dp.aggregate_by_period(df, periode)

# ----------------------------------------------------------------------------
# Banner kualitas data
# ----------------------------------------------------------------------------
n_inactive = (dq["status"] == "Tidak ada data / kamera tidak aktif").sum()
n_partial = (dq["status"] == "Sebagian data kosong").sum()

# ----------------------------------------------------------------------------
# KPI utama
# ----------------------------------------------------------------------------
total_all = df["Total"].sum()
n_hari = df["Tanggal"].nunique()
avg_daily = df.groupby("Tanggal")["Total"].sum().mean()
by_lokasi_sum = df.groupby("Lokasi")["Total"].sum().sort_values(ascending=False)
lokasi_teratas = by_lokasi_sum.index[0] if len(by_lokasi_sum) else "-"
lokasi_teratas_val = by_lokasi_sum.iloc[0] if len(by_lokasi_sum) else 0
n_lokasi_aktif = (dq["status"] != "Tidak ada data / kamera tidak aktif").sum()

k1, k2, k3, k4, k5 = st.columns([1.05, 0.9, 0.75, 0.7, 1.9])
k1.metric("Total Keseluruhan", f"{total_all:,.0f}".replace(",", "."))
k2.metric("Rata-rata / Hari", f"{avg_daily:,.0f}".replace(",", "."))
k3.metric("Lokasi Aktif", f"{n_lokasi_aktif} / {jumlah_lokasi}")
k4.metric("Total Hari", f"{n_hari}")
k5.metric(
    "Lokasi Tersibuk", lokasi_teratas, f"{lokasi_teratas_val:,.0f}".replace(",", ".")
)

st.markdown("---")

# ----------------------------------------------------------------------------
# Tabs dashboard
# ----------------------------------------------------------------------------

tab_tren, tab_lokasi, tab_komposisi, tab_kualitas, tab_data, tab_unduh = st.tabs(
    [
        "Tren",
        "Perbandingan Lokasi",
        "Komposisi",
        "Kualitas Data",
        "Data Detail",
        "Unduh Hasil",
    ]
)

# --- Tab Tren ---

with tab_tren:
    st.subheader(f"Tren Volume Lalu Lintas ({periode})")
    tren = agg.groupby("Periode", as_index=False)["Total"].sum()
    st.line_chart(tren.set_index("Periode")["Total"])

    st.markdown(
        """
    <h3 style="margin-bottom: -20px; margin-top: 0;">
        Tren per Lokasi
    </h3>
    """,
        unsafe_allow_html=True,
    )
    pilihan_tren_lokasi = st.multiselect(
        "",
        sorted(df["Lokasi"].unique().tolist()),
        default=list(by_lokasi_sum.head(5).index),
        key="tren_lokasi",
    )
    if pilihan_tren_lokasi:
        pivot = (
            agg[agg["Lokasi"].isin(pilihan_tren_lokasi)]
            .pivot_table(
                index="Periode", columns="Lokasi", values="Total", aggfunc="sum"
            )
            .fillna(0)
        )
        st.line_chart(pivot)
    else:
        st.caption(
            "Pilih minimal satu lokasi untuk menampilkan grafik perbandingan tren."
        )


import altair as alt

# --- Tab Perbandingan Lokasi ---
with tab_lokasi:
    st.subheader("Perbandingan Volume per Lokasi")

    top_n = st.selectbox(
        "Tampilkan berapa lokasi?",
        options=list(range(1, len(all_lokasi) + 1)),
        index=min(9, len(all_lokasi) - 1),
        key="top_n_lokasi",
    )

    st.subheader(f"Lokasi Teratas berdasarkan Total Volume (Top {top_n})")

    # Terbesar ke terkecil
    rank = by_lokasi_sum.nlargest(top_n).sort_values(ascending=False)
    df_rank = rank.rename("Total").rename_axis("Lokasi").reset_index()

    chart_rank = (
        alt.Chart(df_rank)
        .mark_bar()
        .encode(
            x=alt.X(
                "Lokasi:N",
                sort=df_rank["Lokasi"].tolist(),
                axis=alt.Axis(labelAngle=-90),
                title="Lokasi",
            ),
            y=alt.Y("Total:Q", title="Total Volume"),
            tooltip=["Lokasi", "Total"],
        )
        .properties(width="container", height=450)
    )
    st.altair_chart(chart_rank, use_container_width=True)

    st.subheader("Lokasi dengan Volume Terendah (di antara lokasi aktif)")

    aktif_locs = dq.loc[
        dq["status"] != "Tidak ada data / kamera tidak aktif", "Lokasi"
    ].unique()

    # Terkecil ke terbesar
    rendah = (
        by_lokasi_sum[by_lokasi_sum.index.isin(aktif_locs)]
        .nsmallest(top_n)
        .sort_values(ascending=True)
    )
    df_rendah = rendah.rename("Total").rename_axis("Lokasi").reset_index()

    chart_rendah = (
        alt.Chart(df_rendah)
        .mark_bar()
        .encode(
            x=alt.X(
                "Lokasi:N",
                sort=df_rendah["Lokasi"].tolist(),
                axis=alt.Axis(labelAngle=-90),
                title="Lokasi",
            ),
            y=alt.Y("Total:Q", title="Total Volume"),
            tooltip=["Lokasi", "Total"],
        )
        .properties(width="container", height=450)
    )
    st.altair_chart(chart_rendah, use_container_width=True)


# --- Tab Komposisi ---
with tab_komposisi:
    col1, col2 = st.columns([1, 1.3])

    with col1:
        st.subheader("Komposisi Keseluruhan")

        # Hitung total setiap kategori
        comp = df[dp.CATEGORIES].sum()
        comp = comp[comp > 0]

        # Urutkan kategori berdasarkan total terbanyak ke tersedikit
        comp = comp.sort_values(ascending=False)

        # Siapkan tabel komposisi
        tabel_comp = comp.rename("Total").to_frame()

        tabel_comp["Persentase"] = (
            tabel_comp["Total"] / tabel_comp["Total"].sum() * 100
        ).round(1)

        tabel_comp.index.name = "Jenis"
        tabel_comp = tabel_comp.reset_index()

        # Pastikan total berupa bilangan bulat
        tabel_comp["Total"] = tabel_comp["Total"].round(0).astype("int64")

        # Format persentase
        tabel_comp["Persentase"] = tabel_comp["Persentase"].map(lambda x: f"{x:.1f}%")

        # Styling tabel: header biru dan semua teks rata tengah
        styled_comp = (
            tabel_comp.style.hide(axis="index")
            .set_properties(
                **{
                    "text-align": "center !important",
                    "vertical-align": "middle",
                }
            )
            .set_table_styles(
                [
                    {
                        "selector": "th",
                        "props": [
                            ("background-color", "#0066A6"),
                            ("color", "white"),
                            ("text-align", "center !important"),
                            ("vertical-align", "middle"),
                            ("font-weight", "bold"),
                        ],
                    },
                    {
                        "selector": "td",
                        "props": [
                            ("text-align", "center"),
                            ("vertical-align", "middle"),
                        ],
                    },
                    {
                        "selector": "table",
                        "props": [
                            ("width", "100%"),
                            ("border-collapse", "collapse"),
                            ("text-align", "center"),
                        ],
                    },
                ]
            )
        )

        # Pie chart komposisi keseluruhan
        pie_data = tabel_comp.copy()

        # Label legenda menampilkan kategori dan persentasenya
        pie_data["Legenda"] = pie_data["Jenis"] + " (" + pie_data["Persentase"] + ")"

        # Warna kategori
        warna_kategori = {
            "Orang": "#72B7B2",
            "Motor": "#E45756",
            "Mobil": "#F58518",
            "Truk": "#54A24B",
            "Bus": "#4C78A8",
        }

        pie = (
            alt.Chart(pie_data)
            .mark_arc(outerRadius=105, stroke="white", strokeWidth=1.5)
            .encode(
                theta=alt.Theta(field="Total", type="quantitative", stack=True),
                color=alt.Color(
                    field="Legenda",
                    type="nominal",
                    sort=alt.SortField(field="Total", order="descending"),
                    scale=alt.Scale(
                        domain=pie_data["Legenda"].tolist(),
                        range=[
                            warna_kategori.get(jenis, "#999999")
                            for jenis in pie_data["Jenis"]
                        ],
                    ),
                    legend=alt.Legend(title="Kategori", orient="right"),
                ),
                tooltip=[
                    alt.Tooltip(field="Jenis", type="nominal", title="Jenis"),
                    alt.Tooltip(
                        field="Total", type="quantitative", title="Total", format=",.0f"
                    ),
                    alt.Tooltip(field="Persentase", type="nominal", title="Persentase"),
                ],
            )
            .properties(height=280)
        )

        st.altair_chart(pie, use_container_width=True)

        # Tabel komposisi yang sudah ada
        st.table(styled_comp)

    with col2:
        st.subheader("Distribusi per Kategori")

        # Urutkan kategori dari jumlah terbanyak ke tersedikit
        comp_sorted = comp.sort_values(ascending=False)

        # Ubah menjadi DataFrame agar urutan sumbu X bisa diatur
        df_comp = comp_sorted.rename("Total").rename_axis("Jenis").reset_index()

        # Grafik batang vertikal dengan urutan kategori eksplisit
        chart_comp = (
            alt.Chart(df_comp)
            .mark_bar()
            .encode(
                x=alt.X(
                    "Total:Q",
                    title="Total Volume",
                ),
                y=alt.Y(
                    "Jenis:N",
                    sort=df_comp["Jenis"].tolist(),
                    title="Jenis Kendaraan",
                ),
                tooltip=["Jenis", "Total"],
            )
            .properties(
                width="container",
                height=350,
            )
        )

        st.altair_chart(
            chart_comp,
            use_container_width=True,
        )

    st.subheader("Komposisi Kategori per Lokasi")

    # Dropdown untuk memilih lokasi
    pilihan_comp_lokasi = st.multiselect(
        "Pilih lokasi",
        options=sorted(df["Lokasi"].dropna().unique().tolist()),
        default=(by_lokasi_sum.nlargest(min(top_n, 5)).index.tolist()),
        key="komposisi_lokasi",
    )

    if pilihan_comp_lokasi:
        # Hitung komposisi kategori untuk lokasi yang dipilih
        comp_per_lokasi = (
            df[df["Lokasi"].isin(pilihan_comp_lokasi)]
            .groupby("Lokasi")[dp.CATEGORIES]
            .sum()
        )

        # Hilangkan kategori yang totalnya nol
        comp_per_lokasi = comp_per_lokasi.loc[:, comp_per_lokasi.sum(axis=0) > 0]

        # Ubah data ke format panjang
        df_comp_lokasi = comp_per_lokasi.reset_index().melt(
            id_vars="Lokasi",
            var_name="Jenis",
            value_name="Total",
        )

        # Grafik batang vertikal bertumpuk
        chart_comp_lokasi = (
            alt.Chart(df_comp_lokasi)
            .mark_bar()
            .encode(
                x=alt.X(
                    "Lokasi:N",
                    title="Lokasi",
                    sort=pilihan_comp_lokasi,
                    axis=alt.Axis(labelAngle=-45),
                ),
                y=alt.Y(
                    "Total:Q",
                    title="Total Volume",
                    stack="zero",
                ),
                color=alt.Color(
                    "Jenis:N",
                    title="Jenis Kendaraan",
                ),
                tooltip=[
                    alt.Tooltip("Lokasi:N", title="Lokasi"),
                    alt.Tooltip("Jenis:N", title="Jenis Kendaraan"),
                    alt.Tooltip("Total:Q", title="Total Volume"),
                ],
            )
            .properties(
                width="container",
                height=max(490, 35 * len(pilihan_comp_lokasi)),
            )
        )

        st.altair_chart(
            chart_comp_lokasi,
            use_container_width=True,
        )

    else:
        st.caption("Pilih minimal satu lokasi untuk menampilkan komposisi kategori.")


# --- Tab Kualitas Data ---
with tab_kualitas:
    st.subheader("Status Kualitas Data per Lokasi")

    def badge(s):
        if s == "Data lengkap":
            return "Data lengkap"
        if s == "Sebagian data kosong":
            return "Sebagian data kosong"
        if s == "Total tidak cocok":
            return "Total tidak cocok"
        if s == "Perlu verifikasi nilai nol":
            return "Perlu verifikasi nilai nol"
        return "Tidak ada data / kamera tidak aktif"

    dq_display = dq.copy()

    # Terjemahkan nama status tanpa mengubah data dq asli
    dq_display["status"] = dq_display["status"].apply(badge)

    # Sesuaikan nama kolom berdasarkan hasil data_quality_report()
    rename_kolom = {
        "jumlah_hari_tercatat": "Jumlah Hari Tercatat",
        "total_keseluruhan": "Total Keseluruhan",
        "rata_rata_harian": "Rata-rata Volume Harian",
        "hari_nol": "Jumlah Hari Bernilai Nol",
        "hari_data_kosong": "Jumlah Hari Data Kosong",
        "hari_tanpa_data_valid": "Hari Tanpa Data Valid",
        "hari_total_tidak_cocok": "Hari Total Tidak Cocok",
        "status": "Status",
    }

    dq_display = dq_display.rename(columns=rename_kolom)

    # Tampilkan kolom yang tersedia agar tetap kompatibel
    # jika struktur laporan kualitas berbeda.
    kolom_prioritas = [
        "Lokasi",
        "Jumlah Hari Tercatat",
        "Total Keseluruhan",
        "Rata-rata Volume Harian",
        "Jumlah Hari Bernilai Nol",
        "Jumlah Hari Data Kosong",
        "Hari Tanpa Data Valid",
        "Hari Total Tidak Cocok",
        "Status",
    ]

    kolom_tampil = [kolom for kolom in kolom_prioritas if kolom in dq_display.columns]

    st.dataframe(
        dq_display[kolom_tampil],
        width="stretch",
        hide_index=True,
    )

    st.markdown(
        """
        <div style="
            background-color: #FDECEC;
            border-left: 5px solid #D32F2F;
            padding: 12px 16px;
            border-radius: 6px;
            color: #B71C1C;
            margin-top: 12px;
        ">
            <strong>Keterangan Status Kualitas Data</strong>
            <ul style="margin-bottom: 0;">
                <li><strong>Data lengkap</strong>: seluruh kategori yang diharapkan memiliki nilai valid dan nilai Total sesuai dengan penjumlahan kategori.</li>
                <li><strong>Sebagian data kosong</strong>: terdapat kategori yang kosong atau tidak memiliki nilai numerik valid.</li>
                <li><strong>Total tidak cocok</strong>: nilai Total tidak sesuai dengan hasil penjumlahan kategori.</li>
                <li><strong>Perlu verifikasi nilai nol</strong>: proporsi hari dengan Total bernilai nol cukup tinggi sehingga perlu diperiksa kembali.</li>
                <li><strong>Tidak ada data / kamera tidak aktif</strong>: data tidak memiliki nilai kategori valid atau seluruh nilai Total bernilai nol.</li>
            </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )


# --- Tab Data Detail ---
with tab_data:
    st.subheader(f"Data Teragregasi ({periode})")
    st.dataframe(
        agg.sort_values(["Periode", "Total"], ascending=[True, False]),
        use_container_width=True,
        hide_index=True,
    )
    st.download_button(
        "Unduh data teragregasi (CSV)",
        agg.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"data_agregasi_{periode.lower()}.csv",
        mime="text/csv",
    )

    st.subheader("Data Harian Setelah Cleaning")
    st.dataframe(
        df.sort_values(["Lokasi", "Tanggal"]), use_container_width=True, hide_index=True
    )
    st.download_button(
        "Unduh data harian bersih (CSV)",
        df.to_csv(index=False).encode("utf-8-sig"),
        file_name="data_harian_bersih.csv",
        mime="text/csv",
    )
    st.markdown(
        """
            <div style="
                background-color: #FDECEC;
                border-left: 5px solid #D32F2F;
                padding: 12px 16px;
                border-radius: 6px;
                color: #B71C1C;
                margin-top: 12px;
            ">
                <strong>Keterangan</strong>
                <ul style="margin-bottom: 0;">
                    <li><code>_expected_count</code>: jumlah kategori yang diharapkan tersedia pada setiap baris.</li>
                    <li><code>_valid_count</code>: jumlah kategori yang memiliki nilai numerik valid, termasuk nilai nol.</li>
                    <li><code>_missing_count</code>: jumlah kategori yang kosong atau tidak memiliki nilai numerik valid.</li>
                    <li><code>_total_match</code>: menunjukkan apakah nilai Total sesuai dengan penjumlahan kategori ketika seluruh kategori yang diharapkan memiliki nilai valid.</li>
                </ul>
            </div>
            """,
        unsafe_allow_html=True,
    )


# --- Tab Unduh Hasil ---
with tab_unduh:
    st.subheader("Unduh Hasil Analisis")

    st.write("Jabatan default:", rpt.JABATAN_PENANDATANGAN)
    
    st.markdown(
        "Laporan di bawah dibuat otomatis berdasarkan **filter yang sedang aktif** "
        f"(lokasi terpilih: {len(lokasi_terpilih)} lokasi, periode: "
        f"{rentang_tanggal[0].strftime('%d-%m-%Y')} "
        f"s/d {rentang_tanggal[1].strftime('%d-%m-%Y')}, agregasi: {periode})."
    )

    st.subheader("Identitas Laporan (Opsional)")

    nama_laporan = st.text_input(
        "Nama",
        placeholder="Kosongkan untuk menggunakan nama default",
        key="nama_laporan",
    )

    nip_laporan = st.text_input(
        "NIP",
        placeholder="Kosongkan untuk menggunakan identitas default",
        key="nip_laporan",
    )

    meta = {
        "periode_label": periode,
        "lokasi_terpilih": lokasi_terpilih,
        "sumber_data": sumber_data,
        # Identitas laporan
        "nama": nama_laporan.strip() if nama_laporan.strip() else "Noor Aflah, S.Kom.",
        "nip": nip_laporan.strip() if nip_laporan.strip() else "197609072003121007",
    }

    st.markdown("#### Laporan Dokumen (Word)")
    st.caption(
        "Ringkasan eksekutif otomatis, KPI, grafik tren & komposisi, "
        "tabel lengkap per lokasi, dan catatan kualitas data."
    )

    if st.button("Buat Laporan Word", type="primary", key="btn_docx"):
        with st.spinner("Menyusun dokumen laporan..."):

            # Membuat dokumen Word
            docx_buf = rpt.generate_docx_report(df, dq, meta)
            docx_bytes = docx_buf.getvalue()

            # Menyimpan sementara dan mengonversi Word ke PDF
            with tempfile.TemporaryDirectory() as tmp:
                docx_path = os.path.join(tmp, "laporan.docx")
                pdf_path = os.path.join(tmp, "laporan.pdf")

                with open(docx_path, "wb") as f:
                    f.write(docx_bytes)

                # Mencari aplikasi LibreOffice
                libreoffice_cmd = shutil.which("libreoffice") or shutil.which("soffice")

                if not libreoffice_cmd:
                    raise RuntimeError(
                        "LibreOffice tidak ditemukan. "
                        "Pastikan packages.txt sudah dikonfigurasi."
                    )

                # Konversi DOCX menjadi PDF
                subprocess.run(
                    [
                        libreoffice_cmd,
                        "--headless",
                        "--convert-to",
                        "pdf",
                        "--outdir",
                        tmp,
                        docx_path,
                    ],
                    check=True,
                    capture_output=True,
                    text=True,
                )

                with open(pdf_path, "rb") as f:
                    pdf_bytes = f.read()

            # Menyimpan hasil agar tidak hilang saat Streamlit rerun
            st.session_state["laporan_docx"] = docx_bytes
            st.session_state["laporan_pdf"] = pdf_bytes

    # Tombol unduh Word
    if "laporan_docx" in st.session_state:
        st.download_button(
            "Unduh Laporan (.docx)",
            st.session_state["laporan_docx"],
            file_name=f"laporan_analisis_cctv_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    # Pratinjau PDF
    if "laporan_pdf" in st.session_state:
        st.markdown("#### Pratinjau Laporan")

        pdf_base64 = base64.b64encode(
            st.session_state["laporan_pdf"]
        ).decode("utf-8")

        pdf_display = f"""
            <iframe
                src="data:application/pdf;base64,{pdf_base64}"
                width="100%"
                height="700"
                style="border: none; border-radius: 8px;"
            ></iframe>
        """

        st.markdown(
            pdf_display,
            unsafe_allow_html=True
        )


st.markdown("---")
st.caption(
    "Dashboard Analisis CCTV · Dibuat dengan Streamlit · "
    "Pastikan data yang diunggah berasal dari sumber yang valid untuk hasil analisis yang akurat."
)
