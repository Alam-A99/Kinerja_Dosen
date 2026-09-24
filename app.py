import os
import streamlit as st
import pandas as pd


# =========================================================
# CONFIG
# =========================================================

st.set_page_config(
    page_title="Dashboard Kinerja Dosen - Bisnis Digital FEB UNM",
    page_icon="📊",
    layout="wide"
)


# =========================================================
# FILE
# =========================================================

EXCEL_FILE = "DATA KINERJA DOSEN BISDIG 2022-2025_OK.xlsx"
LOGO_FILE = "logobd.png"


# =========================================================
# STYLE
# =========================================================

st.markdown("""
<style>

.main-title {
    font-size: 32px;
    font-weight: 800;
    color: #123b68;
    margin-bottom: 0;
}

.sub-title {
    font-size: 15px;
    color: #64748b;
    margin-top: 5px;
}

.section-title {
    font-size: 22px;
    font-weight: 800;
    color: #172033;
    margin-top: 30px;
    margin-bottom: 15px;
}

div[data-testid="stMetric"] {
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 15px;
    padding: 18px;
    box-shadow: 0 3px 10px rgba(0,0,0,0.05);
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# HEADER
# =========================================================

col_logo, col_title = st.columns([1, 7])

with col_logo:

    if os.path.exists(LOGO_FILE):

        st.image(
            LOGO_FILE,
            width=90
        )

with col_title:

    st.markdown(
        '<div class="main-title">'
        'Dashboard Kinerja Dosen'
        '<br>'
        'Bisnis Digital FEB UNM'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="sub-title">'
        'Evidence-Based Performance Dashboard '
        '| Pendidikan • Penelitian • PkM • Penunjang'
        '</div>',
        unsafe_allow_html=True
    )


st.divider()


# =========================================================
# LOAD EXCEL
# =========================================================

if not os.path.exists(EXCEL_FILE):

    st.error(
        "File Excel tidak ditemukan."
    )

    st.write(
        "Pastikan file berikut berada satu folder dengan app.py:"
    )

    st.code(EXCEL_FILE)

    st.stop()


try:

    df = pd.read_excel(
        EXCEL_FILE,
        engine="openpyxl"
    )

except Exception as e:

    st.error(
        "Gagal membaca file Excel."
    )

    st.exception(e)

    st.stop()


# =========================================================
# CLEAN COLUMN NAMES
# =========================================================

df.columns = [
    str(c).strip()
    for c in df.columns
]


# =========================================================
# DETECT COLUMN
# =========================================================

def get_column(possible_names):

    for name in possible_names:

        for col in df.columns:

            if str(col).strip().lower() == name.lower():

                return col

    for name in possible_names:

        for col in df.columns:

            if name.lower() in str(col).lower():

                return col

    return None


NAMA = get_column([
    "Nama",
    "Nama Dosen",
    "Dosen"
])


TAHUN = get_column([
    "Tahun",
    "Tahun Akademik"
])


JENIS = get_column([
    "Jenis",
    "Kategori",
    "Bidang",
    "Jenis Kegiatan"
])


KEGIATAN = get_column([
    "Kegiatan",
    "Aktivitas",
    "Uraian"
])


# =========================================================
# SIDEBAR
# =========================================================

st.sidebar.title("🎛️ FILTER DASHBOARD")

st.sidebar.markdown("---")


# Tahun

if TAHUN is not None:

    daftar_tahun = sorted(
        df[TAHUN]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    tahun_pilih = st.sidebar.multiselect(
        "📅 Tahun",
        daftar_tahun,
        default=daftar_tahun
    )

else:

    tahun_pilih = []


# Dosen

if NAMA is not None:

    daftar_dosen = sorted(
        df[NAMA]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    dosen_pilih = st.sidebar.multiselect(
        "👤 Dosen",
        daftar_dosen
    )

else:

    dosen_pilih = []


# =========================================================
# FILTER
# =========================================================

data = df.copy()


if TAHUN is not None and tahun_pilih:

    data = data[
        data[TAHUN]
        .astype(str)
        .isin(tahun_pilih)
    ]


if NAMA is not None and dosen_pilih:

    data = data[
        data[NAMA]
        .astype(str)
        .isin(dosen_pilih)
    ]


# =========================================================
# CLASSIFICATION
# =========================================================

def kategori_kinerja(row):

    text = ""

    for col in df.columns:

        try:

            text += " " + str(
                row[col]
            ).lower()

        except:

            pass

    if (
        "pendidikan" in text
        or "pengajaran" in text
        or "pembelajaran" in text
    ):

        return "Pendidikan"

    if (
        "penelitian" in text
        or "riset" in text
    ):

        return "Penelitian"

    if (
        "pengabdian" in text
        or "pkm" in text
        or "masyarakat" in text
    ):

        return "PkM"

    if (
        "penunjang" in text
        or "kepanitiaan" in text
        or "organisasi" in text
    ):

        return "Penunjang"

    return "Lainnya"


data["Kategori Dashboard"] = data.apply(
    kategori_kinerja,
    axis=1
)


# =========================================================
# KPI
# =========================================================

st.markdown(
    '<div class="section-title">'
    '📊 Executive Performance Summary'
    '</div>',
    unsafe_allow_html=True
)


total = len(data)


if NAMA is not None:

    jumlah_dosen = data[NAMA].nunique()

else:

    jumlah_dosen = 0


pendidikan = (
    data["Kategori Dashboard"]
    == "Pendidikan"
).sum()


penelitian = (
    data["Kategori Dashboard"]
    == "Penelitian"
).sum()


pkm = (
    data["Kategori Dashboard"]
    == "PkM"
).sum()


penunjang = (
    data["Kategori Dashboard"]
    == "Penunjang"
).sum()


c1, c2, c3, c4, c5, c6 = st.columns(6)


with c1:

    st.metric(
        "TOTAL AKTIVITAS",
        f"{total:,}"
    )


with c2:

    st.metric(
        "DOSEN",
        f"{jumlah_dosen:,}"
    )


with c3:

    st.metric(
        "PENDIDIKAN",
        f"{pendidikan:,}"
    )


with c4:

    st.metric(
        "PENELITIAN",
        f"{penelitian:,}"
    )


with c5:

    st.metric(
        "PkM",
        f"{pkm:,}"
    )


with c6:

    st.metric(
        "PENUNJANG",
        f"{penunjang:,}"
    )


# =========================================================
# KOMPOSISI
# =========================================================

st.markdown(
    '<div class="section-title">'
    '📈 Komposisi Kinerja'
    '</div>',
    unsafe_allow_html=True
)


komposisi = (
    data["Kategori Dashboard"]
    .value_counts()
)


st.bar_chart(
    komposisi
)


# =========================================================
# TREND TAHUN
# =========================================================

if TAHUN is not None:

    st.markdown(
        '<div class="section-title">'
        '📅 Tren Kinerja'
        '</div>',
        unsafe_allow_html=True
    )

    trend = pd.crosstab(
        data[TAHUN].astype(str),
        data["Kategori Dashboard"]
    )

    st.line_chart(
        trend
    )


# =========================================================
# KINERJA DOSEN
# =========================================================

if NAMA is not None:

    st.markdown(
        '<div class="section-title">'
        '👥 Kinerja Per Dosen'
        '</div>',
        unsafe_allow_html=True
    )

    tabel_dosen = pd.crosstab(
        data[NAMA],
        data["Kategori Dashboard"]
    )

    tabel_dosen["TOTAL"] = (
        tabel_dosen.sum(axis=1)
    )

    tabel_dosen = tabel_dosen.sort_values(
        "TOTAL",
        ascending=False
    )

    st.dataframe(
        tabel_dosen,
        use_container_width=True,
        height=450
    )


# =========================================================
# DATA DETAIL
# =========================================================

st.markdown(
    '<div class="section-title">'
    '🔎 Evidence Explorer'
    '</div>',
    unsafe_allow_html=True
)


st.dataframe(
    data,
    use_container_width=True,
    height=500
)


# =========================================================
# DOWNLOAD
# =========================================================

csv = data.to_csv(
    index=False
).encode("utf-8")


st.download_button(
    label="⬇️ Download Data CSV",
    data=csv,
    file_name="kinerja_dosen_bisnis_digital.csv",
    mime="text/csv"
)


# =========================================================
# FOOTER
# =========================================================

st.divider()

st.caption(
    "Dashboard Kinerja Dosen | "
    "Program Studi Bisnis Digital FEB UNM"
)
