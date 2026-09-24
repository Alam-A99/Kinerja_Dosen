# ============================================================
# DASHBOARD KINERJA DOSEN
# BISNIS DIGITAL FEB UNM
# ============================================================

import os
import base64

import streamlit as st
import pandas as pd
import numpy as np


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Dashboard Kinerja Dosen | Bisnis Digital FEB UNM",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# FILE CONFIGURATION
# ============================================================

EXCEL_FILE = "DATA KINERJA DOSEN BISDIG 2022-2025_OK.xlsx"
LOGO_FILE = "logobd.png"


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>

    .stApp {
        background-color: #f5f7fb;
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
        max-width: 1500px;
    }

    /* HEADER */

    .dashboard-header {
        display: flex;
        align-items: center;
        gap: 25px;

        background: linear-gradient(
            135deg,
            #ffffff,
            #f2f8ff
        );

        border: 1px solid #dce6f1;
        border-radius: 18px;

        padding: 25px 30px;

        margin-bottom: 25px;

        box-shadow:
            0 5px 18px rgba(0,0,0,0.05);
    }

    .logo-box {
        width: 90px;
        height: 90px;

        display: flex;
        align-items: center;
        justify-content: center;

        flex-shrink: 0;
    }

    .logo-box img {
        max-width: 90px;
        max-height: 90px;
        object-fit: contain;
    }

    .dashboard-title {
        font-size: 30px;
        font-weight: 800;
        color: #123b68;
        line-height: 1.25;
    }

    .dashboard-subtitle {
        font-size: 14px;
        color: #64748b;
        margin-top: 8px;
    }


    /* SECTION */

    .section-title {
        font-size: 21px;
        font-weight: 800;
        color: #172033;

        margin-top: 25px;
        margin-bottom: 15px;
    }


    /* KPI */

    .kpi-card {
        background: white;

        border: 1px solid #e2e8f0;

        border-radius: 16px;

        padding: 20px;

        min-height: 125px;

        box-shadow:
            0 4px 14px rgba(15,23,42,0.05);
    }

    .kpi-title {
        font-size: 12px;
        font-weight: 700;

        color: #64748b;

        text-transform: uppercase;

        letter-spacing: 0.5px;
    }

    .kpi-value {
        font-size: 30px;
        font-weight: 850;

        color: #123b68;

        margin-top: 7px;
    }

    .kpi-note {
        font-size: 11px;
        color: #94a3b8;

        margin-top: 5px;
    }


    /* INFO */

    .info-card {
        background: white;

        border: 1px solid #e2e8f0;

        border-radius: 15px;

        padding: 20px;

        margin-bottom: 15px;
    }


    /* FOOTER */

    .footer {
        text-align: center;

        margin-top: 40px;

        padding: 25px;

        color: #64748b;

        font-size: 12px;

        border-top: 1px solid #e2e8f0;
    }


    @media (max-width: 768px) {

        .dashboard-header {
            flex-direction: column;
            align-items: flex-start;
        }

        .dashboard-title {
            font-size: 23px;
        }

    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# LOGO
# ============================================================

def get_logo():

    if not os.path.exists(LOGO_FILE):
        return None

    try:

        with open(LOGO_FILE, "rb") as f:

            return base64.b64encode(
                f.read()
            ).decode()

    except Exception:

        return None


logo = get_logo()


# ============================================================
# HEADER
# ============================================================

if logo:

    st.markdown(
        f"""
        <div class="dashboard-header">

            <div class="logo-box">

                <img
                    src="data:image/png;base64,{logo}"
                    alt="Logo Bisnis Digital FEB UNM"
                >

            </div>

            <div>

                <div class="dashboard-title">

                    Dashboard Kinerja Dosen
                    <br>
                    Bisnis Digital FEB UNM

                </div>

                <div class="dashboard-subtitle">

                    Evidence-Based Performance Dashboard
                    &nbsp; | &nbsp;
                    Pendidikan · Penelitian · PkM · Penunjang

                </div>

            </div>

        </div>
        """,
        unsafe_allow_html=True
    )

else:

    st.markdown(
        """
        <div class="dashboard-header">

            <div>

                <div class="dashboard-title">

                    Dashboard Kinerja Dosen
                    <br>
                    Bisnis Digital FEB UNM

                </div>

                <div class="dashboard-subtitle">

                    Evidence-Based Performance Dashboard

                </div>

            </div>

        </div>
        """,
        unsafe_allow_html=True
    )

    st.warning(
        "File logobd.png tidak ditemukan."
    )


# ============================================================
# LOAD DATA
# ============================================================

@st.cache_data
def load_data():

    if not os.path.exists(EXCEL_FILE):

        return None, (
            "File Excel tidak ditemukan: "
            + EXCEL_FILE
        )

    try:

        data = pd.read_excel(
            EXCEL_FILE,
            engine="openpyxl"
        )

        return data, None

    except Exception as e:

        return None, str(e)


df, error = load_data()


if error:

    st.error(
        "❌ Gagal membaca dataset"
    )

    st.code(error)

    st.stop()


if df is None or df.empty:

    st.error(
        "Dataset kosong."
    )

    st.stop()


# ============================================================
# NORMALIZE COLUMN
# ============================================================

def normalize_column(col):

    col = str(col).strip().lower()

    for old, new in [
        (" ", "_"),
        ("-", "_"),
        ("/", "_"),
        (".", "_")
    ]:

        col = col.replace(old, new)

    return col


df.columns = [
    normalize_column(c)
    for c in df.columns
]


# ============================================================
# FIND COLUMN
# ============================================================

def find_column(candidates):

    for candidate in candidates:

        candidate = normalize_column(
            candidate
        )

        if candidate in df.columns:

            return candidate

    for candidate in candidates:

        candidate = normalize_column(
            candidate
        )

        for col in df.columns:

            if candidate in col:

                return col

    return None


NAME_COL = find_column([
    "nama",
    "nama_dosen",
    "dosen",
    "nama dosen"
])


TYPE_COL = find_column([
    "jenis",
    "kategori",
    "bidang",
    "jenis_kegiatan",
    "jenis kegiatan"
])


YEAR_COL = find_column([
    "tahun",
    "tahun_akademik",
    "tahun akademik"
])


SEMESTER_COL = find_column([
    "semester",
    "sem"
])


ACTIVITY_COL = find_column([
    "aktivitas",
    "kegiatan",
    "uraian",
    "jenis_aktivitas"
])


EVIDENCE_COL = find_column([
    "bukti",
    "dokumen",
    "file",
    "link",
    "url",
    "evidence"
])


# ============================================================
# CATEGORY
# ============================================================

def classify_activity(value):

    if pd.isna(value):

        return "Lainnya"

    text = str(value).lower()

    if any(
        x in text
        for x in [
            "pendidikan",
            "pengajaran",
            "pembelajaran",
            "kuliah"
        ]
    ):

        return "Pendidikan"

    if any(
        x in text
        for x in [
            "penelitian",
            "research",
            "riset"
        ]
    ):

        return "Penelitian"

    if any(
        x in text
        for x in [
            "pkm",
            "pengabdian",
            "masyarakat"
        ]
    ):

        return "PkM"

    if any(
        x in text
        for x in [
            "penunjang",
            "organisasi",
            "kepanitiaan",
            "seminar",
            "asosiasi"
        ]
    ):

        return "Penunjang"

    return str(value).strip()


if TYPE_COL:

    df["kategori"] = (
        df[TYPE_COL]
        .apply(classify_activity)
    )

else:

    df["kategori"] = "Lainnya"


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title("🎛️ FILTER")

st.sidebar.markdown("---")


# YEAR

if YEAR_COL:

    years = sorted(
        df[YEAR_COL]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    selected_years = st.sidebar.multiselect(
        "📅 Tahun",
        years,
        default=years
    )

else:

    selected_years = []


# SEMESTER

if SEMESTER_COL:

    semesters = sorted(
        df[SEMESTER_COL]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    selected_semesters = st.sidebar.multiselect(
        "🎓 Semester",
        semesters,
        default=semesters
    )

else:

    selected_semesters = []


# CATEGORY

categories = sorted(
    df["kategori"]
    .dropna()
    .unique()
    .tolist()
)


selected_categories = st.sidebar.multiselect(
    "📚 Bidang Kinerja",
    categories,
    default=categories
)


# DOSEN

if NAME_COL:

    lecturers = sorted(
        df[NAME_COL]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    selected_lecturers = st.sidebar.multiselect(
        "👤 Dosen",
        lecturers
    )

else:

    selected_lecturers = []


# RESET

if st.sidebar.button(
    "🔄 Reset Filter",
    use_container_width=True
):

    st.rerun()


# ============================================================
# FILTER DATA
# ============================================================

filtered = df.copy()


if YEAR_COL and selected_years:

    filtered = filtered[
        filtered[YEAR_COL]
        .astype(str)
        .isin(selected_years)
    ]


if SEMESTER_COL and selected_semesters:

    filtered = filtered[
        filtered[SEMESTER_COL]
        .astype(str)
        .isin(selected_semesters)
    ]


if selected_categories:

    filtered = filtered[
        filtered["kategori"]
        .isin(selected_categories)
    ]


if NAME_COL and selected_lecturers:

    filtered = filtered[
        filtered[NAME_COL]
        .astype(str)
        .isin(selected_lecturers)
    ]


# ============================================================
# KPI
# ============================================================

total_activity = len(filtered)


if NAME_COL:

    total_dosen = (
        filtered[NAME_COL]
        .nunique()
    )

else:

    total_dosen = 0


pendidikan = int(
    (
        filtered["kategori"]
        == "Pendidikan"
    ).sum()
)


penelitian = int(
    (
        filtered["kategori"]
        == "Penelitian"
    ).sum()
)


pkm = int(
    (
        filtered["kategori"]
        == "PkM"
    ).sum()
)


penunjang = int(
    (
        filtered["kategori"]
        == "Penunjang"
    ).sum()
)


# ============================================================
# EXECUTIVE SUMMARY
# ============================================================

st.markdown(
    """
    <div class="section-title">
        📊 Executive Performance Summary
    </div>
    """,
    unsafe_allow_html=True
)


cols = st.columns(6)


kpi_data = [

    (
        "TOTAL AKTIVITAS",
        total_activity,
        "rekam kinerja"
    ),

    (
        "DOSEN",
        total_dosen,
        "dosen"
    ),

    (
        "PENDIDIKAN",
        pendidikan,
        "aktivitas"
    ),

    (
        "PENELITIAN",
        penelitian,
        "aktivitas"
    ),

    (
        "PkM",
        pkm,
        "aktivitas"
    ),

    (
        "PENUNJANG",
        penunjang,
        "aktivitas"
    )

]


for col, item in zip(
    cols,
    kpi_data
):

    title, value, note = item

    with col:

        st.markdown(
            f"""
            <div class="kpi-card">

                <div class="kpi-title">
                    {title}
                </div>

                <div class="kpi-value">
                    {value:,}
                </div>

                <div class="kpi-note">
                    {note}
                </div>

            </div>
            """,
            unsafe_allow_html=True
        )


# ============================================================
# KOMPOSISI KINERJA
# ============================================================

st.markdown(
    """
    <div class="section-title">
        📊 Komposisi Kinerja
    </div>
    """,
    unsafe_allow_html=True
)


composition = (

    filtered["kategori"]
    .value_counts()
    .rename("Jumlah")
    .to_frame()
)


st.bar_chart(
    composition,
    use_container_width=True
)


# ============================================================
# TREND
# ============================================================

st.markdown(
    """
    <div class="section-title">
        📈 Tren Kinerja
    </div>
    """,
    unsafe_allow_html=True
)


if YEAR_COL:

    trend = filtered.copy()

    trend["tahun_display"] = (
        trend[YEAR_COL]
        .astype(str)
    )

    trend_table = pd.crosstab(
        trend["tahun_display"],
        trend["kategori"]
    )

    trend_table = trend_table.sort_index()

    st.line_chart(
        trend_table,
        use_container_width=True
    )

else:

    st.info(
        "Kolom tahun tidak ditemukan."
    )


# ============================================================
# KINERJA PER DOSEN
# ============================================================

st.markdown(
    """
    <div class="section-title">
        👥 Kinerja Dosen
    </div>
    """,
    unsafe_allow_html=True
)


if NAME_COL:

    lecturer_table = pd.crosstab(
        filtered[NAME_COL],
        filtered["kategori"]
    )

    lecturer_table["TOTAL"] = (
        lecturer_table.sum(axis=1)
    )

    lecturer_table = (
        lecturer_table
        .sort_values(
            "TOTAL",
            ascending=False
        )
    )

    st.dataframe(
        lecturer_table,
        use_container_width=True,
        height=450
    )

else:

    st.warning(
        "Kolom nama dosen tidak ditemukan."
    )


# ============================================================
# TOP DOSEN
# ============================================================

st.markdown(
    """
    <div class="section-title">
        🏆 Aktivitas Dosen
    </div>
    """,
    unsafe_allow_html=True
)


if NAME_COL:

    top_dosen = (

        filtered[NAME_COL]
        .value_counts()
        .head(10)
        .sort_values()

    )

    st.bar_chart(
        top_dosen,
        horizontal=True,
        use_container_width=True
    )


# ============================================================
# PROFIL TRIDHARMA
# ============================================================

st.markdown(
    """
    <div class="section-title">
        ⚖️ Profil Tridharma dan Penunjang
    </div>
    """,
    unsafe_allow_html=True
)


profil = (

    filtered["kategori"]
    .value_counts()
    .reindex(
        [
            "Pendidikan",
            "Penelitian",
            "PkM",
            "Penunjang"
        ],
        fill_value=0
    )
)


profil_total = profil.sum()


if profil_total > 0:

    profil_persen = (
        profil /
        profil_total *
        100
    )

else:

    profil_persen = profil * 0


profil_df = pd.DataFrame({

    "Bidang": profil.index,

    "Jumlah": profil.values,

    "Persentase (%)":
        profil_persen.round(2).values

})


st.dataframe(
    profil_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# INDIKATOR
# ============================================================

st.markdown(
    """
    <div class="section-title">
        🎯 Indikator Kinerja
    </div>
    """,
    unsafe_allow_html=True
)


if total_dosen > 0:

    rata_rata = (
        total_activity /
        total_dosen
    )

else:

    rata_rata = 0


if total_activity > 0:

    proporsi_penelitian = (
        penelitian /
        total_activity *
        100
    )

    proporsi_pkm = (
        pkm /
        total_activity *
        100
    )

else:

    proporsi_penelitian = 0
    proporsi_pkm = 0


c1, c2, c3 = st.columns(3)


with c1:

    st.metric(
        "Rata-rata Aktivitas / Dosen",
        f"{rata_rata:.1f}"
    )


with c2:

    st.metric(
        "Proporsi Penelitian",
        f"{proporsi_penelitian:.1f}%"
    )


with c3:

    st.metric(
        "Proporsi PkM",
        f"{proporsi_pkm:.1f}%"
    )


# ============================================================
# DATA EVIDENCE
# ============================================================

st.markdown(
    """
    <div class="section-title">
        🔎 Evidence Explorer
    </div>
    """,
    unsafe_allow_html=True
)


evidence_cols = []


if NAME_COL:
    evidence_cols.append(NAME_COL)

if YEAR_COL:
    evidence_cols.append(YEAR_COL)

if SEMESTER_COL:
    evidence_cols.append(SEMESTER_COL)

evidence_cols.append("kategori")


if ACTIVITY_COL:
    evidence_cols.append(ACTIVITY_COL)

if EVIDENCE_COL:
    evidence_cols.append(EVIDENCE_COL)


evidence = filtered[
    evidence_cols
].copy()


# Rename

rename = {}


if NAME_COL:
    rename[NAME_COL] = "Dosen"

if YEAR_COL:
    rename[YEAR_COL] = "Tahun"

if SEMESTER_COL:
    rename[SEMESTER_COL] = "Semester"

rename["kategori"] = "Bidang"


if ACTIVITY_COL:
    rename[ACTIVITY_COL] = "Aktivitas"

if EVIDENCE_COL:
    rename[EVIDENCE_COL] = "Bukti"


evidence = evidence.rename(
    columns=rename
)


# ============================================================
# DISPLAY EVIDENCE
# ============================================================

st.dataframe(
    evidence,
    use_container_width=True,
    height=450,
    hide_index=True
)


# ============================================================
# DOWNLOAD
# ============================================================

st.markdown(
    """
    <div class="section-title">
        ⬇️ Export Data
    </div>
    """,
    unsafe_allow_html=True
)


csv = filtered.to_csv(
    index=False
).encode("utf-8")


st.download_button(
    label="⬇️ Download Data CSV",
    data=csv,
    file_name="kinerja_dosen_bisnis_digital.csv",
    mime="text/csv",
    use_container_width=True
)


# ============================================================
# DATASET INFO
# ============================================================

with st.expander(
    "🔧 Informasi Dataset"
):

    st.write(
        f"Jumlah baris: **{len(df):,}**"
    )

    st.write(
        f"Jumlah kolom: **{len(df.columns):,}**"
    )

    st.write(
        "Kolom dataset:"
    )

    st.code(
        "\n".join(
            df.columns.tolist()
        )
    )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div class="footer">

        <b>
            Dashboard Kinerja Dosen
            Bisnis Digital FEB UNM
        </b>

        <br><br>

        Evidence-Based Performance Dashboard

        <br>

        Pendidikan · Penelitian · PkM · Penunjang

    </div>
    """,
    unsafe_allow_html=True
)
