# ============================================================
# DASHBOARD KINERJA DOSEN
# PROGRAM STUDI BISNIS DIGITAL FEB UNM
# ============================================================
#
# Evidence-Based Performance Dashboard
# Pendidikan | Penelitian | PkM | Penunjang
#
# Struktur folder:
#
# aplikasi/
# ├── app.py
# ├── logobd.png
# └── DATA KINERJA DOSEN BISDIG 2022-2025_OK.xlsx
#
# Jalankan:
# streamlit run app.py
# ============================================================


# ============================================================
# 1. IMPORT LIBRARY
# ============================================================

import os
import base64

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px


# ============================================================
# 2. PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Dashboard Kinerja Dosen | Bisnis Digital FEB UNM",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# 3. KONFIGURASI FILE
# ============================================================

EXCEL_FILE = "DATA KINERJA DOSEN BISDIG 2022-2025_OK.xlsx"
LOGO_FILE = "logobd.png"


# ============================================================
# 4. CUSTOM CSS
# ============================================================

st.html("""
<style>

/* =========================================================
   GLOBAL
========================================================= */

.stApp {
    background: #f5f7fb;
}

.block-container {
    padding-top: 1.5rem;
    padding-bottom: 3rem;
    max-width: 1500px;
}


/* =========================================================
   SIDEBAR
========================================================= */

section[data-testid="stSidebar"] {
    background: #f8fafc;
    border-right: 1px solid #e5e7eb;
}

section[data-testid="stSidebar"] h2 {
    color: #123b68;
}


/* =========================================================
   HEADER
========================================================= */

.dashboard-header {
    width: 100%;
    display: flex;
    align-items: center;
    gap: 24px;

    background: linear-gradient(
        135deg,
        #ffffff 0%,
        #f7fbff 100%
    );

    border: 1px solid #dbe5f0;
    border-radius: 18px;

    padding: 24px 30px;

    box-shadow:
        0 6px 20px rgba(15, 23, 42, 0.06);

    margin-bottom: 25px;
}


/* LOGO */

.logo-box {
    width: 90px;
    height: 90px;

    min-width: 90px;

    display: flex;
    align-items: center;
    justify-content: center;
}

.logo-box img {
    max-width: 90px;
    max-height: 90px;

    width: auto;
    height: auto;

    object-fit: contain;
}


/* HEADER TEXT */

.header-text {
    flex: 1;
}

.dashboard-title {
    font-size: 30px;
    font-weight: 800;

    color: #123b68;

    line-height: 1.2;

    margin-bottom: 8px;
}

.dashboard-subtitle {
    font-size: 14px;

    color: #64748b;

    line-height: 1.5;
}


/* =========================================================
   SECTION TITLE
========================================================= */

.section-title {
    font-size: 21px;

    font-weight: 800;

    color: #172033;

    margin-top: 25px;

    margin-bottom: 14px;
}


/* =========================================================
   KPI
========================================================= */

.kpi-card {
    background: #ffffff;

    border: 1px solid #e2e8f0;

    border-radius: 16px;

    padding: 20px;

    min-height: 125px;

    box-shadow:
        0 4px 14px rgba(15, 23, 42, 0.05);

    transition:
        transform 0.2s ease,
        box-shadow 0.2s ease;
}

.kpi-card:hover {
    transform: translateY(-2px);

    box-shadow:
        0 8px 22px rgba(15, 23, 42, 0.08);
}

.kpi-title {
    font-size: 12px;

    font-weight: 750;

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

    margin-top: 4px;
}


/* =========================================================
   INFO CARD
========================================================= */

.info-card {
    background: #ffffff;

    border: 1px solid #e2e8f0;

    border-radius: 15px;

    padding: 20px;

    box-shadow:
        0 4px 14px rgba(15, 23, 42, 0.04);

    margin-bottom: 15px;
}


/* =========================================================
   EVIDENCE CARD
========================================================= */

.evidence-card {
    background: #ffffff;

    border: 1px solid #e2e8f0;

    border-radius: 15px;

    padding: 18px;

    margin-bottom: 12px;

    box-shadow:
        0 3px 12px rgba(15, 23, 42, 0.04);
}

.evidence-title {
    font-weight: 750;

    color: #123b68;

    font-size: 14px;
}

.evidence-text {
    color: #64748b;

    font-size: 13px;

    margin-top: 5px;
}


/* =========================================================
   FOOTER
========================================================= */

.footer {
    text-align: center;

    margin-top: 40px;

    padding: 25px;

    color: #64748b;

    font-size: 12px;

    border-top: 1px solid #e2e8f0;
}


/* =========================================================
   METRIC
========================================================= */

div[data-testid="stMetric"] {
    background: white;

    padding: 15px;

    border-radius: 12px;

    border: 1px solid #e5e7eb;
}


/* =========================================================
   DATAFRAME
========================================================= */

div[data-testid="stDataFrame"] {
    border-radius: 12px;
}


/* =========================================================
   RESPONSIVE
========================================================= */

@media (max-width: 768px) {

    .dashboard-header {
        flex-direction: column;

        align-items: flex-start;

        padding: 20px;
    }

    .dashboard-title {
        font-size: 23px;
    }

    .logo-box {
        width: 70px;
        height: 70px;
    }

    .logo-box img {
        max-width: 70px;
        max-height: 70px;
    }

}

</style>
""")


# ============================================================
# 5. FUNGSI LOGO
# ============================================================

def image_to_base64(path):

    if not os.path.exists(path):
        return None

    try:

        with open(path, "rb") as file:

            encoded = base64.b64encode(
                file.read()
            ).decode("utf-8")

        return encoded

    except Exception:

        return None


logo_base64 = image_to_base64(LOGO_FILE)


# ============================================================
# 6. HEADER
# ============================================================

if logo_base64:

    st.html(f"""
    <div class="dashboard-header">

        <div class="logo-box">

            <img
                src="data:image/png;base64,{logo_base64}"
                alt="Logo Bisnis Digital FEB UNM"
            >

        </div>

        <div class="header-text">

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
    """)

else:

    st.html("""
    <div class="dashboard-header">

        <div class="header-text">

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
    """)

    st.warning(
        "File logobd.png tidak ditemukan. "
        "Pastikan file berada dalam folder yang sama dengan app.py."
    )


# ============================================================
# 7. FUNGSI NORMALISASI NAMA KOLOM
# ============================================================

def normalize_column_name(column):

    column = str(column).strip().lower()

    replacements = {

        " ": "_",
        "-": "_",
        "/": "_",
        ".": "_",

    }

    for old, new in replacements.items():

        column = column.replace(old, new)

    return column


# ============================================================
# 8. LOAD DATA
# ============================================================

@st.cache_data
def load_data():

    if not os.path.exists(EXCEL_FILE):

        return None, (
            f"File Excel tidak ditemukan: "
            f"{EXCEL_FILE}"
        )

    try:

        data = pd.read_excel(
            EXCEL_FILE
        )

        original_columns = list(
            data.columns
        )

        data.columns = [

            normalize_column_name(c)

            for c in data.columns

        ]

        return data, None

    except Exception as error:

        return None, str(error)


df, load_error = load_data()


# ============================================================
# 9. VALIDASI DATA
# ============================================================

if load_error:

    st.error(
        f"❌ {load_error}"
    )

    st.info(
        "Pastikan file Excel berada dalam folder "
        "yang sama dengan app.py."
    )

    st.stop()


if df is None or df.empty:

    st.error(
        "Dataset kosong atau tidak berhasil dibaca."
    )

    st.stop()


# ============================================================
# 10. IDENTIFIKASI KOLOM
# ============================================================

def find_column(dataframe, candidates):

    columns = list(
        dataframe.columns
    )

    for candidate in candidates:

        candidate = normalize_column_name(
            candidate
        )

        if candidate in columns:

            return candidate

    # pencarian parsial
    for candidate in candidates:

        candidate = normalize_column_name(
            candidate
        )

        for col in columns:

            if candidate in col:

                return col

    return None


NAME_COL = find_column(
    df,
    [
        "nama",
        "nama_dosen",
        "dosen",
        "nama dosen",
        "lecturer"
    ]
)


TYPE_COL = find_column(
    df,
    [
        "jenis",
        "kategori",
        "bidang",
        "jenis_kegiatan",
        "jenis kegiatan"
    ]
)


YEAR_COL = find_column(
    df,
    [
        "tahun_akademik",
        "tahun",
        "tahun akademik",
        "academic_year"
    ]
)


SEMESTER_COL = find_column(
    df,
    [
        "semester",
        "sem"
    ]
)


RUBRIC_COL = find_column(
    df,
    [
        "rubrik",
        "aktivitas",
        "kegiatan",
        "jenis_aktivitas",
        "uraian"
    ]
)


FILE_COL = find_column(
    df,
    [
        "file",
        "bukti",
        "dokumen",
        "link",
        "url",
        "evidence"
    ]
)


# ============================================================
# 11. INFORMASI KOLOM
# ============================================================

with st.expander(
    "🔧 Informasi struktur dataset"
):

    st.write(
        "Kolom yang terdeteksi:"
    )

    detected = pd.DataFrame({

        "Komponen": [

            "Nama Dosen",
            "Jenis/Kategori",
            "Tahun Akademik",
            "Semester",
            "Rubrik/Aktivitas",
            "File/Bukti"

        ],

        "Kolom": [

            NAME_COL,
            TYPE_COL,
            YEAR_COL,
            SEMESTER_COL,
            RUBRIC_COL,
            FILE_COL

        ]

    })

    st.dataframe(
        detected,
        use_container_width=True,
        hide_index=True
    )

    st.write(
        "Seluruh kolom dataset:"
    )

    st.write(
        list(df.columns)
    )


# ============================================================
# 12. NORMALISASI KATEGORI
# ============================================================

if TYPE_COL:

    def normalize_category(value):

        if pd.isna(value):

            return "Lainnya"

        text = str(value).strip().lower()

        # Pendidikan
        if any(
            keyword in text
            for keyword in [
                "pendidikan",
                "pengajaran",
                "ajar",
                "kuliah",
                "pembelajaran"
            ]
        ):

            return "Pendidikan"

        # Penelitian
        if any(
            keyword in text
            for keyword in [
                "penelitian",
                "research",
                "riset"
            ]
        ):

            return "Penelitian"

        # PkM
        if any(
            keyword in text
            for keyword in [
                "pkm",
                "pengabdian",
                "masyarakat"
            ]
        ):

            return "PkM"

        # Penunjang
        if any(
            keyword in text
            for keyword in [
                "penunjang",
                "organisasi",
                "kepanitiaan",
                "seminar",
                "asosiasi",
                "administrasi"
            ]
        ):

            return "Penunjang"

        return str(value).strip()


    df["kategori"] = (

        df[TYPE_COL]

        .apply(
            normalize_category
        )

    )

else:

    df["kategori"] = "Tidak teridentifikasi"


# ============================================================
# 13. SIDEBAR
# ============================================================

st.sidebar.markdown(
    "# 🎛️ FILTER"
)

st.sidebar.markdown(
    "---"
)


# ============================================================
# YEAR FILTER
# ============================================================

if YEAR_COL:

    year_values = (

        df[YEAR_COL]

        .dropna()

        .astype(str)

        .unique()

        .tolist()

    )

    year_values = sorted(
        year_values
    )

else:

    year_values = []


selected_years = st.sidebar.multiselect(

    "📅 Tahun Akademik",

    options=year_values,

    default=year_values

)


# ============================================================
# SEMESTER FILTER
# ============================================================

if SEMESTER_COL:

    semester_values = (

        df[SEMESTER_COL]

        .dropna()

        .unique()

        .tolist()

    )

    semester_values = sorted(
        semester_values,
        key=lambda x: str(x)
    )

else:

    semester_values = []


selected_semesters = st.sidebar.multiselect(

    "🎓 Semester",

    options=semester_values,

    default=semester_values

)


# ============================================================
# CATEGORY FILTER
# ============================================================

category_values = sorted(

    df["kategori"]

    .dropna()

    .astype(str)

    .unique()

    .tolist()

)


selected_categories = st.sidebar.multiselect(

    "📚 Bidang Kinerja",

    options=category_values,

    default=category_values

)


# ============================================================
# DOSEN FILTER
# ============================================================

if NAME_COL:

    lecturer_values = sorted(

        df[NAME_COL]

        .dropna()

        .astype(str)

        .unique()

        .tolist()

    )

else:

    lecturer_values = []


selected_lecturers = st.sidebar.multiselect(

    "👤 Dosen",

    options=lecturer_values

)


# ============================================================
# RESET FILTER
# ============================================================

if st.sidebar.button(
    "🔄 Reset Filter",
    use_container_width=True
):

    st.rerun()


# ============================================================
# 14. APPLY FILTER
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
# 15. KPI CALCULATION
# ============================================================

total_activity = len(
    filtered
)


if NAME_COL:

    total_lecturers = (

        filtered[NAME_COL]

        .dropna()

        .nunique()

    )

else:

    total_lecturers = 0


education_count = int(
    (
        filtered["kategori"]
        == "Pendidikan"
    ).sum()
)


research_count = int(
    (
        filtered["kategori"]
        == "Penelitian"
    ).sum()
)


pkm_count = int(
    (
        filtered["kategori"]
        == "PkM"
    ).sum()
)


support_count = int(
    (
        filtered["kategori"]
        == "Penunjang"
    ).sum()
)


# ============================================================
# 16. EXECUTIVE SUMMARY
# ============================================================

st.html("""
<div class="section-title">
    📊 Executive Performance Summary
</div>
""")


kpi_columns = st.columns(6)


kpis = [

    (
        "TOTAL AKTIVITAS",
        total_activity,
        "seluruh rekam kinerja"
    ),

    (
        "DOSEN",
        total_lecturers,
        "dosen teridentifikasi"
    ),

    (
        "PENDIDIKAN",
        education_count,
        "aktivitas pendidikan"
    ),

    (
        "PENELITIAN",
        research_count,
        "aktivitas penelitian"
    ),

    (
        "PkM",
        pkm_count,
        "aktivitas pengabdian"
    ),

    (
        "PENUNJANG",
        support_count,
        "aktivitas penunjang"
    )

]


for col, (
    title,
    value,
    note
) in zip(
    kpi_columns,
    kpis
):

    with col:

        st.html(f"""
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
        """)


# ============================================================
# 17. TREN DAN KOMPOSISI
# ============================================================

st.html("""
<div class="section-title">
    📈 Ringkasan Kinerja
</div>
""")


col1, col2 = st.columns(2)


# ============================================================
# TREND
# ============================================================

with col1:

    if YEAR_COL and not filtered.empty:

        trend = (

            filtered

            .assign(
                tahun_plot=
                filtered[YEAR_COL]
                .astype(str)
            )

            .groupby(
                [
                    "tahun_plot",
                    "kategori"
                ]
            )

            .size()

            .reset_index(
                name="jumlah"
            )

        )


        if not trend.empty:

            fig = px.line(

                trend,

                x="tahun_plot",

                y="jumlah",

                color="kategori",

                markers=True,

                title="Tren Kinerja Dosen"

            )

            fig.update_layout(

                height=420,

                xaxis_title="Tahun Akademik",

                yaxis_title="Jumlah Aktivitas",

                legend_title="Bidang",

                hovermode="x unified"

            )

            st.plotly_chart(

                fig,

                use_container_width=True

            )

    else:

        st.info(
            "Kolom tahun tidak tersedia."
        )


# ============================================================
# COMPOSITION
# ============================================================

with col2:

    composition = (

        filtered["kategori"]

        .value_counts()

        .reset_index()

    )


    composition.columns = [

        "kategori",
        "jumlah"

    ]


    if not composition.empty:

        fig = px.pie(

            composition,

            names="kategori",

            values="jumlah",

            hole=0.55,

            title="Komposisi Kinerja"

        )

        fig.update_layout(

            height=420,

            legend_title="Bidang"

        )

        st.plotly_chart(

            fig,

            use_container_width=True

        )

    else:

        st.info(
            "Tidak ada data."
        )


# ============================================================
# 18. DISTRIBUSI DOSEN
# ============================================================

st.html("""
<div class="section-title">
    👥 Kinerja per Dosen
</div>
""")


if NAME_COL and not filtered.empty:

    lecturer_perf = (

        filtered

        .groupby(
            [
                NAME_COL,
                "kategori"
            ]
        )

        .size()

        .reset_index(
            name="jumlah"
        )

    )


    pivot_lecturer = (

        lecturer_perf

        .pivot(

            index=NAME_COL,

            columns="kategori",

            values="jumlah"

        )

        .fillna(0)

    )


    pivot_lecturer["TOTAL"] = (

        pivot_lecturer

        .sum(axis=1)

    )


    pivot_lecturer = (

        pivot_lecturer

        .sort_values(
            "TOTAL",
            ascending=False
        )

    )


    st.dataframe(

        pivot_lecturer,

        use_container_width=True,

        height=400

    )

else:

    st.info(
        "Data dosen tidak tersedia."
    )


# ============================================================
# 19. TOP DOSEN & PERBANDINGAN BIDANG
# ============================================================

col1, col2 = st.columns(2)


# ============================================================
# TOP DOSEN
# ============================================================

with col1:

    if NAME_COL and not filtered.empty:

        top_dosen = (

            filtered

            .groupby(NAME_COL)

            .size()

            .reset_index(
                name="jumlah"
            )

            .sort_values(
                "jumlah",
                ascending=True
            )

            .tail(10)

        )


        if not top_dosen.empty:

            fig = px.bar(

                top_dosen,

                x="jumlah",

                y=NAME_COL,

                orientation="h",

                title=
                "10 Dosen dengan Aktivitas Terbanyak"

            )

            fig.update_layout(

                height=450,

                xaxis_title="Jumlah Aktivitas",

                yaxis_title=""

            )

            st.plotly_chart(

                fig,

                use_container_width=True

            )


# ============================================================
# PERBANDINGAN TAHUN
# ============================================================

with col2:

    if YEAR_COL and not filtered.empty:

        category_by_year = (

            filtered

            .assign(
                tahun_plot=
                filtered[YEAR_COL]
                .astype(str)
            )

            .groupby(
                [
                    "tahun_plot",
                    "kategori"
                ]
            )

            .size()

            .reset_index(
                name="jumlah"
            )

        )


        if not category_by_year.empty:

            fig = px.bar(

                category_by_year,

                x="tahun_plot",

                y="jumlah",

                color="kategori",

                barmode="group",

                title=
                "Perbandingan Bidang Kinerja per Tahun"

            )

            fig.update_layout(

                height=450,

                xaxis_title="Tahun Akademik",

                yaxis_title="Jumlah Aktivitas",

                legend_title="Bidang"

            )

            st.plotly_chart(

                fig,

                use_container_width=True

            )


# ============================================================
# 20. ANALISIS SEMESTER
# ============================================================

st.html("""
<div class="section-title">
    🎓 Analisis Semester
</div>
""")


if SEMESTER_COL and not filtered.empty:

    semester_analysis = (

        filtered

        .groupby(
            [
                SEMESTER_COL,
                "kategori"
            ]
        )

        .size()

        .reset_index(
            name="jumlah"
        )

    )


    if not semester_analysis.empty:

        semester_analysis[
            SEMESTER_COL
        ] = semester_analysis[
            SEMESTER_COL
        ].astype(str)


        fig = px.bar(

            semester_analysis,

            x=SEMESTER_COL,

            y="jumlah",

            color="kategori",

            barmode="group",

            title=
            "Distribusi Aktivitas berdasarkan Semester"

        )

        fig.update_layout(

            height=430,

            xaxis_title="Semester",

            yaxis_title="Jumlah Aktivitas",

            legend_title="Bidang"

        )

        st.plotly_chart(

            fig,

            use_container_width=True

        )

else:

    st.info(
        "Kolom semester tidak tersedia."
    )


# ============================================================
# 21. PROFIL TRIDHARMA
# ============================================================

st.html("""
<div class="section-title">
    ⚖️ Profil Kinerja Tridharma
</div>
""")


profile = (

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


profile_total = profile.sum()


if profile_total > 0:

    profile_percentage = (

        profile /

        profile_total *

        100

    )

else:

    profile_percentage = profile.copy()


profile_df = pd.DataFrame({

    "Bidang Kinerja":
        profile.index,

    "Jumlah Aktivitas":
        profile.values,

    "Persentase":
        profile_percentage
        .round(2)
        .values

})


st.dataframe(

    profile_df,

    use_container_width=True,

    hide_index=True

)


# ============================================================
# 22. INDIKATOR KINERJA
# ============================================================

st.html("""
<div class="section-title">
    🎯 Indikator Kinerja
</div>
""")


if total_lecturers > 0:

    average_activity = (

        total_activity /

        total_lecturers

    )

else:

    average_activity = 0


if total_activity > 0:

    research_share = (

        research_count /

        total_activity *

        100

    )

    pkm_share = (

        pkm_count /

        total_activity *

        100

    )

else:

    research_share = 0

    pkm_share = 0


ind1, ind2, ind3 = st.columns(3)


with ind1:

    st.metric(

        "Rata-rata Aktivitas / Dosen",

        f"{average_activity:.1f}"

    )


with ind2:

    st.metric(

        "Proporsi Penelitian",

        f"{research_share:.1f}%"

    )


with ind3:

    st.metric(

        "Proporsi PkM",

        f"{pkm_share:.1f}%"

    )


# ============================================================
# 23. RUBRIK / JENIS AKTIVITAS
# ============================================================

st.html("""
<div class="section-title">
    📚 Analisis Jenis Aktivitas
</div>
""")


if RUBRIC_COL and not filtered.empty:

    rubric_summary = (

        filtered

        .groupby(
            [
                "kategori",
                RUBRIC_COL
            ]
        )

        .size()

        .reset_index(
            name="jumlah"
        )

        .sort_values(
            "jumlah",
            ascending=False
        )

    )


    st.dataframe(

        rubric_summary,

        use_container_width=True,

        height=450,

        hide_index=True

    )

else:

    st.info(
        "Kolom rubrik/aktivitas tidak ditemukan."
    )


# ============================================================
# 24. MATRIX DOSEN × BIDANG
# ============================================================

st.html("""
<div class="section-title">
    🧩 Matriks Kinerja Dosen × Bidang
</div>
""")


if NAME_COL and not filtered.empty:

    matrix = pd.crosstab(

        filtered[NAME_COL],

        filtered["kategori"]

    )


    matrix["TOTAL"] = (

        matrix.sum(axis=1)

    )


    matrix = (

        matrix

        .sort_values(
            "TOTAL",
            ascending=False
        )

    )


    st.dataframe(

        matrix,

        use_container_width=True,

        height=450

    )

else:

    st.info(
        "Data dosen tidak tersedia."
    )


# ============================================================
# 25. EVIDENCE EXPLORER
# ============================================================

st.html("""
<div class="section-title">
    🔎 Evidence Explorer
</div>
""")


st.html("""
<div class="info-card">

    <b>Evidence-Based Performance</b>

    <br><br>

    Dashboard ini menghubungkan rekam kinerja
    dosen dengan data aktivitas yang tersedia
    dalam dataset.

    <br><br>

    Data dapat digunakan untuk menelusuri
    Pendidikan, Penelitian, PkM, dan Penunjang.

</div>
""")


# ============================================================
# EVIDENCE TABLE
# ============================================================

evidence_columns = []


if NAME_COL:

    evidence_columns.append(
        NAME_COL
    )


if YEAR_COL:

    evidence_columns.append(
        YEAR_COL
    )


if SEMESTER_COL:

    evidence_columns.append(
        SEMESTER_COL
    )


evidence_columns.append(
    "kategori"
)


if RUBRIC_COL:

    evidence_columns.append(
        RUBRIC_COL
    )


if FILE_COL:

    evidence_columns.append(
        FILE_COL
    )


evidence_df = filtered[
    evidence_columns
].copy()


# Rename untuk tampilan

rename_map = {}


if NAME_COL:

    rename_map[NAME_COL] = "Dosen"


if YEAR_COL:

    rename_map[YEAR_COL] = "Tahun Akademik"


if SEMESTER_COL:

    rename_map[SEMESTER_COL] = "Semester"


rename_map["kategori"] = "Bidang"


if RUBRIC_COL:

    rename_map[RUBRIC_COL] = "Aktivitas"


if FILE_COL:

    rename_map[FILE_COL] = "Bukti"


evidence_df = evidence_df.rename(
    columns=rename_map
)


# ============================================================
# EVIDENCE LINK
# ============================================================

if "Bukti" in evidence_df.columns:

    def format_evidence(value):

        if pd.isna(value):

            return ""

        value = str(value).strip()

        if (

            value.startswith(
                "http://"
            )

            or

            value.startswith(
                "https://"
            )

        ):

            return (
                f'<a href="{value}" '
                f'target="_blank">'
                f'📄 Lihat Bukti'
                f'</a>'
            )

        return value


    evidence_html = (

        evidence_df

        .copy()

    )


    evidence_html["Bukti"] = (

        evidence_html["Bukti"]

        .apply(
            format_evidence
        )

    )


    st.write(

        evidence_html.to_html(

            escape=False,

            index=False

        ),

        unsafe_allow_html=True

    )

else:

    st.dataframe(

        evidence_df,

        use_container_width=True,

        hide_index=True

    )


# ============================================================
# 26. DATA DETAIL
# ============================================================

st.html("""
<div class="section-title">
    🗂️ Data Detail
</div>
""")


with st.expander(
    "Tampilkan seluruh data hasil filter"
):

    st.dataframe(

        filtered,

        use_container_width=True,

        height=500

    )


# ============================================================
# 27. DOWNLOAD CSV
# ============================================================

st.html("""
<div class="section-title">
    ⬇️ Export Data
</div>
""")


csv_data = (

    filtered

    .to_csv(
        index=False
    )

    .encode("utf-8")

)


st.download_button(

    label="⬇️ Download Data CSV",

    data=csv_data,

    file_name=(
        "kinerja_dosen_bisnis_digital_filtered.csv"
    ),

    mime="text/csv",

    use_container_width=True

)


# ============================================================
# 28. FOOTER
# ============================================================

st.html("""
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
""")


# ============================================================
# END
# ============================================================
