# ============================================================
# DASHBOARD KINERJA DOSEN
# PROGRAM STUDI BISNIS DIGITAL FEB UNM
# VERSI STABIL STREAMLIT CLOUD
# ============================================================

import os
import pandas as pd
import streamlit as st

# ============================================================
# OPTIONAL PLOTLY
# ============================================================
try:
    import plotly.express as px

    PLOTLY_AVAILABLE = True

except Exception:

    PLOTLY_AVAILABLE = False


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Dashboard Kinerja Dosen | Bisnis Digital FEB UNM",
    page_icon="❇️",
    layout="wide",
    initial_sidebar_state="expanded")

# ============================================================
# FILE CONFIGURATION
# ============================================================

EXCEL_FILE = "DATA KINERJA DOSEN BISDIG 2022-2025_OK.xlsx"
LOGO_FILE = "logobd.png"

# ============================================================
# SIMPLE CSS
# Tidak menggunakan HTML content untuk dashboard.
# CSS hanya untuk memperhalus tampilan native Streamlit.
# ============================================================

st.markdown(
    """
    <style>

    .stApp {
        background-color: #f6f8fb;
    }

    .block-container {
        max-width: 1500px;
        padding-top: 1.5rem;
        padding-bottom: 3rem;
    }

    section[data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e5e7eb;
    }

    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 15px;
        box-shadow: 0 3px 12px rgba(15, 23, 42, 0.05);
    }

    div[data-testid="stMetricLabel"] {
        font-weight: 700;
    }

    div[data-testid="stDataFrame"] {
        border-radius: 10px;
    }

    button[data-baseweb="tab"] {
        font-weight: 700;
    }

    </style>
    """,
    unsafe_allow_html=True
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

    except Exception as error:

        return None, str(error)


df, load_error = load_data()


# ============================================================
# ERROR CHECK
# ============================================================

if load_error:

    st.error("❌ Dataset gagal dibaca.")

    st.write(
        "Pastikan file berikut berada dalam folder yang sama dengan app.py:"
    )

    st.code(EXCEL_FILE)

    st.exception(
        Exception(load_error)
    )

    st.stop()


if df is None or df.empty:

    st.error(
        "Dataset kosong atau tidak dapat dibaca."
    )

    st.stop()


# ============================================================
# CLEAN COLUMN NAMES
# ============================================================

def normalize_column(column):

    text = str(column).strip().lower()

    text = text.replace(" ", "_")
    text = text.replace("-", "_")
    text = text.replace("/", "_")
    text = text.replace(".", "_")

    return text
df.columns = [
    normalize_column(column)
    for column in df.columns]
# ============================================================
# FIND COLUMN
# ============================================================
def find_column(candidates):
    # Exact matching
    for candidate in candidates:
        candidate = normalize_column(
            candidate
        )

        if candidate in df.columns:

            return candidate


    # Partial matching

    for candidate in candidates:

        candidate = normalize_column(
            candidate
        )

        for column in df.columns:

            if candidate in column:

                return column


    return None


NAME_COL = find_column(
    [
        "nama",
        "nama_dosen",
        "dosen",
        "nama dosen"
    ]
)


YEAR_COL = find_column(
    [
        "tahun",
        "tahun_akademik",
        "tahun akademik"
    ]
)


SEMESTER_COL = find_column(
    [
        "semester",
        "sem"
    ]
)


TYPE_COL = find_column(
    [
        "jenis",
        "kategori",
        "bidang",
        "jenis_kegiatan"
    ]
)


ACTIVITY_COL = find_column(
    [
        "aktivitas",
        "kegiatan",
        "uraian",
        "rubrik",
        "jenis_aktivitas"
    ]
)


EVIDENCE_COL = find_column(
    [
        "bukti",
        "file",
        "dokumen",
        "link",
        "url",
        "evidence"
    ]
)


# ============================================================
# CATEGORY CLASSIFICATION
# ============================================================

def classify_category(value):

    if pd.isna(value):

        return "Lainnya"


    text = str(value).lower()


    # Pendidikan

    if any(
        keyword in text
        for keyword in [
            "pendidikan",
            "pengajaran",
            "pembelajaran",
            "kuliah",
            "ajar"
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
            "asosiasi",
            "administrasi",
            "seminar"
        ]
    ):

        return "Penunjang"


    return str(value).strip()


if TYPE_COL is not None:

    df["kategori_dashboard"] = (
        df[TYPE_COL]
        .apply(classify_category)
    )

else:

    df["kategori_dashboard"] = "Lainnya"


# ============================================================
# HEADER
# ============================================================

logo_col, title_col = st.columns(
    [1, 7],
    vertical_alignment="center"
)


with logo_col:

    if os.path.exists(LOGO_FILE):

        st.image(
            LOGO_FILE,
            width=95
        )

    else:

        st.write("📊")


with title_col:

    st.title(
        "Dashboard Kinerja Dosen"
    )

    st.subheader(
        "Program Studi Bisnis Digital FEB UNM"
    )

    st.caption(
        "Evidence-Based Performance Dashboard | "
        "Pendidikan • Penelitian • PkM • Penunjang"
    )

    st.caption(
        "DATA KINERJA DOSEN 2022–2025"
    )


st.divider()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header(
        "🎛️ Dashboard Control"
    )

    st.caption(
        "Gunakan filter untuk melakukan "
        "drill-down terhadap data kinerja."
    )

    st.divider()


    # --------------------------------------------------------
    # YEAR
    # --------------------------------------------------------

    if YEAR_COL is not None:

        years = sorted(
            df[YEAR_COL]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

    else:

        years = []


    selected_years = st.multiselect(
        "📅 Tahun",
        options=years,
        default=years
    )


    # --------------------------------------------------------
    # SEMESTER
    # --------------------------------------------------------

    if SEMESTER_COL is not None:

        semesters = sorted(
            df[SEMESTER_COL]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

    else:

        semesters = []


    selected_semesters = st.multiselect(
        "🎓 Semester",
        options=semesters,
        default=semesters
    )


    # --------------------------------------------------------
    # CATEGORY
    # --------------------------------------------------------

    categories = sorted(
        df["kategori_dashboard"]
        .dropna()
        .unique()
        .tolist()
    )


    selected_categories = st.multiselect(
        "📚 Bidang Kinerja",
        options=categories,
        default=categories
    )


    # --------------------------------------------------------
    # DOSEN
    # --------------------------------------------------------

    if NAME_COL is not None:

        lecturers = sorted(
            df[NAME_COL]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

    else:

        lecturers = []


    selected_lecturers = st.multiselect(
        "👤 Dosen",
        options=lecturers
    )


    st.divider()


    st.caption(
        f"Total dataset: {len(df):,} record"
    )


# ============================================================
# APPLY FILTER
# ============================================================

filtered = df.copy()


# YEAR

if YEAR_COL is not None and selected_years:

    filtered = filtered[
        filtered[YEAR_COL]
        .astype(str)
        .isin(selected_years)
    ]


# SEMESTER

if (
    SEMESTER_COL is not None
    and selected_semesters
):

    filtered = filtered[
        filtered[SEMESTER_COL]
        .astype(str)
        .isin(selected_semesters)
    ]


# CATEGORY

if selected_categories:

    filtered = filtered[
        filtered["kategori_dashboard"]
        .isin(selected_categories)
    ]


# DOSEN

if (
    NAME_COL is not None
    and selected_lecturers
):

    filtered = filtered[
        filtered[NAME_COL]
        .astype(str)
        .isin(selected_lecturers)
    ]


# ============================================================
# KPI CALCULATION
# ============================================================

total_activity = len(filtered)


if NAME_COL is not None:

    total_dosen = (
        filtered[NAME_COL]
        .dropna()
        .nunique()
    )

else:

    total_dosen = 0


pendidikan = int(
    (
        filtered["kategori_dashboard"]
        == "Pendidikan"
    ).sum()
)


penelitian = int(
    (
        filtered["kategori_dashboard"]
        == "Penelitian"
    ).sum()
)


pkm = int(
    (
        filtered["kategori_dashboard"]
        == "PkM"
    ).sum()
)


penunjang = int(
    (
        filtered["kategori_dashboard"]
        == "Penunjang"
    ).sum()
)


# ============================================================
# MAIN TABS
# ============================================================

tab_overview, tab_dosen, tab_tridharma, tab_evidence = st.tabs(
    [
        "📊 Overview",
        "👥 Kinerja Dosen",
        "⚖️ Tridharma",
        "🔎 Evidence"
    ]
)


# ============================================================
# TAB 1 — OVERVIEW
# ============================================================

with tab_overview:

    st.header(
        "Executive Performance Summary"
    )

    st.caption(
        "Ringkasan kinerja berdasarkan filter yang aktif."
    )


    # --------------------------------------------------------
    # KPI
    # --------------------------------------------------------

    c1, c2, c3, c4, c5, c6 = st.columns(6)


    with c1:

        st.metric(
            "Total Aktivitas",
            f"{total_activity:,}"
        )


    with c2:

        st.metric(
            "Dosen",
            f"{total_dosen:,}"
        )


    with c3:

        st.metric(
            "Pendidikan",
            f"{pendidikan:,}"
        )


    with c4:

        st.metric(
            "Penelitian",
            f"{penelitian:,}"
        )


    with c5:

        st.metric(
            "PkM",
            f"{pkm:,}"
        )


    with c6:

        st.metric(
            "Penunjang",
            f"{penunjang:,}"
        )


    st.divider()


    # --------------------------------------------------------
    # CHARTS
    # --------------------------------------------------------

    chart_left, chart_right = st.columns(2)


    # TREND

    with chart_left:

        st.subheader(
            "📈 Tren Kinerja Dosen"
        )


        if (
            YEAR_COL is not None
            and not filtered.empty
        ):

            trend = pd.crosstab(
                filtered[YEAR_COL].astype(str),
                filtered["kategori_dashboard"]
            )


            if PLOTLY_AVAILABLE:

                trend_long = (
                    filtered
                    .assign(
                        tahun=
                        filtered[YEAR_COL]
                        .astype(str)
                    )
                    .groupby(
                        [
                            "tahun",
                            "kategori_dashboard"
                        ]
                    )
                    .size()
                    .reset_index(
                        name="Jumlah"
                    )
                )


                fig = px.line(
                    trend_long,
                    x="tahun",
                    y="Jumlah",
                    color="kategori_dashboard",
                    markers=True,
                    title="Tren Kinerja Tahunan"
                )


                fig.update_layout(
                    height=420,
                    template="plotly_white",
                    xaxis_title="Tahun",
                    yaxis_title="Jumlah Aktivitas",
                    legend_title="Bidang"
                )


                st.plotly_chart(
                    fig,
                    use_container_width=True
                )

            else:

                st.line_chart(
                    trend,
                    use_container_width=True
                )

        else:

            st.info(
                "Data tahun tidak tersedia."
            )


    # COMPOSITION

    with chart_right:

        st.subheader(
            "📊 Komposisi Kinerja"
        )


        composition = (
            filtered["kategori_dashboard"]
            .value_counts()
        )


        if (
            PLOTLY_AVAILABLE
            and not composition.empty
        ):

            fig = px.pie(
                values=composition.values,
                names=composition.index,
                hole=0.55,
                title="Komposisi Bidang Kinerja"
            )


            fig.update_layout(
                height=420,
                template="plotly_white"
            )


            st.plotly_chart(
                fig,
                use_container_width=True
            )

        else:

            st.bar_chart(
                composition,
                use_container_width=True
            )


    # --------------------------------------------------------
    # KPI TAMBAHAN
    # --------------------------------------------------------

    st.subheader(
        "🎯 Key Performance Indicators"
    )


    if total_dosen > 0:

        average_activity = (
            total_activity /
            total_dosen
        )

    else:

        average_activity = 0


    if total_activity > 0:

        research_percentage = (
            penelitian /
            total_activity *
            100
        )

        pkm_percentage = (
            pkm /
            total_activity *
            100
        )

    else:

        research_percentage = 0
        pkm_percentage = 0


    k1, k2, k3 = st.columns(3)


    with k1:

        st.metric(
            "Rata-rata Aktivitas / Dosen",
            f"{average_activity:.1f}"
        )


    with k2:

        st.metric(
            "Proporsi Penelitian",
            f"{research_percentage:.1f}%"
        )


    with k3:

        st.metric(
            "Proporsi PkM",
            f"{pkm_percentage:.1f}%"
        )


# ============================================================
# TAB 2 — KINERJA DOSEN
# ============================================================

with tab_dosen:

    st.header(
        "👥 Kinerja Individual Dosen"
    )

    st.caption(
        "Distribusi aktivitas berdasarkan dosen."
    )


    if (
        NAME_COL is not None
        and not filtered.empty
    ):

        # ----------------------------------------------------
        # MATRIX
        # ----------------------------------------------------

        lecturer_matrix = pd.crosstab(
            filtered[NAME_COL],
            filtered["kategori_dashboard"]
        ).fillna(0)


        lecturer_matrix["TOTAL"] = (
            lecturer_matrix.sum(axis=1)
        )


        lecturer_matrix = (
            lecturer_matrix
            .sort_values(
                "TOTAL",
                ascending=False
            )
        )


        st.dataframe(
            lecturer_matrix,
            use_container_width=True,
            height=450
        )


        # ----------------------------------------------------
        # TOP DOSEN
        # ----------------------------------------------------

        st.subheader(
            "🏆 10 Dosen dengan Aktivitas Terbanyak"
        )


        top_dosen = (
            filtered[NAME_COL]
            .value_counts()
            .head(10)
            .sort_values()
        )


        if PLOTLY_AVAILABLE:

            fig = px.bar(
                top_dosen,
                orientation="h",
                title="Top 10 Aktivitas Dosen"
            )


            fig.update_layout(
                height=500,
                template="plotly_white",
                xaxis_title="Jumlah Aktivitas",
                yaxis_title="Dosen"
            )


            st.plotly_chart(
                fig,
                use_container_width=True
            )

        else:

            st.bar_chart(
                top_dosen,
                horizontal=True,
                use_container_width=True
            )


    else:

        st.info(
            "Kolom nama dosen tidak tersedia "
            "atau data kosong."
        )


# ============================================================
# TAB 3 — TRIDHARMA
# ============================================================

with tab_tridharma:

    st.header(
        "⚖️ Profil Tridharma dan Penunjang"
    )


    # --------------------------------------------------------
    # PROFILE
    # --------------------------------------------------------

    profile = (
        filtered["kategori_dashboard"]
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

        profile_percentage = (
            profile * 0
        )


    profile_table = pd.DataFrame({

        "Bidang Kinerja":
            profile.index,

        "Jumlah Aktivitas":
            profile.values,

        "Persentase (%)":
            profile_percentage
            .round(2)
            .values

    })


    st.dataframe(
        profile_table,
        use_container_width=True,
        hide_index=True
    )


    # --------------------------------------------------------
    # YEAR COMPARISON
    # --------------------------------------------------------

    st.subheader(
        "📅 Perbandingan Kinerja per Tahun"
    )


    if (
        YEAR_COL is not None
        and not filtered.empty
    ):

        year_table = pd.crosstab(
            filtered[YEAR_COL].astype(str),
            filtered["kategori_dashboard"]
        )


        if PLOTLY_AVAILABLE:

            year_long = (
                filtered
                .assign(
                    Tahun=
                    filtered[YEAR_COL]
                    .astype(str)
                )
                .groupby(
                    [
                        "Tahun",
                        "kategori_dashboard"
                    ]
                )
                .size()
                .reset_index(
                    name="Jumlah"
                )
            )


            fig = px.bar(
                year_long,
                x="Tahun",
                y="Jumlah",
                color="kategori_dashboard",
                barmode="group",
                title="Kinerja per Tahun"
            )


            fig.update_layout(
                height=450,
                template="plotly_white",
                xaxis_title="Tahun",
                yaxis_title="Jumlah Aktivitas",
                legend_title="Bidang"
            )


            st.plotly_chart(
                fig,
                use_container_width=True
            )

        else:

            st.bar_chart(
                year_table,
                use_container_width=True
            )


    # --------------------------------------------------------
    # SEMESTER
    # --------------------------------------------------------

    if SEMESTER_COL is not None:

        st.subheader(
            "🎓 Aktivitas berdasarkan Semester"
        )


        semester_table = pd.crosstab(
            filtered[SEMESTER_COL].astype(str),
            filtered["kategori_dashboard"]
        )


        st.bar_chart(
            semester_table,
            use_container_width=True
        )


# ============================================================
# TAB 4 — EVIDENCE
# ============================================================

with tab_evidence:

    st.header("🔎 Evidence Explorer")

    st.info(
        """
        **Evidence-Based Performance**

        Dashboard menghubungkan rekam aktivitas dosen
        dengan data Pendidikan, Penelitian,
        Pengabdian kepada Masyarakat, dan Penunjang.

        Gunakan filter di sebelah kiri untuk melakukan
        **drill-down** berdasarkan tahun, semester,
        bidang kinerja, dan dosen.
        """
    )

    # --------------------------------------------------------
    # EVIDENCE COLUMNS
    # --------------------------------------------------------

    evidence_columns = []

    if NAME_COL is not None:
        evidence_columns.append(NAME_COL)

    if YEAR_COL is not None:
        evidence_columns.append(YEAR_COL)

    if SEMESTER_COL is not None:
        evidence_columns.append(SEMESTER_COL)

    evidence_columns.append(
        "kategori_dashboard"
    )

    if ACTIVITY_COL is not None:
        evidence_columns.append(ACTIVITY_COL)

    if EVIDENCE_COL is not None:
        evidence_columns.append(EVIDENCE_COL)


    # --------------------------------------------------------
    # CREATE EVIDENCE TABLE
    # --------------------------------------------------------

    evidence = filtered[
        evidence_columns
    ].copy()


    # --------------------------------------------------------
    # RENAME COLUMNS
    # --------------------------------------------------------

    rename_map = {}

    if NAME_COL is not None:
        rename_map[NAME_COL] = "Dosen"

    if YEAR_COL is not None:
        rename_map[YEAR_COL] = "Tahun"

    if SEMESTER_COL is not None:
        rename_map[SEMESTER_COL] = "Semester"

    rename_map[
        "kategori_dashboard"
    ] = "Bidang"

    if ACTIVITY_COL is not None:
        rename_map[ACTIVITY_COL] = "Aktivitas"

    if EVIDENCE_COL is not None:
        rename_map[EVIDENCE_COL] = "Evidence"


    evidence = evidence.rename(
        columns=rename_map
    )


    # --------------------------------------------------------
    # EVIDENCE ICON
    # --------------------------------------------------------

    if "Evidence" in evidence.columns:

        def create_evidence_link(value):

            if pd.isna(value):

                return ""

            value = str(value).strip()

            if value == "":
                return ""

            # Jika evidence sudah berupa URL
            if (
                value.startswith("http://")
                or value.startswith("https://")
            ):

                return value

            # Jika berupa path/file lokal,
            # tetap ditampilkan sebagai teks kosong
            return ""


        evidence["Download"] = (
            evidence["Evidence"]
            .apply(create_evidence_link)
        )


        # Hapus kolom Evidence asli
        evidence = evidence.drop(
            columns=["Evidence"]
        )


        # ----------------------------------------------------
        # MOVE DOWNLOAD TO LAST COLUMN
        # ----------------------------------------------------

        columns = [
            col
            for col in evidence.columns
            if col != "Download"
        ]

        columns.append("Download")

        evidence = evidence[
            columns
        ]


    # --------------------------------------------------------
    # DISPLAY TABLE
    # --------------------------------------------------------

    if "Download" in evidence.columns:

        st.dataframe(
            evidence,
            use_container_width=True,
            height=520,
            hide_index=True,

            column_config={

                "Download":
                    st.column_config.LinkColumn(
                        "Evidence",
                        help="Klik ikon untuk membuka / mengunduh evidence",
                        display_text="⬇️",
                        width="small"
                    )
            }
        )

    else:

        st.dataframe(
            evidence,
            use_container_width=True,
            height=520,
            hide_index=True
        )


    # --------------------------------------------------------
    # DOWNLOAD DATA FILTER
    # --------------------------------------------------------

    st.subheader(
        "⬇️ Export Data"
    )


    csv_data = (
        filtered
        .to_csv(index=False)
        .encode("utf-8")
    )


    st.download_button(
        label="⬇️ Download Data Hasil Filter",
        data=csv_data,
        file_name=(
            "kinerja_dosen_bisnis_digital_filtered.csv"
        ),
        mime="text/csv",
        use_container_width=True
    )


# ============================================================
# DATASET INFORMATION
# ============================================================

with st.expander(
    "🔧 Informasi Dataset"
):

    info1, info2, info3 = st.columns(3)


    with info1:

        st.metric(
            "Total Record",
            f"{len(df):,}"
        )


    with info2:

        st.metric(
            "Jumlah Kolom",
            f"{len(df.columns):,}"
        )


    with info3:

        st.metric(
            "Record Terfilter",
            f"{len(filtered):,}"
        )


    st.write(
        "Struktur kolom dataset:"
    )


    st.code(
        "\n".join(
            df.columns.tolist()
        )
    )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "Dashboard Kinerja Dosen — "
    "Program Studi Bisnis Digital FEB UNM"
)

st.caption(
    "Evidence-Based Performance Dashboard | "
    "Pendidikan • Penelitian • PkM • Penunjang"
)

st.caption(
    "Universitas Negeri Makassar"
)
