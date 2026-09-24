# ============================================================
# DASHBOARD KINERJA DOSEN
# PROGRAM STUDI BISNIS DIGITAL FEB UNM
# ============================================================

import os
import pandas as pd
import streamlit as st

# Plotly dibuat optional
try:
    import plotly.express as px
    PLOTLY_AVAILABLE = True
except Exception:
    PLOTLY_AVAILABLE = False


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Kinerja Dosen | Bisnis Digital FEB UNM",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# FILE CONFIG
# ============================================================

EXCEL_FILE = "DATA KINERJA DOSEN BISDIG 2022-2025_OK.xlsx"
LOGO_FILE = "logobd.png"


# ============================================================
# GLOBAL CSS
# ============================================================

st.markdown(
    """
    <style>

    /* =========================
       GLOBAL
    ========================= */

    .stApp {
        background-color: #f5f7fb;
    }

    .block-container {
        max-width: 1500px;
        padding-top: 1.5rem;
        padding-bottom: 3rem;
    }


    /* =========================
       SIDEBAR
    ========================= */

    section[data-testid="stSidebar"] {
        background: #ffffff;
        border-right: 1px solid #e2e8f0;
    }

    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        color: #123b68;
    }


    /* =========================
       HEADER
    ========================= */

    .main-header {
        background: linear-gradient(
            135deg,
            #ffffff 0%,
            #eef6ff 100%
        );

        border: 1px solid #d8e5f2;

        border-radius: 20px;

        padding: 28px 32px;

        box-shadow:
            0 8px 25px rgba(15, 23, 42, 0.06);

        margin-bottom: 25px;
    }

    .header-title {
        font-size: 31px;
        font-weight: 800;
        color: #123b68;
        line-height: 1.2;
    }

    .header-subtitle {
        color: #64748b;
        font-size: 14px;
        margin-top: 8px;
        line-height: 1.6;
    }

    .header-badge {
        display: inline-block;

        margin-top: 12px;

        padding: 6px 12px;

        background: #e8f2ff;

        color: #145a94;

        border-radius: 20px;

        font-size: 12px;

        font-weight: 700;
    }


    /* =========================
       SECTION
    ========================= */

    .section-title {
        font-size: 21px;
        font-weight: 800;

        color: #172033;

        margin-top: 25px;
        margin-bottom: 15px;
    }

    .section-description {
        color: #64748b;
        font-size: 13px;
        margin-bottom: 15px;
    }


    /* =========================
       KPI
    ========================= */

    div[data-testid="stMetric"] {

        background: #ffffff;

        border: 1px solid #e2e8f0;

        border-radius: 16px;

        padding: 18px;

        box-shadow:
            0 4px 15px rgba(15, 23, 42, 0.05);

        min-height: 120px;
    }

    div[data-testid="stMetricLabel"] {
        color: #64748b !important;
        font-size: 12px !important;
        font-weight: 700 !important;
    }

    div[data-testid="stMetricValue"] {
        color: #123b68 !important;
        font-weight: 800 !important;
    }


    /* =========================
       CARDS
    ========================= */

    .info-card {

        background: #ffffff;

        border: 1px solid #e2e8f0;

        border-radius: 16px;

        padding: 20px;

        box-shadow:
            0 4px 15px rgba(15, 23, 42, 0.04);

        margin-bottom: 15px;
    }

    .card-title {
        color: #123b68;
        font-size: 15px;
        font-weight: 800;
        margin-bottom: 6px;
    }

    .card-text {
        color: #64748b;
        font-size: 13px;
        line-height: 1.6;
    }


    /* =========================
       TABS
    ========================= */

    button[data-baseweb="tab"] {
        font-weight: 700;
    }


    /* =========================
       DATAFRAME
    ========================= */

    div[data-testid="stDataFrame"] {
        border-radius: 12px;
        overflow: hidden;
    }


    /* =========================
       FOOTER
    ========================= */

    .footer {

        margin-top: 45px;

        padding-top: 20px;

        border-top: 1px solid #e2e8f0;

        text-align: center;

        color: #64748b;

        font-size: 12px;
    }


    /* =========================
       MOBILE
    ========================= */

    @media (max-width: 768px) {

        .header-title {
            font-size: 23px;
        }

        .main-header {
            padding: 20px;
        }

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
            f"File tidak ditemukan: {EXCEL_FILE}"
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

    st.error("❌ Dataset gagal dibaca.")

    st.exception(
        Exception(error)
    )

    st.stop()


if df is None or df.empty:

    st.error(
        "Dataset kosong."
    )

    st.stop()


# ============================================================
# NORMALIZE COLUMN
# ============================================================

def normalize_column(column):

    text = str(column).strip().lower()

    replacements = {
        " ": "_",
        "-": "_",
        "/": "_",
        ".": "_"
    }

    for old, new in replacements.items():

        text = text.replace(
            old,
            new
        )

    return text


df.columns = [
    normalize_column(c)
    for c in df.columns
]


# ============================================================
# FIND COLUMN
# ============================================================

def find_column(candidates):

    columns = list(df.columns)

    # Exact
    for candidate in candidates:

        candidate = normalize_column(
            candidate
        )

        if candidate in columns:

            return candidate

    # Partial
    for candidate in candidates:

        candidate = normalize_column(
            candidate
        )

        for col in columns:

            if candidate in col:

                return col

    return None


NAME_COL = find_column([
    "nama",
    "nama_dosen",
    "dosen",
    "nama dosen"
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


TYPE_COL = find_column([
    "jenis",
    "kategori",
    "bidang",
    "jenis_kegiatan"
])


ACTIVITY_COL = find_column([
    "aktivitas",
    "kegiatan",
    "uraian",
    "rubrik",
    "jenis_aktivitas"
])


EVIDENCE_COL = find_column([
    "bukti",
    "file",
    "dokumen",
    "link",
    "url",
    "evidence"
])


# ============================================================
# CATEGORY
# ============================================================

def classify_category(value):

    if pd.isna(value):

        return "Lainnya"

    text = str(value).lower()

    if any(
        x in text
        for x in [
            "pendidikan",
            "pengajaran",
            "pembelajaran",
            "kuliah",
            "ajar"
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
            "asosiasi",
            "administrasi",
            "seminar"
        ]
    ):

        return "Penunjang"

    return str(value).strip()


if TYPE_COL:

    df["kategori"] = (
        df[TYPE_COL]
        .apply(classify_category)
    )

else:

    df["kategori"] = "Lainnya"


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

        st.markdown(
            "### 📊"
        )


with title_col:
st.title("📊 Dashboard Kinerja Dosen")
st.subheader("Program Studi Bisnis Digital FEB UNM")

st.write(
    "Evidence-Based Performance Dashboard untuk pemantauan "
    "kinerja Pendidikan, Penelitian, Pengabdian kepada Masyarakat, "
    "dan Penunjang."
)

st.caption(
    "DATA KINERJA DOSEN 2022–2025"
)

st.divider()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown(
        "## 🎛️ Dashboard Control"
    )

    st.caption(
        "Gunakan filter untuk melakukan "
        "drill-down terhadap data kinerja."
    )

    st.divider()


    # Tahun

    if YEAR_COL:

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
        "📅 Tahun Akademik",
        options=years,
        default=years
    )


    # Semester

    if SEMESTER_COL:

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


    # Category

    categories = sorted(
        df["kategori"]
        .dropna()
        .unique()
        .tolist()
    )


    selected_categories = st.multiselect(
        "📚 Bidang Kinerja",
        options=categories,
        default=categories
    )


    # Lecturer

    if NAME_COL:

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
        f"Total record dataset: {len(df):,}"
    )


# ============================================================
# APPLY FILTER
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
        .dropna()
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
# TABS
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

    st.markdown(
        '<div class="section-title">'
        'Executive Performance Summary'
        '</div>',
        unsafe_allow_html=True
    )

    st.caption(
        "Ringkasan kinerja berdasarkan filter aktif."
    )


    c1, c2, c3, c4, c5, c6 = st.columns(6)


    with c1:

        st.metric(
            "TOTAL AKTIVITAS",
            f"{total_activity:,}"
        )


    with c2:

        st.metric(
            "DOSEN",
            f"{total_dosen:,}"
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


    st.markdown(
        '<div class="section-title">'
        '📈 Performance Analytics'
        '</div>',
        unsafe_allow_html=True
    )


    col1, col2 = st.columns(2)


    # TREND

    with col1:

        if YEAR_COL and not filtered.empty:

            trend = pd.crosstab(
                filtered[YEAR_COL].astype(str),
                filtered["kategori"]
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
                        ["tahun", "kategori"]
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
                    color="kategori",
                    markers=True,
                    title="Tren Kinerja Tahunan"
                )

                fig.update_layout(
                    height=420,
                    template="plotly_white",
                    legend_title="Bidang",
                    xaxis_title="Tahun",
                    yaxis_title="Jumlah Aktivitas"
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


    # COMPOSITION

    with col2:

        composition = (
            filtered["kategori"]
            .value_counts()
            .rename("Jumlah")
        )


        if PLOTLY_AVAILABLE and not composition.empty:

            fig = px.pie(
                values=composition.values,
                names=composition.index,
                hole=0.58,
                title="Komposisi Kinerja"
            )

            fig.update_layout(
                height=420,
                template="plotly_white",
                legend_title="Bidang"
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


    # INDICATORS

    st.markdown(
        '<div class="section-title">'
        '🎯 Key Performance Indicators'
        '</div>',
        unsafe_allow_html=True
    )


    if total_dosen > 0:

        avg_activity = (
            total_activity /
            total_dosen
        )

    else:

        avg_activity = 0


    if total_activity > 0:

        research_share = (
            penelitian /
            total_activity *
            100
        )

        pkm_share = (
            pkm /
            total_activity *
            100
        )

    else:

        research_share = 0
        pkm_share = 0


    i1, i2, i3 = st.columns(3)


    with i1:

        st.metric(
            "Rata-rata Aktivitas / Dosen",
            f"{avg_activity:.1f}"
        )


    with i2:

        st.metric(
            "Proporsi Penelitian",
            f"{research_share:.1f}%"
        )


    with i3:

        st.metric(
            "Proporsi PkM",
            f"{pkm_share:.1f}%"
        )


# ============================================================
# TAB 2 — KINERJA DOSEN
# ============================================================

with tab_dosen:

    st.markdown(
        '<div class="section-title">'
        '👥 Kinerja Individual Dosen'
        '</div>',
        unsafe_allow_html=True
    )


    if NAME_COL and not filtered.empty:

        lecturer_matrix = pd.crosstab(
            filtered[NAME_COL],
            filtered["kategori"]
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


        st.markdown(
            '<div class="section-title">'
            '🏆 Distribusi Aktivitas Dosen'
            '</div>',
            unsafe_allow_html=True
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
                title="10 Dosen dengan Aktivitas Terbanyak"
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
            "Data dosen tidak tersedia."
        )


# ============================================================
# TAB 3 — TRIDHARMA
# ============================================================

with tab_tridharma:

    st.markdown(
        '<div class="section-title">'
        '⚖️ Profil Tridharma dan Penunjang'
        '</div>',
        unsafe_allow_html=True
    )


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

        profile_percentage = (
            profile * 0
        )


    profile_df = pd.DataFrame({

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
        profile_df,
        use_container_width=True,
        hide_index=True
    )


    st.markdown(
        '<div class="section-title">'
        '📅 Perbandingan Kinerja per Tahun'
        '</div>',
        unsafe_allow_html=True
    )


    if YEAR_COL and not filtered.empty:

        year_category = pd.crosstab(
            filtered[YEAR_COL].astype(str),
            filtered["kategori"]
        )


        if PLOTLY_AVAILABLE:

            long_year = (
                filtered
                .assign(
                    Tahun=
                    filtered[YEAR_COL]
                    .astype(str)
                )
                .groupby(
                    ["Tahun", "kategori"]
                )
                .size()
                .reset_index(
                    name="Jumlah"
                )
            )


            fig = px.bar(
                long_year,
                x="Tahun",
                y="Jumlah",
                color="kategori",
                barmode="group",
                title="Kinerja per Tahun"
            )


            fig.update_layout(
                height=450,
                template="plotly_white",
                legend_title="Bidang"
            )


            st.plotly_chart(
                fig,
                use_container_width=True
            )

        else:

            st.bar_chart(
                year_category,
                use_container_width=True
            )


    # Semester

    if SEMESTER_COL:

        st.markdown(
            '<div class="section-title">'
            '🎓 Aktivitas berdasarkan Semester'
            '</div>',
            unsafe_allow_html=True
        )


        semester = pd.crosstab(
            filtered[SEMESTER_COL].astype(str),
            filtered["kategori"]
        )


        st.bar_chart(
            semester,
            use_container_width=True
        )


# ============================================================
# TAB 4 — EVIDENCE
# ============================================================

with tab_evidence:

    st.markdown(
        '<div class="section-title">'
        '🔎 Evidence Explorer'
        '</div>',
        unsafe_allow_html=True
    )


    st.markdown(
        """
        <div class="info-card">

            <div class="card-title">
                Evidence-Based Performance
            </div>

            <div class="card-text">

                Dashboard menghubungkan rekam aktivitas
                dosen dengan data Pendidikan, Penelitian,
                Pengabdian kepada Masyarakat, dan Penunjang.

                <br><br>

                Gunakan filter di sebelah kiri untuk melakukan
                <b>drill-down</b> berdasarkan tahun, semester,
                bidang kinerja, dan dosen.

            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


    # Evidence columns

    evidence_cols = []


    if NAME_COL:

        evidence_cols.append(NAME_COL)


    if YEAR_COL:

        evidence_cols.append(YEAR_COL)


    if SEMESTER_COL:

        evidence_cols.append(SEMESTER_COL)


    evidence_cols.append(
        "kategori"
    )


    if ACTIVITY_COL:

        evidence_cols.append(
            ACTIVITY_COL
        )


    if EVIDENCE_COL:

        evidence_cols.append(
            EVIDENCE_COL
        )


    evidence = filtered[
        evidence_cols
    ].copy()


    # Rename

    rename_map = {}


    if NAME_COL:

        rename_map[NAME_COL] = "Dosen"


    if YEAR_COL:

        rename_map[YEAR_COL] = "Tahun"


    if SEMESTER_COL:

        rename_map[
            SEMESTER_COL
        ] = "Semester"


    rename_map[
        "kategori"
    ] = "Bidang"


    if ACTIVITY_COL:

        rename_map[
            ACTIVITY_COL
        ] = "Aktivitas"


    if EVIDENCE_COL:

        rename_map[
            EVIDENCE_COL
        ] = "Bukti"


    evidence = evidence.rename(
        columns=rename_map
    )


    st.dataframe(
        evidence,
        use_container_width=True,
        height=500,
        hide_index=True
    )


    # ========================================================
    # DOWNLOAD
    # ========================================================

    st.markdown(
        '<div class="section-title">'
        '⬇️ Export Evidence'
        '</div>',
        unsafe_allow_html=True
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
    "🔧 Informasi Dataset & Struktur Data"
):

    c1, c2, c3 = st.columns(3)


    with c1:

        st.metric(
            "Jumlah Record",
            f"{len(df):,}"
        )


    with c2:

        st.metric(
            "Jumlah Kolom",
            f"{len(df.columns):,}"
        )


    with c3:

        st.metric(
            "Record Terfilter",
            f"{len(filtered):,}"
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
            — Bisnis Digital FEB UNM
        </b>

        <br>

        Evidence-Based Performance Dashboard

        <br><br>

        Pendidikan · Penelitian · PkM · Penunjang

        <br><br>

        Universitas Negeri Makassar

    </div>
    """,
    unsafe_allow_html=True
)
