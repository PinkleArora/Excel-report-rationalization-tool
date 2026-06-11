"""
Excel Report Rationalization & BAU Standardization Tool
Entry point — run with: streamlit run app.py
"""

import streamlit as st

st.set_page_config(
    page_title="Excel Report Rationalizer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

_ARCHIVE_PAGES = [
    "Upload & Ingest",
    "Profile Reports",
    "Workbook Analysis",
    "Schema Matching",
    "Rationalization",
    "Consolidation",
    "KPI Engine",
    "Workbook Generator",
    "Documentation",
    "Reconciliation",
]


def main() -> None:
    st.title("📊 Excel Report Rationalization & BAU Standardization Tool")
    st.caption("AI-assisted analysis, deduplication, and standardization of Excel reports.")

    st.sidebar.title("Navigation")

    # Primary navigation
    primary_choice = st.sidebar.radio(
        "Go to",
        ["Home", "Guided Rationalization"],
        key="nav_primary",
    )

    # Archive modules — collapsed by default
    archive_choice = None
    with st.sidebar.expander("Archive Modules", expanded=False):
        archive_choice = st.radio(
            "Archive",
            _ARCHIVE_PAGES,
            key="nav_archive",
            label_visibility="collapsed",
        )
        if st.button("Open module →", key="nav_archive_open"):
            st.session_state["nav_active_archive"] = archive_choice

    # Determine active page: archive open button takes precedence
    active_archive = st.session_state.get("nav_active_archive")
    if active_archive and primary_choice in ("Home", "Guided Rationalization"):
        # Reset archive selection when user clicks a primary link
        pass

    # Route
    page = primary_choice
    if st.session_state.get("nav_active_archive"):
        page = st.session_state["nav_active_archive"]

    # Primary links reset archive selection
    if primary_choice in ("Home", "Guided Rationalization"):
        if f"_prev_primary_{primary_choice}" not in st.session_state:
            st.session_state["nav_active_archive"] = None
        st.session_state[f"_prev_primary_{primary_choice}"] = True
        page = primary_choice

    if page == "Home":
        _render_home()
    elif page == "Guided Rationalization":
        from src.ui.pages.guided_rationalization import render as render_guided
        render_guided()
    elif page in ("Upload & Ingest", "Profile Reports"):
        from src.ui.pages.profile_reports import render as render_profile
        render_profile()
    elif page == "Workbook Analysis":
        from src.ui.pages.workbook_analysis import render as render_analysis
        render_analysis()
    elif page == "Workbook Generator":
        from src.ui.pages.generate_master import render as render_master
        render_master()
    else:
        st.info(f"**{page}** — coming soon.")


def _render_home() -> None:
    st.markdown(
        """
        ## Welcome

        This tool helps organisations rationalise and standardise their Excel-based
        Business-As-Usual (BAU) reports.

        ### Primary Workflow

        Use **Guided Rationalization** in the sidebar to:

        1. Upload source workbooks
        2. Configure source data tabs and KPI summary tabs
        3. Review the source column analysis
        4. Generate two output workbooks:
           - **Future_State_Workbook.xlsx** — production-ready: master source data + recreated KPI tabs
           - **Rationalization_Analysis_Pack.xlsx** — diagnostics: mappings, audit, reconciliation, documentation

        ---

        ### Archive Modules

        The individual step-by-step modules (Ingest, Profile, Schema Matching, etc.) are
        available under **Archive Modules** in the sidebar for reference.
        """
    )
    st.subheader("Sample Data")
    st.write("Example files are available in the `sample_data/` directory.")


if __name__ == "__main__":
    main()
