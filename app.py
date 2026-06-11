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


def main() -> None:
    st.title("📊 Excel Report Rationalization & BAU Standardization Tool")
    st.caption("AI-assisted analysis, deduplication, and standardization of Excel reports.")

    st.sidebar.title("Navigation")
    page = st.sidebar.radio(
        "Go to",
        [
            "Home",
            "Upload & Ingest",
            "Profile Reports",
            "Workbook Analysis",
            "Guided Rationalization",
            "Schema Matching",
            "Rationalization",
            "Consolidation",
            "KPI Engine",
            "Workbook Generator",
            "Documentation",
            "Reconciliation",
        ],
    )

    if page == "Home":
        _render_home()
    elif page in ("Upload & Ingest", "Profile Reports"):
        from src.ui.pages.profile_reports import render as render_profile
        render_profile()
    elif page == "Workbook Analysis":
        from src.ui.pages.workbook_analysis import render as render_analysis
        render_analysis()
    elif page == "Guided Rationalization":
        from src.ui.pages.guided_rationalization import render as render_guided
        render_guided()
    elif page == "Workbook Generator":
        from src.ui.pages.generate_master import render as render_master
        render_master()
    else:
        st.info(f"**{page}** — coming soon. Business logic will be implemented in a subsequent sprint.")


def _render_home() -> None:
    st.markdown(
        """
        ## Welcome

        This tool helps organisations rationalise and standardise their Excel-based
        Business-As-Usual (BAU) reports by:

        | Step | Module | Description |
        |------|--------|-------------|
        | 1 | **Ingest** | Upload and parse one or more Excel workbooks |
        | 2 | **Profile** | Summarise columns, data types, completeness, and statistics |
        | 3 | **Workbook Analysis** | Classify tabs, extract KPIs, detect schema collisions — download `workbook_analysis.xlsx` |
        | 4 | **Schema Matching** | Detect duplicate or semantically equivalent columns across reports |
        | 5 | **Rationalize** | Flag redundant reports and propose a consolidated inventory |
        | 6 | **Consolidate** | Merge aligned reports into a single master dataset |
        | 7 | **KPI Engine** | Define and compute standardised KPIs across consolidated data |
        | 8 | **Workbook Generator** | Produce a clean, formatted Excel output workbook |
        | 9 | **Documentation** | Auto-generate data dictionaries and lineage metadata |
        | 10 | **Reconcile** | Validate output totals against source reports |

        ---
        ### Getting Started
        1. Navigate to **Upload & Ingest** in the sidebar.
        2. Upload your Excel files (`.xlsx` / `.xls`).
        3. Follow the guided workflow from left to right.
        """
    )

    st.subheader("Sample Data")
    st.write("Example files are available in the `sample_data/` directory.")


if __name__ == "__main__":
    main()
