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
        ["Home", "Agentic Rationalization"],
        key="nav_primary",
    )

    if page == "Home":
        _render_home()
    elif page == "Agentic Rationalization":
        from src.ui.pages.agentic_rationalization import render as render_agentic
        render_agentic()


def _render_home() -> None:
    st.markdown(
        """
        ## Welcome

        This tool helps organisations rationalise and standardise their Excel-based
        Business-As-Usual (BAU) reports.

        ### Agentic Rationalization

        Use **Agentic Rationalization** in the sidebar to:

        1. Upload source workbooks
        2. Review and confirm discovered source data tabs and KPI summary tabs
        3. Run AI-driven consolidation and schema analysis
        4. Review KPI alignment and file redundancy recommendations
        5. Generate output workbooks:
           - **Future_State_Workbook.xlsx** — production-ready master source data + recreated KPI tabs
           - **Rationalization_Analysis_Pack.xlsx** — diagnostics: mappings, audit, reconciliation, documentation
        """
    )
    st.subheader("Sample Data")
    st.write("Example files are available in the `sample_data/` directory.")


if __name__ == "__main__":
    main()
