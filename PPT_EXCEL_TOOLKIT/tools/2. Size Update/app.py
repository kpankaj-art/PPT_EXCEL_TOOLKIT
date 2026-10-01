import streamlit as st
import pandas as pd
import re
import io

from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="PPT SIZE FIXER V9",
    page_icon="📐",
    layout="wide"
)


# =========================================================
# TITLE
# =========================================================

st.title("📐 PPT SIZE FIXER")

st.caption(
    "Excel ke Width/Height columns automatically detect karke "
    "PPT ke existing Size field ko update karega."
)


# =========================================================
# SESSION STATE
# =========================================================

if "fixed_ppt_bytes" not in st.session_state:
    st.session_state.fixed_ppt_bytes = None

if "result_df" not in st.session_state:
    st.session_state.result_df = None

if "output_filename" not in st.session_state:
    st.session_state.output_filename = (
        "PPT_SIZE_FIXED_V9.pptx"
    )


# =========================================================
# WIDTH ALIASES
# =========================================================

WIDTH_ALIASES = {
    "w",
    "width",
    "wight",
    "wid",
    "widths",
    "width inch",
    "width inches",
    "width inchs",
    "width in inch",
    "width in inches",
    "width (inch)",
    "width (inches)",
    "w inch",
    "w inches",
    "w (inch)",
    "w (inches)",
    "board width",
    "size width"
}


# =========================================================
# HEIGHT ALIASES
# =========================================================

HEIGHT_ALIASES = {
    "h",
    "height",
    "hight",
    "heigt",
    "ht",
    "height inch",
    "height inches",
    "height inchs",
    "height in inch",
    "height in inches",
    "height (inch)",
    "height (inches)",
    "h inch",
    "h inches",
    "h (inch)",
    "h (inches)",
