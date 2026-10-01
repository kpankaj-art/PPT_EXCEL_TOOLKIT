import importlib
import pkgutil
import streamlit as st

st.set_page_config(
    page_title="PPT & Excel Toolkit",
    page_icon="🧰",
    layout="wide",
    initial_sidebar_state="collapsed",
)

MAX_UPLOAD_MB = 2048

st.markdown("""
<style>
.block-container {max-width: 1400px; padding-top: 2rem;}
.tool-card {
    border: 1px solid #334155;
    border-radius: 14px;
    padding: 18px;
    min-height: 150px;
    background: #0f172a;
}
.tool-title {font-size: 20px; font-weight: 700; margin-bottom: 6px;}
.tool-desc {color: #94a3b8; font-size: 14px; line-height: 1.5;}
.small-note {color:#94a3b8;font-size:13px;}
</style>
""", unsafe_allow_html=True)

TOOL_MODULES = [
    ("01_smart_matcher", "🔗 Smart PPT + Excel Matcher", "Match Excel rows with PPT slides and create a review report."),
    ("02_size_fixer", "📐 PPT Size Fixer", "Update W × H size values from Excel while keeping the PPT layout aligned."),
    ("03_sap_brand", "🏷️ SAP Code + Brand", "Transfer SAP code and brand values from Excel into matching PPT slides."),
    ("04_remark_extractor", "📝 Remark Extractor", "Extract slide remarks/text into an Excel report."),
    ("05_slide_deleter", "🗑️ Excel Based Slide Deleter", "Delete PPT slides using slide numbers supplied in Excel."),
    ("06_image_markup", "🖼️ PPT Image Markup", "Extract PPT images and create an image review package."),
]

def load_tool(module_name):
    return importlib.import_module(f"tools.{module_name}")

if "selected_tool" not in st.session_state:
    st.session_state.selected_tool = None

st.title("🧰 PPT & Excel Toolkit")
st.caption("Select what you want to do first. Upload options appear only after you select a tool.")

if st.session_state.selected_tool is None:
    st.subheader("What do you want to do?")
    cols = st.columns(2)

    for i, (module_name, title, desc) in enumerate(TOOL_MODULES):
        with cols[i % 2]:
            st.markdown(
                f'<div class="tool-card"><div class="tool-title">{title}</div>'
                f'<div class="tool-desc">{desc}</div></div>',
                unsafe_allow_html=True,
            )
            if st.button(f"Open {title}", key=f"open_{module_name}", use_container_width=True):
                st.session_state.selected_tool = module_name
                st.rerun()

    st.divider()
    st.info("Upload limit configured to 2 GB. Your hosting provider may impose its own request/file-size limit.")
else:
    selected = next(x for x in TOOL_MODULES if x[0] == st.session_state.selected_tool)
    module_name, title, desc = selected

    top_left, top_right = st.columns([1, 5])
    with top_left:
        if st.button("← Home", use_container_width=True):
            st.session_state.selected_tool = None
            st.rerun()

    st.header(title)
    st.caption(desc)

    try:
        module = load_tool(module_name)
        if not hasattr(module, "run_tool"):
            st.error(f"{module_name}.py does not expose run_tool().")
        else:
            module.run_tool(max_upload_mb=MAX_UPLOAD_MB)
    except Exception as exc:
        st.error("Tool could not be loaded.")
        st.exception(exc)
