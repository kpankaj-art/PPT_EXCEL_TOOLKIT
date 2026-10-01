import streamlit as st

from tools import smart_matcher
from tools import size_fixer
from tools import slide_deleter
from tools import image_markup
from tools import sap_brand
from tools import remark_extractor

st.set_page_config(
    page_title="PPT & Excel Toolkit",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

TOOLS = {
    "matcher": {
        "title": "Smart PPT + Excel Matcher",
        "icon": "🔄",
        "desc": "Excel aur PPT ko smart matching ke saath sync, reorder aur remarks update karein.",
        "module": smart_matcher,
    },
    "size": {
        "title": "PPT Size Fixer",
        "icon": "📐",
        "desc": "Excel ke Width/Height ko automatically detect karke PPT ke existing Size fields update karein.",
        "module": size_fixer,
    },
    "delete": {
        "title": "PPT Slide Deleter",
        "icon": "🗑️",
        "desc": "Excel me delete mark ki gayi rows ke corresponding PPT slides remove karein.",
        "module": slide_deleter,
    },
    "images": {
        "title": "PPT Image + Markup Extractor",
        "icon": "🖼️",
        "desc": "PPT se images aur merged markup ko extract karke organized output banayein.",
        "module": image_markup,
    },
    "sap": {
        "title": "SAP / Brand Transfer",
        "icon": "📊",
        "desc": "Excel records ko PPT se safely match karke SAP Code aur/or Brand transfer karein.",
        "module": sap_brand,
    },
    "remarks": {
        "title": "PPT → Excel Remarks",
        "icon": "📝",
        "desc": "PPT client remarks ko smart matching ke through Excel me add karein.",
        "module": remark_extractor,
    },
}


def inject_css():
    st.markdown(
        """
        <style>
        [data-testid="stHeader"] {display:none;}
        #MainMenu {visibility:hidden;}
        footer {visibility:hidden;}
        .block-container {padding: 1.2rem 2rem 2.5rem 2rem; max-width: 1500px;}

        .hero {
            padding: 28px 30px;
            border-radius: 20px;
            background: linear-gradient(135deg, #111827 0%, #1e293b 55%, #172554 100%);
            border: 1px solid #334155;
            margin-bottom: 24px;
        }
        .hero h1 {margin:0; font-size: 2.1rem; color:#f8fafc;}
        .hero p {margin:8px 0 0; color:#cbd5e1; font-size:1rem;}

        .tool-card {
            min-height: 175px;
            padding: 22px;
            border-radius: 16px;
            border: 1px solid #334155;
            background: #111827;
            margin-bottom: 8px;
        }
        .tool-icon {font-size: 30px; margin-bottom: 8px;}
        .tool-title {font-size: 1.08rem; font-weight:700; color:#f8fafc;}
        .tool-desc {font-size: .86rem; line-height:1.45; color:#94a3b8; margin-top:7px;}

        .section-title {font-size:1.15rem; font-weight:700; color:#e2e8f0; margin: 8px 0 14px;}
        .topbar {
            padding: 10px 14px;
            border:1px solid #334155;
            border-radius:12px;
            background:#0f172a;
            margin-bottom:18px;
        }
        .topbar-title {font-weight:700; color:#f8fafc; font-size:1.05rem;}
        .topbar-sub {color:#94a3b8; font-size:.8rem; margin-top:2px;}
        </style>
        """,
        unsafe_allow_html=True,
    )


def go_home():
    st.session_state["selected_tool"] = None


def select_tool(key):
    st.session_state["selected_tool"] = key


def show_home():
    st.markdown(
        """
        <div class="hero">
            <h1>⚡ PPT & Excel Toolkit</h1>
            <p>Ek hi dashboard se apne saare PPT / Excel automation tools run karein.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown('<div class="section-title">Aapko kya karwana hai?</div>', unsafe_allow_html=True)

    keys = list(TOOLS.keys())
    for row_start in range(0, len(keys), 3):
        cols = st.columns(3, gap="medium")
        for col, key in zip(cols, keys[row_start:row_start + 3]):
            tool = TOOLS[key]
            with col:
                st.markdown(
                    f'''<div class="tool-card"><div class="tool-icon">{tool["icon"]}</div><div class="tool-title">{tool["title"]}</div><div class="tool-desc">{tool["desc"]}</div></div>''',
                    unsafe_allow_html=True,
                )
                st.button(
                    f"Open {tool['title']}",
                    key=f"open_{key}",
                    use_container_width=True,
                    on_click=select_tool,
                    args=(key,),
                )

    st.caption("Tip: Pehle task select karein. Uske baad sirf usi tool ke required upload options dikhेंगे.")


def show_tool(key):
    tool = TOOLS[key]
    left, right = st.columns([1, 7])
    with left:
        st.button("← Home", key="back_home", use_container_width=True, on_click=go_home)
    with right:
        st.markdown(
            f'''<div class="topbar"><div class="topbar-title">{tool["icon"]} {tool["title"]}</div><div class="topbar-sub">Selected task — ab sirf isi tool ke options neeche dikh rahe hain.</div></div>''',
            unsafe_allow_html=True,
        )
    st.divider()
    tool["module"].render()


def main():
    inject_css()
    st.session_state.setdefault("selected_tool", None)
    key = st.session_state["selected_tool"]
    if key is None:
        show_home()
    elif key in TOOLS:
        show_tool(key)
    else:
        go_home()
        st.rerun()


if __name__ == "__main__":
    main()
