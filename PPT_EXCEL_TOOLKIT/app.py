import importlib.util
import re
import runpy
from pathlib import Path
import streamlit as st

st.set_page_config(
    page_title="PPT & Excel Toolkit",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = Path(__file__).resolve().parent
TOOLS_DIR = BASE_DIR / "tools"

BUILTIN_META = {
    "smart_matcher": ("🔄", "Smart PPT + Excel Matcher", "Excel aur PPT ko smart matching ke saath sync, reorder aur remarks update karein."),
    "size_fixer": ("📐", "PPT Size Fixer", "Excel ke Width/Height ko automatically detect karke PPT ke existing Size fields update karein."),
    "slide_deleter": ("🗑️", "PPT Slide Deleter", "Excel me delete mark ki gayi rows ke corresponding PPT slides remove karein."),
    "image_markup": ("🖼️", "PPT Image + Markup Extractor", "PPT se images aur merged markup ko extract karke organized output banayein."),
    "sap_brand": ("📊", "SAP / Brand Transfer", "Excel records ko PPT se safely match karke SAP Code aur/or Brand transfer karein."),
    "remark_extractor": ("📝", "PPT → Excel Remarks", "PPT client remarks ko smart matching ke through Excel me add karein."),
}


def clean_display_name(name: str) -> str:
    # Remove optional numeric ordering prefix such as 1., 01-, 2_, etc.
    name = re.sub(r"^\s*\d+\s*[._-]\s*", "", name)
    name = re.sub(r"[_-]+", " ", name).strip()
    name = re.sub(r"\s+", " ", name)
    return name or "New Tool"


def nice_title(name: str) -> str:
    return clean_display_name(name).title()


def order_key(path: Path):
    m = re.match(r"^\s*(\d+)\s*[._-]", path.name)
    if m:
        return (0, int(m.group(1)), path.name.lower())
    return (1, 999999, path.name.lower())


def load_module(path: Path):
    module_name = "ppt_tool_" + re.sub(r"[^a-zA-Z0-9_]", "_", str(path.relative_to(BASE_DIR)))
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Could not create module spec for {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def discover_tools():
    """Discover tools automatically.

    Preferred format for future tools:
        tools/<Tool Name>/app.py

    Existing refactored tools with render() in tools/*.py are also supported.
    """
    discovered = []
    seen_paths = set()

    if not TOOLS_DIR.exists():
        return discovered

    # 1) Preferred: every folder containing app.py becomes a tool.
    folders = [p for p in TOOLS_DIR.iterdir() if p.is_dir() and not p.name.startswith("__")]
    for folder in sorted(folders, key=order_key):
        path = folder / "app.py"
        if not path.is_file():
            continue
        seen_paths.add(path.resolve())

        meta_key = folder.name
        # Support old names that match the built-in metadata after prefix removal.
        normalized = re.sub(r"[^a-z0-9]+", "_", clean_display_name(folder.name).lower()).strip("_")
        icon, title, desc = BUILTIN_META.get(
            normalized,
            ("🧩", clean_display_name(folder.name), "Custom Streamlit app.py tool."),
        )
        discovered.append({
            "id": f"legacy:{path.relative_to(TOOLS_DIR).as_posix()}",
            "title": title,
            "icon": icon,
            "desc": desc,
            "kind": "legacy",
            "path": path,
            "sort": order_key(folder),
        })

    # 2) Backward compatibility: tools/*.py modules exposing render().
    for path in sorted(TOOLS_DIR.glob("*.py"), key=order_key):
        if path.name.startswith("_"):
            continue
        if path.resolve() in seen_paths:
            continue
        try:
            module = load_module(path)
            if not callable(getattr(module, "render", None)):
                continue
            icon, title, desc = BUILTIN_META.get(
                path.stem,
                ("🧩", nice_title(path.stem), "Custom Streamlit tool."),
            )
            discovered.append({
                "id": f"module:{path.stem}",
                "title": title,
                "icon": icon,
                "desc": desc,
                "kind": "module",
                "module": module,
                "path": path,
                "sort": order_key(path),
            })
        except Exception as exc:
            discovered.append({
                "id": f"error:{path.stem}",
                "title": nice_title(path.stem),
                "icon": "⚠️",
                "desc": "Tool load error — open it to see details.",
                "kind": "error",
                "error": exc,
                "path": path,
                "sort": order_key(path),
            })

    discovered.sort(key=lambda x: x["sort"])
    return discovered


def inject_css():
    st.markdown("""
    <style>
    [data-testid="stHeader"] {display:none;}
    #MainMenu {visibility:hidden;}
    footer {visibility:hidden;}
    .block-container {padding:1.2rem 2rem 2.5rem 2rem; max-width:1500px;}
    .hero {padding:28px 30px; border-radius:20px; background:linear-gradient(135deg,#111827 0%,#1e293b 55%,#172554 100%); border:1px solid #334155; margin-bottom:24px;}
    .hero h1 {margin:0; font-size:2.1rem; color:#f8fafc;}
    .hero p {margin:8px 0 0; color:#cbd5e1; font-size:1rem;}
    .tool-card {min-height:175px; padding:22px; border-radius:16px; border:1px solid #334155; background:#111827; margin-bottom:8px;}
    .tool-icon {font-size:30px; margin-bottom:8px;}
    .tool-title {font-size:1.08rem; font-weight:700; color:#f8fafc;}
    .tool-desc {font-size:.86rem; line-height:1.45; color:#94a3b8; margin-top:7px;}
    .section-title {font-size:1.15rem; font-weight:700; color:#e2e8f0; margin:8px 0 14px;}
    .topbar {padding:10px 14px; border:1px solid #334155; border-radius:12px; background:#0f172a; margin-bottom:18px;}
    .topbar-title {font-weight:700; color:#f8fafc; font-size:1.05rem;}
    .topbar-sub {color:#94a3b8; font-size:.8rem; margin-top:2px;}
    </style>
    """, unsafe_allow_html=True)


def show_home(tools):
    st.markdown('<div class="hero"><h1>⚡ PPT & Excel Toolkit</h1><p>Ek hi dashboard se apne saare PPT / Excel automation tools run karein.</p></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Aapko kya karwana hai?</div>', unsafe_allow_html=True)
    if not tools:
        st.warning("tools folder me koi valid tool nahi mila.")
        return
    for row_start in range(0, len(tools), 3):
        cols = st.columns(3, gap="medium")
        for col, tool in zip(cols, tools[row_start:row_start + 3]):
            with col:
                st.markdown(f'<div class="tool-card"><div class="tool-icon">{tool["icon"]}</div><div class="tool-title">{tool["title"]}</div><div class="tool-desc">{tool["desc"]}</div></div>', unsafe_allow_html=True)
                st.button(f"Open {tool['title']}", key=f"open_{tool['id']}", use_container_width=True, on_click=lambda tid=tool["id"]: st.session_state.update(selected_tool=tid))
    st.caption("New tool add karna: tools ke andar ek naya folder banao aur usme us tool ki app.py paste karo. Dashboard automatically detect karega.")


def run_legacy_app(path):
    original_set_page_config = st.set_page_config
    original_file_uploader = st.file_uploader

    def safe_set_page_config(*args, **kwargs):
        return None

    def large_file_uploader(*args, **kwargs):
        kwargs["max_upload_size"] = 2048
        return original_file_uploader(*args, **kwargs)

    st.set_page_config = safe_set_page_config
    st.file_uploader = large_file_uploader
    try:
        runpy.run_path(str(path), run_name="__main__")
    finally:
        st.set_page_config = original_set_page_config
        st.file_uploader = original_file_uploader


def show_tool(tool):
    if st.button("← Home", key="back_home"):
        st.session_state["selected_tool"] = None
        st.rerun()
    st.markdown(f'<div class="topbar"><div class="topbar-title">{tool["icon"]} {tool["title"]}</div><div class="topbar-sub">Selected task — ab sirf isi tool ke options neeche dikh rahe hain.</div></div>', unsafe_allow_html=True)
    st.divider()
    if tool["kind"] == "module":
        tool["module"].render()
    elif tool["kind"] == "legacy":
        try:
            run_legacy_app(tool["path"])
        except Exception as exc:
            st.error(f"Tool error: {exc}")
            st.exception(exc)
    else:
        st.error(f"Could not load: {tool['path'].name}")
        st.exception(tool["error"])


def main():
    inject_css()
    tools = discover_tools()
    st.session_state.setdefault("selected_tool", None)
    selected = st.session_state["selected_tool"]
    if selected is None:
        show_home(tools)
        return
    selected_tool = next((t for t in tools if t["id"] == selected), None)
    if selected_tool is None:
        st.session_state["selected_tool"] = None
        st.rerun()
    show_tool(selected_tool)


if __name__ == "__main__":
    main()
