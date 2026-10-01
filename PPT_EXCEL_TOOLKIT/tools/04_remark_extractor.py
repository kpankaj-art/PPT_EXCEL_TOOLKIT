import streamlit as st
from pptx import Presentation
from .common import save_uploaded, dataframe_download, ppt_text
import pandas as pd


def run_tool(max_upload_mb=2048):
    st.subheader("1. Upload PowerPoint")
    ppt = st.file_uploader("PowerPoint File", type=["pptx"], key="remark_ppt")
    if not ppt:
        st.info("Upload PPT to continue.")
        return

    if st.button("🚀 Extract Remarks", type="primary", key="remark_start"):
        prs = Presentation(str(save_uploaded(ppt)))
        rows = []
        for no, slide in enumerate(prs.slides, start=1):
            rows.append({
                "Slide": no,
                "Client Remark": ppt_text(slide).strip()
            })

        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True)
        st.download_button(
            "⬇️ Download Remarks Excel",
            dataframe_download(df),
            "ppt_remarks.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
