    st.download_button(
        label="⬇️ DOWNLOAD FIXED PPT",
        data=st.session_state.fixed_ppt_bytes,
        file_name=st.session_state.output_filename,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "presentationml.presentation"
        ),
        type="primary",
        use_container_width=True,
        on_click="ignore"
    )


# =========================================================
# REPORT
# =========================================================

if (
    st.session_state.result_df
    is not None
):

    result_df = (
        st.session_state.result_df
    )

    st.markdown(
        "### 📋 Processing Report"
    )

    updated = len(
        result_df[
            result_df[
                "Status"
            ].str.contains(
                "Updated",
                na=False
            )
        ]
    )

    not_found = len(
        result_df[
            result_df[
                "Status"
            ].str.contains(
                "Not Found",
                na=False
            )
        ]
    )

    missing = len(
        result_df[
            result_df[
                "Status"
            ].str.contains(
                "Missing",
                na=False
            )
        ]
    )

    c1, c2, c3 = st.columns(
        3
    )

    with c1:

        st.metric(
            "Updated",
            updated
        )

    with c2:

        st.metric(
            "Size Not Found",
            not_found
        )

    with c3:

        st.metric(
            "Width/Height Missing",
            missing
        )

    st.dataframe(
        result_df,
        use_container_width=True,
        height=500
    )


# =========================================================
# FOOTER
# =========================================================

st.markdown(
    "---"
)

st.caption(
    "PPT SIZE FIXER V9 | "
    "Excel Row 2 → Slide 1 | "
    "Excel Row 3 → Slide 2 | "
    "Automatic W/H detection | "
    "Existing Size field only"
)
