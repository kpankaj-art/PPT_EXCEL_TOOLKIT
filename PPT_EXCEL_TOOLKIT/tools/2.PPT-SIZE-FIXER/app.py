import os
import tempfile
import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches


# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="PPT Size Update Tool",
    page_icon="📊",
    layout="centered"
)

st.title("📊 PowerPoint Size Update Tool")
st.write(
    "Upload your Excel and PowerPoint files to automatically "
    "update the W × H size on each PowerPoint slide."
)


# =========================================================
# FILE UPLOADERS
# =========================================================
excel_file = st.file_uploader(
    "Select Excel File (.xlsx, .xls)",
    type=["xlsx", "xls"]
)

ppt_file = st.file_uploader(
    "Select PowerPoint File (.pptx)",
    type=["pptx"]
)


# =========================================================
# UPDATE SIZE FUNCTION
# =========================================================
def update_ppt_sizes(ppt_path, excel_path, output_path):

    # Read Excel
    df = pd.read_excel(excel_path)

    # Check required columns
    required_columns = ["W", "H"]

    missing_columns = [
        col for col in required_columns
        if col not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            f"Missing Excel column(s): {', '.join(missing_columns)}"
        )

    # Remove completely empty rows
    df = df.dropna(how="all").reset_index(drop=True)

    # Load PowerPoint
    prs = Presentation(ppt_path)

    # Number of usable Excel rows
    excel_rows = len(df)

    # Number of PPT slides
    ppt_slides = len(prs.slides)

    if excel_rows < ppt_slides:
        raise ValueError(
            f"Excel has only {excel_rows} data rows, "
            f"but PowerPoint has {ppt_slides} slides."
        )

    # -----------------------------------------------------
    # Helper functions
    # -----------------------------------------------------
    def clean_number(value):
        """Convert Excel number to clean text."""
        if pd.isna(value):
            return ""

        try:
            number = float(value)

            if number.is_integer():
                return str(int(number))

            return str(number).rstrip("0").rstrip(".")

        except Exception:
            return str(value).strip()

    def get_shape_text(shape):
        """Safely get shape text."""
        try:
            return shape.text.strip()
        except Exception:
            return ""

    # -----------------------------------------------------
    # Process every slide
    # -----------------------------------------------------
    for slide_index, slide in enumerate(prs.slides):

        # Excel row corresponding to PPT slide
        row = df.iloc[slide_index]

        width_value = clean_number(row["W"])
        height_value = clean_number(row["H"])

        if not width_value or not height_value:
            continue

        target_height_shape = None
        target_width_x_shape = None

        # -------------------------------------------------
        # Find size-related text boxes
        # -------------------------------------------------
        text_shapes = []

        for shape in slide.shapes:

            if not hasattr(shape, "text_frame"):
                continue

            text = get_shape_text(shape)

            if not text:
                continue

            text_shapes.append(shape)

        # -------------------------------------------------
        # First find a height-only shape.
        #
        # Example:
        #     48
        #     60
        #     96
        # -------------------------------------------------
        for shape in text_shapes:

            text = get_shape_text(shape)

            # Ignore width/x type shapes
            if "x" in text.lower():
                continue

            # Compare with existing numeric text
            try:
                float(text)
                target_height_shape = shape
                break
            except Exception:
                pass

        # -------------------------------------------------
        # Find existing "width x" shape.
        #
        # Examples:
        #     120x
        #     240x
        #     180x
        # -------------------------------------------------
        for shape in text_shapes:

            text = get_shape_text(shape).lower().replace(" ", "")

            if text.endswith("x"):
                target_width_x_shape = shape
                break

        # -------------------------------------------------
        # If width-x shape exists, update it
        # -------------------------------------------------
        if target_width_x_shape is not None:

            target_width_x_shape.text = f"{width_value}x"

            # Keep the existing position and size
            # so the PPT design remains unchanged.

        # -------------------------------------------------
        # Otherwise find a separate "x" shape
        # -------------------------------------------------
        else:

            x_shape = None

            for shape in text_shapes:

                text = get_shape_text(shape).lower().strip()

                if text == "x":
                    x_shape = shape
                    break

            if x_shape is not None:

                # Change the existing x box to width + x
                x_shape.text = f"{width_value}x"

                # Keep it close to the original x position.
                #
                # This gives enough space for values such as
                # 120x, 240x, 360x, etc.
                x_shape.left = Inches(2.98)
                x_shape.width = Inches(0.72)

                target_width_x_shape = x_shape

        # -------------------------------------------------
        # Update height
        # -------------------------------------------------
        if target_height_shape is not None:

            target_height_shape.text = height_value

        # -------------------------------------------------
        # If height shape wasn't found, try finding a
        # numeric shape near the size area.
        # -------------------------------------------------
        else:

            numeric_shapes = []

            for shape in text_shapes:

                text = get_shape_text(shape)

                try:
                    float(text)
                except Exception:
                    continue

                # Size information is generally located
                # toward the lower portion of the slide.
                if shape.top > prs.slide_height * 0.55:
                    numeric_shapes.append(shape)

            if numeric_shapes:

                # Use the lowest numeric text box
                target_height_shape = sorted(
                    numeric_shapes,
                    key=lambda s: s.top,
                    reverse=True
                )[0]

                target_height_shape.text = height_value

    # -----------------------------------------------------
    # Save output
    # -----------------------------------------------------
    prs.save(output_path)


# =========================================================
# PROCESS BUTTON
# =========================================================
if excel_file and ppt_file:

    st.success("Both files uploaded successfully.")

    if st.button(
        "Update PPT Sizes",
        type="primary",
        use_container_width=True
    ):

        try:

            with st.spinner("Updating PowerPoint sizes..."):

                # Temporary directory
                with tempfile.TemporaryDirectory() as temp_dir:

                    excel_path = os.path.join(
                        temp_dir,
                        "input.xlsx"
                    )

                    ppt_path = os.path.join(
                        temp_dir,
                        "input.pptx"
                    )

                    output_path = os.path.join(
                        temp_dir,
                        "PPT_SIZE_UPDATED.pptx"
                    )

                    # Save uploaded Excel
                    with open(excel_path, "wb") as f:
                        f.write(excel_file.getbuffer())

                    # Save uploaded PPT
                    with open(ppt_path, "wb") as f:
                        f.write(ppt_file.getbuffer())

                    # Update PPT
                    update_ppt_sizes(
                        ppt_path,
                        excel_path,
                        output_path
                    )

                    # Read final PPT
                    with open(output_path, "rb") as f:
                        output_data = f.read()

            st.success(
                "✅ PowerPoint size update completed successfully!"
            )

            st.download_button(
                label="⬇️ Download Updated PowerPoint",
                data=output_data,
                file_name="PPT_SIZE_UPDATED.pptx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "presentationml.presentation"
                ),
                use_container_width=True
            )

        except Exception as e:

            st.error(
                f"❌ Error while processing the files:\n\n{str(e)}"
            )

else:

    st.info(
        "Please upload both the Excel file and PowerPoint file."
    )
