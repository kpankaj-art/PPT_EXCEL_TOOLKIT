import streamlit as st
import pandas as pd
import io
import re
import os
import tempfile
from copy import deepcopy

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor


# =========================================================
# STREAMLIT CONFIG
# =========================================================

st.set_page_config(
    page_title="PPT Size Update",
    page_icon="📐",
    layout="wide"
)


# =========================================================
# CONSTANTS
# =========================================================

MAX_FILE_MB = 2048

WIDTH_ALIASES = [
    "w",
    "width",
    "size w",
    "size_w",
    "w.",
    "width.",
    "sizewidth",
    "item width",
    "product width"
]

HEIGHT_ALIASES = [
    "h",
    "height",
    "size h",
    "size_h",
    "h.",
    "height.",
    "sizeheight",
    "item height",
    "product height"
]


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def normalize_text(value):
    if value is None:
        return ""

    value = str(value).strip().lower()

    value = value.replace("\n", " ")
    value = value.replace("\r", " ")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def normalize_column_name(value):
    value = normalize_text(value)

    value = value.replace("(", "")
    value = value.replace(")", "")
    value = value.replace("-", " ")
    value = value.replace("_", " ")

    value = re.sub(r"\s+", " ", value)

    return value.strip()


def find_column(df, aliases):
    normalized = {
        normalize_column_name(col): col
        for col in df.columns
    }

    # Exact match first
    for alias in aliases:
        alias_norm = normalize_column_name(alias)

        if alias_norm in normalized:
            return normalized[alias_norm]

    # Partial match
    for col_norm, original in normalized.items():

        for alias in aliases:
            alias_norm = normalize_column_name(alias)

            if alias_norm == col_norm:
                return original

            if alias_norm in col_norm:
                return original

    return None


def clean_number(value):
    if pd.isna(value):
        return None

    text = str(value).strip()

    if not text:
        return None

    # Remove units and unnecessary characters
    text = text.replace(",", "")
    text = re.sub(
        r"(?i)(mm|cm|inch|inches|in|\"|')",
        "",
        text
    )

    match = re.search(
        r"-?\d+(?:\.\d+)?",
        text
    )

    if not match:
        return None

    try:
        number = float(match.group())

        if number.is_integer():
            return int(number)

        return number

    except Exception:
        return None


def get_shape_text(shape):
    try:
        if hasattr(shape, "text"):
            return shape.text or ""
    except Exception:
        pass

    return ""


def set_shape_text(shape, text):
    try:
        if not shape.has_text_frame:
            return False

        tf = shape.text_frame

        # Preserve approximate first paragraph formatting
        first_run = None

        if tf.paragraphs:
            p = tf.paragraphs[0]

            if p.runs:
                first_run = p.runs[0]

        tf.clear()

        paragraph = tf.paragraphs[0]
        run = paragraph.add_run()
        run.text = str(text)

        if first_run is not None:

            try:
                run.font.name = first_run.font.name
            except Exception:
                pass

            try:
                if first_run.font.size:
                    run.font.size = first_run.font.size
            except Exception:
                pass

            try:
                run.font.bold = first_run.font.bold
            except Exception:
                pass

            try:
                run.font.italic = first_run.font.italic
            except Exception:
                pass

            try:
                if first_run.font.color.type is not None:
                    run.font.color.rgb = first_run.font.color.rgb
            except Exception:
                pass

        return True

    except Exception:
        return False


def iter_shapes(shapes):
    """
    Recursively iterate normal shapes + grouped shapes.
    """

    for shape in shapes:

        # Group
        if hasattr(shape, "shapes"):

            try:
                for child in iter_shapes(shape.shapes):
                    yield child
            except Exception:
                pass

        yield shape


def get_all_text_shapes(slide):
    result = []

    for shape in iter_shapes(slide.shapes):

        try:
            if hasattr(shape, "has_text_frame") and shape.has_text_frame:
                result.append(shape)
        except Exception:
            pass

    return result


# =========================================================
# W / H DETECTION
# =========================================================

def detect_size_shapes(slide):
    """
    Find likely W and H text boxes.

    Returns:
        width_shape
        height_shape
        combined_shape
    """

    shapes = get_all_text_shapes(slide)

    width_candidates = []
    height_candidates = []
    combined_candidates = []

    for shape in shapes:

        text = normalize_text(get_shape_text(shape))

        if not text:
            continue

        # Combined W X H
        if (
            re.search(r"\bw\b.*[x×].*\bh\b", text)
            or
            re.search(r"\bwidth\b.*[x×].*\bheight\b", text)
            or
            re.search(r"\bw\s*[:=].*[x×].*h\s*[:=]", text)
        ):
            combined_candidates.append(shape)

        # Width
        if re.search(r"(^|\s)w(\s|[:=]|$)", text):
            width_candidates.append(shape)

        if "width" in text:
            width_candidates.append(shape)

        # Height
        if re.search(r"(^|\s)h(\s|[:=]|$)", text):
            height_candidates.append(shape)

        if "height" in text:
            height_candidates.append(shape)

    # Remove duplicates
    width_candidates = list(dict.fromkeys(width_candidates))
    height_candidates = list(dict.fromkeys(height_candidates))
    combined_candidates = list(dict.fromkeys(combined_candidates))

    width_shape = width_candidates[0] if width_candidates else None
    height_shape = height_candidates[0] if height_candidates else None
    combined_shape = combined_candidates[0] if combined_candidates else None

    return width_shape, height_shape, combined_shape


# =========================================================
# POSITION HELPERS
# =========================================================

def shape_center_x(shape):
    try:
        return shape.left + (shape.width / 2)
    except Exception:
        return 0


def shape_center_y(shape):
    try:
        return shape.top + (shape.height / 2)
    except Exception:
        return 0


def shape_distance(a, b):
    try:
        dx = shape_center_x(a) - shape_center_x(b)
        dy = shape_center_y(a) - shape_center_y(b)

        return abs(dx) + abs(dy)

    except Exception:
        return 999999999


def find_best_height_shape(width_shape, height_candidates):
    if not height_candidates:
        return None

    if width_shape is None:
        return height_candidates[0]

    return min(
        height_candidates,
        key=lambda s: shape_distance(width_shape, s)
    )


# =========================================================
# COMBINED W × H CREATION
# =========================================================

def create_combined_size_box(
    slide,
    width_value,
    height_value,
    reference_shape=None
):
    """
    Create a single W × H textbox.

    Example:
        120 × 48

    This prevents W and H from appearing on separate lines.
    """

    if reference_shape is not None:

        left = reference_shape.left
        top = reference_shape.top

        # Make enough room for both values
        width = max(
            reference_shape.width,
            Inches(0.9)
        )

        height = max(
            reference_shape.height,
            Inches(0.25)
        )

    else:

        # Safe fallback position
        left = Inches(0.5)
        top = Inches(0.5)
        width = Inches(1.2)
        height = Inches(0.3)

    textbox = slide.shapes.add_textbox(
        left,
        top,
        width,
        height
    )

    tf = textbox.text_frame

    tf.clear()

    tf.word_wrap = False
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0

    paragraph = tf.paragraphs[0]

    paragraph.alignment = PP_ALIGN.CENTER

    run = paragraph.add_run()

    run.text = f"{width_value} × {height_value}"

    run.font.size = Pt(11)
    run.font.bold = True

    return textbox


# =========================================================
# UPDATE EXISTING COMBINED BOX
# =========================================================

def update_combined_box(shape, width_value, height_value):

    text = f"{width_value} × {height_value}"

    try:

        tf = shape.text_frame

        # Do not allow wrapping
        tf.word_wrap = False

        tf.clear()

        paragraph = tf.paragraphs[0]
        paragraph.alignment = PP_ALIGN.CENTER

        run = paragraph.add_run()
        run.text = text

        run.font.bold = True

        try:
            run.font.size = Pt(11)
        except Exception:
            pass

        return True

    except Exception:
        return False


# =========================================================
# PROCESS ONE SLIDE
# =========================================================

def update_slide_size(
    slide,
    width_value,
    height_value
):

    width_shape, height_shape, combined_shape = detect_size_shapes(
        slide
    )

    # -----------------------------------------------------
    # CASE 1: Existing combined W × H field
    # -----------------------------------------------------

    if combined_shape is not None:

        update_combined_box(
            combined_shape,
            width_value,
            height_value
        )

        return {
            "status": "Updated",
            "action": "Updated existing W × H field",
            "width": width_value,
            "height": height_value
        }

    # -----------------------------------------------------
    # CASE 2: W + H shapes exist
    # -----------------------------------------------------

    if width_shape is not None and height_shape is not None:

        # IMPORTANT:
        # Keep W and H on SAME LINE by converting W field
        # into a single combined field.

        try:
            combined_text = f"{width_value} × {height_value}"

            set_shape_text(
                width_shape,
                combined_text
            )

            # Hide H text instead of leaving it on another line
            set_shape_text(
                height_shape,
                ""
            )

            return {
                "status": "Updated",
                "action": "Combined W + H into one line",
                "width": width_value,
                "height": height_value
            }

        except Exception:

            return {
                "status": "Error",
                "action": "Could not combine W/H",
                "width": width_value,
                "height": height_value
            }

    # -----------------------------------------------------
    # CASE 3: Only W exists
    # -----------------------------------------------------

    if width_shape is not None:

        set_shape_text(
            width_shape,
            f"{width_value} × {height_value}"
        )

        return {
            "status": "Updated",
            "action": "W field used for combined W × H",
            "width": width_value,
            "height": height_value
        }

    # -----------------------------------------------------
    # CASE 4: Only H exists
    # -----------------------------------------------------

    if height_shape is not None:

        create_combined_size_box(
            slide,
            width_value,
            height_value,
            reference_shape=height_shape
        )

        set_shape_text(
            height_shape,
            ""
        )

        return {
            "status": "Updated",
            "action": "Created combined W × H near H field",
            "width": width_value,
            "height": height_value
        }

    # -----------------------------------------------------
    # CASE 5: Nothing exists
    # -----------------------------------------------------

    create_combined_size_box(
        slide,
        width_value,
        height_value
    )

    return {
        "status": "Created",
        "action": "Created new W × H field",
        "width": width_value,
        "height": height_value
    }


# =========================================================
# PROCESS PPT
# =========================================================

def process_ppt(
    ppt_bytes,
    excel_df
):

    prs = Presentation(
        io.BytesIO(ppt_bytes)
    )

    width_col = find_column(
        excel_df,
        WIDTH_ALIASES
    )

    height_col = find_column(
        excel_df,
        HEIGHT_ALIASES
    )

    if width_col is None:
        raise ValueError(
            "Excel mein W / Width column nahi mila."
        )

    if height_col is None:
        raise ValueError(
            "Excel mein H / Height column nahi mila."
        )

    report = []

    slide_count = len(prs.slides)

    # -----------------------------------------------------
    # Slide-by-slide mapping
    # -----------------------------------------------------

    for index, slide in enumerate(prs.slides):

        slide_number = index + 1

        # Excel row
        if index >= len(excel_df):

            report.append({
                "Slide": slide_number,
                "Status": "No Excel Row",
                "Width": "",
                "Height": "",
                "Action": "Excel row not available"
            })

            continue

        row = excel_df.iloc[index]

        width_value = clean_number(
            row[width_col]
        )

        height_value = clean_number(
            row[height_col]
        )

        if width_value is None or height_value is None:

            report.append({
                "Slide": slide_number,
                "Status": "Invalid Size",
                "Width": width_value,
                "Height": height_value,
                "Action": "W/H value missing or invalid"
            })

            continue

        result = update_slide_size(
            slide,
            width_value,
            height_value
        )

        report.append({
            "Slide": slide_number,
            "Status": result["status"],
            "Width": width_value,
            "Height": height_value,
            "Action": result["action"]
        })

    # -----------------------------------------------------
    # Save
    # -----------------------------------------------------

    output = io.BytesIO()

    prs.save(output)

    output.seek(0)

    return (
        output.getvalue(),
        pd.DataFrame(report),
        width_col,
        height_col
    )


# =========================================================
# UI
# =========================================================

st.markdown(
    """
    <style>

    .main-title {
        font-size: 32px;
        font-weight: 800;
        margin-bottom: 5px;
    }

    .subtitle {
        color: #9ca3af;
        margin-bottom: 25px;
    }

    .info-box {
        padding: 18px;
        border-radius: 12px;
        background: #111827;
        border: 1px solid #334155;
        margin-bottom: 20px;
    }

    </style>
    """,
    unsafe_allow_html=True
)


st.markdown(
    '<div class="main-title">📐 PPT Size Update</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'Excel ke W/H values ko PPT mein automatically update karein.'
    '</div>',
    unsafe_allow_html=True
)


st.markdown(
    """
    <div class="info-box">
    <b>How it works</b><br>
    Excel upload → PPT upload → W/H detect → 
    W × H same line → Fixed PPT download
    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# UPLOADS
# =========================================================

excel_file = st.file_uploader(
    "📊 Excel Master File",
    type=["xlsx", "xls"],
    key="size_excel"
)

ppt_file = st.file_uploader(
    "📄 PowerPoint File",
    type=["pptx"],
    key="size_ppt"
)


# =========================================================
# FILE INFO
# =========================================================

if excel_file is not None:

    size_mb = excel_file.size / (1024 * 1024)

    if size_mb > MAX_FILE_MB:
        st.error(
            f"Excel file {size_mb:.1f} MB hai. "
            f"Maximum allowed size {MAX_FILE_MB} MB hai."
        )
        excel_file = None

    else:

        st.success(
            f"Excel loaded: {excel_file.name} "
            f"({size_mb:.1f} MB)"
        )


if ppt_file is not None:

    size_mb = ppt_file.size / (1024 * 1024)

    if size_mb > MAX_FILE_MB:

        st.error(
            f"PPT file {size_mb:.1f} MB hai. "
            f"Maximum allowed size {MAX_FILE_MB} MB hai."
        )

        ppt_file = None

    else:

        st.success(
            f"PPT loaded: {ppt_file.name} "
            f"({size_mb:.1f} MB)"
        )


# =========================================================
# PROCESS
# =========================================================

if excel_file is not None and ppt_file is not None:

    st.divider()

    if st.button(
        "🚀 UPDATE PPT SIZE",
        type="primary",
        use_container_width=True
    ):

        progress = st.progress(0)

        status_box = st.empty()

        try:

            status_box.info(
                "📊 Excel read ho rahi hai..."
            )

            excel_bytes = excel_file.getvalue()

            excel_df = pd.read_excel(
                io.BytesIO(excel_bytes)
            )

            progress.progress(20)

            status_box.info(
                "📄 PPT process ho rahi hai..."
            )

            ppt_bytes = ppt_file.getvalue()

            fixed_ppt, report_df, width_col, height_col = process_ppt(
                ppt_bytes,
                excel_df
            )

            progress.progress(100)

            st.success(
                f"✅ Processing complete — "
                f"W column: `{width_col}` | "
                f"H column: `{height_col}`"
            )

            # -------------------------------------------------
            # Statistics
            # -------------------------------------------------

            total = len(report_df)

            updated = len(
                report_df[
                    report_df["Status"].isin(
                        ["Updated", "Created"]
                    )
                ]
            )

            errors = total - updated

            c1, c2, c3 = st.columns(3)

            with c1:
                st.metric(
                    "Total Slides",
                    total
                )

            with c2:
                st.metric(
                    "Updated",
                    updated
                )

            with c3:
                st.metric(
                    "Needs Review",
                    errors
                )

            # -------------------------------------------------
            # Report
            # -------------------------------------------------

            st.subheader(
                "📋 Processing Report"
            )

            st.dataframe(
                report_df,
                use_container_width=True,
                hide_index=True
            )

            # -------------------------------------------------
            # Download PPT
            # -------------------------------------------------

            st.download_button(
                label="📥 Download Fixed PPT",
                data=fixed_ppt,
                file_name=(
                    os.path.splitext(
                        ppt_file.name
                    )[0]
                    + "_SIZE_UPDATED.pptx"
                ),
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.presentationml.presentation"
                ),
                use_container_width=True
            )

            # -------------------------------------------------
            # Download Report
            # -------------------------------------------------

            report_buffer = io.BytesIO()

            with pd.ExcelWriter(
                report_buffer,
                engine="openpyxl"
            ) as writer:

                report_df.to_excel(
                    writer,
                    index=False,
                    sheet_name="Processing Report"
                )

            report_buffer.seek(0)

            st.download_button(
                label="📊 Download Processing Report",
                data=report_buffer.getvalue(),
                file_name="PPT_SIZE_UPDATE_REPORT.xlsx",
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True
            )

        except Exception as e:

            st.error(
                f"❌ Processing Error: {str(e)}"
            )

            st.exception(e)

else:

    st.info(
        "👆 Pehle Excel Master File aur PowerPoint File upload karein."
    )
