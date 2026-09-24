import io
import re
import zipfile
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET

import streamlit as st
from PIL import Image

st.set_page_config(
    page_title="PPT Image + Markup Extractor",
    page_icon="🖼️",
    layout="centered",
)

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

EMU_PER_INCH = 914400


def local_name(tag):
    return tag.rsplit("}", 1)[-1]


def get_attr(element, name, default=None):
    return element.attrib.get(name, default)


def parse_xfrm(pic):
    """Return x, y, width, height from a PowerPoint picture."""
    sppr = pic.find("p:spPr", NS)
    if sppr is None:
        return None

    xfrm = sppr.find("a:xfrm", NS)
    if xfrm is None:
        return None

    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)

    if off is None or ext is None:
        return None

    try:
        return (
            int(off.attrib["x"]),
            int(off.attrib["y"]),
            int(ext.attrib["cx"]),
            int(ext.attrib["cy"]),
        )
    except (KeyError, ValueError):
        return None


def bbox_contains(outer, inner, tolerance=0.02):
    """Whether inner's center is inside outer, with a small tolerance."""
    ox, oy, ow, oh = outer
    ix, iy, iw, ih = inner

    cx = ix + iw / 2
    cy = iy + ih / 2

    tx = ow * tolerance
    ty = oh * tolerance

    return (
        ox - tx <= cx <= ox + ow + tx
        and oy - ty <= cy <= oy + oh + ty
    )


def bbox_area(box):
    return box[2] * box[3]


def parse_relationships(xml_bytes):
    root = ET.fromstring(xml_bytes)
    result = {}

    for rel in root:
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        rel_type = rel.attrib.get("Type", "")
        if rid and target:
            result[rid] = {
                "target": target,
                "type": rel_type,
            }

    return result


def normalize_target(target):
    # Slide relationship targets normally look like ../media/image1.png.
    target = target.replace("\\", "/")
    while target.startswith("../"):
        target = target[3:]
    if not target.startswith("ppt/"):
        target = "ppt/" + target
    return target


def get_picture_info(pic):
    nv = pic.find("p:nvPicPr", NS)
    c_nv_pr = nv.find("p:cNvPr", NS) if nv is not None else None

    name = c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else ""

    blip = pic.find(".//a:blip", NS)
    rid = blip.attrib.get(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
    ) if blip is not None else None

    bbox = parse_xfrm(pic)

    return {
        "name": name,
        "rid": rid,
        "bbox": bbox,
        "pic": pic,
    }


def image_from_zip(zf, target):
    data = zf.read(target)
    return Image.open(io.BytesIO(data)).convert("RGBA")


def safe_filename(text):
    text = re.sub(r"[^\w\-. ]+", "_", text, flags=re.UNICODE)
    text = re.sub(r"\s+", "_", text).strip("._ ")
    return text or "image"


def overlay_markup_on_image(base_img, markup_img, base_box, markup_box):
    """
    Place the markup image on the extracted base image using the same
    relative position/size that the objects have on the PowerPoint slide.

    The markup fallback image has transparent pixels, so alpha compositing
    keeps the original photograph intact.
    """
    bx, by, bw, bh = base_box
    mx, my, mw, mh = markup_box

    if bw <= 0 or bh <= 0 or mw <= 0 or mh <= 0:
        return base_img

    # Relative geometry on the base picture.
    rel_x = (mx - bx) / bw
    rel_y = (my - by) / bh
    rel_w = mw / bw
    rel_h = mh / bh

    # Convert relative geometry to actual pixels.
    px = round(rel_x * base_img.width)
    py = round(rel_y * base_img.height)
    pw = max(1, round(rel_w * base_img.width))
    ph = max(1, round(rel_h * base_img.height))

    # The fallback ink image is intended to fill its PowerPoint bbox.
    markup = markup_img.resize((pw, ph), Image.Resampling.LANCZOS)

    # Crop if the overlay extends slightly outside the photo.
    canvas = Image.new("RGBA", base_img.size, (0, 0, 0, 0))
    canvas.alpha_composite(markup, (px, py))

    return Image.alpha_composite(base_img, canvas)



def extract_text_shapes(root):
    """Return visible text boxes with their PowerPoint position/size."""
    items = []
    for sp in root.findall(".//p:sp", NS):
        texts = [t.text.strip() for t in sp.findall(".//a:t", NS) if t.text and t.text.strip()]
        if not texts:
            continue

        xfrm = sp.find(".//a:xfrm", NS)
        off = xfrm.find("a:off", NS) if xfrm is not None else None
        ext = xfrm.find("a:ext", NS) if xfrm is not None else None
        if off is None or ext is None:
            continue

        try:
            x = int(off.attrib["x"])
            y = int(off.attrib["y"])
            w = int(ext.attrib["cx"])
            h = int(ext.attrib["cy"])
        except (KeyError, ValueError):
            continue

        items.append({
            "text": " ".join(texts),
            "x": x,
            "y": y,
            "w": w,
            "h": h,
            "cx": x + w / 2,
            "cy": y + h / 2,
        })
    return items


def extract_slide_text(root):
    """Extract visible text from a slide XML in reading order."""
    return [
        t.text.strip()
        for t in root.findall(".//a:t", NS)
        if t.text and t.text.strip()
    ]


def clean_filename_part(value):
    """Make a safe filename part while preserving useful outlet/type/size text."""
    value = str(value or "").strip()

    # Normalize multiplication/spacing forms: 10 X 2 -> 10x2.
    value = re.sub(r"\s*[xX×]\s*", "x", value)
    value = re.sub(r"\s+", "_", value)

    # Keep letters, numbers, underscore, dot and hyphen only.
    value = re.sub(r"[^A-Za-z0-9_.-]+", "_", value)
    value = re.sub(r"_+", "_", value).strip("._-")

    return value or "UNKNOWN"


def _combined_label_value(text, label, next_labels=()):
    """Extract a value from text such as 'TYPE:-NL SIZE:-10X2'."""
    pattern = re.escape(label) + r"\s*[:\-]*\s*(.*?)"
    if next_labels:
        stop = "|".join(re.escape(x) for x in next_labels)
        pattern += rf"(?=\s+(?:{stop})\s*[:\-]|$)"
    else:
        pattern += r"$"

    m = re.search(pattern, text, re.I)
    return m.group(1).strip(" -:") if m else ""


def extract_slide_metadata(root):
    """
    Extract ONLY the four fields required for output filenames:

        OUTLETNAME_MOBILE_TYPE_SIZE_01.png

    Address, city, quantity, dates, URLs and other slide text are deliberately
    ignored.

    The PPT in use has two common layouts:
      1. Labels and values are separate text boxes.
      2. Labels and values are combined in the same text box.
    """
    shapes = extract_text_shapes(root)
    texts = [s["text"] for s in shapes]
    full = " | ".join(texts)

    # ---------- MOBILE ----------
    # Prefer a 10-digit number near the Contact No area. This prevents a
    # PIN/serial number from being used as the mobile number.
    mobile = ""
    contact_shapes = [
        s for s in shapes if re.search(r"CONTACT\s*NO", s["text"], re.I)
    ]
    if contact_shapes:
        ref = contact_shapes[0]
        candidates = []
        for s in shapes:
            digits = re.sub(r"\D", "", s["text"])
            if len(digits) >= 10:
                candidates.append((abs(s["cx"] - ref["cx"]) + abs(s["cy"] - ref["cy"]), digits[-10:]))
        if candidates:
            mobile = min(candidates, key=lambda x: x[0])[1]

    if not mobile:
        m = re.search(r"CONTACT\s*NO\s*[:\-]*\s*(\+?\d[\d\s\-]{8,})", full, re.I)
        if m:
            digits = re.sub(r"\D", "", m.group(1))
            if len(digits) >= 10:
                mobile = digits[-10:]

    # ---------- TYPE ----------
    typ = ""

    # Combined format, e.g. TYPE:-NL.
    for s in shapes:
        m = re.search(r"\bTYPE\s*[:\-]+\s*([A-Za-z0-9_-]+)", s["text"], re.I)
        if m:
            typ = m.group(1).strip()
            break

    # Separate label/value format. The media type is in the bottom row and
    # sits close to the Type label horizontally.
    if not typ:
        type_labels = [s for s in shapes if re.fullmatch(r"TYPE\s*:?", s["text"], re.I)]
        if type_labels:
            ref = type_labels[0]
            candidates = []
            for s in shapes:
                value = s["text"].strip()
                if not re.fullmatch(r"[A-Za-z]{1,10}", value):
                    continue
                if re.fullmatch(r"(?:QTY|SIZE|TYPE|ADDRESS|CITY|CONTACT|NO|VIEW)", value, re.I):
                    continue
                # Same bottom band as Type label and close in X.
                if s["cy"] >= ref["cy"] - 1_000_000:
                    candidates.append((abs(s["cx"] - ref["cx"]) + abs(s["cy"] - ref["cy"]) * 2, value))
            if candidates:
                typ = min(candidates, key=lambda x: x[0])[1]

    # ---------- SIZE ----------
    size = ""

    # Combined format, e.g. SIZE:-9X1.5.
    for s in shapes:
        m = re.search(r"\bSIZE\s*[:\-]*\s*(\d+(?:\.\d+)?\s*[xX×]\s*\d+(?:\.\d+)?)", s["text"], re.I)
        if m:
            size = m.group(1)
            break

    # Separate format: the dimension is represented by three text boxes,
    # e.g. '10' + 'x' + '2'. Use the Size label's horizontal area only.
    if not size:
        size_labels = [s for s in shapes if re.fullmatch(r"SIZE\s*:?", s["text"], re.I)]
        if size_labels:
            ref = size_labels[0]
            # Search numeric/x tokens in the same bottom band and around the
            # Size label. Sort left-to-right, then build the dimension string.
            tokens = []
            for s in shapes:
                value = s["text"].strip()
                if not re.fullmatch(r"(?:\d+(?:\.\d+)?|[xX×])", value):
                    continue
                if abs(s["cy"] - ref["cy"]) > 1_000_000:
                    continue
                if s["cx"] < ref["x"] - 500_000 or s["cx"] > ref["x"] + ref["w"] + 500_000:
                    continue
                tokens.append(s)

            tokens.sort(key=lambda s: s["x"])
            raw = "".join(s["text"] for s in tokens)
            m = re.search(r"(\d+(?:\.\d+)?)[xX×](\d+(?:\.\d+)?)", raw)
            if m:
                size = f"{m.group(1)}x{m.group(2)}"

    # ---------- OUTLET NAME ----------
    outlet = ""

    # Combined format: explicitly stop at Address so the address can NEVER
    # become part of the outlet name.
    for s in shapes:
        if re.search(r"OUTLET\s*NAME", s["text"], re.I):
            m = re.search(
                r"OUTLET\s*NAME\s*[:\-]\s*(.*?)\s+(?=ADDRESS\s*[:\-]|CITY\s*[:\-]|CONTACT\s*NO\s*[:\-]|INSTALLATION\s*DATE)",
                s["text"],
                re.I,
            )
            if m and m.group(1).strip():
                outlet = m.group(1).strip()
                break

    # Separate format: choose the text box nearest to the Outlet Name label,
    # but only from the top header area and only alphabetic text. This avoids
    # selecting Address, City or the contact number.
    if not outlet:
        label_shapes = [
            s for s in shapes if re.search(r"OUTLET\s*NAME", s["text"], re.I)
        ]
        if label_shapes:
            ref = label_shapes[0]
            candidates = []
            for s in shapes:
                value = s["text"].strip()
                if not value or re.search(r"(?:ADDRESS|CITY|CONTACT|INSTALLATION|OUTLET\s*NAME)", value, re.I):
                    continue
                if re.search(r"\d", value):
                    continue
                if s["y"] > 800_000:  # keep only the outlet-name line; ignore city/footer
                    continue
                # Outlet name is on the left side of the header in this layout.
                if s["x"] > 9_000_000:
                    continue
                score = abs(s["cx"] - ref["cx"]) + abs(s["cy"] - ref["cy"])
                candidates.append((score, value))
            if candidates:
                outlet = min(candidates, key=lambda x: x[0])[1]

    return {
        "outlet": clean_filename_part(outlet),
        "mobile": clean_filename_part(mobile),
        "type": clean_filename_part(typ).upper(),
        "size": clean_filename_part(size),
    }


def make_output_filename(metadata, image_number):
    """Required naming: OUTLETNAME_MOBILE_TYPE_SIZE_01.png"""
    return (
        f"{metadata['outlet']}_"
        f"{metadata['mobile']}_"
        f"{metadata['type']}_"
        f"{metadata['size']}_"
        f"{image_number:02d}.png"
    )


def process_pptx(uploaded_file, image_selection="Both", progress_callback=None):
    """
    Extract every photo from every slide and merge any PowerPoint Ink
    fallback image whose center lies over that photo.

    Returns:
        results: list of dicts with slide number, image number, name, bytes
        stats: dict
    """
    ppt_bytes = uploaded_file.getvalue()

    results = []
    total_slides = 0
    markup_count = 0
    image_count = 0
    merged_count = 0
    used_filenames = set()

    with zipfile.ZipFile(io.BytesIO(ppt_bytes), "r") as zf:
        slide_names = [
            n for n in zf.namelist()
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)
        ]
        slide_names.sort(key=lambda x: int(re.search(r"slide(\d+)", x).group(1)))
        total_slides = len(slide_names)

        for slide_index, slide_path in enumerate(slide_names, start=1):
            root = ET.fromstring(zf.read(slide_path))
            slide_metadata = extract_slide_metadata(root)

            rel_path = (
                "ppt/slides/_rels/"
                + Path(slide_path).name
                + ".rels"
            )

            if rel_path not in zf.namelist():
                continue

            rels = parse_relationships(zf.read(rel_path))

            pics = [
                get_picture_info(pic)
                for pic in root.findall(".//p:pic", NS)
            ]

            # Candidate normal pictures. Ink fallback pictures are excluded.
            normal_pics = [
                p for p in pics
                if p["bbox"] is not None
                and p["rid"] is not None
                and not p["name"].lower().startswith("ink")
            ]

            ink_pics = [
                p for p in pics
                if p["bbox"] is not None
                and p["rid"] is not None
                and p["name"].lower().startswith("ink")
            ]

            markup_count += len(ink_pics)

            # Build a list of markups that belong to each base picture.
            assigned = {id(p): [] for p in normal_pics}

            for ink in ink_pics:
                ink_box = ink["bbox"]

                # Find the smallest normal picture containing the ink center.
                candidates = [
                    p for p in normal_pics
                    if bbox_contains(p["bbox"], ink_box)
                ]

                if candidates:
                    owner = min(candidates, key=lambda p: bbox_area(p["bbox"]))
                    assigned[id(owner)].append(ink)

            # Sort photos from left to right so the first photo is the left image
            # and the second photo is the right image.
            normal_pics.sort(key=lambda p: (p["bbox"][0], p["bbox"][1]))

            # Select which photo(s) to export.
            if image_selection == "Left image":
                selected_pics = normal_pics[:1]
            elif image_selection == "Right image":
                selected_pics = normal_pics[-1:]
            else:
                selected_pics = normal_pics

            output_image_no = 0

            for base in selected_pics:
                rid_info = rels.get(base["rid"])
                if not rid_info:
                    continue

                target = normalize_target(rid_info["target"])

                if target not in zf.namelist():
                    continue

                # Only process actual image relationships.
                if "image" not in rid_info["type"].lower():
                    continue

                try:
                    base_img = image_from_zip(zf, target)
                except Exception:
                    continue

                image_count += 1
                output_image_no += 1

                # Merge all ink objects assigned to this picture.
                output_img = base_img
                assigned_inks = assigned.get(id(base), [])

                for ink in assigned_inks:
                    ink_rel = rels.get(ink["rid"])
                    if not ink_rel:
                        continue

                    ink_target = normalize_target(ink_rel["target"])
                    if ink_target not in zf.namelist():
                        continue

                    try:
                        markup_img = image_from_zip(zf, ink_target)
                        output_img = overlay_markup_on_image(
                            output_img,
                            markup_img,
                            base["bbox"],
                            ink["bbox"],
                        )
                        merged_count += 1
                    except Exception:
                        # Keep the original photo if a particular markup
                        # image cannot be decoded.
                        pass

                buffer = io.BytesIO()
                output_img.convert("RGB").save(
                    buffer,
                    format="PNG",
                    optimize=True,
                )

                # Keep the requested naming format while avoiding duplicate
                # ZIP filenames when two slides contain identical metadata.
                final_image_no = output_image_no
                filename = make_output_filename(
                    slide_metadata,
                    final_image_no,
                )
                while filename in used_filenames:
                    final_image_no += 1
                    filename = make_output_filename(
                        slide_metadata,
                        final_image_no,
                    )
                used_filenames.add(filename)

                results.append({
                    "slide": slide_index,
                    "image": final_image_no,
                    "filename": filename,
                    "data": buffer.getvalue(),
                    "markup_count": len(assigned_inks),
                    "outlet": slide_metadata["outlet"],
                    "mobile": slide_metadata["mobile"],
                    "type": slide_metadata["type"],
                    "size": slide_metadata["size"],
                })

            if progress_callback:
                progress_callback(slide_index / max(total_slides, 1))

    return results, {
        "slides": total_slides,
        "images": image_count,
        "markups": markup_count,
        "merged": merged_count,
        "outputs": len(results),
    }


def make_zip(results):
    out = io.BytesIO()

    with zipfile.ZipFile(
        out,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=6,
    ) as z:
        for item in results:
            z.writestr(item["filename"], item["data"])

    return out.getvalue()


st.title("🖼️ PPT Image + Markup Extractor")
st.write(
    "Upload a PowerPoint file to extract photos as separate PNG images "
    "with any PowerPoint markup merged onto the corresponding photo."
)

st.info(
    "This tool detects PowerPoint Ink/markup fallback images and "
    "merges them onto the corresponding photo at the correct position."
)

uploaded_file = st.file_uploader(
    "PowerPoint File (.pptx)",
    type=["pptx"],
)

if uploaded_file:
    st.success(f"Selected: {uploaded_file.name}")

    image_selection = st.radio(
        "Which image do you want to extract?",
        options=["Left image", "Right image", "Both"],
        index=2,
        horizontal=True,
    )

    if st.button("🚀 Process PPT", type="primary", use_container_width=True):
        progress = st.progress(0)
        status = st.empty()

        def update_progress(value):
            progress.progress(min(max(value, 0.0), 1.0))
            status.text(
                f"Processing slides... {int(value * 100)}%"
            )

        try:
            with st.spinner("Processing PowerPoint file..."):
                results, stats = process_pptx(
                    uploaded_file,
                    image_selection=image_selection,
                    progress_callback=update_progress,
                )

            progress.progress(1.0)
            status.empty()

            if not results:
                st.error(
                    "No images were found. Please make sure the uploaded "
                    "PPTX file is valid and contains images."
                )
            else:
                col1, col2, col3 = st.columns(3)
                col1.metric("Slides", stats["slides"])
                col2.metric("Images", stats["images"])
                col3.metric("Markup merged", stats["merged"])

                st.success(
                    f"Successfully generated {len(results)} images."
                )

                zip_bytes = make_zip(results)

                st.download_button(
                    label="⬇️ Download All Images (ZIP)",
                    data=zip_bytes,
                    file_name="PPT_Images_With_Markup.zip",
                    mime="application/zip",
                    use_container_width=True,
                )

                st.subheader("Preview")

                # Show first few generated images.
                preview_items = results[:6]

                for item in preview_items:
                    st.image(
                        item["data"],
                        caption=(
                            f"{item['filename']} — "
                            f"Markup objects: {item['markup_count']}"
                        ),
                        use_container_width=True,
                    )

                if len(results) > 6:
                    st.caption(
                        f"Showing the first 6 images in the preview. "
                        f"The ZIP file contains all {len(results)} images."
                    )

        except zipfile.BadZipFile:
            st.error("The uploaded file does not appear to be a valid PPTX file.")
        except Exception as e:
            st.error(f"Processing error: {e}")
            st.exception(e)
