from __future__ import annotations
from pathlib import Path
import io
import re
import tempfile
import zipfile
import pandas as pd
from pptx import Presentation


def save_uploaded(uploaded_file, suffix=None) -> Path:
    suffix = suffix or Path(uploaded_file.name).suffix
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getbuffer())
    tmp.close()
    return Path(tmp.name)


def read_excel(uploaded_file) -> pd.DataFrame:
    data = uploaded_file.getvalue()
    return pd.read_excel(io.BytesIO(data))


def normalize(value) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()


def find_column(df, aliases):
    normalized = {normalize(c): c for c in df.columns}
    for alias in aliases:
        a = normalize(alias)
        if a in normalized:
            return normalized[a]
    for c in df.columns:
        nc = normalize(c)
        if any(normalize(alias) in nc or nc in normalize(alias) for alias in aliases):
            return c
    return None


def ppt_text(slide) -> str:
    parts = []
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.text:
            parts.append(shape.text)
    return " ".join(parts)


def all_slide_text(prs):
    return [ppt_text(slide) for slide in prs.slides]


def make_download_bytes(path: Path) -> bytes:
    return path.read_bytes()


def safe_name(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name)


def dataframe_download(df: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Report")
    return out.getvalue()
