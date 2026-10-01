# PPT & Excel Toolkit

A clean Streamlit hub for multiple PPT/Excel utilities.

## Tools
1. Smart PPT + Excel Matcher
2. PPT Size Fixer
3. SAP Code + Brand
4. Remark Extractor
5. Excel Based Slide Deleter
6. PPT Image Markup

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Adding a new tool
Create a Python file inside `tools/` that exposes:

```python
def run_tool(max_upload_mb=2048):
    ...
```

Then add one entry to `TOOL_MODULES` in `app.py`.

## Upload size
`.streamlit/config.toml` is configured for 2048 MB. Streamlit Community Cloud or another host may impose an independent platform/request limit.
