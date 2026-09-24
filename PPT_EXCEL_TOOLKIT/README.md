# PPT & Excel Toolkit

## Add a new tool without editing the main app

You do **not** need to modify the main `app.py` when adding a new standalone Streamlit tool.

Create a new folder inside `tools/` and put the original tool's file there as `app.py`:

```text
tools/
  My New Tool/
    app.py
```

That's it. On the next Streamlit restart/deploy, the dashboard automatically finds `tools/**/app.py` and adds it to the home screen.

### Example
If your new tool is called `PPT Compressor`, make:

```text
tools/
  PPT Compressor/
    app.py
```

You do **not** need to edit the main `app.py`, import the tool, or add it to a list.

- Keep the original filename as `app.py`.
- The folder name becomes the tool name automatically.
- Add as many tool folders as you want.
- Legacy `app.py` files are executed only after that tool is selected.
- Legacy uploaders automatically receive a 2GB per-file limit.

Existing refactored tools can remain as `tools/*.py` files with a `render()` function.

## 2GB upload limit
The project uses:

```toml
[server]
maxUploadSize = 2048
maxMessageSize = 2048
```

and legacy file uploaders are automatically patched to use `max_upload_size=2048`.
