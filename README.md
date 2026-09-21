# PSO Outlet Improvement App V7

A Streamlit app for creating professional PSO outlet Before/After improvement PowerPoint reports.

## V7 reliability changes
- Uses only Streamlit's native `st.file_uploader` for Gallery uploads.
- No custom Streamlit component.
- No external JavaScript or CDN dependency.
- Shows only one photo uploader at a time.
- Saves a photo immediately after Streamlit receives it.
- Uses SHA-256 to avoid recompressing the same uploaded file on every rerun.
- Compresses photos automatically for the PowerPoint.
- Never writes to widget-owned session-state keys after widget creation.
- Uses fresh, stable widget keys for each improvement and replacement photo.
- Generates the PowerPoint automatically after `Save & Finish`.

## Deploy on Streamlit Community Cloud
1. Delete the old app files from your GitHub repository, or upload these V7 files over them.
2. Keep this exact structure:
   - `streamlit_app.py`
   - `app_core.py`
   - `requirements.txt`
   - `.streamlit/config.toml`
   - `assets/pso_logo.png`
   - `assets/cover.jpg`
   - `assets/footer.jpg`
3. In Streamlit Cloud, set the main file path to `streamlit_app.py`.
4. Reboot/redeploy the app after GitHub updates.

## Mobile use
On Android Chrome, tap Browse files and choose the image from Gallery/Photos. The actual phone picker is controlled by Android/Chrome, while the app processes the photo immediately after the browser uploads it to Streamlit.
