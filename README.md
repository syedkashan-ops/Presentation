# PSO Outlet Improvement Report – V5

A mobile-friendly Streamlit app for collecting outlet improvement Before/After photos and generating a professional PowerPoint.

## V5 changes
- Automatic photo optimization after upload: EXIF orientation correction, resize to max 1600 px and JPEG quality 82.
- Shows the original and optimized file size after each upload.
- Draft heading, Before Description and After Description are explicitly persisted in session state.
- Final validation reads the current Streamlit widget values directly, preventing false "complete heading/description" errors caused by stale widget state.
- Clear validation messages identify exactly which field is missing.
- PowerPoint generation is wrapped with a visible error message instead of the generic Streamlit failure.
- Large phone photos are converted to compact JPEGs before being retained in the session and inserted into the PPT.

## Deploy
Upload `streamlit_app.py`, `requirements.txt`, and the entire `assets` folder to GitHub. Streamlit Cloud should use `streamlit_app.py` as the main file.
