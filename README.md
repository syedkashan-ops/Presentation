# PSO Outlet Improvement Report – Streamlit App

A browser-based Streamlit app that collects an outlet name/code and any number of Before/After improvement items, then generates a professional PowerPoint presentation.

## Files
- `streamlit_app.py` – main application
- `requirements.txt` – dependencies
- `assets/` – PSO presentation visuals extracted from the supplied template

## Streamlit Cloud
1. Create a GitHub repository.
2. Upload `streamlit_app.py`, `requirements.txt`, and the `assets` folder.
3. In Streamlit Community Cloud choose the repository and set the main file to `streamlit_app.py`.
4. Deploy.

The app requires no API key.

## Workflow
1. Enter Outlet Name and Outlet Code.
2. Enter improvement heading.
3. Upload Before photo and description.
4. Upload After photo and description.
5. Choose Save & Continue for another improvement, or Save & Generate PPT to finish.
6. Download the generated `.pptx`.
