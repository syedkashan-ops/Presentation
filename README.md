# PSO Outlet Improvement Report – V6

Mobile-first Streamlit application for collecting Before/After outlet improvement photos and generating a professional PowerPoint.

## V6 fixes
- Replaced Streamlit's standard `st.file_uploader` with a native mobile Gallery picker custom component.
- Photos are resized and JPEG-compressed **inside the phone browser before they are sent to Streamlit**, reducing upload size and improving reliability on Android/mobile networks.
- Target image dimension: maximum 1600 px; browser compression targets about 1.5 MB or less where practical.
- Python applies a second validation/compression pass before storing the image in session state and the PowerPoint.
- Fixed `StreamlitWidgetAlreadyInstantiatedError` by never modifying a widget-owned session-state key after the widget is instantiated. Each new improvement uses a fresh `draft_id` and fresh widget keys.
- Final heading/description validation reads the current unique widget keys, preventing the previous false validation problem.
- PowerPoint generation remains compatible with the existing PSO design and assets.

## Files
- `streamlit_app.py` – main application
- `mobile_photo_picker/index.html` – native mobile photo picker and browser-side compression
- `assets/` – PSO logo and presentation assets
- `requirements.txt` – Python dependencies

## Deployment
Upload all files/folders to the GitHub repository used by Streamlit Cloud. Keep the `mobile_photo_picker` folder at the same level as `streamlit_app.py`.
