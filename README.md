# PSO Outlet Improvement Report App V9

## V9 changes
- Before Picture and Before Description are on the same page.
- After Picture and After Description are on the same page.
- The complete improvement is entered on one screen: Heading → Before Photo → Before Description → After Photo → After Description → Save & Finish.
- Removed the old staged photo/description navigation that could cause the Before Description to disappear from the final validation.
- Save & Finish reads the live text widgets directly while they are still present on the same page.
- Uses only native Streamlit file uploaders; no custom component and no external JavaScript/CDN.
- Selected images are automatically compressed before being stored in the session.

## Deploy
1. Download the ZIP.
2. Replace the complete contents of the GitHub repository with the ZIP contents.
3. Keep the `.streamlit/config.toml` file.
4. In Streamlit Cloud, reboot the app.

## Mobile upload
Use **Browse files / Select file** and choose the image from the Android Gallery/File Picker. The app cannot control Android's Gallery application itself, but V9 removes the application-side staged uploader logic and accepts the selected file through Streamlit's native uploader.
