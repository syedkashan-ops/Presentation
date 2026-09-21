import hashlib
from pathlib import Path

import streamlit as st

from app_core import build_ppt, compress_image, safe_filename

st.set_page_config(page_title="PSO Outlet Improvement Report", page_icon="📊", layout="centered")

ASSETS = Path(__file__).parent / "assets"
ALLOWED = ["jpg", "jpeg", "png", "webp"]


def init_state():
    defaults = {
        "step": 0,
        "outlet": {"name": "", "code": "", "presented_by": "", "designation": ""},
        "improvements": [],
        "current_uploads": {},
        "photo_stage": "before",
        "app_epoch": 1,
        "before_replace": 0,
        "after_replace": 0,
        "ppt_bytes": None,
        "ppt_name": None,
        "draft": {"heading": "", "before_desc": "", "after_desc": ""},
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def reset_current_improvement():
    # Only app-owned state is changed here. No widget key is ever overwritten.
    st.session_state.current_uploads = {}
    st.session_state.photo_stage = "before"
    st.session_state.before_replace = 0
    st.session_state.after_replace = 0
    st.session_state.draft = {"heading": "", "before_desc": "", "after_desc": ""}


def start_over():
    # Incrementing the epoch gives every new report fresh widget keys without
    # mutating existing widgets in the current Streamlit run.
    st.session_state.step = 0
    st.session_state.outlet = {"name": "", "code": "", "presented_by": "", "designation": ""}
    st.session_state.improvements = []
    st.session_state.current_uploads = {}
    st.session_state.photo_stage = "before"
    st.session_state.before_replace = 0
    st.session_state.after_replace = 0
    st.session_state.ppt_bytes = None
    st.session_state.ppt_name = None
    st.session_state.draft = {"heading": "", "before_desc": "", "after_desc": ""}
    st.session_state.app_epoch += 1


def store_upload(uploaded_file, slot):
    raw = uploaded_file.getvalue()
    digest = hashlib.sha256(raw).hexdigest()
    existing = st.session_state.current_uploads.get(slot)
    if existing and existing.get("sha256") == digest:
        return existing

    data, size = compress_image(raw)
    record = {
        "data": data,
        "name": uploaded_file.name,
        "size": size,
        "original_bytes": len(raw),
        "compressed_bytes": len(data),
        "sha256": digest,
    }
    st.session_state.current_uploads[slot] = record
    return record


def size_caption(record):
    old_mb = record["original_bytes"] / 1024 / 1024
    new_mb = record["compressed_bytes"] / 1024 / 1024
    w, h = record["size"]
    return f"Optimized automatically: {w}×{h}px • {old_mb:.1f} MB → {new_mb:.1f} MB"


def current_keys():
    epoch = st.session_state.app_epoch
    number = len(st.session_state.improvements) + 1
    prefix = f"e{epoch}_i{number}"
    return {
        "heading": f"{prefix}_heading",
        "before_desc": f"{prefix}_before_desc",
        "after_desc": f"{prefix}_after_desc",
        "before_file": f"{prefix}_before_file_r{st.session_state.before_replace}",
        "after_file": f"{prefix}_after_file_r{st.session_state.after_replace}",
        "before_slot": f"{prefix}_before",
        "after_slot": f"{prefix}_after",
    }


def sync_draft(keys):
    # Mirror widget values into app-owned state before a stage change.
    # This prevents a description from being lost when its widget disappears
    # on the next rerun.
    st.session_state.draft["heading"] = st.session_state.get(keys["heading"], "") or ""
    st.session_state.draft["before_desc"] = st.session_state.get(keys["before_desc"], "") or ""
    st.session_state.draft["after_desc"] = st.session_state.get(keys["after_desc"], "") or ""


init_state()

st.markdown("""
<style>
.block-container {max-width: 1050px; padding-top: 1.1rem; padding-bottom: 2rem;}
.pso-title {font-size: 2rem; font-weight: 800; color:#0052A5; margin-bottom:0.1rem;}
.pso-sub {color:#4b5563; margin-bottom:1.2rem;}
[data-testid="stFileUploader"] {border: 1px solid #dbe3ea; border-radius: 12px; padding: 0.5rem;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="pso-title">PSO Outlet Improvement Report</div>', unsafe_allow_html=True)
st.markdown('<div class="pso-sub">Create a professional Before & After PowerPoint presentation from your outlet improvement photographs.</div>', unsafe_allow_html=True)

# ---------------- STEP 1: OUTLET DETAILS ----------------
if st.session_state.step == 0:
    st.subheader("1. Outlet Details")
    epoch = st.session_state.app_epoch

    with st.form(f"outlet_form_{epoch}", clear_on_submit=False):
        name = st.text_input("Outlet Name *", key=f"outlet_name_{epoch}", placeholder="e.g. PSO Shahrah-e-Faisal")
        code = st.text_input("Outlet Code *", key=f"outlet_code_{epoch}", placeholder="e.g. 123456")
        presented_by = st.text_input("Presented By", key=f"presented_by_{epoch}", placeholder="Optional")
        designation = st.text_input("Designation", key=f"designation_{epoch}", placeholder="Optional")
        submitted = st.form_submit_button("Continue →", type="primary", use_container_width=True)

    if submitted:
        if not name.strip():
            st.error("Please enter the Outlet Name.")
        elif not code.strip():
            st.error("Please enter the Outlet Code.")
        else:
            st.session_state.outlet = {
                "name": name.strip(),
                "code": code.strip(),
                "presented_by": presented_by.strip(),
                "designation": designation.strip(),
            }
            reset_current_improvement()
            st.session_state.step = 1
            st.rerun()

# ---------------- STEP 2: IMPROVEMENTS ----------------
elif st.session_state.step == 1:
    number = len(st.session_state.improvements) + 1
    keys = current_keys()

    st.subheader(f"2. Improvement #{number}")
    st.caption(f"Outlet: {st.session_state.outlet['name']}  |  Code: {st.session_state.outlet['code']}")

    heading = st.text_input(
        "Improvement Heading *",
        key=keys["heading"],
        placeholder="e.g. Forecourt Painting & Branding Improvement"
    )
    st.session_state.draft["heading"] = heading or ""

    # BEFORE PHOTO
    if st.session_state.photo_stage == "before":
        st.markdown("### Before Picture")
        st.info("Tap **Browse files** and choose the photo from Gallery. V8 accepts the file picker without a browser-side extension filter, then validates and compresses the image on the server.")
        uploaded = st.file_uploader(
            "Choose Before Photo",
            type=None,
            accept_multiple_files=False,
            key=keys["before_file"],
            help="Supported: JPG, JPEG, PNG and WEBP."
        )

        record = st.session_state.current_uploads.get(keys["before_slot"])
        if uploaded is not None:
            try:
                record = store_upload(uploaded, keys["before_slot"])
            except Exception as exc:
                st.error(f"This selected file is not a readable image: {type(exc).__name__}: {exc}")
                record = None

        if record:
            st.image(record["data"], caption="Before picture saved", use_container_width=True)
            st.success("Before picture is saved.")
            st.caption(size_caption(record))
            if st.button("Continue to Before Description →", type="primary", use_container_width=True,
                         key=f"{keys['before_slot']}_continue"):
                if not heading.strip():
                    st.error("Please enter the Improvement Heading first.")
                else:
                    st.session_state.photo_stage = "before_description"
                    st.rerun()

    # BEFORE DESCRIPTION
    elif st.session_state.photo_stage == "before_description":
        record = st.session_state.current_uploads.get(keys["before_slot"])
        if not record:
            st.session_state.photo_stage = "before"
            st.rerun()

        st.markdown("### Before Description")
        st.image(record["data"], caption="Before picture", use_container_width=True)
        before_desc = st.text_area(
            "Before Description *",
            key=keys["before_desc"],
            height=130,
            placeholder="Describe the condition before improvement..."
        )
        st.session_state.draft["before_desc"] = before_desc or ""

        c1, c2 = st.columns(2)
        with c1:
            if st.button("← Replace Before Photo", use_container_width=True,
                         key=f"{keys['before_slot']}_replace"):
                st.session_state.current_uploads.pop(keys["before_slot"], None)
                st.session_state.before_replace += 1
                st.session_state.photo_stage = "before"
                st.rerun()
        with c2:
            if st.button("Continue to After Photo →", type="primary", use_container_width=True,
                         key=f"{keys['before_slot']}_next"):
                if not heading.strip():
                    st.error("Please enter the Improvement Heading.")
                elif not before_desc.strip():
                    st.error("Please enter the Before Description.")
                else:
                    st.session_state.photo_stage = "after"
                    st.rerun()

    # AFTER PHOTO
    elif st.session_state.photo_stage == "after":
        st.markdown("### After Picture")
        st.info("Choose the After photo from your Gallery. Your Before photo and description remain saved in this session.")
        uploaded = st.file_uploader(
            "Choose After Photo",
            type=None,
            accept_multiple_files=False,
            key=keys["after_file"],
            help="Supported: JPG, JPEG, PNG and WEBP."
        )

        record = st.session_state.current_uploads.get(keys["after_slot"])
        if uploaded is not None:
            try:
                record = store_upload(uploaded, keys["after_slot"])
            except Exception as exc:
                st.error(f"This selected file is not a readable image: {type(exc).__name__}: {exc}")
                record = None

        if record:
            st.image(record["data"], caption="After picture saved", use_container_width=True)
            st.success("After picture is saved.")
            st.caption(size_caption(record))
            if st.button("Continue to After Description →", type="primary", use_container_width=True,
                         key=f"{keys['after_slot']}_continue"):
                st.session_state.photo_stage = "after_description"
                st.rerun()

    # AFTER DESCRIPTION / SAVE
    elif st.session_state.photo_stage == "after_description":
        record = st.session_state.current_uploads.get(keys["after_slot"])
        if not record:
            st.session_state.photo_stage = "after"
            st.rerun()

        st.markdown("### After Description")
        st.image(record["data"], caption="After picture", use_container_width=True)
        after_desc = st.text_area(
            "After Description *",
            key=keys["after_desc"],
            height=130,
            placeholder="Describe the improvement after completion..."
        )
        st.session_state.draft["after_desc"] = after_desc or ""

        if st.button("← Replace After Photo", use_container_width=True,
                     key=f"{keys['after_slot']}_replace"):
            st.session_state.current_uploads.pop(keys["after_slot"], None)
            st.session_state.after_replace += 1
            st.session_state.photo_stage = "after"
            st.rerun()

        st.markdown("---")
        c1, c2 = st.columns(2)

        def collect_current():
            sync_draft(keys)
            h = st.session_state.draft.get("heading", "").strip()
            bd = st.session_state.draft.get("before_desc", "").strip()
            ad = st.session_state.draft.get("after_desc", "").strip()
            b = st.session_state.current_uploads.get(keys["before_slot"])
            a = st.session_state.current_uploads.get(keys["after_slot"])
            return h, bd, ad, b, a

        def validate(values):
            h, bd, ad, b, a = values
            if not h: return "Please enter the Improvement Heading."
            if not b: return "Please upload the Before Picture."
            if not bd: return "Please enter the Before Description."
            if not a: return "Please upload the After Picture."
            if not ad: return "Please enter the After Description."
            return None

        def append_improvement(values):
            h, bd, ad, b, a = values
            st.session_state.improvements.append({
                "heading": h,
                "before": b["data"],
                "before_name": b["name"],
                "before_desc": bd,
                "after": a["data"],
                "after_name": a["name"],
                "after_desc": ad,
            })

        with c1:
            if st.button("➕ Save & Continue", use_container_width=True,
                         key=f"e{st.session_state.app_epoch}_save_continue_{number}"):
                values = collect_current()
                error = validate(values)
                if error:
                    st.error(error)
                else:
                    append_improvement(values)
                    reset_current_improvement()
                    st.rerun()

        with c2:
            if st.button("✅ Save & Finish", type="primary", use_container_width=True,
                         key=f"e{st.session_state.app_epoch}_save_finish_{number}"):
                values = collect_current()
                error = validate(values)
                if error:
                    st.error(error)
                else:
                    append_improvement(values)
                    st.session_state.step = 2
                    st.session_state.ppt_bytes = None
                    st.session_state.ppt_name = None
                    st.rerun()

    if st.session_state.improvements:
        st.markdown("### Improvements already added")
        for i, item in enumerate(st.session_state.improvements, 1):
            st.write(f"**{i}. {item['heading']}**")

    if st.button("↩ Start Over", use_container_width=True, key=f"start_over_{st.session_state.app_epoch}_step1"):
        start_over()
        st.rerun()

# ---------------- STEP 3: GENERATE ----------------
elif st.session_state.step == 2:
    st.subheader("3. Presentation Ready")
    st.success(f"{len(st.session_state.improvements)} improvement(s) captured for {st.session_state.outlet['name']}.")

    if st.session_state.ppt_bytes is None:
        try:
            st.session_state.ppt_bytes = build_ppt(
                st.session_state.outlet,
                st.session_state.improvements,
                ASSETS,
            )
            st.session_state.ppt_name = f"{safe_filename(st.session_state.outlet['name'])}_Improvement_Report.pptx"
        except Exception as exc:
            st.error(f"PowerPoint generation failed: {type(exc).__name__}: {exc}")

    if st.session_state.ppt_bytes:
        st.download_button(
            "⬇️ Download PowerPoint",
            data=st.session_state.ppt_bytes,
            file_name=st.session_state.ppt_name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
            use_container_width=True,
            key=f"download_{st.session_state.app_epoch}",
        )
        st.info("The PowerPoint is generated and ready to download.")

    if st.button("＋ Create Another Report", use_container_width=True,
                 key=f"another_{st.session_state.app_epoch}"):
        start_over()
        st.rerun()
