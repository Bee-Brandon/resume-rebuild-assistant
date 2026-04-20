"""Resume Rebuild Assistant — Streamlit UI."""
import json
import copy
import time
from pathlib import Path

import streamlit as st

from config import TEMPLATES_DIR, OUTPUT_DIR, DEFAULT_TEMPLATE
from rendering import render_docx, render_docx_styled, render_pdf, render_pdf_styled, DEFAULT_FORMAT
import db

# ── Page config ──
st.set_page_config(page_title="Resume Rebuild Assistant", page_icon="📄", layout="wide")

# ── API Key handling (sidebar) ──
# Allows users without a .env file to paste their key in the UI
import os
with st.sidebar:
    st.markdown("### Settings")
    _env_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not _env_key:
        from config import ANTHROPIC_API_KEY as _cfg_key
        _env_key = _cfg_key
    if _env_key:
        st.success("API key loaded", icon="✅")
    else:
        _ui_key = st.text_input("Anthropic API Key", type="password",
                                help="Get yours at console.anthropic.com/settings/keys")
        if _ui_key:
            os.environ["ANTHROPIC_API_KEY"] = _ui_key
            st.success("Key set for this session", icon="✅")
            st.rerun()
        else:
            st.warning("Enter your Anthropic API key to use extraction", icon="⚠️")
    st.markdown("---")

    # Template import — always visible in sidebar
    st.markdown("### Import Template")
    st.caption("Upload a resume template — PDF, image, or Word. "
               "PDFs and images are analyzed by AI to match the style.")
    tpl_upload = st.file_uploader("Upload template",
                                  type=["pdf", "png", "jpg", "jpeg", "docx"],
                                  key="tpl_upload")
    tpl_name = st.text_input("Template name", placeholder="e.g. modern_blue", key="tpl_import_name")

    if tpl_upload:
        is_visual = not tpl_upload.name.lower().endswith(".docx")

        if is_visual:
            st.caption("AI will analyze the layout, fonts, colors, and spacing to create a matching style preset.")

        if st.button("Import Template", key="tpl_import_btn", use_container_width=True, type="primary"):
            from config import TEMPLATES_DIR, OUTPUT_DIR
            tmp = OUTPUT_DIR / f"_tpl_{tpl_upload.name}"
            tmp.write_bytes(tpl_upload.getvalue())

            try:
                if is_visual:
                    # PDF/image → Claude Vision analysis → save as JSON preset
                    with st.spinner("Analyzing template layout with AI..."):
                        from analyze_template import analyze_template_image, load_template_image
                        img = load_template_image(str(tmp))
                        fmt_settings, description = analyze_template_image(img)

                    # Save the source image as a thumbnail for reference
                    presets_dir = TEMPLATES_DIR / "presets"
                    presets_dir.mkdir(exist_ok=True)
                    import re as _re
                    raw_name = tpl_name.strip() or tpl_upload.name.rsplit(".", 1)[0]
                    safe_name = _re.sub(r'[^a-zA-Z0-9_-]', '_', raw_name.replace(" ", "_"))

                    # Save preset JSON
                    preset_path = presets_dir / f"{safe_name}.json"
                    import json as _json
                    preset_data = {
                        "name": safe_name.replace("_", " ").title(),
                        "description": description,
                        "format": fmt_settings,
                    }
                    preset_path.write_text(_json.dumps(preset_data, indent=2), encoding="utf-8")

                    # Save thumbnail
                    thumb = img.copy()
                    thumb.thumbnail((200, 280))
                    thumb.save(str(presets_dir / f"{safe_name}.png"))

                    st.success(f"Style preset saved: **{safe_name}**")
                    st.caption(f"Description: {description}")
                    st.image(thumb, caption="Analyzed template", width=180)
                else:
                    # .docx → copy to templates/
                    from import_template import import_template as do_import
                    out = do_import(str(tmp), tpl_name.strip() or None)
                    st.success(f"Word template saved: **{out.name}**")

            except Exception as e:
                st.error(f"Import failed: {e}")
            finally:
                tmp.unlink(missing_ok=True)

    # Show saved templates and presets
    from config import TEMPLATES_DIR
    tpls = sorted(TEMPLATES_DIR.glob("*.docx"))
    presets_dir = TEMPLATES_DIR / "presets"
    presets = sorted(presets_dir.glob("*.json")) if presets_dir.exists() else []

    if tpls or presets:
        st.markdown("#### Saved")
        for t in tpls:
            st.caption(f"📄 {t.name}")
        for p in presets:
            import json as _json
            try:
                pd = _json.loads(p.read_text(encoding="utf-8"))
                st.caption(f"🎨 {pd.get('name', p.stem)}")
            except Exception:
                st.caption(f"🎨 {p.stem}")

    st.markdown("---")
    st.caption("Resume Rebuild Assistant v1.0")

# ── Session state defaults ──
_DEFAULTS = {
    "resume_data": None,       # the editable dict
    "resume_id": None,         # DB row id
    "original_filename": None,
    "extract_error": None,
    "last_save_hash": None,
}
for k, v in _DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


def _data_hash(d):
    return hash(json.dumps(d, sort_keys=True)) if d else None


def _auto_save():
    """Persist edited JSON to DB if changed since last save."""
    d = st.session_state.resume_data
    rid = st.session_state.resume_id
    if d is None or rid is None:
        return
    h = _data_hash(d)
    if h != st.session_state.last_save_hash:
        db.update_edited(rid, d)
        st.session_state.last_save_hash = h


# ── Tabs ──
tab_upload, tab_edit, tab_generate, tab_history = st.tabs(
    ["📤 Upload & Extract", "✏️ Edit", "📄 Generate", "🕘 History"]
)

# ══════════════════════════════════════════════
#  TAB 1: Upload & Extract
# ══════════════════════════════════════════════
with tab_upload:
    st.header("Upload & Extract")
    uploaded = st.file_uploader(
        "Drop a resume file", type=["pdf", "docx", "png", "jpg", "jpeg"],
        help="Accepts PDF (text or scanned), Word, or image files",
    )

    # Fast-path notice for .docx
    if uploaded and uploaded.name.lower().endswith(".docx"):
        st.info("💡 This is a Word file — a future version will let you edit it directly without re-templating. (Coming soon)", icon="ℹ️")

    if uploaded and st.button("🔍 Extract Resume Data", type="primary"):
        # Save upload to temp file
        tmp = OUTPUT_DIR / f"_upload_{uploaded.name}"
        tmp.write_bytes(uploaded.getvalue())

        try:
            with st.status("Processing resume...", expanded=True) as status:
                st.write("📂 Loading file...")
                from ingestion import load_resume
                content = load_resume(tmp)

                if content["type"] == "images":
                    st.write(f"🖼️ Scanned PDF — converted {len(content['content'])} page(s) to images at 300 DPI")
                else:
                    st.write(f"📝 Extracted {len(content['content']):,} characters of text")

                st.write("🤖 Sending to Claude for structured extraction...")
                from extraction import extract_structured
                data = extract_structured(content)

                st.write("💾 Saving to database...")
                rid = db.save_resume(uploaded.name, data)

                status.update(label="Extraction complete!", state="complete")

            st.session_state.resume_data = data
            st.session_state.resume_id = rid
            st.session_state.original_filename = uploaded.name
            st.session_state.last_save_hash = _data_hash(data)
            st.session_state.extract_error = None

            name = data.get("contact", {}).get("name", "Unknown")
            st.success(f"Extracted resume for **{name}** — switch to the Edit tab to review and modify.")

            with st.expander("Raw extracted JSON", expanded=False):
                st.json(data)

        except RuntimeError as e:
            st.session_state.extract_error = str(e)
            st.error(f"Extraction failed: {e}")
            if "API error" in str(e):
                if st.button("🔄 Retry"):
                    st.rerun()
        except Exception as e:
            err = str(e)
            if "poppler" in err.lower() or "pdftoppm" in err.lower():
                st.error(f"PDF image conversion failed — is Poppler installed and on PATH?\n\nDetail: {err}")
            else:
                st.error(f"Unexpected error: {err}")
        finally:
            tmp.unlink(missing_ok=True)


# ══════════════════════════════════════════════
#  TAB 2: Edit
# ══════════════════════════════════════════════
with tab_edit:
    st.header("Edit Resume Data")

    if st.session_state.resume_data is None:
        st.info("Upload and extract a resume first, or load one from History.")
    else:
        data = st.session_state.resume_data

        # ── Contact ──
        st.subheader("Contact Information")
        c = data.setdefault("contact", {})
        col1, col2 = st.columns(2)
        c["name"] = col1.text_input("Full Name", value=c.get("name", ""), key="ed_name")
        c["email"] = col2.text_input("Email", value=c.get("email", ""), key="ed_email")
        col3, col4 = st.columns(2)
        c["phone"] = col3.text_input("Phone", value=c.get("phone", ""), key="ed_phone")
        c["location"] = col4.text_input("Location", value=c.get("location", ""), key="ed_loc")
        c["linkedin"] = st.text_input("LinkedIn", value=c.get("linkedin", ""), key="ed_li")

        # ── Summary ──
        st.subheader("Professional Summary")
        data["summary"] = st.text_area(
            "Summary", value=data.get("summary", ""), height=100, key="ed_summary",
            label_visibility="collapsed",
        )

        # ── Skills ──
        st.subheader("Skills")
        skills = data.setdefault("skills", [])
        skills_str = st.text_input(
            "Skills (comma-separated)",
            value=", ".join(skills),
            key="ed_skills",
            help="Type skills separated by commas. Each becomes a removable tag.",
        )
        data["skills"] = [s.strip() for s in skills_str.split(",") if s.strip()]
        # Show as chips
        if data["skills"]:
            chips_html = " ".join(
                f"<span style='background:#1e3a4a;color:#8dd8f8;padding:3px 10px;border-radius:12px;"
                f"font-size:13px;margin:3px;display:inline-block;border:1px solid #2a5a6a'>{skill}</span>"
                for skill in data["skills"]
            )
            st.markdown(chips_html, unsafe_allow_html=True)

        # ── Work History ──
        st.subheader("Work History")
        work = data.setdefault("work_history", [])

        for i, job in enumerate(work):
            with st.expander(f"**{job.get('role', 'Role')}** — {job.get('employer', 'Employer')}", expanded=(i == 0)):
                col1, col2 = st.columns(2)
                job["role"] = col1.text_input("Job Title", value=job.get("role", ""), key=f"wr_{i}")
                job["employer"] = col2.text_input("Employer", value=job.get("employer", ""), key=f"we_{i}")
                col3, col4, col5 = st.columns(3)
                job["start_date"] = col3.text_input("Start", value=job.get("start_date", ""), key=f"ws_{i}")
                job["end_date"] = col4.text_input("End", value=job.get("end_date", ""), key=f"wend_{i}")
                job["location"] = col5.text_input("Location", value=job.get("location", ""), key=f"wl_{i}")

                st.markdown("**Bullet Points**")
                bullets = job.setdefault("bullets", [])
                new_bullets = []
                for j, b in enumerate(bullets):
                    bc1, bc2 = st.columns([10, 1])
                    val = bc1.text_input(f"Bullet {j+1}", value=b, key=f"wb_{i}_{j}", label_visibility="collapsed")
                    if bc2.button("🗑", key=f"wbd_{i}_{j}", help="Remove bullet"):
                        continue  # skip — effectively deletes
                    new_bullets.append(val)
                job["bullets"] = new_bullets

                bc1, bc2 = st.columns([1, 1])
                if bc1.button("➕ Add Bullet", key=f"wba_{i}"):
                    job["bullets"].append("")
                    st.rerun()
                if bc2.button("🗑 Remove Job", key=f"wdel_{i}", type="secondary"):
                    work.pop(i)
                    st.rerun()

        col_add, col_reorder = st.columns(2)
        if col_add.button("➕ Add Job"):
            work.append({"role": "", "employer": "", "location": "", "start_date": "", "end_date": "", "bullets": [""]})
            st.rerun()

        # ── Education ──
        st.subheader("Education")
        edu = data.setdefault("education", [])

        for i, entry in enumerate(edu):
            with st.expander(f"{entry.get('institution', 'Institution')}", expanded=(i == 0)):
                col1, col2 = st.columns(2)
                entry["institution"] = col1.text_input("Institution", value=entry.get("institution", ""), key=f"ei_{i}")
                entry["degree"] = col2.text_input("Degree", value=entry.get("degree", ""), key=f"ed_{i}")
                col3, col4, col5 = st.columns(3)
                entry["field"] = col3.text_input("Field of Study", value=entry.get("field", ""), key=f"ef_{i}")
                entry["graduation_year"] = col4.text_input("Year", value=entry.get("graduation_year", ""), key=f"ey_{i}")
                entry["honors"] = col5.text_input("Honors", value=entry.get("honors", ""), key=f"eh_{i}")
                if st.button("🗑 Remove", key=f"edel_{i}"):
                    edu.pop(i)
                    st.rerun()

        if st.button("➕ Add Education"):
            edu.append({"institution": "", "degree": "", "field": "", "graduation_year": "", "honors": ""})
            st.rerun()

        # ── Certifications ──
        st.subheader("Certifications")
        certs = data.setdefault("certifications", [])

        for i, cert in enumerate(certs):
            col1, col2, col3, col4 = st.columns([3, 3, 2, 1])
            cert["name"] = col1.text_input("Certification", value=cert.get("name", ""), key=f"cn_{i}")
            cert["issuer"] = col2.text_input("Issuer", value=cert.get("issuer", ""), key=f"ci_{i}")
            cert["date"] = col3.text_input("Date", value=cert.get("date", ""), key=f"cd_{i}")
            if col4.button("🗑", key=f"cdel_{i}"):
                certs.pop(i)
                st.rerun()

        if st.button("➕ Add Certification"):
            certs.append({"name": "", "issuer": "", "date": ""})
            st.rerun()

        # ── Auto-save ──
        _auto_save()
        st.caption("Changes are auto-saved to the database.")


# ══════════════════════════════════════════════
#  TAB 3: Generate
# ══════════════════════════════════════════════
with tab_generate:
    st.header("Generate Output")

    if st.session_state.resume_data is None:
        st.info("Upload, extract, and edit a resume first.")
    else:
        data = st.session_state.resume_data
        name = data.get("contact", {}).get("name", "resume").replace(" ", "_")

        # Initialize resume_format in session state
        if "resume_format" not in st.session_state:
            st.session_state["resume_format"] = copy.deepcopy(DEFAULT_FORMAT)

        rf = st.session_state["resume_format"]

        col_left, col_mid, col_right = st.columns([1, 1, 1], gap="large")

        # Section display names
        SECTION_LABELS = {
            "summary": "Summary",
            "skills": "Skills",
            "experience": "Experience",
            "education": "Education",
            "certifications": "Certifications",
        }

        # ── Left column: Format Controls ──
        with col_left:
            # Layout style selector at the top
            layout_opts = ["single", "two-column"]
            layout_labels_map = {"single": "Single Column", "two-column": "Two Column Sidebar"}
            current_ls = rf.get("layout_style", "single")
            rf["layout_style"] = st.radio(
                "Layout Style",
                layout_opts,
                index=layout_opts.index(current_ls),
                format_func=lambda x: layout_labels_map[x],
                key="fmt_layout_style",
                horizontal=True,
            )

            # Tagline (two-column only but always stored)
            if rf["layout_style"] == "two-column":
                rf["tagline"] = st.text_input("Tagline (under name)",
                    value=rf.get("tagline", ""),
                    placeholder="e.g. SOFTWARE ENGINEER, ATTORNEY",
                    key="fmt_tagline")
                rf["header_banner_color"] = st.color_picker(
                    "Banner Color", value=rf.get("header_banner_color", "#d4c5a9"), key="fmt_banner")

            st.subheader("Typography")

            rf["font_family"] = st.selectbox(
                "Font Family",
                ["Calibri", "Arial", "Garamond", "Times New Roman", "Georgia", "Cambria"],
                index=["Calibri", "Arial", "Garamond", "Times New Roman", "Georgia", "Cambria"].index(rf["font_family"]),
                key="fmt_font",
            )

            sz1, sz2 = st.columns(2)
            rf["name_size"] = sz1.number_input("Name Size", min_value=14, max_value=24,
                                                value=int(rf["name_size"]), step=1, key="fmt_name_sz")
            rf["body_size"] = sz2.number_input("Body Size", min_value=9.0, max_value=13.0,
                                                value=float(rf["body_size"]), step=0.5, key="fmt_body_sz")

            sz3, sz4 = st.columns(2)
            rf["section_header_size"] = sz3.number_input("Header Size", min_value=10, max_value=16,
                                                          value=int(rf["section_header_size"]), step=1, key="fmt_hdr_sz")
            rf["date_size"] = sz4.number_input("Date Size", min_value=8, max_value=12,
                                                value=int(rf["date_size"]), step=1, key="fmt_date_sz")

            rf["line_spacing"] = st.selectbox(
                "Line Spacing",
                [1.0, 1.05, 1.1, 1.15, 1.2],
                index=[1.0, 1.05, 1.1, 1.15, 1.2].index(rf["line_spacing"]),
                key="fmt_ls",
            )

            margin_val = st.slider("Margins (inches)", min_value=0.4, max_value=1.0,
                                    value=float(rf["margin_left"]), step=0.05, key="fmt_margin")
            rf["margin_top"] = margin_val
            rf["margin_bottom"] = margin_val
            rf["margin_left"] = margin_val
            rf["margin_right"] = margin_val

            rf["show_section_lines"] = st.checkbox("Show section lines", value=rf["show_section_lines"],
                                                    key="fmt_lines")

            with st.expander("Colors"):
                rf["name_color"] = st.color_picker("Name", value=rf["name_color"], key="fmt_c_name")
                rf["header_color"] = st.color_picker("Headers", value=rf["header_color"], key="fmt_c_hdr")
                rf["body_color"] = st.color_picker("Body", value=rf["body_color"], key="fmt_c_body")
                rf["accent_color"] = st.color_picker("Accent", value=rf["accent_color"], key="fmt_c_accent")
                rf["section_line_color"] = st.color_picker("Lines", value=rf["section_line_color"], key="fmt_c_line")

            if st.button("Reset All to Defaults", key="fmt_reset"):
                st.session_state["resume_format"] = copy.deepcopy(DEFAULT_FORMAT)
                st.rerun()

        # ── Middle column: Layout & Section Order ──
        with col_mid:
            is_two_col = rf.get("layout_style") == "two-column"

            st.subheader("Layout")

            if not is_two_col:
                # Single-column contact layout
                layout_options = ["center", "left", "right"]
                layout_labels = {"center": "Centered under name", "left": "Left-aligned", "right": "Name left, contact right"}
                current_layout = rf.get("contact_layout", "center")
                rf["contact_layout"] = st.radio(
                    "Contact Info Position",
                    layout_options,
                    index=layout_options.index(current_layout),
                    format_func=lambda x: layout_labels[x],
                    key="fmt_contact_layout",
                )

            st.markdown("---")
            st.subheader("Section Order")

            order = rf.setdefault("section_order", list(DEFAULT_FORMAT["section_order"]))
            sidebar_secs = rf.setdefault("sidebar_sections", list(DEFAULT_FORMAT["sidebar_sections"]))

            if is_two_col:
                st.caption("Check sections to place in the **sidebar** (left column). Unchecked sections go in the main area (right column).")
                movable_secs = ["summary", "skills", "experience", "education", "certifications"]
                for sk in movable_secs:
                    in_sidebar = sk in sidebar_secs
                    label = SECTION_LABELS.get(sk, sk)
                    col_check = f"sidebar" if in_sidebar else "main"
                    checked = st.checkbox(
                        f"{label} → {'Sidebar' if in_sidebar else 'Main'}",
                        value=in_sidebar,
                        key=f"sb_{sk}",
                    )
                    if checked and sk not in sidebar_secs:
                        sidebar_secs.append(sk)
                    elif not checked and sk in sidebar_secs:
                        sidebar_secs.remove(sk)

                # Sidebar width
                rf["sidebar_width_pct"] = st.slider(
                    "Sidebar Width %", min_value=25, max_value=45,
                    value=int(rf.get("sidebar_width_pct", 35)), step=1, key="fmt_sidebar_w")
            else:
                st.caption("Move sections up or down to reorder.")

            # Section order with up/down buttons (applies to both layouts)
            for i, section_key in enumerate(order):
                c1, c2, c3 = st.columns([1, 6, 1])
                if i > 0:
                    if c1.button("▲", key=f"sec_up_{i}", help="Move up"):
                        order[i], order[i - 1] = order[i - 1], order[i]
                        st.rerun()
                else:
                    c1.markdown("")
                loc_tag = ""
                if is_two_col:
                    loc_tag = " 🔲" if section_key in sidebar_secs else " 📄"
                c2.markdown(f"**{i + 1}.** {SECTION_LABELS.get(section_key, section_key)}{loc_tag}")
                if i < len(order) - 1:
                    if c3.button("▼", key=f"sec_dn_{i}", help="Move down"):
                        order[i], order[i + 1] = order[i + 1], order[i]
                        st.rerun()

            # ── Visual Layout Preview ──
            st.markdown("---")
            st.caption("Preview")

            hdr_color = rf.get("header_color", "#1a1a2e")
            accent = rf.get("accent_color", "#444444")
            line_color = rf.get("section_line_color", "#999999")
            show_lines = rf.get("show_section_lines", True)
            line_css = f"border-bottom:1px solid {line_color};padding-bottom:2px;" if show_lines else ""
            name_color = rf.get("name_color", "#1a1a1a")
            font_fam = rf.get("font_family", "Calibri")

            def _section_block(label):
                return (
                    f"<div style='font-size:9px;font-weight:700;color:{hdr_color};{line_css}"
                    f"margin-top:5px;margin-bottom:2px'>{label}</div>"
                    f"<div style='font-size:7px;color:#888;margin-bottom:1px'>Content...</div>"
                )

            if is_two_col:
                # Two-column preview
                banner = rf.get("header_banner_color", "#d4c5a9")
                tagline = rf.get("tagline", "")
                sw = rf.get("sidebar_width_pct", 35)
                mw = 100 - sw

                sidebar_html = "".join(
                    _section_block(SECTION_LABELS.get(sk, sk).upper())
                    for sk in order if sk in sidebar_secs
                )
                main_html = "".join(
                    _section_block(SECTION_LABELS.get(sk, sk).upper())
                    for sk in order if sk not in sidebar_secs
                )

                tagline_html = f"<div style='font-size:8px;color:{name_color};letter-spacing:2px;margin-top:2px'>{tagline.upper()}</div>" if tagline else ""

                preview_html = (
                    f"<div style='background:white;border:1px solid #ccc;border-radius:4px;"
                    f"max-width:280px;box-shadow:0 2px 8px rgba(0,0,0,0.15);font-family:{font_fam},sans-serif;overflow:hidden'>"
                    # Banner
                    f"<div style='background:{banner};padding:10px 12px'>"
                    f"<div style='font-size:16px;font-weight:700;color:{name_color};letter-spacing:3px'>NAME</div>"
                    f"{tagline_html}</div>"
                    # Two columns
                    f"<div style='display:flex'>"
                    f"<div style='width:{sw}%;padding:8px 10px;border-right:1px solid #eee'>"
                    f"{_section_block('CONTACT')}{sidebar_html}</div>"
                    f"<div style='width:{mw}%;padding:8px 10px'>{main_html}</div>"
                    f"</div></div>"
                )
            else:
                # Single-column preview
                contact_pos = rf.get("contact_layout", "center")
                if contact_pos == "right":
                    header_html = (
                        f"<div style='display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px'>"
                        f"<div style='font-size:14px;font-weight:700;color:{name_color}'>Name</div>"
                        f"<div style='font-size:8px;color:{accent};text-align:right'>email<br>phone<br>city</div></div>"
                    )
                elif contact_pos == "left":
                    header_html = (
                        f"<div style='margin-bottom:8px'>"
                        f"<div style='font-size:14px;font-weight:700;color:{name_color}'>Name</div>"
                        f"<div style='font-size:8px;color:{accent}'>email | phone | city</div></div>"
                    )
                else:
                    header_html = (
                        f"<div style='text-align:center;margin-bottom:8px'>"
                        f"<div style='font-size:14px;font-weight:700;color:{name_color}'>Name</div>"
                        f"<div style='font-size:8px;color:{accent}'>email | phone | city</div></div>"
                    )
                sections_html = "".join(_section_block(SECTION_LABELS.get(sk, sk).upper()) for sk in order)
                preview_html = (
                    f"<div style='background:white;border:1px solid #ccc;border-radius:4px;"
                    f"padding:12px 14px;max-width:260px;box-shadow:0 2px 8px rgba(0,0,0,0.15);"
                    f"font-family:{font_fam},sans-serif'>{header_html}{sections_html}</div>"
                )

            st.markdown(preview_html, unsafe_allow_html=True)

        # ── Right column: Generate & Download ──
        with col_right:
            st.subheader("Generate")

            # Load available style presets
            _presets_dir = TEMPLATES_DIR / "presets"
            _preset_files = sorted(_presets_dir.glob("*.json")) if _presets_dir.exists() else []
            _presets = []
            for pf in _preset_files:
                try:
                    pd = json.loads(pf.read_text(encoding="utf-8"))
                    pd["_path"] = pf
                    _presets.append(pd)
                except Exception:
                    pass

            # Build mode options dynamically
            mode_options = ["Styled (custom)"]
            if _presets:
                mode_options += [f"🎨 {p['name']}" for p in _presets]
            mode_options.append("Template-based (Word)")

            render_mode = st.selectbox("Rendering Mode", mode_options, key="fmt_mode")

            # Show preset thumbnail if a preset is selected
            chosen_tpl = None
            preset_fmt = None

            if render_mode.startswith("🎨"):
                # A style preset was selected
                preset_idx = [f"🎨 {p['name']}" for p in _presets].index(render_mode)
                preset = _presets[preset_idx]
                preset_fmt = preset.get("format", {})
                st.caption(preset.get("description", ""))

                # Show thumbnail if available
                thumb_path = preset["_path"].with_suffix(".png")
                if thumb_path.exists():
                    st.image(str(thumb_path), caption="Template reference", width=200)

                # Apply preset to format settings
                if st.button("Load into format controls", key="load_preset"):
                    for k, v in preset_fmt.items():
                        rf[k] = v
                    st.success("Format settings loaded from preset!")
                    st.rerun()

            elif render_mode == "Template-based (Word)":
                templates = sorted(TEMPLATES_DIR.glob("*.docx"))
                if not templates:
                    st.warning("No templates found. Import one in the sidebar →")
                    chosen_tpl = None
                else:
                    tpl_names = [t.name for t in templates]
                    chosen_idx = st.selectbox("Template", range(len(tpl_names)),
                                              format_func=lambda i: tpl_names[i], key="fmt_tpl")
                    chosen_tpl = templates[chosen_idx]
                st.caption("Uses the Word template's own styling and Jinja2 tokens.")

            st.markdown("---")

            # Determine which format to use for styled rendering
            use_fmt = {**rf, **(preset_fmt or {})} if preset_fmt else rf
            is_template_mode = render_mode == "Template-based (Word)"

            # Generate Word
            if st.button("📝 Generate Word Document", type="primary", key="gen_docx", use_container_width=True):
                out_path = OUTPUT_DIR / f"{name}_resume.docx"
                try:
                    if is_template_mode:
                        if chosen_tpl is None:
                            st.error("No template selected.")
                            st.stop()
                        render_docx(data, chosen_tpl, out_path)
                    else:
                        render_docx_styled(data, use_fmt, out_path)
                    if st.session_state.resume_id:
                        db.update_output_path(st.session_state.resume_id, str(out_path))
                    st.success("Word document generated!")
                    with open(out_path, "rb") as f:
                        st.download_button(
                            "⬇️ Download .docx", f.read(),
                            file_name=out_path.name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="dl_docx", use_container_width=True,
                        )
                except Exception as e:
                    st.error(f"Word generation failed: {e}")

            # Generate PDF
            if st.button("📄 Generate PDF", key="gen_pdf", use_container_width=True):
                out_path = OUTPUT_DIR / f"{name}_resume.pdf"
                try:
                    with st.spinner("Generating PDF (requires Microsoft Word)..."):
                        if is_template_mode:
                            if chosen_tpl is None:
                                st.error("No template selected.")
                                st.stop()
                            render_pdf(data, chosen_tpl, out_path)
                        else:
                            render_pdf_styled(data, use_fmt, out_path)
                    if st.session_state.resume_id:
                        db.update_output_path(st.session_state.resume_id, str(out_path))
                    st.success("PDF generated!")
                    with open(out_path, "rb") as f:
                        st.download_button(
                            "⬇️ Download .pdf", f.read(),
                            file_name=out_path.name,
                            mime="application/pdf",
                            key="dl_pdf", use_container_width=True,
                        )
                except Exception as e:
                    err = str(e)
                    if "word" in err.lower() or "com" in err.lower():
                        st.error("PDF conversion requires Microsoft Word. "
                                 "Generate Word instead and export to PDF from Word.")
                    else:
                        st.error(f"PDF generation failed: {e}")


# ══════════════════════════════════════════════
#  TAB 4: History
# ══════════════════════════════════════════════
with tab_history:
    st.header("Resume History")
    resumes = db.list_resumes()

    if not resumes:
        st.info("No resumes processed yet. Upload one to get started.")
    else:
        for r in resumes:
            col1, col2, col3 = st.columns([4, 2, 2])
            label = r["participant_name"] or r["original_filename"]
            col1.markdown(f"**{label}**  \n`{r['original_filename']}`")
            col2.caption(f"Updated: {r['updated_at'][:16]}")

            if col3.button("📂 Load", key=f"hload_{r['id']}"):
                full = db.load_resume(r["id"])
                if full:
                    # Use edited_json if available, otherwise extracted
                    st.session_state.resume_data = full.get("edited_json") or full["extracted_json"]
                    st.session_state.resume_id = r["id"]
                    st.session_state.original_filename = r["original_filename"]
                    st.session_state.last_save_hash = _data_hash(st.session_state.resume_data)
                    st.success(f"Loaded {label} — switch to Edit tab.")
                    st.rerun()
