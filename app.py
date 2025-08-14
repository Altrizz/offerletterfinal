# app.py - Streamlit (Cloud-ready) with robust PPTX replacement
# - Replaces placeholders even when split across runs/tables/grouped shapes
# - Supports both {{CURLY}} tokens and legacy X-style tokens
# - Includes Join Date, Offer Date, City, salary formatting
# - Output file name: "Offer Letter - <Candidate Name>.pptx"
# - Hogarth Worldwide brand preset + customizable theme
# - Optional Windows Outlook draft helper (ignored on Cloud/macOS)

import re
import os
import sys
import base64
from datetime import date
from io import BytesIO

import streamlit as st
from pptx import Presentation
from pptx.shapes.group import GroupShape

# ==============================
# Helpers: dates, formatting, tokens
# ==============================
MESES_ES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]

def fecha_es(d: date) -> str:
    """Return 'DD de <mes> de YYYY' in Spanish."""
    return f"{d.day} de {MESES_ES[d.month-1]} de {d.year}"

def format_ars_dots(value: str) -> str:
    digits = "".join(ch for ch in str(value) if ch.isdigit())
    if not digits:
        return str(value)
    return f"{int(digits):,}".replace(",", ".")

# Curly placeholders like {{CANDIDATE_NAME}}
PLACEHOLDER = re.compile(r"{{\s*([A-Z0-9_]+)\s*}}")

# Legacy X-style tokens from the provided template
PAT_NAME = re.compile(r"\{X{6}\}")            # {XXXXXX}
PAT_POS  = re.compile(r"(?<!\{)X{8}(?!\})")   # XXXXXXXX not inside {}
PAT_DATE = re.compile(r"\bX{2}\s+de\s+X{4,5}\s+de\s+\d{4}\b")
PAT_SAL  = re.compile(r"\bX\.XXX\.XXX\b")

def apply_x_style(text: str, mapping: dict) -> str:
    name = mapping.get("CANDIDATE_NAME")
    position = mapping.get("POSITION")
    join_date_es = mapping.get("JOIN_DATE")
    salary = mapping.get("SALARY")
    city = mapping.get("CITY", "Buenos Aires")
    offer_date_es = mapping.get("DATE")

    out = text
    if name:
        out = PAT_NAME.sub(name, out)
    if position:
        out = PAT_POS.sub(position, out)
    if join_date_es:
        out = PAT_DATE.sub(join_date_es, out)
    if salary:
        out = PAT_SAL.sub(format_ars_dots(salary), out)

    # A line that is just ", Buenos Aires" should become "<DATE>, <CITY>"
    striped = out.strip()
    if striped == ", Buenos Aires" or striped.endswith(", Buenos Aires"):
        out = f"{offer_date_es}, {city}" if offer_date_es else city
    return out

def replace_placeholders_in_text(text: str, mapping: dict) -> str:
    """Resolve {{KEY}} placeholders, then legacy X-tokens."""
    def repl(m):
        key = m.group(1).upper()
        return str(mapping.get(key, m.group(0)))
    text2 = PLACEHOLDER.sub(repl, text)
    return apply_x_style(text2, mapping)

# ==============================
# PPTX replacement across runs / tables / grouped shapes
# ==============================
def _replace_in_text_frame(tf, mapping: dict):
    for para in tf.paragraphs:
        if not para.runs:
            if para.text:
                new = replace_placeholders_in_text(para.text, mapping)
                if new != para.text:
                    para.text = new
            continue
        full = "".join(run.text for run in para.runs)
        new = replace_placeholders_in_text(full, mapping)
        if new == full:
            continue
        # Put the new text in the first run, clear the rest (handles split tokens)
        para.runs[0].text = new
        for r in para.runs[1:]:
            r.text = ""

def _replace_in_table(tbl, mapping: dict):
    for r in tbl.rows:
        for c in r.cells:
            if c.text_frame:
                _replace_in_text_frame(c.text_frame, mapping)

def _walk_shapes(shapes, mapping: dict):
    for shape in shapes:
        # Text frames
        if getattr(shape, "has_text_frame", False) and shape.text_frame:
            _replace_in_text_frame(shape.text_frame, mapping)
        # Tables
        if getattr(shape, "has_table", False):
            _replace_in_table(shape.table, mapping)
        # Grouped shapes
        if isinstance(shape, GroupShape):
            _walk_shapes(shape.shapes, mapping)

def render_pptx(pptx_bytes: bytes, mapping: dict) -> BytesIO:
    prs = Presentation(BytesIO(pptx_bytes))
    for slide in prs.slides:
        _walk_shapes(slide.shapes, mapping)
    out = BytesIO()
    prs.save(out)
    out.seek(0)
    return out

# ==============================
# Optional: Windows Outlook draft helper (safe on Cloud/macOS)
# ==============================
def _open_outlook_with_attachment(to_addr: str, subject: str, html_body: str,
                                  attachment_bytes: bytes, attachment_name: str) -> str:
    """Create a draft in desktop Outlook (Windows) with attachment. Requires pywin32 & Outlook."""
    if sys.platform != "win32":
        return "Outlook automation is only available on Windows."
    try:
        import win32com.client as win32  # type: ignore
    except Exception:
        return "pywin32 not installed. Run: pip install pywin32"
    import tempfile
    tmpdir = tempfile.mkdtemp(prefix="offer_letter_")
    tmp_path = os.path.join(tmpdir, attachment_name)
    with open(tmp_path, "wb") as f:
        f.write(attachment_bytes)
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = to_addr
        mail.Subject = subject
        mail.HTMLBody = html_body
        mail.Attachments.Add(tmp_path)
        mail.Display()
        return "Opened Outlook with a new draft."
    except Exception as e:
        return f"Failed to open Outlook draft: {e}"

# ==============================
# THEME / UI
# ==============================
st.set_page_config(page_title="Offer Letter Generator", page_icon="📄", layout="centered")

with st.sidebar:
    st.subheader("🎨 Appearance")

    preset = st.selectbox(
        "Brand preset",
        ["Hogarth Worldwide", "Custom"],
        index=0,
        help="Pick Hogarth colors or switch to Custom to choose your own.",
    )

    HOGARTH = {
        "primary": "#FF527E",   # Wild Watermelon
        "accent":  "#27C79A",   # Shamrock
        "indigo":  "#4F51C9",   # Indigo
        "star":    "#DEF034",   # Starship
        "mirage":  "#191528",   # Mirage
    }

    if preset == "Hogarth Worldwide":
        primary = HOGARTH["primary"]
        accent  = HOGARTH["accent"]
        rounded = 16
        bg_style = "Hogarth gradient"
        bg_file = None
    else:
        primary = st.color_picker("Primary color", value="#2563EB")
        accent  = st.color_picker("Accent color", value="#10B981")
        rounded = st.slider("Corner radius (px)", 4, 24, 14, 1)
        bg_style = st.selectbox(
            "Background",
            ["Soft gradient", "Solid light", "Solid dark"],
            index=0,
        )
        bg_file = st.file_uploader(
            "Custom background image (optional)",
            type=["png", "jpg", "jpeg"],
            accept_multiple_files=False
        )

# Build background CSS
if preset == "Hogarth Worldwide":
    bg_css = (
        f"background: radial-gradient(1200px 600px at 85% -10%, {HOGARTH['star']}20, transparent 60%), "
        f"linear-gradient(135deg, {HOGARTH['mirage']} 0%, {HOGARTH['indigo']} 60%, {HOGARTH['mirage']} 100%);"
    )
else:
    if 'bg_file' in locals() and bg_file is not None:
        b64 = base64.b64encode(bg_file.read()).decode("utf-8")
        bg_css = f"background: url(data:image/png;base64,{b64}) center/cover fixed no-repeat;"
    else:
        if bg_style == "Soft gradient":
            bg_css = "background: linear-gradient(135deg, #f8fafc 0%, #eef2ff 45%, #fff7ed 100%);"
        elif bg_style == "Solid light":
            bg_css = "background: #f8fafc;"
        else:
            bg_css = "background: #0b1220;"

text_on_primary = "#0b1220"

st.markdown(
    f"""
    <style>
    .stApp {{ {bg_css} }}
    html, body, [class*="css"]  {{ -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; }}
    h1, h2, h3, .stMarkdown h1, .stMarkdown h2 {{ letter-spacing: .2px; color: {'#FFFFFF' if preset=='Hogarth Worldwide' else 'inherit'}; }}
    div.stButton > button {{ background: {primary}; color: {text_on_primary}; border: 0; padding: .7rem 1.1rem; border-radius: 16px; box-shadow: 0 6px 14px rgba(0,0,0,.10); transition: transform .02s, box-shadow .2s; }}
    div.stButton > button:hover {{ transform: translateY(-1px); box-shadow: 0 10px 24px rgba(0,0,0,.18); filter: brightness(1.02); }}
    div.stButton > button:disabled {{ opacity: .6; cursor: not-allowed; }}
    .stDownloadButton > button {{ background: {accent}; color: #0b1220; border: 0; padding: .7rem 1.1rem; border-radius: 16px; box-shadow: 0 6px 14px rgba(0,0,0,.10); }}
    .stTextInput > div > div > input, .stTextArea > div > div > textarea, .stDateInput > div > div input {{ border-radius: 16px !important; }}
    .block-container {{ padding-top: 2rem; padding-bottom: 3rem; color: {'#e6e7ee' if preset=='Hogarth Worldwide' else 'inherit'}; }}
    a, .egzxvld2 {{ color: {HOGARTH['star'] if preset=='Hogarth Worldwide' else primary}; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================
# APP CONTENT
# ==============================
st.title("📄 Offer Letter Generator")

uploaded_file = st.file_uploader("Upload PPTX Template", type=["pptx"])

col1, col2 = st.columns(2)
with col1:
    name = st.text_input("Candidate Name", placeholder="Jane Doe")
    position = st.text_input("Position", placeholder="Software Engineer")
    salary = st.text_input("Salary", placeholder="1500000")
with col2:
    join_date = st.date_input("Join Date", date.today())
    offer_date = st.date_input("Offer Date", date.today())
    city = st.text_input("City", value="Buenos Aires")

st.subheader("✉️ Email (optional)")
to_email = st.text_input("Recipient email (To)", placeholder="candidate@example.com")
subject = st.text_input("Subject", value="Offer Letter")
body_html = st.text_area(
    "Email body (HTML supported)",
    value=(
        "<p>Hi,</p>"
        "<p>Please find attached your offer letter. Let us know if you have any questions.</p>"
        "<p>Best regards,<br/>HR</p>"
    ),
    height=140,
)

extras = []  # hook for additional placeholders if you add a table editor

c1, c2 = st.columns([1, 1])
with c1:
    generate_clicked = st.button("⚙️ Generate Offer Letter", type="primary", disabled=(uploaded_file is None))
with c2:
    clear_clicked = st.button("🧹 Clear fields")

if clear_clicked:
    st.experimental_rerun()

if generate_clicked:
    if not uploaded_file:
        st.warning("Please upload a PPTX template first.")
    else:
        mapping = {
            "CANDIDATE_NAME": name,
            "POSITION": position,
            "SALARY": salary,
            "JOIN_DATE": fecha_es(join_date),
            "DATE": fecha_es(offer_date),
            "CITY": city,
        }
        for row in extras:
            k = (row.get("key") or "").strip()
            v = (row.get("value") or "").strip()
            if k:
                mapping[k.upper()] = v
        try:
            edited = render_pptx(uploaded_file.read(), mapping)
            safe_name = " ".join((name or "").strip().split())
            file_name_out = f"Offer Letter - {safe_name}.pptx" if safe_name else "Offer Letter.pptx"

            st.session_state["offer_bytes"] = edited.getvalue()
            st.session_state["offer_name"] = file_name_out

            st.download_button(
                "⬇️ Download Updated PPTX",
                data=st.session_state["offer_bytes"],
                file_name=file_name_out,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            st.success(f"Done! Generated {file_name_out} with all placeholders replaced.")
        except Exception as e:
            st.exception(e)

st.markdown("---")
st.subheader("Send via Outlook (Windows desktop)")
if sys.platform != "win32":
    st.info("Outlook automation is only available on Windows. You can still download the file and send it manually.")
else:
    disabled = not (st.session_state.get("offer_bytes") and to_email)
    if st.button("📧 Open Outlook draft with attachment", disabled=disabled):
        if not st.session_state.get("offer_bytes"):
            st.warning("Generate the offer first.")
        elif not to_email:
            st.warning("Enter a recipient email.")
        else:
            msg = _open_outlook_with_attachment(
                to_addr=to_email,
                subject=subject,
                html_body=body_html,
                attachment_bytes=st.session_state["offer_bytes"],
                attachment_name=st.session_state.get("offer_name", "Offer Letter.pptx"),
            )
            if msg.startswith("Opened Outlook"):
                st.success(msg)
            else:
                st.error(msg)
