# app.py — modern UI + robust PPTX replacement (Streamlit Cloud ready)
# - White / soft gradient background, glass card, refined inputs & buttons
# - Replaces placeholders across split runs, tables, grouped shapes
# - Supports {{CURLY}} tokens and legacy X-style tokens
# - Output filename: "Offer Letter - <Candidate Name>.pptx"
# - Optional footer logo: add "hogarth_split_black.png" to repo root

import re
import os
from datetime import date
from io import BytesIO

import streamlit as st
from pptx import Presentation
from pptx.shapes.group import GroupShape

# ==============================
# THEME COLORS (tweak freely)
# ==============================
PRIMARY = "#FF527E"   # Wild Watermelon
ACCENT  = "#27C79A"   # Shamrock
BG1     = "#ffffff"   # background top
BG2     = "#f6f7fb"   # background bottom
TEXT    = "#0b1220"   # main text color
RADIUS  = 14          # corner radius

# ==============================
# Helpers: dates, formatting, tokens
# ==============================
MESES_ES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]

def fecha_es(d: date) -> str:
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

    # lines like ", Buenos Aires" -> "<DATE>, <CITY>"
    striped = out.strip()
    if striped == ", Buenos Aires" or striped.endswith(", Buenos Aires"):
        out = f"{offer_date_es}, {city}" if offer_date_es else city
    return out

def replace_placeholders_in_text(text: str, mapping: dict) -> str:
    def repl(m):
        key = m.group(1).upper()
        return str(mapping.get(key, m.group(0)))
    return apply_x_style(PLACEHOLDER.sub(repl, text), mapping)

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
        if getattr(shape, "has_text_frame", False) and shape.text_frame:
            _replace_in_text_frame(shape.text_frame, mapping)
        if getattr(shape, "has_table", False):
            _replace_in_table(shape.table, mapping)
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
# PAGE SETUP + STYLES
# ==============================
st.set_page_config(page_title="Offer Letter Generator", page_icon="📄", layout="centered")

st.markdown(f"""
<style>
:root {{
  --primary: {PRIMARY};
  --accent:  {ACCENT};
  --text:    {TEXT};
  --radius:  {RADIUS}px;
}}
/* Background */
.stApp {{
  background: linear-gradient(180deg, {BG1} 0%, {BG2} 100%);
  color: var(--text);
}}
/* Center column width & spacing */
.block-container {{
  max-width: 980px;
  padding-top: 2.5rem;
  padding-bottom: 4rem;
}}
/* Glass card container */
.card {{
  background: rgba(255,255,255,0.85);
  backdrop-filter: blur(6px);
  border: 1px solid rgba(12,12,32,0.06);
  border-radius: var(--radius);
  padding: 1.4rem 1.6rem;
  box-shadow: 0 10px 30px rgba(12,12,32,.06);
}}
/* Headings */
h1, h2, h3 {{ color: var(--text); letter-spacing: .2px; }}
h1.title-gradient {{
  background: linear-gradient(90deg, var(--text) 0%, #3b3b55 40%, #7a7a90 100%);
  -webkit-background-clip: text; background-clip: text; color: transparent;
  font-weight: 800;
}}
/* Inputs */
.stTextInput > div > div > input,
.stDateInput > div > div input {{
  border-radius: var(--radius) !important;
  border: 1px solid rgba(12,12,32,0.12);
  background: #fff;
}}
/* Data editor */
[data-testid="stDataEditor"] .st-dx, [data-testid="stDataEditor"] .st-dk {{
  border-radius: var(--radius);
}}
/* Buttons */
div.stButton > button {{
  border: 0;
  border-radius: var(--radius);
  padding: .75rem 1.15rem;
  font-weight: 600;
  box-shadow: 0 10px 28px rgba(12,12,32,.10);
  transition: transform .02s ease, box-shadow .2s ease, filter .2s ease;
}}
div.stButton > button[kind="primary"] {{
  background: var(--primary);
  color: #0b1220;
}}
div.stButton > button:hover {{ transform: translateY(-1px); filter: brightness(1.02); }}
/* Download button */
.stDownloadButton > button {{
  background: var(--accent);
  color: #0b1220;
  border: 0;
  border-radius: var(--radius);
  padding: .75rem 1.15rem;
  font-weight: 600;
  box-shadow: 0 10px 28px rgba(12,12,32,.10);
}}
/* Footer logo area */
.footer-logo {{
  display: flex; align-items: center; justify-content: center;
  margin-top: 48px; opacity: .9;
}}
.footer-logo img {{
  max-width: 520px; width: 50%; height: auto;
  filter: contrast(110%);
}}
</style>
""", unsafe_allow_html=True)

# ==============================
# APP CONTENT (wrapped in a "card")
# ==============================
st.markdown("<h1 class='title-gradient'>Offer Letter Generator</h1>", unsafe_allow_html=True)
st.markdown("<div class='card'>", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Upload PPTX Template",
    type=["pptx"],
    help="Use placeholders like {{CANDIDATE_NAME}}, {{POSITION}}, {{SALARY}}, {{JOIN_DATE}}, {{DATE}}, {{CITY}} "
         "or the legacy X tokens used in your template."
)

col1, col2 = st.columns(2)
with col1:
    name = st.text_input("Candidate Name", placeholder="Jane Doe")
    position = st.text_input("Position", placeholder="Software Engineer")
    salary = st.text_input("Salary (ARS)", placeholder="1500000")
with col2:
    offer_date = st.date_input("Offer Date", value=date.today())
    join_date  = st.date_input("Join Date",  value=date.today())
    city       = st.text_input("City", value="Buenos Aires")

st.subheader("Extra placeholders (optional)")
extras = st.data_editor([{"key": "", "value": ""}], num_rows="dynamic", hide_index=True)

c1, c2 = st.columns([1, 1])
with c1:
    generate_clicked = st.button("⚙️ Generate Offer Letter", type="primary", disabled=(uploaded_file is None))
with c2:
    clear_clicked = st.button("🧹 Clear fields")

st.markdown("</div>", unsafe_allow_html=True)  # close .card

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
            "DATE":      fecha_es(offer_date),
            "CITY":      city,
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

            st.download_button(
                "⬇️ Download Updated PPTX",
                data=edited.getvalue(),
                file_name=file_name_out,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            st.success(f"Done! Generated {file_name_out} with all placeholders replaced.")
        except Exception as e:
            st.exception(e)

# ==============================
# Optional Hogarth footer logo
# ==============================
logo_candidates = ["hogarth_split_black.png", "hogarth_split.png"]
logo_path = next((p for p in logo_candidates if os.path.exists(p)), None)
if logo_path:
    st.markdown("<div class='footer-logo'>", unsafe_allow_html=True)
    st.image(logo_path)
    st.markdown("</div>", unsafe_allow_html=True)
