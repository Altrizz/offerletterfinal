# app.py — Streamlit (Cloud) with robust PPTX replacement
# Branding: White background + subtle Hogarth split band across the middle
# Email features removed

import re
from datetime import date
from io import BytesIO

import streamlit as st
from pptx import Presentation
from pptx.shapes.group import GroupShape

# ==============================
# Helpers: dates, formatting, tokens
# ==============================
MESES_ES = [
    "enero","febrero","marzo","abril","mayo","junio",
    "julio","agosto","septiembre","octubre","noviembre","diciembre",
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
# THEME — White bg + Hogarth split band
# ==============================
HOGARTH = {
    "primary": "#FF527E",  # Wild Watermelon
    "accent":  "#27C79A",  # Shamrock
    "indigo":  "#4F51C9",  # Indigo
    "star":    "#DEF034",  # Starship
    "mirage":  "#191528",  # Mirage
}

st.set_page_config(page_title="Offer Letter Generator", page_icon="📄", layout="centered")

st.markdown(
    f"""
    <style>
      .stApp {{ background:#ffffff; }}
      /* Subtle mid-page 'Hogarth split' band */
      .stApp::before {{
        content:"";
        position: fixed; left:0; right:0; top:50%;
        transform: translateY(-50%);
        height: 240px;      /* thickness of the band */
        background:
          radial-gradient(800px 240px at 85% 0%, {HOGARTH['star']}22, transparent 60%),
          linear-gradient(135deg, {HOGARTH['indigo']} 0%, {HOGARTH['mirage']} 100%);
        opacity: .12;       /* keep it subtle */
        pointer-events: none;
        z-index: 0;
      }}
      .block-container {{ position: relative; z-index: 1; }}

      /* Buttons */
      div.stButton > button {{
        background: {HOGARTH['primary']};
        color: #0b1220;
        border: 0; padding: .7rem 1.1rem; border-radius: 14px;
        box-shadow: 0 6px 14px rgba(0,0,0,.08);
        transition: transform .02s, box-shadow .2s;
      }}
      div.stButton > button:hover {{
        transform: translateY(-1px);
        box-shadow: 0 10px 20px rgba(0,0,0,.12);
      }}
      .stDownloadButton > button {{
        background: {HOGARTH['accent']};
        color: #0b1220; border:0; padding:.7rem 1.1rem; border-radius:14px;
        box-shadow: 0 6px 14px rgba(0,0,0,.08);
      }}

      /* Inputs rounding */
      .stTextInput > div > div > input,
      .stDateInput > div > div input {{ border-radius: 14px !important; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================
# APP CONTENT (no email)
# ==============================
st.title("📄 Offer Letter Generator")

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
