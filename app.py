# app.py - Streamlit Cloud ready
# Default Streamlit background, cleaned inputs, robust PPTX placeholder replacement
# - Supports {{CURLY}} tokens and legacy X-style tokens
# - Replaces even when tokens are split across runs, tables, grouped shapes
# - File name: "Offer Letter - <First> <Last>.pptx"
# - Optional footer logo: add "hogarth_split_black.png" (or hogarth_split.png) to repo root

import re
import os
import base64
from pathlib import Path
from datetime import date
from io import BytesIO

import streamlit as st
from pptx import Presentation
from pptx.shapes.group import GroupShape

# ------------------------------
# Helpers: dates, salary format
# ------------------------------
MESES_ES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]

def fecha_es(d: date) -> str:
    return f"{d.day} de {MESES_ES[d.month-1]} de {d.year}"

def format_ars_dots(value) -> str:
    # accepts int/str, returns ARS with dot thousands
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
    # First resolve {{KEY}} placeholders, then legacy X tokens
    def repl(m):
        key = m.group(1).upper()
        return str(mapping.get(key, m.group(0)))
    return apply_x_style(PLACEHOLDER.sub(repl, text), mapping)

# ------------------------------
# PPTX replacement across runs / tables / groups
# ------------------------------
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

# ------------------------------
# Optional footer logo
# ------------------------------
def footer_logo():
    for name in ("hogarth_split_black.png", "hogarth_split.png"):
        p = Path(__file__).with_name(name)
        if p.exists():
            b64 = base64.b64encode(p.read_bytes()).decode("utf-8")
            st.markdown(
                f"<div style='display:flex;justify-content:center;margin-top:36px'>"
                f"<img src='data:image/png;base64,{b64}' style='max-width:520px;width:60%;height:auto' />"
                f"</div>",
                unsafe_allow_html=True,
            )
            break

# ------------------------------
# Token detector (debug helper)
# ------------------------------
def detect_tokens(pptx_bytes: bytes):
    prs = Presentation(BytesIO(pptx_bytes))
    found_curly = set()
    x_name = x_pos = x_date = x_sal = False

    def scan_text(text: str):
        nonlocal x_name, x_pos, x_date, x_sal
        for m in PLACEHOLDER.finditer(text or ""):
            found_curly.add(m.group(1).upper())
        if PAT_NAME.search(text or ""): x_name = True
        if PAT_POS.search(text or ""):  x_pos  = True
        if PAT_DATE.search(text or ""): x_date = True
        if PAT_SAL.search(text or ""):  x_sal  = True

    def visit_shapes(shapes):
        for shape in shapes:
            if getattr(shape, "has_text_frame", False) and shape.text_frame:
                for para in shape.text_frame.paragraphs:
                    if para.runs:
                        scan_text("".join(run.text for run in para.runs))
                    else:
                        scan_text(para.text)
            if getattr(shape, "has_table", False):
                for r in shape.table.rows:
                    for c in r.cells:
                        if c.text_frame:
                            for para in c.text_frame.paragraphs:
                                if para.runs:
                                    scan_text("".join(run.text for run in para.runs))
                                else:
                                    scan_text(para.text)
            if isinstance(shape, GroupShape):
                visit_shapes(shape.shapes)

    for slide in prs.slides:
        visit_shapes(slide.shapes)

    legacy = [t for t, ok in {
        "NAME {XXXXXX}": x_name, "POSITION XXXXXXXX": x_pos,
        "JOIN_DATE pattern": x_date, "SALARY X.XXX.XXX": x_sal
    }.items() if ok]
    return found_curly, legacy

# ------------------------------
# UI - back to Streamlit defaults
# ------------------------------
st.set_page_config(page_title="Offer Letter Generator", page_icon="📄", layout="centered")
st.title("Offer Letter Generator")

uploaded_file = st.file_uploader(
    "Upload PPTX Template",
    type=["pptx"],
    help="Use placeholders like {{CANDIDATE_NAME}}, {{FIRST_NAME}}, {{LAST_NAME}}, {{POSITION}}, "
         "{{SALARY}}, {{JOIN_DATE}}, {{DATE}}, {{CITY}} or the legacy X tokens used in your template."
)

# Better field options
colA, colB = st.columns(2)
with colA:
    first_name = st.text_input("First name", placeholder="Jane")
    position   = st.text_input("Position", placeholder="Software Engineer")
    salary_num = st.number_input("Salary (ARS)", min_value=0, value=2000000, step=5000)
with colB:
    last_name  = st.text_input("Last name", placeholder="Doe")
    offer_date = st.date_input("Offer date", value=date.today())
    join_date  = st.date_input("Join date",  value=date.today())

# City selector with custom option
common_cities = ["Buenos Aires", "Córdoba", "Rosario", "Mendoza", "Other..."]
city_pick = st.selectbox("City", options=common_cities, index=0)
city = st.text_input("City (custom)", value="", placeholder="Type city") if city_pick == "Other..." else city_pick

# Show formatted previews
st.caption(
    f"Preview - Candidate: **{(first_name + ' ' + last_name).strip()}**, "
    f"Salary formatted: **{format_ars_dots(salary_num)}**, "
    f"Offer date: **{fecha_es(offer_date)}**, Join date: **{fecha_es(join_date)}**, City: **{city}**"
)

# Extras mapping
st.subheader("Extra placeholders (optional)")
extras = st.data_editor([{"key": "", "value": ""}], num_rows="dynamic", hide_index=True)

# Token detector
if uploaded_file:
    curly, legacy = detect_tokens(uploaded_file.read())
    st.info(
        "Placeholders detected in template: " +
        (", ".join(sorted(curly)) if curly else "none") +
        (" | Legacy tokens: " + ", ".join(legacy) if legacy else "")
    )
    uploaded_file.seek(0)  # rewind for later read

# Actions
col1, col2 = st.columns(2)
with col1:
    generate_clicked = st.button("Generate Offer Letter", type="primary", disabled=(uploaded_file is None))
with col2:
    if st.button("Clear fields"):
        st.experimental_rerun()

if generate_clicked:
    if not uploaded_file:
        st.warning("Please upload a PPTX template first.")
    else:
        full_name = f"{first_name} {last_name}".strip()
        mapping = {
            "CANDIDATE_NAME": full_name,
            "FIRST_NAME": first_name,
            "LAST_NAME": last_name,
            "POSITION": position,
            "SALARY": format_ars_dots(salary_num),
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
            safe_name = " ".join(full_name.split()) or "Offer Letter"
            file_name_out = f"Offer Letter - {safe_name}.pptx"

            st.download_button(
                "Download Updated PPTX",
                data=edited.getvalue(),
                file_name=file_name_out,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            st.success(f"Done! Generated {file_name_out}.")
        except Exception as e:
            st.exception(e)

# Optional footer logo if present
footer_logo()
