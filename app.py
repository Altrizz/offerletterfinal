# app.py — Streamlit (Cloud‑ready) with robust PPTX replacement
# - Replaces placeholders even when split across runs/tables/grouped shapes
# - Supports both {{CURLY}} tokens and legacy X‑style tokens
# - Adds missing fields: join date, offer date, city, salary formatting

import re
import os
from datetime import date
from io import BytesIO

import streamlit as st
from pptx import Presentation
from pptx.shapes.group import GroupShape

# ------------------------------
# Helpers: formatting and tokens
# ------------------------------
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

# Legacy X‑style tokens from the provided template
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
    # First resolve {{KEY}} placeholders, then legacy X tokens
    def repl(m):
        key = m.group(1).upper()
        return str(mapping.get(key, m.group(0)))
    text2 = PLACEHOLDER.sub(repl, text)
    return apply_x_style(text2, mapping)

# ------------------------------
# PPTX replacement across runs / tables / groups
# ------------------------------

def _replace_in_text_frame(tf, mapping: dict):
    for para in tf.paragraphs:
        if not para.runs:
            # Empty paragraph — still handle legacy line token
            if para.text:
                new = replace_placeholders_in_text(para.text, mapping)
                if new != para.text:
                    para.text = new
            continue
        full = "".join(run.text for run in para.runs)
        new = replace_placeholders_in_text(full, mapping)
        if new == full:
            continue
        # Put the new text in the first run, clear the rest to avoid split tokens
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

# ------------------------------
# Streamlit UI
# ------------------------------
st.set_page_config(page_title="Offer Letter Generator", page_icon="📄", layout="centered")
st.title("Offer Letter Generator")

uploaded_file = st.file_uploader("Upload PPTX Template", type=["pptx"], help="Use placeholders like {{CANDIDATE_NAME}} or legacy X‑tokens")

col1, col2 = st.columns(2)
with col1:
    name = st.text_input("Candidate Name")
    position = st.text_input("Position")
    salary = st.text_input("Salary (ARS)", value="1500000")
with col2:
    offer_date = st.date_input("Offer Date", value=date.today())
    join_date = st.date_input("Join Date", value=date.today())
    city = st.text_input("City", value="Buenos Aires")

st.subheader("Extra placeholders (optional)")
extras = st.data_editor([{"key": "", "value": ""}], num_rows="dynamic", hide_index=True)

if st.button("Generate Offer Letter", disabled=(uploaded_file is None)):
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
            st.download_button(
                "Download Updated PPTX",
                data=edited.getvalue(),
                file_name="Offer_Letter_Filled.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            st.success("Done! Placeholders replaced across text boxes and tables.")
        except Exception as e:
            st.exception(e)

# ------------------------------
# Self‑tests (so we don't ship regressions)
# ------------------------------
with st.expander("Run self tests"):
    def _t_apply_x():
        m = {
            "CANDIDATE_NAME": "Juan Perez",
            "POSITION": "Software Engineer",
            "JOIN_DATE": "22 de agosto de 2025",
            "SALARY": "1500000",
            "DATE": "08 de agosto de 2025",
            "CITY": "Buenos Aires",
        }
        s = (
            "Hola {XXXXXX}\n"
            "Cargo: XXXXXXXX\n"
            "Fecha: XX de XXXXX de 2025\n"
            "Salario: X.XXX.XXX\n"
            ", Buenos Aires\n"
        )
        out = apply_x_style(s, m)
        assert "Juan Perez" in out
        assert "Software Engineer" in out
        assert "22 de agosto de 2025" in out
        assert "1.500.000" in out
        assert "08 de agosto de 2025, Buenos Aires" in out

    def _t_curly_simple():
        m = {"NAME": "Ana", "CITY": "BA"}
        txt = "Hola {{ NAME }}, bienvenid@ a {{CITY}}"
        out = replace_placeholders_in_text(txt, m)
        assert out == "Hola Ana, bienvenid@ a BA"

    run = st.button("Run tests")
    if run:
        try:
            _t_apply_x(); _t_curly_simple()
            st.success("OK — tests passed.")
        except AssertionError as e:
            st.error(f"Test failed: {e}")

"""
requirements.txt (make sure this is in your repo root):

streamlit
python-pptx
"""

