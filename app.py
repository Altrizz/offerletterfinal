# app.py — Streamlit Cloud ready
# Default Streamlit theme, cleaner inputs, robust PPTX replacement
# + Session History: replay, download, delete, export-zip

import re
import os
import io
import base64
from pathlib import Path
from datetime import date, datetime
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED

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
# Session History utils
# ------------------------------
HISTORY_KEY = "offer_history"
MAX_HISTORY = 50

def ensure_history():
    if HISTORY_KEY not in st.session_state:
        st.session_state[HISTORY_KEY] = []  # list of dicts

def push_history(entry: dict):
    ensure_history()
    st.session_state[HISTORY_KEY].append(entry)
    # cap size
    if len(st.session_state[HISTORY_KEY]) > MAX_HISTORY:
        st.session_state[HISTORY_KEY] = st.session_state[HISTORY_KEY][-MAX_HISTORY:]

def delete_history(idx: int):
    ensure_history()
    if 0 <= idx < len(st.session_state[HISTORY_KEY]):
        st.session_state[HISTORY_KEY].pop(idx)

def zip_all_history() -> bytes:
    ensure_history()
    mem = io.BytesIO()
    with ZipFile(mem, "w", ZIP_DEFLATED) as zf:
        for i, item in enumerate(st.session_state[HISTORY_KEY], start=1):
            fname = item["file_name"]
            zf.writestr(fname, item["pptx_bytes"])
    mem.seek(0)
    return mem.getvalue()

# ------------------------------
# UI
# ------------------------------
st.set_page_config(page_title="Offer Letter Generator", page_icon="📄", layout="centered")
st.title("Offer Letter Generator")

uploaded_file = st.file_uploader(
    "Upload PPTX Template",
    type=["pptx"],
    help="Use placeholders like {{CANDIDATE_NAME}}, {{FIRST_NAME}}, {{LAST_NAME}}, {{POSITION}}, "
         "{{SALARY}}, {{JOIN_DATE}}, {{DATE}}, {{CITY}} or the legacy X tokens used in your template."
)

# Inputs
colA, colB = st.columns(2)
with colA:
    first_name = st.text_input("First name", placeholder="Jane")
    position   = st.text_input("Position", placeholder="Software Engineer")
    salary_num = st.number_input("Salary (ARS)", min_value=0, value=2000000, step=5000)
with colB:
    last_name  = st.text_input("Last name", placeholder="Doe")
    offer_date = st.date_input("Offer date", value=date.today())
    join_date  = st.date_input("Join date",  value=date.today())

common_cities = ["Buenos Aires", "Córdoba", "Rosario", "Mendoza", "Other..."]
city_pick = st.selectbox("City", options=common_cities, index=0)
city = st.text_input("City (custom)", value="", placeholder="Type city") if city_pick == "Other..." else city_pick

# Extras mapping
st.subheader("Extra placeholders (optional)")
extras = st.data_editor([{"key": "", "value": ""}], num_rows="dynamic", hide_index=True)

# Actions
col1, col2 = st.columns(2)
with col1:
    generate_clicked = st.button("Generate Offer Letter", type="primary", disabled=(uploaded_file is None))
with col2:
    if st.button("Clear fields"):
        st.experimental_rerun()

# Generate
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
        extra_pairs = []
        for row in extras:
            k = (row.get("key") or "").strip()
            v = (row.get("value") or "").strip()
            if k:
                mapping[k.upper()] = v
                extra_pairs.append((k.upper(), v))

        try:
            pptx_bytes = uploaded_file.read()
            edited = render_pptx(pptx_bytes, mapping)
            safe_name = " ".join(full_name.split()) or "Offer Letter"
            file_name_out = f"Offer Letter - {safe_name}.pptx"

            # Offer for immediate download
            st.download_button(
                "Download Updated PPTX",
                data=edited.getvalue(),
                file_name=file_name_out,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            st.success(f"Done! Generated {file_name_out}.")

            # Save to session history
            push_history({
                "ts": datetime.now(),
                "file_name": file_name_out,
                "pptx_bytes": edited.getvalue(),
                "fields": {
                    "first_name": first_name, "last_name": last_name, "position": position,
                    "salary_num": int(salary_num), "offer_date": offer_date.isoformat(),
                    "join_date": join_date.isoformat(), "city": city,
                },
                "extras": extra_pairs,
            })
        except Exception as e:
            st.exception(e)

st.divider()

# ------------------------------
# History panel
# ------------------------------
ensure_history()
st.subheader(f"History (this session) — {len(st.session_state[HISTORY_KEY])} item(s)")

if st.session_state[HISTORY_KEY]:
    # Export all as ZIP
    zip_bytes = zip_all_history()
    st.download_button(
        "Export all offers (ZIP)",
        data=zip_bytes,
        file_name="offer_letters_history.zip",
        mime="application/zip",
        type="secondary",
    )

    # Show newest first
    for idx, item in reversed(list(enumerate(st.session_state[HISTORY_KEY]))):
        meta = item["fields"]
        ts = item["ts"].strftime("%Y-%m-%d %H:%M")
        label = f"🗂️ {item['file_name']} · {meta['position']} · {meta['city']} · {ts}"
        with st.expander(label, expanded=False):
            cA, cB, cC = st.columns([1,1,1])
            with cA:
                st.write(f"**Candidate:** {meta['first_name']} {meta['last_name']}")
                st.write(f"**Position:** {meta['position']}")
                st.write(f"**City:** {meta['city']}")
            with cB:
                st.write(f"**Offer date:** {fecha_es(date.fromisoformat(meta['offer_date']))}")
                st.write(f"**Join date:** {fecha_es(date.fromisoformat(meta['join_date']))}")
                st.write(f"**Salary (ARS):** {format_ars_dots(meta['salary_num'])}")
            with cC:
                if item["extras"]:
                    st.write("**Extras:**")
                    for k, v in item["extras"]:
                        st.write(f"- {k}: {v}")

            d1, d2, d3 = st.columns([1,1,1])
            with d1:
                st.download_button(
                    "Download again",
                    data=item["pptx_bytes"],
                    file_name=item["file_name"],
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key=f"dl_{idx}",
                )
            with d2:
                if st.button("Restore to form", key=f"restore_{idx}"):
                    # Repopulate inputs and rerun
                    st.session_state["First name"] = meta["first_name"]
                    st.session_state["Last name"]  = meta["last_name"]
                    st.session_state["Position"]   = meta["position"]
                    st.session_state["Salary (ARS)"] = meta["salary_num"]
                    st.session_state["Offer date"] = date.fromisoformat(meta["offer_date"])
                    st.session_state["Join date"]  = date.fromisoformat(meta["join_date"])
                    st.session_state["City"]       = meta["city"]
                    # Recreate extras grid rows
                    st.session_state["_extras_prefill"] = [{"key": k, "value": v} for k, v in item["extras"]]
                    st.experimental_rerun()
            with d3:
                if st.button("Delete", key=f"del_{idx}"):
                    delete_history(idx)
                    st.experimental_rerun()
else:
    st.info("No offers generated yet. When you create one, it will appear here for quick download and restore.")

# Prefill extras after restore, if present
if "_extras_prefill" in st.session_state:
    st.session_state.pop("_extras_prefill")  # consumed; editor cannot be updated post-render

# Footer logo (optional)
footer_logo()
