import streamlit as st
from io import BytesIO
from datetime import date
import tempfile
import os
import sys
import base64

# Assumes render_pptx(name->pptx bytes) and fecha_es(date->"DD de mes de YYYY") exist in this file or are imported.
# If they live elsewhere in your repo, import them accordingly.

# ------------------------------
# Page / Theme Controls (with Hogarth preset)
# ------------------------------
st.set_page_config(page_title="Offer Letter Generator", page_icon="📄", layout="centered")

with st.sidebar:
    st.subheader("🎨 Appearance")

    preset = st.selectbox(
        "Brand preset",
        ["Hogarth Worldwide", "Custom"],
        index=0,
        help="Pick Hogarth colors or switch to Custom to choose your own.",
    )

    # Hogarth palette
    HOGARTH = {
        "primary": "#FF527E",   # Wild Watermelon
        "accent":  "#27C79A",   # Shamrock
        "indigo":  "#4F51C9",   # Indigo
        "star":    "#DEF034",   # Starship
        "mirage":  "#191528",   # Mirage (deep background)
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
            [
                "Soft gradient",
                "Solid light",
                "Solid dark",
            ],
            index=0,
        )
        bg_file = st.file_uploader("Custom background image (optional)", type=["png", "jpg", "jpeg"], accept_multiple_files=False)

# Build background CSS
if preset == "Hogarth Worldwide":
    # Layered gradients using brand colors (Mirage base, Indigo overlay, Starship accent glow)
    bg_css = (
        f"background: radial-gradient(1200px 600px at 85% -10%, {HOGARTH['star']}20, transparent 60%), "
        f"linear-gradient(135deg, {HOGARTH['mirage']} 0%, {HOGARTH['indigo']} 60%, {HOGARTH['mirage']} 100%);"
    )
else:
    if bg_file is not None:
        b64 = base64.b64encode(bg_file.read()).decode("utf-8")
        bg_css = f"background: url(data:image/png;base64,{b64}) center/cover fixed no-repeat;"
    else:
        if bg_style == "Soft gradient":
            bg_css = "background: linear-gradient(135deg, #f8fafc 0%, #eef2ff 45%, #fff7ed 100%);"
        elif bg_style == "Solid light":
            bg_css = "background: #f8fafc;"
        else:
            bg_css = "background: #0b1220;"

# Accessible text color for dark backgrounds
text_on_primary = "#0b1220" if preset == "Hogarth Worldwide" else "#0b1220"

# Inject global CSS theme
st.markdown(
    f"""
    <style>
    /* Page background */
    .stApp {{ {bg_css} }}

    /* Fonts + smoothing */
    html, body, [class*="css"]  {{
      -webkit-font-smoothing: antialiased;
      -moz-osx-font-smoothing: grayscale;
    }}

    /* Headers */
    h1, h2, h3, .stMarkdown h1, .stMarkdown h2 {{
      letter-spacing: .2px;
      color: {'#FFFFFF' if preset=='Hogarth Worldwide' else 'inherit'};
    }}

    /* Primary buttons */
    div.stButton > button {{
      background: {primary};
      color: {text_on_primary};
      border: 0;
      padding: 0.7rem 1.1rem;
      border-radius: {rounded}px;
      box-shadow: 0 6px 14px rgba(0,0,0,.10);
      transition: transform .02s ease-in-out, box-shadow .2s ease;
    }}
    div.stButton > button:hover {{
      transform: translateY(-1px);
      box-shadow: 0 10px 24px rgba(0,0,0,.18);
      filter: brightness(1.02);
    }}
    div.stButton > button:disabled {{
      opacity: .6; cursor: not-allowed;
    }}

    /* Download button */
    .stDownloadButton > button {{
      background: {accent};
      color: #0b1220;
      border: 0;
      padding: 0.7rem 1.1rem;
      border-radius: {rounded}px;
      box-shadow: 0 6px 14px rgba(0,0,0,.10);
    }}

    /* Inputs */
    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea,
    .stDateInput > div > div input {{
      border-radius: {rounded}px !important;
    }}

    /* Cards */
    .block-container {{
      padding-top: 2rem; padding-bottom: 3rem;
      color: {'#e6e7ee' if preset=='Hogarth Worldwide' else 'inherit'};
    }}

    /* Links */
    a, .egzxvld2 {{ color: {HOGARTH['star'] if preset=='Hogarth Worldwide' else primary}; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ------------------------------
# App Content
# ------------------------------
st.title("📄 Offer Letter Generator")

uploaded_file = st.file_uploader("Upload PPTX Template", type=["pptx"])

# Input grid
col1, col2 = st.columns(2)
with col1:
    name = st.text_input("Candidate Name", placeholder="Jane Doe")
    position = st.text_input("Position", placeholder="Software Engineer")
    salary = st.text_input("Salary", placeholder="1500000")
with col2:
    join_date = st.date_input("Join Date", date.today())
    offer_date = st.date_input("Offer Date", date.today())
    city = st.text_input("City", value="Buenos Aires")

# Email compose inputs
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

extras = []  # Add logic if you have additional placeholders

# Action row
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

# Outlook compose button (Windows only)
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
