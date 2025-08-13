import streamlit as st
from pptx import Presentation
from io import BytesIO

st.title("📄 Offer Letter Generator")

uploaded_template = st.file_uploader("Upload PPTX Template", type=["pptx"])

name = st.text_input("Candidate Name", "John Doe")
position = st.text_input("Position", "Software Engineer")
salary = st.text_input("Salary", "$100,000")

if uploaded_template:
    if st.button("Generate Offer Letter"):
        prs = Presentation(uploaded_template)
        replacements = {
            "{{NAME}}": name,
            "{{POSITION}}": position,
            "{{SALARY}}": salary,
        }

        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for key, val in replacements.items():
                        if key in shape.text:
                            shape.text = shape.text.replace(key, val)

        output = BytesIO()
        prs.save(output)
        st.download_button("Download Offer Letter", output.getvalue(), file_name="offer_letter.pptx")
