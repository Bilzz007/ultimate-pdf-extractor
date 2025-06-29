import streamlit as st
import pdfplumber
import pandas as pd
from docx import Document
import io
import os
from PIL import Image

st.set_page_config(page_title="Ultimate PDF Data Extractor", layout="wide")
st.title("Ultimate PDF Data Extractor")

# Ensure feedback directory exists
os.makedirs("feedback", exist_ok=True)

uploaded_file = st.file_uploader("Upload a PDF", type="pdf")

if uploaded_file:
    st.success("PDF uploaded! Processing...")
    with pdfplumber.open(uploaded_file) as pdf:
        all_text = ""
        tables = []
        images = []

        for page_num, page in enumerate(pdf.pages, 1):
            # Extract text
            text = page.extract_text() or ""
            all_text += f"\n--- Page {page_num} ---\n{text}"
            
            # Extract tables
            page_tables = page.extract_tables() or []
            tables.extend(page_tables)

            # Extract images
            for img_obj in page.images:
                try:
                    im = page.to_image(resolution=150)
                    cropped = im.crop((img_obj["x0"], img_obj["top"], img_obj["x1"], img_obj["bottom"]))
                    images.append(cropped.original)
                except Exception as e:
                    st.warning(f"Image extraction failed on page {page_num}: {e}")

    # --- Display and edit extracted text ---
    st.subheader("Extracted Text (editable)")
    text_area = st.text_area("Edit extracted text:", all_text, height=300)

    # --- Display and edit extracted tables ---
    st.subheader("Extracted Tables")
    edited_tables = []
    if tables:
        for i, table in enumerate(tables):
            try:
                df = pd.DataFrame(table[1:], columns=table[0])
            except Exception:
                # Fallback: treat all as strings
                df = pd.DataFrame(table)
            edited_df = st.data_editor(df, num_rows="dynamic", key=f"table_{i}")
            edited_tables.append(edited_df)
            st.write("---")
    else:
        st.info("No tables found in this PDF.")

    # --- Display extracted images ---
    st.subheader("Extracted Images")
    if images:
        for idx, img in enumerate(images):
            st.image(img, caption=f"Image {idx+1}")
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            st.download_button(
                label=f"Download Image {idx+1}",
                data=buf.getvalue(),
                file_name=f"pdf_image_{idx+1}.png",
                mime="image/png"
            )
    else:
        st.info("No images found in this PDF.")

    # --- Export options ---
    st.subheader("Export Extracted Data")

    # Export to Word (.docx)
    doc_buf = io.BytesIO()
    if st.button("Export to Word (.docx)"):
        doc = Document()
        doc.add_heading("Extracted PDF Text", level=1)
        doc.add_paragraph(text_area)
        if edited_tables:
            for idx, df in enumerate(edited_tables):
                doc.add_heading(f"Table {idx+1}", level=2)
                t = doc.add_table(rows=df.shape[0]+1, cols=df.shape[1])
                # Add headers
                for j, col in enumerate(df.columns):
                    t.cell(0, j).text = str(col)
                # Add rows
                for i, row in df.iterrows():
                    for j, val in enumerate(row):
                        t.cell(i+1, j).text = str(val)
        doc.save(doc_buf)
        st.download_button(
            "Download Word File", doc_buf.getvalue(), "output.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    # Export to Excel (.xlsx)
    excel_buf = io.BytesIO()
    if st.button("Export to Excel (.xlsx)"):
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            for idx, df in enumerate(edited_tables):
                sheet_name = f"Table_{idx+1}"
                df.to_excel(writer, index=False, sheet_name=sheet_name)
        st.download_button(
            "Download Excel File", excel_buf.getvalue(), "output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Export text as .txt
    if st.button("Export Text (.txt)"):
        st.download_button("Download Text File", text_area.encode(), "output.txt")

    # --- User feedback ---
    st.subheader("Submit Feedback / Corrections")
    feedback = st.text_area("How can we improve this extraction?", key="feedback")
    if st.button("Submit Feedback"):
        with open("feedback/feedback.txt", "a", encoding="utf-8") as f:
            f.write(f"--- Feedback for file: {uploaded_file.name} ---\n{text_area}\n{feedback}\n\n")
        st.success("Thank you for your feedback! This will help improve future extractions.")

else:
    st.info("Please upload a PDF to get started.")

st.markdown("---")
st.caption("v0.1 | Code Generator GPT · This is a foundation MVP. Expand with OCR, more formats, smarter learning as needed.")

