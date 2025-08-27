# app.py
import io, zipfile, streamlit as st, pandas as pd
from generator import generate

st.set_page_config(page_title="Applicant Reply Generator", page_icon="📧")
st.markdown("## 📧 Applicant Reply Generator")

uploaded = st.file_uploader("Upload applicant spreadsheet (.xlsx)", type=["xlsx"])
if uploaded:
    try:
        zip_bytes, updated_df = generate(uploaded.getvalue())
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
        st.stop()

    # Show success + counts
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        st.success(f"Draft emails generated: **{len(zf.namelist())}**")

    # Download buttons
    st.download_button(
        "Download ZIP of email drafts",
        data=zip_bytes,
        file_name="applicant_emails.zip",
        mime="application/zip",
    )

    buf = io.BytesIO(); updated_df.to_excel(buf, index=False)
    st.download_button(
        "Download updated spreadsheet",
        data=buf.getvalue(),
        file_name="applicants_updated.xlsx",
        mime=("application/"
              "vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    )
