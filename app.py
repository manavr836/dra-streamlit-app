import streamlit as st
import os
import pandas as pd
import tempfile
from datetime import date
import time
import glob
import zipfile
from PIL import Image

# Import helper module
from DRA_Distance_to_Boundary_v17 import (
    process_coords,
    CreatingWordFile,
    TEMPLATE_DOCX_PATH,
    OUTPUT_DIR,
)

# ---------------------------------------------------
# Streamlit Page Setup
# ---------------------------------------------------
st.set_page_config(
    page_title="DRA Distance to Boundary",
    layout="wide",
    page_icon="📍"
)

st.title("📍 Intelligent Desktop Risk Assessment (IDRA)")
st.markdown("Upload CSV, process coordinates, and generate Word files automatically.")

# ---------------------------------------------------
# Sidebar Controls
# ---------------------------------------------------
with st.sidebar:
    st.header("⚙️ Settings & Upload")
    input_csv_file = st.file_uploader("Upload input CSV", type=["csv"])
    input_docx_file = st.file_uploader("Upload DOCX template", type=["docx"])
    output_dir = st.text_input("Output directory", "Output")
    run_button = st.button("▶️ Run Processing")

# ---------------------------------------------------
# Session state
# ---------------------------------------------------
for key in ["processed", "output_csv", "word_files"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ---------------------------------------------------
# Helper function
# ---------------------------------------------------
def zip_folder(folder_path, zip_name):
    """Zip all files in a folder."""
    zip_path = os.path.join(folder_path, zip_name)
    with zipfile.ZipFile(zip_path, "w") as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                zipf.write(os.path.join(root, file), arcname=file)
    return zip_path


# ---------------------------------------------------
# Main Processing
# ---------------------------------------------------
if run_button:
    if input_csv_file is None:
        st.error("⚠️ Please upload an input CSV file.")
    else:
        st.info("⏳ Starting processing...")

        # Temporary save uploaded files
        temp_csv = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
        temp_csv.write(input_csv_file.read())
        temp_csv.close()

        if input_docx_file:
            temp_docx = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            temp_docx.write(input_docx_file.read())
            temp_docx.close()
            template_path = temp_docx.name
        else:
            template_path = TEMPLATE_DOCX_PATH

        # Ensure output folders exist
        os.makedirs(output_dir, exist_ok=True)
        output_csv_path = os.path.join(output_dir, f"output_{date.today().strftime('%Y%m%d')}.csv")
        word_output_dir = os.path.join(output_dir, "WordFiles")
        os.makedirs(word_output_dir, exist_ok=True)

        # Step 1: Process CSV
        st.subheader("1️⃣ Processing CSV file")
        with st.spinner("Processing coordinates..."):
            try:
                df = process_coords(temp_csv.name, output_csv_path)
                st.session_state.processed = True
                st.session_state.output_csv = output_csv_path
                st.success(f"✅ Processed CSV saved to `{output_csv_path}`")
                st.dataframe(df.head())
            except Exception as e:
                st.error(f"❌ CSV processing failed: {e}")

        # Step 2: Generate Word Files
        if st.session_state.processed:
            st.subheader("2️⃣ Generating Word files")
            with st.spinner("Creating Word files..."):
                try:
                    CreatingWordFile(output_csv_path, template_path, word_output_dir)
                    word_files = glob.glob(os.path.join(word_output_dir, "*.docx"))
                    st.session_state.word_files = word_files
                    st.success(f"✅ {len(word_files)} Word files generated in `{word_output_dir}`")
                except Exception as e:
                    st.error(f"❌ Word file generation failed: {e}")

        # Step 3: Summary & Downloads
        if st.session_state.word_files:
            st.subheader("📦 Download Results")

            # Download processed CSV
            if os.path.exists(output_csv_path):
                with open(output_csv_path, "rb") as f:
                    st.download_button(
                        "📥 Download Processed CSV",
                        data=f,
                        file_name=os.path.basename(output_csv_path),
                        mime="text/csv"
                    )

            # ZIP and Download all Word files
            zip_path = os.path.join(output_dir, "WordFiles_All.zip")
            with zipfile.ZipFile(zip_path, "w") as zipf:
                for wf in st.session_state.word_files:
                    zipf.write(wf, os.path.basename(wf))
            with open(zip_path, "rb") as f:
                st.download_button(
                    "📦 Download All Word Files (ZIP)",
                    data=f,
                    file_name="WordFiles_All.zip",
                    mime="application/zip"
                )

            st.success("✅ Processing complete!")


# ---------------------------------------------------
# Add Branding Logos (Optional)
# ---------------------------------------------------
def make_transparent(img_path):
    """Make white pixels transparent for PNG logos."""
    img = Image.open(img_path).convert("RGBA")
    datas = img.getdata()
    new_data = []
    for item in datas:
        if item[0] > 240 and item[1] > 240 and item[2] > 240:
            new_data.append((255, 255, 255, 0))
        else:
            new_data.append(item)
    img.putdata(new_data)
    import io, base64
    buffered = io.BytesIO()
    img.save(buffered, format="PNG")
    encoded = base64.b64encode(buffered.getvalue()).decode()
    return encoded


logo_path = "Input/Logos/SOPRASTERIA_logo_RVB_exe.png"
if os.path.exists(logo_path):
    encoded_logo = make_transparent(logo_path)
    st.markdown(
        f"""
        <style>
        .logo-bottom {{
            position: fixed;
            bottom: 10px;
            right: 10px;
            width: clamp(120px, 20vw, 220px);
            opacity: 0.85;
            z-index: 100;
        }}
        </style>
        <img src="data:image/png;base64,{encoded_logo}" class="logo-bottom">
        """,
        unsafe_allow_html=True,
    )
