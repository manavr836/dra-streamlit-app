import streamlit as st
import os
import tempfile
from datetime import date
import glob
import zipfile
import pandas as pd
import time
from PIL import Image

# Import your main functions from DRA_Distance_to_Boundary_v17.py
from DRA_Distance_to_Boundary_v17 import (
    process_coords,
    CreatingWordFile,
    TEMPLATE_DOCX_PATH,
)

# ---------------------------------------------------
# Streamlit Page Setup
# ---------------------------------------------------
st.set_page_config(page_title="DRA Distance to Boundary", layout="wide", page_icon="📍")
st.title("📍 Intelligent Desktop Risk Assessment (IDRA)")
st.markdown("Upload CSV, process coordinates, generate Word files, and visualize maps.")

# ---------------------------------------------------
# Sidebar Controls
# ---------------------------------------------------
with st.sidebar:
    st.header("⚙️ Settings & Upload")
    input_csv_file = st.file_uploader("Upload input CSV", type=["csv"])
    input_docx_file = st.file_uploader("Upload DOCX template", type=["docx"])
    output_dir = st.text_input("Output directory", "Output")
    run_button = st.button("▶️ Run DRA Processing")

# ---------------------------------------------------
# Initialize session state
# ---------------------------------------------------
for key in ["processed", "output_csv", "word_files", "map_files", "df"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ---------------------------------------------------
# Helper functions
# ---------------------------------------------------
def zip_folder(folder_path, zip_name):
    """Zip all files in a folder."""
    zip_path = os.path.join(folder_path, zip_name)
    with zipfile.ZipFile(zip_path, "w") as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                zipf.write(os.path.join(root, file), arcname=file)
    return zip_path

def make_transparent(img_path):
    """Convert white background to transparent and return base64."""
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

# ---------------------------------------------------
# Main Processing
# ---------------------------------------------------
if run_button:
    if input_csv_file is None:
        st.error("⚠️ Please upload a CSV file to proceed.")
    else:
        st.info("⏳ Starting DRA processing...")

        # Save uploads temporarily
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

        os.makedirs(output_dir, exist_ok=True)
        output_csv_path = os.path.join(output_dir, f"output_{date.today().strftime('%Y%m%d')}.csv")

        map_output_dir = os.path.join(output_dir, "Maps")
        word_output_dir = os.path.join(output_dir, "WordFiles")
        os.makedirs(map_output_dir, exist_ok=True)
        os.makedirs(word_output_dir, exist_ok=True)

        # ---------- STEP 1: Process CSV ----------
        st.subheader("1️⃣ Processing CSV File")
        with st.spinner("Processing coordinates..."):
            try:
                df = process_coords(temp_csv.name, output_csv_path)
                st.session_state.df = df
                st.session_state.output_csv = output_csv_path
                st.session_state.processed = True
                st.success(f"✅ Processed CSV saved: `{output_csv_path}`")
                st.dataframe(df.head())
            except Exception as e:
                st.error(f"❌ CSV processing failed: {e}")

        # ---------- STEP 2: Generate Word Files ----------
        if st.session_state.processed:
            st.subheader("2️⃣ Generating Word Files")
            with st.spinner("Creating Word documents..."):
                try:
                    CreatingWordFile(output_csv_path, template_path, word_output_dir)
                    word_files = glob.glob(os.path.join(word_output_dir, "*.docx"))
                    st.session_state.word_files = word_files
                    st.success(f"✅ Generated {len(word_files)} Word files in `{word_output_dir}`")
                except Exception as e:
                    st.error(f"❌ Word file generation failed: {e}")

        # ---------- STEP 3: Generate Maps ----------
        if st.session_state.processed:
            st.subheader("3️⃣ Generating Maps")
            try:
                # Assuming your `process_coords` already creates maps and saves to map_output_dir
                map_files = glob.glob(os.path.join(map_output_dir, "*.png"))
                st.session_state.map_files = map_files
                if map_files:
                    st.success(f"✅ {len(map_files)} map(s) generated.")
                else:
                    st.warning("⚠️ No map files found.")
            except Exception as e:
                st.error(f"❌ Map generation failed: {e}")

        # ---------- STEP 4: Summary & Downloads ----------
        if st.session_state.processed:
            st.subheader("📦 Download Results")

            # Download processed CSV
            if os.path.exists(output_csv_path):
                with open(output_csv_path, "rb") as f:
                    st.download_button(
                        "📥 Download Processed CSV",
                        data=f,
                        file_name=os.path.basename(output_csv_path),
                        mime="text/csv",
                    )

            # Download all Word files
            if st.session_state.word_files:
                word_zip_path = zip_folder(word_output_dir, "WordFiles_All.zip")
                with open(word_zip_path, "rb") as f:
                    st.download_button(
                        "📦 Download All Word Files (ZIP)",
                        data=f,
                        file_name="WordFiles_All.zip",
                        mime="application/zip",
                    )

            # Download all Maps
            if st.session_state.map_files:
                map_zip_path = zip_folder(map_output_dir, "Maps_All.zip")
                with open(map_zip_path, "rb") as f:
                    st.download_button(
                        "🗺️ Download All Maps (ZIP)",
                        data=f,
                        file_name="Maps_All.zip",
                        mime="application/zip",
                    )

                # Display maps in grid
                st.subheader("🗺️ Generated Maps Preview")
                maps_per_row = 3
                for i in range(0, len(st.session_state.map_files), maps_per_row):
                    cols = st.columns(maps_per_row)
                    for j, mf in enumerate(st.session_state.map_files[i:i+maps_per_row]):
                        with cols[j]:
                            st.image(mf, caption=os.path.basename(mf), use_container_width=True)

            st.success("✅ Processing complete!")

# ---------------------------------------------------
# Branding Logo (Optional)
# ---------------------------------------------------
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
