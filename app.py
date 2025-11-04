import streamlit as st
import os
import tempfile
from datetime import date
import threading
import pythoncom
import glob
from PIL import Image
import pandas as pd
import time
import zipfile
import base64

from DRA_Distance_to_Boundary_v17 import process_coords, CreatingWordFile, TEMPLATE_DOCX_PATH
import DRA_Distance_to_Boundary_v17

# -------------------- PAGE CONFIG --------------------
st.set_page_config(page_title="DRA Distance to Boundary", layout="wide", page_icon="📍")

# -------------------- CUSTOM CSS --------------------
st.markdown("""
    <style>
        .block-container { padding: 2rem 3rem; }

        /* Default Streamlit buttons */
        .stButton>button, .stDownloadButton>button {
            font-weight: bold;
            color: white;
            transition: all 0.3s ease;
        }

        /* Specific button colors */
        .stButton>button {
            background-color: #0052cc;
        }
        .stDownloadButton>button {
            background-color: #008000;
        }

        /* Hover effect */
        .stButton>button:hover {
            background-color: #003d99;
            color: white;
        }
        .stDownloadButton>button:hover {
            background-color: #006400;
            color: white;
        }

        .logo-bottom-right { position: fixed; bottom: 10px; right: 10px; width: 120px; opacity: 0.8; }
        .card { padding: 1rem; border-radius: 10px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); background-color: #f9f9f9; margin-bottom: 1rem; }
        .card h3 { margin-top: 0; }
    </style>
""", unsafe_allow_html=True)

# -------------------- APP TITLE --------------------
st.title("📍 Intelligent Desktop Risk Assessment (IDRA)")
st.markdown("Upload CSV and DOCX template. Process coordinates, generate Word files, and visualize maps.")

# -------------------- SIDEBAR --------------------
with st.sidebar:
    st.header("Upload & Settings")
    input_csv_file = st.file_uploader("Upload input CSV", type=["csv"])
    input_docx_file = st.file_uploader("Upload DOCX template", type=["docx"])
    output_dir = st.text_input("Output directory", os.getcwd())
    run_button = st.button("Run DRA Processing")

# -------------------- SESSION STATE INIT --------------------
for key in ["output_csv_path", "word_files", "map_files", "processed", "df"]:
    if key not in st.session_state:
        st.session_state[key] = None if key not in ["processed"] else False

# -------------------- HELPER FUNCTION --------------------
def create_word_file_thread(output_csv, template_path, output_dir, progress_placeholder):
    pythoncom.CoInitialize()
    try:
        CreatingWordFile(output_csv, template_path, output_dir)
        progress_placeholder.success("✅ Word file generation complete!")
    except Exception as e:
        progress_placeholder.error(f"Word creation failed: {e}")
    finally:
        pythoncom.CoUninitialize()

# -------------------- MAIN PROCESS --------------------
if run_button:
    if input_csv_file is None:
        st.error("Please upload an input CSV file to proceed.")
    else:
        csv_status = st.empty()
        word_status = st.empty()
        map_status = st.empty()

        # ---------- CSV PROCESSING ----------
        csv_card = st.container()
        with csv_card:
            st.markdown('<div class="card"><h3>1️⃣ CSV Processing</h3></div>', unsafe_allow_html=True)
            csv_progress = st.progress(0)

        with st.spinner("Processing CSV..."):
            temp_csv = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
            temp_csv.write(input_csv_file.read())
            temp_csv.close()

            if input_docx_file:
                temp_docx = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
                temp_docx.write(input_docx_file.read())
                temp_docx.close()
                template_path = temp_docx.name
            else:
                template_path = TEMPLATE_DOCX_PATH if os.path.exists(TEMPLATE_DOCX_PATH) else None

            os.makedirs(output_dir, exist_ok=True)
            output_csv_path = os.path.join(output_dir, f"output_{date.today().strftime('%Y%m%d')}.csv")
            map_output_dir = os.path.join(output_dir, "Maps")
            os.makedirs(map_output_dir, exist_ok=True)
            DRA_Distance_to_Boundary_v17.MAP_OUTPUT_DIR = map_output_dir
            output_word_dir = os.path.join(output_dir, "WordFiles")
            os.makedirs(output_word_dir, exist_ok=True)

            for i in range(0, 101, 20):
                csv_progress.progress(i)
                time.sleep(0.1)

            try:
                process_coords(temp_csv.name, output_csv_path, template_docx=template_path, console_print=True)
                csv_progress.progress(100)
                csv_status.success(f"✅ CSV processed: {output_csv_path}")
                st.session_state.output_csv_path = output_csv_path
                st.session_state.processed = True
            except Exception as e:
                csv_status.error(f"CSV processing failed: {e}")

        # ---------- WORD FILE GENERATION ----------
        word_card = st.container()
        with word_card:
            st.markdown('<div class="card"><h3>2️⃣ Generating Desktop Risk Assessment Form</h3></div>', unsafe_allow_html=True)
            word_progress = st.progress(0)
        with st.spinner("Generating Desktop Risk Assessment Form..."):
            temp_csv = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
            temp_csv.write(input_csv_file.read())
            temp_csv.close()
            try:
                df = pd.read_csv(output_csv_path)
                df.rename(columns={"Converted Lat": "Latitude", "Converted Lon": "Longitude", "Asset Name": "Name"}, inplace=True)
                if 'AssetID' not in df.columns:
                    df['AssetID'] = df['Job Reference']
                df.to_csv(output_csv_path, index=False)
                st.session_state.df = df

                required_columns = ["AssetID", "Latitude", "Longitude", "Name"]
                missing_columns = [col for col in required_columns if col not in df.columns]

                if missing_columns:
                    word_status.error(f"Cannot generate Word files. Missing columns: {missing_columns}")
                else:
                    for i in range(0, 101, 25):
                        word_progress.progress(i)
                        time.sleep(0.1)

                    word_thread = threading.Thread(
                        target=create_word_file_thread,
                        args=(output_csv_path, template_path, output_word_dir, word_status)
                    )
                    word_thread.start()
                    word_thread.join()

                    word_files = glob.glob(os.path.join(output_word_dir, "*.docx"))
                    st.session_state.word_files = word_files

            except Exception as e:
                word_status.error(f"Word generation failed: {e}")

        # ---------- MAP VISUALIZATION ----------
        map_card = st.container()
        with map_card:
            st.markdown('<div class="card"><h3>3️⃣ Generating Maps for assets within the Boundary</h3></div>', unsafe_allow_html=True)
            map_progress = st.progress(0)
            map_files = glob.glob(os.path.join(map_output_dir, "*.png"))
            st.session_state.map_files = map_files

            if map_files:
                for i in range(0, 101, 20):
                    map_progress.progress(i)
                    time.sleep(0.05)
                map_progress.progress(100)
                map_status.success("✅ Map generation complete!")
            else:
                map_status.warning("No map images found.")

# -------------------- SUMMARY + DOWNLOADS --------------------
if st.session_state.processed:
    # ---------- SUMMARY ----------
    summary_card = st.container()
    with summary_card:
        st.markdown('<div class="card"><h3>📊 Processing Summary</h3></div>', unsafe_allow_html=True)
        try:
            df = st.session_state.df if st.session_state.df is not None else pd.DataFrame()
            total_assets = len(df)
            total_word_files = len(st.session_state.word_files or [])
            total_map_files = len(st.session_state.map_files or [])

            col1, col2, col3 = st.columns(3)
            col1.metric("✅ Total Rows Processed", total_assets)
            col2.metric("📄 Word Files Generated", total_word_files)
            col3.metric("🗺️ Maps Generated", total_map_files)
        except Exception as e:
            st.error(f"Failed to generate summary: {e}")

    # ---------- DOWNLOADS ----------
    st.markdown('<div class="card"><h3>📥 Download Files</h3></div>', unsafe_allow_html=True)

    # CSV
    if st.session_state.output_csv_path and os.path.exists(st.session_state.output_csv_path):
        with open(st.session_state.output_csv_path, "rb") as f:
            st.download_button(
                "📥 Download Processed CSV",
                data=f,
                file_name=os.path.basename(st.session_state.output_csv_path),
                mime="text/csv",
                key="dl_csv"
            )

    # Word files
    if st.session_state.word_files:
        for wf in st.session_state.word_files:
            with open(wf, "rb") as f:
                st.download_button(
                    f"📥 Download {os.path.basename(wf)}",
                    data=f,
                    file_name=os.path.basename(wf),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_word_{os.path.basename(wf)}"
                )

        # ZIP all Word files
        word_zip_path = os.path.join(os.path.dirname(st.session_state.word_files[0]), "WordFiles_All.zip")
        with zipfile.ZipFile(word_zip_path, 'w') as zipf:
            for wf in st.session_state.word_files:
                zipf.write(wf, os.path.basename(wf))
        with open(word_zip_path, "rb") as f:
            st.download_button("📦 Download All Word Files (ZIP)", data=f, file_name="WordFiles_All.zip", mime="application/zip", key="dl_word_zip")

    # Maps
    if st.session_state.map_files:
        # ZIP all maps
        map_zip_path = os.path.join(os.path.dirname(st.session_state.map_files[0]), "Maps_All.zip")
        with zipfile.ZipFile(map_zip_path, 'w') as zipf:
            for mf in st.session_state.map_files:
                zipf.write(mf, os.path.basename(mf))
        with open(map_zip_path, "rb") as f:
            st.download_button("🗺️ Download All Maps (ZIP)", data=f, file_name="Maps_All.zip", mime="application/zip", key="dl_map_zip")

        # Display maps in a grid (side by side)
        st.markdown('<div class="card"><h3>🗺️ Generated Maps</h3></div>', unsafe_allow_html=True)
        maps_per_row = 3  # Number of maps per row
        map_files = st.session_state.map_files

        for i in range(0, len(map_files), maps_per_row):
            cols = st.columns(maps_per_row)
            for j, mf in enumerate(map_files[i:i+maps_per_row]):
                with cols[j]:
                    st.image(mf, caption=os.path.basename(mf), use_container_width=True)
                    with open(mf, "rb") as f:
                        st.download_button(
                            f"📥 Download {os.path.basename(mf)}",
                            data=f,
                            file_name=os.path.basename(mf),
                            mime="image/png",
                            key=f"dl_map_{os.path.basename(mf)}"
                        )

# -------------------- LOGOS --------------------
from PIL import Image

def make_transparent(img_path):
    """Convert white background to transparent and return as base64."""
    img = Image.open(img_path).convert("RGBA")
    datas = img.getdata()

    new_data = []
    for item in datas:
        # Change all white (or near white) pixels to transparent
        if item[0] > 240 and item[1] > 240 and item[2] > 240:
            new_data.append((255, 255, 255, 0))
        else:
            new_data.append(item)
    img.putdata(new_data)

    # Save to a bytes buffer
    import io, base64
    buffered = io.BytesIO()
    img.save(buffered, format="PNG")
    encoded = base64.b64encode(buffered.getvalue()).decode()
    return encoded

# Paths
logo_path_bottom = r"C:\Users\mmogha\OneDrive - Sopra Steria\Desktop\SafeDig DRA\Input\Logos\SOPRASTERIA_logo_RVB_exe.png"
logo_path_top = r"C:\Users\mmogha\OneDrive - Sopra Steria\Desktop\SafeDig DRA\Input\Logos\image (1).png"

# Bottom logo
if os.path.exists(logo_path_bottom):
    encoded_logo_bottom = make_transparent(logo_path_bottom)
    st.markdown(
        f"""
        <style>
        .logo-bottom-right {{
            position: fixed;
            bottom: 0px;
            right: 10px;
            width: clamp(120px, 15vw, 250px);
            height: auto;
            z-index: 100;
            background: transparent;
        }}
        </style>
        <img src="data:image/png;base64,{encoded_logo_bottom}" class="logo-bottom-right">
        """,
        unsafe_allow_html=True
    )

# Top logo
if os.path.exists(logo_path_top):
    encoded_logo_top = make_transparent(logo_path_top)
    st.markdown(
        f"""
        <style>
        .logo-top-right {{
            position: fixed;
            top: 60px;
            right: 10px;
            width: clamp(120px, 15vw, 250px);
            height: auto;
            z-index: 101;
            background: transparent;
        }}
        </style>
        <img src="data:image/png;base64,{encoded_logo_top}" class="logo-top-right">
        """,
        unsafe_allow_html=True
    )
