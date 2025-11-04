import os
import sys
import pandas as pd

# Try to import COM libraries only if on Windows
if sys.platform.startswith("win"):
    try:
        import pythoncom
        import win32com.client as win32
        HAS_WIN32 = True
    except ImportError:
        HAS_WIN32 = False
else:
    HAS_WIN32 = False

TEMPLATE_DOCX_PATH = "Input/Template.docx"
OUTPUT_DIR = "Output/WordFiles"


def process_coords(input_csv, output_csv, console_print=False):
    """Example coordinate processor (you can replace with your real logic)."""
    df = pd.read_csv(input_csv)
    if "AssetID" not in df.columns:
        df["AssetID"] = range(1, len(df) + 1)
    if "Latitude" not in df.columns:
        df["Latitude"] = 0
    if "Longitude" not in df.columns:
        df["Longitude"] = 0
    df.to_csv(output_csv, index=False)
    if console_print:
        print(f"✅ Processed file saved to: {output_csv}")
    return df


# --------------------------------------------------------------------
# WORD FILE GENERATION — Windows COM Version
# --------------------------------------------------------------------
def CreatingWordFile_COM(output_csv, template_path, output_dir):
    import pythoncom
    pythoncom.CoInitialize()
    try:
        df = pd.read_csv(output_csv)
        os.makedirs(output_dir, exist_ok=True)
        word = win32.Dispatch("Word.Application")
        word.Visible = False

        for _, row in df.iterrows():
            doc = word.Documents.Open(template_path)
            for field, value in row.items():
                try:
                    word.Selection.Find.Execute(
                        f"<<{field}>>", False, False, False, False, False, True, 1, False, str(value), 2
                    )
                except Exception:
                    pass
            doc.SaveAs(os.path.join(output_dir, f"{row['AssetID']}.docx"))
            doc.Close()
        word.Quit()
        print("✅ Word files created with MS Word COM.")
    finally:
        pythoncom.CoUninitialize()


# --------------------------------------------------------------------
# WORD FILE GENERATION — Cloud-Safe Version (python-docx)
# --------------------------------------------------------------------
def CreatingWordFile_DOCX(output_csv, template_path, output_dir):
    from docx import Document
    df = pd.read_csv(output_csv)
    os.makedirs(output_dir, exist_ok=True)

    for _, row in df.iterrows():
        doc = Document(template_path)

        # Replace placeholders in all paragraphs
        for p in doc.paragraphs:
            for field, value in row.items():
                placeholder = f"<<{field}>>"
                if placeholder in p.text:
                    p.text = p.text.replace(placeholder, str(value))

        # Replace placeholders in tables
        for table in doc.tables:
            for row_cells in table.rows:
                for cell in row_cells.cells:
                    for field, value in row.items():
                        placeholder = f"<<{field}>>"
                        if placeholder in cell.text:
                            cell.text = cell.text.replace(placeholder, str(value))

        doc.save(os.path.join(output_dir, f"{row['AssetID']}.docx"))
    print("✅ Word files created with python-docx.")


def CreatingWordFile(output_csv, template_path, output_dir):
    """Auto-selects COM or DOCX generator."""
    if HAS_WIN32:
        CreatingWordFile_COM(output_csv, template_path, output_dir)
    else:
        print("⚠️ Using python-docx fallback (no win32 available).")
        CreatingWordFile_DOCX(output_csv, template_path, output_dir)
