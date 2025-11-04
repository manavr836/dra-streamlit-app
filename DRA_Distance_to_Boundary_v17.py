import os
import sys
import pandas as pd

# ✅ For Windows users with MS Word
if sys.platform.startswith("win"):
    import pythoncom
    import win32com.client as win32

TEMPLATE_DOCX_PATH = "Input/Template.docx"
MAP_OUTPUT_DIR = "Output/Maps"


def process_coords(input_csv, output_csv, template_docx=None, console_print=False):
    """Dummy coordinate processing — replace with your actual logic"""
    df = pd.read_csv(input_csv)
    df["AssetID"] = df.get("AssetID", range(1, len(df) + 1))
    df["Latitude"] = df.get("Latitude", 0)
    df["Longitude"] = df.get("Longitude", 0)
    df.to_csv(output_csv, index=False)
    if console_print:
        print(f"✅ Processed CSV saved: {output_csv}")


def CreatingWordFile_COM(output_csv, template_path, output_dir):
    """Windows COM automation version"""
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
                    word.Selection.Find.Execute(f"<<{field}>>", False, False, False, False, False, True, 1, False, str(value), 2)
                except Exception:
                    pass
            doc.SaveAs(os.path.join(output_dir, f"{row['AssetID']}.docx"))
            doc.Close()
        word.Quit()
        print("✅ Word documents generated via COM")
    finally:
        pythoncom.CoUninitialize()


def CreatingWordFile_DOCX(output_csv, template_path, output_dir):
    """Cloud-safe fallback using python-docx"""
    from docx import Document
    df = pd.read_csv(output_csv)
    os.makedirs(output_dir, exist_ok=True)

    for _, row in df.iterrows():
        doc = Document(template_path)
        for p in doc.paragraphs:
            for field, value in row.items():
                placeholder = f"<<{field}>>"
                if placeholder in p.text:
                    p.text = p.text.replace(placeholder, str(value))
        doc.save(os.path.join(output_dir, f"{row['AssetID']}.docx"))
    print("✅ Word documents generated via python-docx")


def CreatingWordFile(output_csv, template_path, output_dir):
    """Automatically selects version"""
    if sys.platform.startswith("win"):
        CreatingWordFile_COM(output_csv, template_path, output_dir)
    else:
        CreatingWordFile_DOCX(output_csv, template_path, output_dir)
