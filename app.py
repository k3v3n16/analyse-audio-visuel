import streamlit as st
import fitz
import re
import pytesseract
from PIL import Image
from pypdf import PdfWriter, PdfReader
import openpyxl
import io

# -----------------------------
# CONFIGURATION
# -----------------------------

KEYWORDS = [
    "audio", "audiovisuel", "audio visuel", "A/V", "Multiprise", "Extension"
    "sonorisation", "projecteur", "écran", "micro", "ensemble de projection", "projection", "ensemble", "encore", "Encore"
]

DATE_REGEX = r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{1,2} [A-Za-zéû]+ 20\d{2})\b"
SALLE_REGEX = r"(Salle\s?[A-Za-z0-9-]+|EMPLACEMENT\s?[A-Za-z0-9-]+|HALL|Bloc\s?[A-Z0-9]+)"
AV_REGEX = r"(micro(?:phone)?|projection|écran|haut-?parleurs?|HDMI|ordinateur)"

# -----------------------------
# OCR
# -----------------------------

def ocr_page(page):
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return pytesseract.image_to_string(img, lang="fra")

# -----------------------------
# ANALYSE PDF
# -----------------------------

def analyze_pdf(pdf_bytes, use_ocr=False):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()

    results = []

    for page_num, page in enumerate(doc):
        text = page.get_text()

        if use_ocr and len(text.strip()) < 20:
            text = ocr_page(page)

        if any(k.lower() in text.lower() for k in KEYWORDS):

            dates = re.findall(DATE_REGEX, text)
            salles = re.findall(SALLE_REGEX, text, flags=re.IGNORECASE)
            besoins = re.findall(AV_REGEX, text, flags=re.IGNORECASE)

            results.append({
                "page": page_num + 1,
                "dates": dates,
                "salles": salles,
                "besoins": besoins
            })

            writer.add_page(reader.pages[page_num])

    # PDF filtré en mémoire
    pdf_output = io.BytesIO()
    writer.write(pdf_output)
    pdf_output.seek(0)

    return results, pdf_output

# -----------------------------
# EXPORT EXCEL
# -----------------------------

def export_excel(results):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Analyse AV"

    ws.append(["Page", "Dates", "Salles", "Besoins AV"])

    for r in results:
        ws.append([
            r["page"],
            ", ".join(r["dates"]),
            ", ".join(r["salles"]),
            ", ".join(r["besoins"])
        ])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -----------------------------
# INTERFACE WEB STREAMLIT
# -----------------------------

st.title("🔎 Analyse Audio-Visuel PDF (Web)")
st.write("Importe un PDF et extrait automatiquement les pages contenant une section audio-visuel.")

uploaded_file = st.file_uploader("Choisir un PDF", type=["pdf"])
use_ocr = st.checkbox("Activer OCR (pour PDF scannés)")

if uploaded_file:
    if st.button("Analyser le document"):
        pdf_bytes = uploaded_file.read()

        with st.spinner("Analyse en cours..."):
            results, filtered_pdf = analyze_pdf(pdf_bytes, use_ocr)

        st.success("Analyse terminée !")

        # Affichage des résultats
        st.subheader("📊 Résultats")
        st.write(results)

        # Téléchargement Excel
        excel_file = export_excel(results)
        st.download_button(
            label="📥 Télécharger Excel",
            data=excel_file,
            file_name="analyse_audio_visuel.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Téléchargement PDF filtré
        st.download_button(
            label="📥 Télécharger PDF filtré",
            data=filtered_pdf,
            file_name="pages_audio_visuel.pdf",
            mime="application/pdf"

        )

