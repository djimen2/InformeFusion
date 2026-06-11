import io
import re
import unicodedata
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st
from pypdf import PdfReader, PdfWriter
from rapidfuzz import fuzz, process


st.set_page_config(
    page_title="Unificador d'informes PDF",
    page_icon="📄",
    layout="wide",
)


def normalitza_text(text: str) -> str:
    """Converteix un text a una forma fàcil de comparar."""
    if text is None:
        return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def nom_seguretat_fitxer(text: str) -> str:
    """Neteja el nom final del PDF perquè sigui vàlid com a nom de fitxer."""
    text = str(text).strip()
    text = re.sub(r"[\\/:*?\"<>|]", "-", text)
    text = re.sub(r"\s+", " ", text)
    return text[:180]


def llegeix_alumnes(uploaded_file):
    if uploaded_file is None:
        return None

    nom = uploaded_file.name.lower()
    if nom.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)


def extreu_nom_alumne(row, mode, columna_nom, columna_cognoms, columna_nom_complet):
    if mode == "Una columna amb nom complet":
        return str(row[columna_nom_complet]).strip()
    nom = str(row[columna_nom]).strip()
    cognoms = str(row[columna_cognoms]).strip()
    return f"{nom} {cognoms}".strip()


def prepara_fitxers(files):
    resultat = []
    for f in files or []:
        contingut = f.read()
        nom_original = f.name
        nom_normalitzat = normalitza_text(nom_original.rsplit(".", 1)[0])
        resultat.append({
            "nom_original": nom_original,
            "nom_normalitzat": nom_normalitzat,
            "bytes": contingut,
        })
    return resultat


def busca_millor_pdf(nom_alumne, fitxers, llindar):
    if not fitxers:
        return None, 0

    nom_norm = normalitza_text(nom_alumne)
    candidats = {f["nom_normalitzat"]: f for f in fitxers}
    millor = process.extractOne(
        nom_norm,
        list(candidats.keys()),
        scorer=fuzz.token_set_ratio,
    )
    if not millor:
        return None, 0

    nom_fitxer_norm, puntuacio, _ = millor
    if puntuacio < llindar:
        return None, puntuacio
    return candidats[nom_fitxer_norm], puntuacio


def afegeix_pdf(writer, pdf_bytes):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    for page in reader.pages:
        writer.add_page(page)


def genera_zip(taula, ordre, titol_base):
    buffer_zip = io.BytesIO()
    errors = []
    generats = 0

    with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for registre in taula:
            writer = PdfWriter()
            afegit = False

            for bloc in ordre:
                fitxer = registre.get(bloc)
                if fitxer is None:
                    continue
                try:
                    afegeix_pdf(writer, fitxer["bytes"])
                    afegit = True
                except Exception as e:
                    errors.append(f"{registre['alumne']}: error llegint {bloc} ({fitxer['nom_original']}): {e}")

            if afegit:
                pdf_buffer = io.BytesIO()
                writer.write(pdf_buffer)
                nom_final = nom_seguretat_fitxer(f"{titol_base} - {registre['alumne']}.pdf")
                zf.writestr(nom_final, pdf_buffer.getvalue())
                generats += 1
            else:
                errors.append(f"{registre['alumne']}: no s'ha trobat cap PDF per unificar.")

        if errors:
            zf.writestr("AVISOS.txt", "\n".join(errors))

    buffer_zip.seek(0)
    return buffer_zip, generats, errors


st.title("📄 Unificador d'informes PDF")
st.caption("Uneix informes oficials, PSI i observacions en un únic PDF per alumne.")

with st.expander("Funcionament i protecció de dades", expanded=True):
    st.write(
        "Aquesta aplicació processa els arxius durant la sessió i genera un ZIP final. "
        "No desa els PDFs en cap carpeta permanent. Tot i així, cal utilitzar-la només en un entorn autoritzat pel centre."
    )

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Llistat d'alumnes")
    excel_alumnes = st.file_uploader(
        "Puja un Excel o CSV amb el llistat de la classe",
        type=["xlsx", "xls", "csv"],
    )

    df = None
    alumnes = []
    if excel_alumnes:
        try:
            df = llegeix_alumnes(excel_alumnes)
            st.success("Llistat carregat correctament.")
            st.dataframe(df.head(10), use_container_width=True)

            columnes = list(df.columns)
            mode_nom = st.radio(
                "Com està escrit el nom de l'alumne?",
                ["Una columna amb nom complet", "Una columna de nom i una de cognoms"],
            )

            if mode_nom == "Una columna amb nom complet":
                columna_nom_complet = st.selectbox("Columna amb el nom complet", columnes)
                columna_nom = columna_cognoms = None
            else:
                columna_nom = st.selectbox("Columna del nom", columnes)
                columna_cognoms = st.selectbox("Columna dels cognoms", columnes)
                columna_nom_complet = None

            alumnes = [
                extreu_nom_alumne(row, mode_nom, columna_nom, columna_cognoms, columna_nom_complet)
                for _, row in df.iterrows()
                if extreu_nom_alumne(row, mode_nom, columna_nom, columna_cognoms, columna_nom_complet)
                and extreu_nom_alumne(row, mode_nom, columna_nom, columna_cognoms, columna_nom_complet).lower() != "nan"
            ]
            st.info(f"Alumnes detectats: {len(alumnes)}")
        except Exception as e:
            st.error(f"No he pogut llegir el llistat: {e}")

with col2:
    st.subheader("2. PDFs")
    pdf_notes = st.file_uploader(
        "Puja els PDFs d'informe oficial / notes / creuetes",
        type=["pdf"],
        accept_multiple_files=True,
    )
    pdf_psi = st.file_uploader(
        "Puja els PDFs de PSI, només si n'hi ha",
        type=["pdf"],
        accept_multiple_files=True,
    )
    pdf_observacions = st.file_uploader(
        "Puja els PDFs d'observacions",
        type=["pdf"],
        accept_multiple_files=True,
    )

st.subheader("3. Configuració")
conf1, conf2, conf3 = st.columns([2, 2, 1])
with conf1:
    titol_base = st.text_input("Nom base dels PDFs finals", value="Informe final")
with conf2:
    ordre = st.multiselect(
        "Ordre d'unificació",
        options=["notes", "psi", "observacions"],
        default=["notes", "psi", "observacions"],
    )
with conf3:
    llindar = st.slider("Exigència coincidència", 50, 100, 72)

st.subheader("4. Revisió i generació")

if alumnes and ordre:
    fitxers_notes = prepara_fitxers(pdf_notes)
    fitxers_psi = prepara_fitxers(pdf_psi)
    fitxers_observacions = prepara_fitxers(pdf_observacions)

    registres = []
    files_revisio = []
    for alumne in alumnes:
        nota, p_nota = busca_millor_pdf(alumne, fitxers_notes, llindar)
        psi, p_psi = busca_millor_pdf(alumne, fitxers_psi, llindar)
        obs, p_obs = busca_millor_pdf(alumne, fitxers_observacions, llindar)

        registres.append({
            "alumne": alumne,
            "notes": nota,
            "psi": psi,
            "observacions": obs,
        })
        files_revisio.append({
            "Alumne": alumne,
            "Notes": nota["nom_original"] if nota else "—",
            "Coinc. notes": p_nota,
            "PSI": psi["nom_original"] if psi else "—",
            "Coinc. PSI": p_psi,
            "Observacions": obs["nom_original"] if obs else "—",
            "Coinc. observacions": p_obs,
        })

    st.dataframe(pd.DataFrame(files_revisio), use_container_width=True)

    if st.button("Generar ZIP amb els informes unificats", type="primary"):
        zip_buffer, generats, errors = genera_zip(registres, ordre, titol_base)
        data_actual = datetime.now().strftime("%Y%m%d_%H%M")
        st.success(f"PDFs generats: {generats}")
        if errors:
            st.warning("Hi ha avisos. També trobaràs un fitxer AVISOS.txt dins del ZIP.")
        st.download_button(
            "Descarregar ZIP",
            data=zip_buffer,
            file_name=f"informes_unificats_{data_actual}.zip",
            mime="application/zip",
        )
else:
    st.info("Puja el llistat d'alumnes i selecciona l'ordre per començar.")
