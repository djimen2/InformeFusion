import io
import re
import unicodedata
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st
from pypdf import PdfReader, PdfWriter


st.set_page_config(
    page_title="InformeFusion",
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


def tokens_text(text: str) -> list[str]:
    """Retorna paraules normalitzades útils per comparar noms."""
    text_norm = normalitza_text(text)
    return [token for token in text_norm.split() if token]


def nom_seguretat_fitxer(text: str) -> str:
    """Neteja el nom final perquè sigui vàlid com a nom de fitxer."""
    text = str(text).strip()
    text = re.sub(r"[\\/:*?\"<>|]", "-", text)
    text = re.sub(r"\s+", " ", text)
    return text[:180]


def llegeix_alumnes(uploaded_file, sense_capcalera=False):
    if uploaded_file is None:
        return None

    nom = uploaded_file.name.lower()
    header = None if sense_capcalera else 0

    if nom.endswith(".csv"):
        return pd.read_csv(uploaded_file, header=header)
    return pd.read_excel(uploaded_file, header=header)


def extreu_nom_alumne(row, mode, columna_nom, columna_cognoms, columna_nom_complet):
    if mode == "Una columna amb nom complet":
        return str(row[columna_nom_complet]).strip()
    nom = str(row[columna_nom]).strip()
    cognoms = str(row[columna_cognoms]).strip()
    return f"{cognoms}, {nom}".strip(", ")


def prepara_fitxers(files):
    resultat = []
    for f in files or []:
        contingut = f.read()
        nom_original = f.name
        nom_sense_extensio = nom_original.rsplit(".", 1)[0]
        resultat.append({
            "nom_original": nom_original,
            "nom_normalitzat": normalitza_text(nom_sense_extensio),
            "tokens": set(tokens_text(nom_sense_extensio)),
            "bytes": contingut,
        })
    return resultat


def busca_pdf_per_nom(nom_alumne, fitxers):
    """
    Relaciona un PDF amb un alumne només si totes les parts del nom de l'alumne
    apareixen al nom de l'arxiu, independentment de l'ordre.

    Exemple:
    - Alumne: "Admer, Yahya"
    - PDF: "Yahya Admer.pdf"
    -> Coincideix.
    """
    if not fitxers:
        return None, ""

    tokens_alumne = set(tokens_text(nom_alumne))
    if not tokens_alumne:
        return None, ""

    candidats = []
    for fitxer in fitxers:
        if tokens_alumne.issubset(fitxer["tokens"]):
            tokens_extres = len(fitxer["tokens"] - tokens_alumne)
            llargada_nom = len(fitxer["nom_normalitzat"])
            candidats.append((tokens_extres, llargada_nom, fitxer))

    if not candidats:
        return None, ""

    candidats.sort(key=lambda item: (item[0], item[1], item[2]["nom_original"]))
    millor = candidats[0][2]

    if len(candidats) > 1:
        altres = ", ".join(c[2]["nom_original"] for c in candidats[1:])
        avis = f"Més d'una coincidència possible. S'ha triat {millor['nom_original']}. Altres candidats: {altres}"
        return millor, avis

    return millor, ""


def afegeix_pdf(writer, pdf_bytes):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    for page in reader.pages:
        writer.add_page(page)


def construeix_nom_pdf(alumne, titol_base, format_nom_sortida):
    if format_nom_sortida == "Alumne - Document":
        return nom_seguretat_fitxer(f"{alumne} - {titol_base}.pdf")
    if format_nom_sortida == "Document - Alumne":
        return nom_seguretat_fitxer(f"{titol_base} - {alumne}.pdf")
    return nom_seguretat_fitxer(f"{alumne}.pdf")


def genera_zip(taula, ordre, titol_base, format_nom_sortida):
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
                    errors.append(
                        f"{registre['alumne']}: error llegint {bloc} ({fitxer['nom_original']}): {e}"
                    )

            if afegit:
                pdf_buffer = io.BytesIO()
                writer.write(pdf_buffer)
                nom_final = construeix_nom_pdf(registre["alumne"], titol_base, format_nom_sortida)
                zf.writestr(nom_final, pdf_buffer.getvalue())
                generats += 1
            else:
                errors.append(f"{registre['alumne']}: no s'ha trobat cap PDF per unificar.")

        avisos = []
        for registre in taula:
            for camp in ["avis_notes", "avis_psi", "avis_observacions"]:
                if registre.get(camp):
                    avisos.append(f"{registre['alumne']}: {registre[camp]}")

        contingut_avisos = avisos + errors
        if contingut_avisos:
            zf.writestr("AVISOS.txt", "\n".join(contingut_avisos))

    buffer_zip.seek(0)
    return buffer_zip, generats, contingut_avisos if 'contingut_avisos' in locals() else []


st.title("📄 InformeFusion")
st.caption("Unifica informes oficials, PSI i observacions en un únic PDF per alumne.")

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

    sense_capcalera = st.checkbox(
        "El llistat no té capçalera i la primera fila ja és un alumne",
        value=False,
    )

    df = None
    alumnes = []
    if excel_alumnes:
        try:
            df = llegeix_alumnes(excel_alumnes, sense_capcalera)
            if sense_capcalera:
                df.columns = [f"Columna {i + 1}" for i in range(len(df.columns))]

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

            alumnes = []
            for _, row in df.iterrows():
                alumne = extreu_nom_alumne(
                    row,
                    mode_nom,
                    columna_nom,
                    columna_cognoms,
                    columna_nom_complet,
                )
                if alumne and alumne.lower() != "nan":
                    alumnes.append(alumne)

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
conf1, conf2 = st.columns([1, 1])
with conf1:
    titol_base = st.text_input("Nom base dels PDFs finals", value="Informe final")
    nom_zip_base = st.text_input("Nom del ZIP final", value="informes_unificats")
with conf2:
    ordre = st.multiselect(
        "Ordre d'unificació",
        options=["notes", "psi", "observacions"],
        default=["notes", "psi", "observacions"],
    )
    format_nom_sortida = st.radio(
        "Format del nom dels PDFs finals",
        ["Alumne - Document", "Document - Alumne", "Només alumne"],
        index=0,
        horizontal=True,
    )

st.info(
    "La coincidència es fa pel nom de l'arxiu: si totes les parts del nom de l'alumne apareixen al PDF, "
    "es relacionen automàticament encara que estiguin en un altre ordre."
)

st.subheader("4. Revisió i generació")

if alumnes and ordre:
    fitxers_notes = prepara_fitxers(pdf_notes)
    fitxers_psi = prepara_fitxers(pdf_psi)
    fitxers_observacions = prepara_fitxers(pdf_observacions)

    registres = []
    files_revisio = []
    for alumne in alumnes:
        nota, avis_nota = busca_pdf_per_nom(alumne, fitxers_notes)
        psi, avis_psi = busca_pdf_per_nom(alumne, fitxers_psi)
        obs, avis_obs = busca_pdf_per_nom(alumne, fitxers_observacions)

        registres.append({
            "alumne": alumne,
            "notes": nota,
            "psi": psi,
            "observacions": obs,
            "avis_notes": avis_nota,
            "avis_psi": avis_psi,
            "avis_observacions": avis_obs,
        })
        files_revisio.append({
            "Alumne": alumne,
            "Notes": nota["nom_original"] if nota else "—",
            "PSI": psi["nom_original"] if psi else "—",
            "Observacions": obs["nom_original"] if obs else "—",
        })

    st.dataframe(pd.DataFrame(files_revisio), use_container_width=True)

    if st.button("Generar ZIP amb els informes unificats", type="primary"):
        zip_buffer, generats, avisos = genera_zip(registres, ordre, titol_base, format_nom_sortida)
        data_actual = datetime.now().strftime("%Y%m%d_%H%M")
        st.success(f"PDFs generats: {generats}")
        if avisos:
            st.warning("Hi ha avisos. També trobaràs un fitxer AVISOS.txt dins del ZIP.")
        st.download_button(
            "Descarregar ZIP",
            data=zip_buffer,
            file_name=nom_seguretat_fitxer(f"{nom_zip_base}_{data_actual}.zip"),
            mime="application/zip",
        )
else:
    st.info("Puja el llistat d'alumnes i selecciona l'ordre per començar.")
