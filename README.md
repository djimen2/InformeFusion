# Unificador d'informes PDF

Aplicació web senzilla per unificar informes PDF d'alumnes: informe oficial/notes, PSI i observacions.

## Què fa?

- Llegeix un llistat d'alumnes en Excel o CSV.
- Permet pujar diferents blocs de PDFs:
  - informes oficials / notes / creuetes
  - PSI
  - observacions
- Busca automàticament el PDF corresponent a cada alumne segons el nom del fitxer.
- Uneix els documents en l'ordre triat.
- Genera un ZIP amb un PDF final per alumne.

## Privacitat

L'aplicació no desa els documents de manera permanent. Processa els PDFs durant la sessió i genera un ZIP descarregable.

Tot i això, com que es treballa amb dades d'alumnes, s'ha d'utilitzar només en un entorn autoritzat pel centre i seguint les normes de protecció de dades.

## Estructura recomanada dels noms dels PDFs

Per millorar la detecció automàtica, és recomanable que els PDFs incloguin el nom i cognoms de l'alumne.

Exemples:

- `Sara Ait El Mhatef - notes.pdf`
- `Sara Ait El Mhatef - observacions.pdf`
- `Sara Ait El Mhatef - PSI.pdf`

## Instal·lació local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicació a Streamlit Community Cloud

1. Crea un repositori a GitHub.
2. Puja aquests fitxers:
   - `app.py`
   - `requirements.txt`
   - `README.md`
   - `.gitignore`
3. Entra a Streamlit Community Cloud.
4. Crea una nova app connectada al repositori.
5. Indica que el fitxer principal és `app.py`.
6. Comparteix l'enllaç amb els mestres.

## Recomanació important

No pugis mai PDFs reals d'alumnes al repositori de GitHub.
