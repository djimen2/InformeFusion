# InformeFusion

Unifica automàticament informes de notes, observacions i PSI en un únic PDF per alumne. Generació massiva i descàrrega en ZIP.

## Funcionament

1. Puja un Excel o CSV amb el llistat d'alumnes.
2. Puja els PDFs de notes, observacions i, si cal, PSI.
3. Tria l'ordre d'unificació.
4. Tria el format del nom final dels PDFs.
5. Genera i descarrega el ZIP.

## Coincidència de noms

L'aplicació relaciona cada alumne amb els PDFs a partir del nom de l'arxiu. Si totes les parts del nom de l'alumne apareixen al nom del PDF, es considera coincidència encara que estiguin en un altre ordre.

Per exemple, aquests casos coincideixen:

- Alumne: `Admer, Yahya`
- PDF: `Yahya Admer.pdf`
- PDF: `Admer, Yahya.pdf`

## Protecció de dades

L'aplicació processa els documents durant la sessió i genera un ZIP final. No s'han de pujar documents reals al repositori de GitHub.
