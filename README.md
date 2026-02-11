# Document Beautifier

Web-Tool für DOCX-Whitepaper:
- Dokument hochladen
- Logo hochladen
- Farben definieren
- Vorlage wählen
- **direkt im Tool DOCX erzeugen + herunterladen**

## Lokal testen (Frontend)

```bash
cd "/Users/davidwegener/Desktop/Dokument-hübsch-Macher"
python3 -m http.server 4310
```

Dann: `http://localhost:4310`

## Lokal testen (CLI)

```bash
python3 "/Users/davidwegener/Desktop/Dokument-hübsch-Macher/style_whitepaper.py" \
  "/ABSOLUTER/PFAD/input.docx" \
  "/ABSOLUTER/PFAD/output_styled.docx" \
  --template executive \
  --primary-color "#F50000" \
  --text-color "#111111" \
  --org-name "Your Organization" \
  --logo "/ABSOLUTER/PFAD/logo.png"
```

## Deploy auf Vercel

Projekt ist auf Vercel für statisches Frontend + Python API (`/api/style`) vorbereitet.
