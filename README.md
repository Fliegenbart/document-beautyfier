# Document Beautifier

Generischer Whitepaper-Styler mit Frontend-Konfigurator:
- Dokument hochladen
- Logo hochladen
- Farben definieren
- Vorlage wählen (`minimal`, `executive`, `bold`)

## Lokal starten (Frontend)

```bash
cd "/Users/davidwegener/Desktop/Dokument-hübsch-Macher"
python3 -m http.server 4310
```

Dann im Browser: `http://localhost:4310`

## DOCX-Styling ausführen (CLI)

```bash
python3 "/Users/davidwegener/Desktop/Dokument-hübsch-Macher/style_whitepaper.py" \
  "/ABSOLUTER/PFAD/input.docx" \
  "/ABSOLUTER/PFAD/output_styled.docx" \
  --template executive \
  --primary-color "#F50000" \
  --text-color "#111111" \
  --org-name "GRUENEWALD GmbH" \
  --logo "/ABSOLUTER/PFAD/logo.png"
```

## Deploy auf Vercel

Das Frontend ist statisch und direkt Vercel-kompatibel.
Push nach GitHub, danach in Vercel `New Project` -> Repo wählen -> `Deploy`.
