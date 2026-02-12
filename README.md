# Document Beautifier

Web-Tool fuer Whitepaper-Transformation:
- DOCX Upload
- optionales Logo
- Farbdefinitionen
- Template-Auswahl (`minimal`, `executive`, `bold`)
- Ausgabe als **DOCX oder hochwertiges PDF**

## Live-Flow
1. Dokument hochladen
2. Design konfigurieren
3. Output-Format waehlen (PDF/DOCX)
4. Generieren -> Download startet direkt

## Lokale CLI-Nutzung

```bash
python3 "/Users/davidwegener/Desktop/Dokument-hübsch-Macher/style_whitepaper.py" \
  "/ABSOLUTER/PFAD/input.docx" \
  "/ABSOLUTER/PFAD/output_styled.pdf" \
  --template executive \
  --primary-color "#F50000" \
  --text-color "#111111" \
  --org-name "Your Organization" \
  --logo "/ABSOLUTER/PFAD/logo.png"
```

## API
`POST /api/style` (multipart/form-data)
- `document` (.docx, required)
- `logo` (optional)
- `outputFormat`: `docx` | `pdf`
- `template`, `primaryColor`, `textColor`, `orgName`, `outputName`

Antwort: Binardatei als Download.
