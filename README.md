# Gruenewald Whitepaper Frontend

Schlankes Frontend für deine Whitepaper-Styling-Maschine (`style_whitepaper.py`), optimiert für GitHub + Vercel.

## Lokal testen

```bash
cd "/Users/davidwegener/Desktop/Dokument-hübsch-Macher"
python3 -m http.server 4310
```

Dann im Browser: `http://localhost:4310`

## Deployment mit GitHub + Vercel

```bash
cd "/Users/davidwegener/Desktop/Dokument-hübsch-Macher"
git init
git add .
git commit -m "Add slim frontend for whitepaper styler"
git branch -M main
git remote add origin <DEIN_GITHUB_REPO_URL>
git push -u origin main
```

Danach in Vercel:
1. New Project
2. GitHub Repo auswählen
3. Deploy klicken

## Whitepaper Styling ausführen

Die Website erzeugt nur den Terminal-Befehl; das eigentliche Styling läuft lokal mit Python:

```bash
python3 "/Users/davidwegener/Desktop/Dokument-hübsch-Macher/style_whitepaper.py" \
  "/Users/davidwegener/Desktop/Whitepaper_KI_Validierung_Siemens_Gruenewald_final_DW.docx" \
  "/Users/davidwegener/Desktop/Dokument-hübsch-Macher/Whitepaper_styled.docx" \
  --logo "/Users/davidwegener/Desktop/Dokument-hübsch-Macher/gruenewald_logo.png"
```
