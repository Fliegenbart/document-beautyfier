const docFile = document.getElementById('docFile');
const logoFile = document.getElementById('logoFile');
const docFileName = document.getElementById('docFileName');
const logoFileName = document.getElementById('logoFileName');
const logoPreview = document.getElementById('logoPreview');

const outputFormat = document.getElementById('outputFormat');
const pdfTheme = document.getElementById('pdfTheme');
const outputName = document.getElementById('outputName');
const orgName = document.getElementById('orgName');
const lineSpacing = document.getElementById('lineSpacing');
const readingWidthCh = document.getElementById('readingWidthCh');
const includeSummaryPage = document.getElementById('includeSummaryPage');
const useThemePalette = document.getElementById('useThemePalette');
const primaryColor = document.getElementById('primaryColor');
const textColor = document.getElementById('textColor');
const primaryColorLabel = document.getElementById('primaryColorLabel');
const textColorLabel = document.getElementById('textColorLabel');

const generateBtn = document.getElementById('generateBtn');
const statusText = document.getElementById('statusText');
const templateGrid = document.getElementById('templateGrid');
const previewGrid = document.getElementById('previewGrid');

let previewTimer = null;

function selectedTemplate() {
  return document.querySelector('input[name="template"]:checked')?.value || 'minimal';
}

function refreshTemplateCardState() {
  templateGrid.querySelectorAll('.template').forEach((card) => {
    const radio = card.querySelector('input[type="radio"]');
    card.classList.toggle('active', radio.checked);
  });
}

function updateVisualTheme() {
  document.documentElement.style.setProperty('--accent', primaryColor.value);
  document.documentElement.style.setProperty('--ink', textColor.value);
  primaryColorLabel.textContent = primaryColor.value;
  textColorLabel.textContent = textColor.value;
}

function ensureExtension(filename, ext) {
  const dotExt = `.${ext}`;
  if (filename.toLowerCase().endsWith(dotExt)) return filename;
  return `${filename.replace(/\.[a-z0-9]+$/i, '')}${dotExt}`;
}

function updateFileLabels() {
  const doc = docFile.files?.[0];
  const logo = logoFile.files?.[0];

  docFileName.textContent = doc ? doc.name : 'Keine Datei gewählt';
  logoFileName.textContent = logo ? logo.name : 'Kein Logo gewählt';

  if (doc && (!outputName.value || /document_styled\.(pdf|docx)/i.test(outputName.value))) {
    const ext = outputFormat.value === 'pdf' ? 'pdf' : 'docx';
    outputName.value = `${doc.name.replace(/\.docx$/i, '')}_styled.${ext}`;
  }
}

function updateOutputExtension() {
  const ext = outputFormat.value === 'pdf' ? 'pdf' : 'docx';
  const pdfMode = outputFormat.value === 'pdf';
  pdfTheme.disabled = outputFormat.value !== 'pdf';
  const lockColors = pdfMode && useThemePalette.checked;
  primaryColor.disabled = lockColors;
  textColor.disabled = lockColors;
  outputName.value = ensureExtension(outputName.value.trim() || `document_styled.${ext}`, ext);
}

function updateLogoPreview() {
  const file = logoFile.files?.[0];
  if (!file) {
    logoPreview.removeAttribute('src');
    logoPreview.style.display = 'none';
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    logoPreview.src = reader.result;
    logoPreview.style.display = 'block';
  };
  reader.readAsDataURL(file);
}

function setStatus(message, isError = false) {
  statusText.textContent = message;
  statusText.classList.toggle('error', isError);
}

function buildFormData() {
  const doc = docFile.files?.[0];
  const format = outputFormat.value;
  const form = new FormData();

  if (doc) form.append('document', doc);

  const logo = logoFile.files?.[0];
  if (logo) form.append('logo', logo);

  form.append('outputName', outputName.value.trim() || `document_styled.${format}`);
  form.append('outputFormat', format);
  form.append('pdfTheme', pdfTheme.value || 'consulting');
  form.append('orgName', orgName.value.trim() || 'Your Organization');
  form.append('template', selectedTemplate());

  const usePalette = useThemePalette.checked && format === 'pdf';
  form.append('primaryColor', usePalette ? 'auto' : primaryColor.value);
  form.append('textColor', usePalette ? 'auto' : textColor.value);

  form.append('lineSpacing', lineSpacing.value || '1.55');
  form.append('readingWidthCh', readingWidthCh.value || '72');
  form.append('includeSummaryPage', includeSummaryPage.checked ? 'true' : 'false');
  return form;
}

function renderPreview(pages) {
  previewGrid.innerHTML = '';
  if (!pages || pages.length === 0) return;
  pages.forEach((p) => {
    const item = document.createElement('div');
    item.className = 'preview-item';
    const img = document.createElement('img');
    img.src = `data:image/png;base64,${p.pngBase64}`;
    img.alt = `Preview Seite ${p.page}`;
    const cap = document.createElement('div');
    cap.className = 'preview-caption';
    cap.textContent = `Seite ${p.page}`;
    item.appendChild(img);
    item.appendChild(cap);
    previewGrid.appendChild(item);
  });
}

async function updatePreviewNow() {
  const doc = docFile.files?.[0];
  if (!doc) {
    previewGrid.innerHTML = '';
    return;
  }
  if (outputFormat.value !== 'pdf') {
    previewGrid.innerHTML = '';
    return;
  }

  try {
    const res = await fetch('/api/preview', {
      method: 'POST',
      body: buildFormData()
    });
    if (!res.ok) {
      let msg = 'Preview fehlgeschlagen.';
      try {
        const payload = await res.json();
        if (payload?.error) msg = payload.error;
      } catch {}
      throw new Error(msg);
    }
    const payload = await res.json();
    renderPreview(payload.pages);
  } catch (e) {
    previewGrid.innerHTML = '';
    setStatus(e.message || 'Preview Fehler.', true);
  }
}

function schedulePreview() {
  if (previewTimer) window.clearTimeout(previewTimer);
  previewTimer = window.setTimeout(updatePreviewNow, 650);
}

async function generateDocument() {
  const doc = docFile.files?.[0];
  if (!doc) {
    setStatus('Bitte zuerst ein DOCX hochladen.', true);
    return;
  }

  const format = outputFormat.value;

  generateBtn.disabled = true;
  setStatus('Erzeuge Dokument...');

  try {
    const response = await fetch('/api/style', {
      method: 'POST',
      body: buildFormData()
    });

    if (!response.ok) {
      let message = 'Erzeugung fehlgeschlagen.';
      try {
        const payload = await response.json();
        if (payload?.error) message = payload.error;
      } catch {
        // ignore parse errors
      }
      throw new Error(message);
    }

    const blob = await response.blob();
    const downloadName = ensureExtension(
      outputName.value.trim() || `document_styled.${format}`,
      format === 'pdf' ? 'pdf' : 'docx'
    );
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = downloadName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    setStatus(`Fertig. ${format.toUpperCase()} Download gestartet.`);
  } catch (error) {
    setStatus(error.message || 'Fehler bei der Verarbeitung.', true);
  } finally {
    generateBtn.disabled = false;
  }
}

[docFile, logoFile].forEach((el) =>
  el.addEventListener('change', () => {
    updateFileLabels();
    updateLogoPreview();
    schedulePreview();
  })
);

[primaryColor, textColor].forEach((el) =>
  el.addEventListener('input', () => {
    updateVisualTheme();
    schedulePreview();
  })
);

outputFormat.addEventListener('change', () => {
  updateOutputExtension();
  updateFileLabels();
  schedulePreview();
});

document.querySelectorAll('input[name="template"]').forEach((el) => {
  el.addEventListener('change', () => {
    refreshTemplateCardState();
    schedulePreview();
  });
});

pdfTheme.addEventListener('change', schedulePreview);
lineSpacing.addEventListener('input', schedulePreview);
readingWidthCh.addEventListener('input', schedulePreview);
includeSummaryPage.addEventListener('change', schedulePreview);
useThemePalette.addEventListener('change', () => {
  updateOutputExtension();
  schedulePreview();
});

generateBtn.addEventListener('click', generateDocument);

updateVisualTheme();
refreshTemplateCardState();
updateFileLabels();
updateLogoPreview();
updateOutputExtension();
schedulePreview();
