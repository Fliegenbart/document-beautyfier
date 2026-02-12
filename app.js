const docFile = document.getElementById('docFile');
const logoFile = document.getElementById('logoFile');
const docFileName = document.getElementById('docFileName');
const logoFileName = document.getElementById('logoFileName');
const logoPreview = document.getElementById('logoPreview');

const outputFormat = document.getElementById('outputFormat');
const outputName = document.getElementById('outputName');
const orgName = document.getElementById('orgName');
const primaryColor = document.getElementById('primaryColor');
const textColor = document.getElementById('textColor');
const primaryColorLabel = document.getElementById('primaryColorLabel');
const textColorLabel = document.getElementById('textColorLabel');

const generateBtn = document.getElementById('generateBtn');
const statusText = document.getElementById('statusText');
const templateGrid = document.getElementById('templateGrid');

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

async function generateDocument() {
  const doc = docFile.files?.[0];
  if (!doc) {
    setStatus('Bitte zuerst ein DOCX hochladen.', true);
    return;
  }

  const format = outputFormat.value;
  const form = new FormData();
  form.append('document', doc);

  const logo = logoFile.files?.[0];
  if (logo) {
    form.append('logo', logo);
  }

  form.append('outputName', outputName.value.trim() || `document_styled.${format}`);
  form.append('outputFormat', format);
  form.append('orgName', orgName.value.trim() || 'Your Organization');
  form.append('template', selectedTemplate());
  form.append('primaryColor', primaryColor.value);
  form.append('textColor', textColor.value);

  generateBtn.disabled = true;
  setStatus('Erzeuge Dokument...');

  try {
    const response = await fetch('/api/style', {
      method: 'POST',
      body: form
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
  })
);

[primaryColor, textColor].forEach((el) =>
  el.addEventListener('input', () => {
    updateVisualTheme();
  })
);

outputFormat.addEventListener('change', () => {
  updateOutputExtension();
  updateFileLabels();
});

document.querySelectorAll('input[name="template"]').forEach((el) => {
  el.addEventListener('change', refreshTemplateCardState);
});

generateBtn.addEventListener('click', generateDocument);

updateVisualTheme();
refreshTemplateCardState();
updateFileLabels();
updateLogoPreview();
updateOutputExtension();
