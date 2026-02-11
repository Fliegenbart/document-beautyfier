const docFile = document.getElementById('docFile');
const logoFile = document.getElementById('logoFile');
const docFileName = document.getElementById('docFileName');
const logoFileName = document.getElementById('logoFileName');
const logoPreview = document.getElementById('logoPreview');

const inputDocxPath = document.getElementById('inputDocxPath');
const outputDocxPath = document.getElementById('outputDocxPath');
const logoPath = document.getElementById('logoPath');
const orgName = document.getElementById('orgName');

const primaryColor = document.getElementById('primaryColor');
const textColor = document.getElementById('textColor');
const primaryColorLabel = document.getElementById('primaryColorLabel');
const textColorLabel = document.getElementById('textColorLabel');

const commandBox = document.getElementById('commandBox');
const copyBtn = document.getElementById('copyBtn');
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

function ensurePathDefaultsFromFiles() {
  const doc = docFile.files?.[0];
  const logo = logoFile.files?.[0];

  if (doc) {
    docFileName.textContent = doc.name;
    if (!inputDocxPath.value.trim()) {
      inputDocxPath.value = `/ABSOLUTER/PFAD/${doc.name}`;
    }
    if (!outputDocxPath.value.includes('_styled.docx')) {
      outputDocxPath.value = `/ABSOLUTER/PFAD/${doc.name.replace(/\.docx$/i, '')}_styled.docx`;
    }
  }

  if (logo) {
    logoFileName.textContent = logo.name;
    if (!logoPath.value.trim()) {
      logoPath.value = `/ABSOLUTER/PFAD/${logo.name}`;
    }
  }
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

function updateVisualTheme() {
  document.documentElement.style.setProperty('--accent', primaryColor.value);
  document.documentElement.style.setProperty('--ink', textColor.value);
  primaryColorLabel.textContent = primaryColor.value;
  textColorLabel.textContent = textColor.value;
}

function updateCommand() {
  const input = inputDocxPath.value.trim() || '/ABSOLUTER/PFAD/input.docx';
  const output = outputDocxPath.value.trim() || '/ABSOLUTER/PFAD/output_styled.docx';
  const logo = logoPath.value.trim();
  const template = selectedTemplate();
  const org = orgName.value.trim() || 'Meine Organisation';

  const lines = [
    'python3 "/Users/davidwegener/Desktop/Dokument-hübsch-Macher/style_whitepaper.py"',
    `  "${input}"`,
    `  "${output}"`,
    `  --template ${template}`,
    `  --primary-color "${primaryColor.value}"`,
    `  --text-color "${textColor.value}"`,
    `  --org-name "${org.replace(/"/g, '\\"')}"`
  ];

  if (logo) {
    lines.push(`  --logo "${logo}"`);
  }

  commandBox.textContent = lines.join(' \\\n');
}

function handleInputChange() {
  ensurePathDefaultsFromFiles();
  updateLogoPreview();
  refreshTemplateCardState();
  updateVisualTheme();
  updateCommand();
}

[
  docFile,
  logoFile,
  inputDocxPath,
  outputDocxPath,
  logoPath,
  orgName,
  primaryColor,
  textColor
].forEach((el) => el.addEventListener('input', handleInputChange));

document.querySelectorAll('input[name="template"]').forEach((el) => {
  el.addEventListener('change', handleInputChange);
});

copyBtn.addEventListener('click', async () => {
  try {
    await navigator.clipboard.writeText(commandBox.textContent);
    copyBtn.textContent = 'Kopiert';
    setTimeout(() => {
      copyBtn.textContent = 'Befehl kopieren';
    }, 1200);
  } catch {
    copyBtn.textContent = 'Manuell kopieren';
  }
});

handleInputChange();
