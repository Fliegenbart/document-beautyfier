const inputDocx = document.getElementById('inputDocx');
const outputDocx = document.getElementById('outputDocx');
const logoPath = document.getElementById('logoPath');
const commandBox = document.getElementById('commandBox');
const copyBtn = document.getElementById('copyBtn');

function updateCommand() {
  const input = inputDocx.value.trim() || '/path/to/input.docx';
  const output = outputDocx.value.trim() || '/path/to/output.docx';
  const logo = logoPath.value.trim();

  const base = [
    'python3 "/Users/davidwegener/Desktop/Dokument-hübsch-Macher/style_whitepaper.py"',
    `  "${input}"`,
    `  "${output}"`
  ];

  if (logo) {
    base.push(`  --logo "${logo}"`);
  }

  commandBox.textContent = base.join(' \\\n');
}

[inputDocx, outputDocx, logoPath].forEach((el) => el.addEventListener('input', updateCommand));

copyBtn.addEventListener('click', async () => {
  try {
    await navigator.clipboard.writeText(commandBox.textContent);
    copyBtn.textContent = 'Kopiert';
    setTimeout(() => {
      copyBtn.textContent = 'Befehl kopieren';
    }, 1300);
  } catch (_) {
    copyBtn.textContent = 'Bitte manuell kopieren';
  }
});

updateCommand();
