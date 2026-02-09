const fileInput = document.getElementById('file-input');
const fileMeta = document.getElementById('file-meta');
const mapTo = document.getElementById('map-to');
const mapSubject = document.getElementById('map-subject');
const mapBody = document.getElementById('map-body');
const subjectTemplate = document.getElementById('subject-template');
const bodyTemplate = document.getElementById('body-template');
const generateButton = document.getElementById('generate');
const previewBody = document.getElementById('preview-body');
const exportJsonButton = document.getElementById('export-json');
const exportCsvButton = document.getElementById('export-csv');
const status = document.getElementById('status');

let rows = [];
let headers = [];
let drafts = [];

const setStatus = (message) => {
  status.textContent = message;
};

const resetState = () => {
  rows = [];
  headers = [];
  drafts = [];
  previewBody.innerHTML = '';
  mapTo.innerHTML = '';
  mapSubject.innerHTML = '';
  mapBody.innerHTML = '';
  mapTo.disabled = true;
  mapSubject.disabled = true;
  mapBody.disabled = true;
  generateButton.disabled = true;
  exportJsonButton.disabled = true;
  exportCsvButton.disabled = true;
  setStatus('');
};

const fillSelects = () => {
  const fragment = document.createDocumentFragment();
  const emptyOption = new Option('Select column', '');
  fragment.appendChild(emptyOption);
  headers.forEach((header) => {
    fragment.appendChild(new Option(header, header));
  });

  [mapTo, mapSubject, mapBody].forEach((select) => {
    select.innerHTML = '';
    select.appendChild(fragment.cloneNode(true));
    select.disabled = false;
  });
  generateButton.disabled = false;
};

const buildRowObjects = (sheetRows) => {
  return sheetRows
    .filter((row) => row.some((cell) => cell !== undefined && cell !== null && `${cell}`.trim() !== ''))
    .map((row) => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] ?? '';
      });
      return obj;
    });
};

const applyTemplate = (template, row) => {
  return template.replace(/{{\s*([^}]+)\s*}}/g, (_, key) => {
    const value = row[key.trim()];
    return value === undefined || value === null ? '' : String(value);
  });
};

const generateDrafts = () => {
  const toColumn = mapTo.value;
  const subjectColumn = mapSubject.value;
  const bodyColumn = mapBody.value;

  if (!toColumn || !subjectColumn || !bodyColumn) {
    setStatus('Please map all columns before generating drafts.');
    return;
  }

  drafts = rows.map((row) => {
    const subjectSource = subjectTemplate.value.trim() || String(row[subjectColumn] ?? '');
    const bodySource = bodyTemplate.value.trim() || String(row[bodyColumn] ?? '');

    return {
      to: row[toColumn] ?? '',
      subject: applyTemplate(subjectSource, row),
      body: applyTemplate(bodySource, row)
    };
  });

  renderPreview();
  exportJsonButton.disabled = drafts.length === 0;
  exportCsvButton.disabled = drafts.length === 0;
  setStatus(`Generated ${drafts.length} drafts.`);
};

const renderPreview = () => {
  previewBody.innerHTML = '';
  drafts.slice(0, 50).forEach((draft) => {
    const row = document.createElement('tr');
    ['to', 'subject', 'body'].forEach((key) => {
      const cell = document.createElement('td');
      cell.textContent = draft[key];
      row.appendChild(cell);
    });
    previewBody.appendChild(row);
  });
};

const downloadBlob = (blob, filename) => {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
};

const exportJson = () => {
  const blob = new Blob([JSON.stringify(drafts, null, 2)], { type: 'application/json' });
  downloadBlob(blob, 'email-drafts.json');
};

const exportCsv = () => {
  const headerRow = ['to', 'subject', 'body'];
  const escapeValue = (value) => {
    const text = String(value ?? '');
    if (text.includes(',') || text.includes('\n') || text.includes('"')) {
      return `"${text.replace(/"/g, '""')}"`;
    }
    return text;
  };

  const csvLines = [headerRow.join(',')];
  drafts.forEach((draft) => {
    csvLines.push(headerRow.map((key) => escapeValue(draft[key])).join(','));
  });

  const blob = new Blob([csvLines.join('\n')], { type: 'text/csv' });
  downloadBlob(blob, 'email-drafts.csv');
};

fileInput.addEventListener('change', async (event) => {
  resetState();
  const file = event.target.files?.[0];
  if (!file) {
    fileMeta.textContent = '';
    return;
  }

  fileMeta.textContent = `${file.name} (${Math.round(file.size / 1024)} KB)`;

  const arrayBuffer = await file.arrayBuffer();
  const workbook = window.xlsx.read(arrayBuffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const sheetRows = window.xlsx.utils.sheetToJson(sheet, { header: 1 });

  if (!sheetRows.length) {
    setStatus('No rows found in the spreadsheet.');
    return;
  }

  headers = sheetRows[0].map((header, index) => header || `Column ${index + 1}`);
  rows = buildRowObjects(sheetRows.slice(1));
  fillSelects();
  setStatus(`Loaded ${rows.length} rows from ${sheetName}.`);
});

generateButton.addEventListener('click', generateDrafts);
exportJsonButton.addEventListener('click', exportJson);
exportCsvButton.addEventListener('click', exportCsv);
