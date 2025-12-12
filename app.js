// Static, browser-only Excel-style Credit Checklist
// - No external libraries
// - Works when opening index.html
// - Uses fetch for data/*.json when possible, with embedded fallbacks for file://

const CHECKLIST_URL = './data/checklist.credit.json';
const SAMPLE_DEAL_URL = './data/sample.deal.json';

// Embedded fallbacks (so file:// works even if fetch is blocked)
const FALLBACK_CHECKLIST = {
  sheet: 'Credit Checklist',
  version: '1.0',
  columns: ['Item/Description', 'Status', 'CHECK', 'COMMENTS'],
  rows: [
    { id: 'deal_presentation_provided', section: 'Application', label: 'Deal presentation provided', defaultStatus: 'PENDING', rules: { dealField: 'dealPresentationProvided' } },
    { id: 'executed_application_form', section: 'Application', label: 'Executed Application form', defaultStatus: 'PENDING', rules: { dealField: 'executedApplicationForm' } },
    { id: 'credit_consent_signed', section: 'Application', label: 'Credit check consent signed', defaultStatus: 'PENDING', rules: { dealField: 'creditConsentSigned' } },

    { id: 'drivers_license_held', section: 'KYC / Identification', label: "Driver's licence held", defaultStatus: 'PENDING', rules: { dealField: 'driversLicenseHeld' } },
    { id: 'passport_held', section: 'KYC / Identification', label: 'Passport held', defaultStatus: '', rules: { dealField: 'passportHeld' } },
    { id: 'medicare_held', section: 'KYC / Identification', label: 'Medicare card held', defaultStatus: '', rules: { dealField: 'medicareHeld' } },

    { id: 'abn_checked', section: 'Business Checks', label: 'ABN checked', defaultStatus: 'PENDING', rules: { dealField: 'abnChecked' } },
    { id: 'gst_registered', section: 'Business Checks', label: 'GST registered', defaultStatus: '', rules: { dealField: 'gstRegistered' } },
    { id: 'company_search_done', section: 'Business Checks', label: 'Company search completed', defaultStatus: 'PENDING', rules: { dealField: 'companySearchDone' } },

    { id: 'equifax_clear', section: 'Credit Checks', label: 'Equifax clear', defaultStatus: 'PENDING', rules: { dealField: 'equifaxClear' } },
    { id: 'adverse_media_checked', section: 'Credit Checks', label: 'Adverse media check', defaultStatus: 'PENDING', rules: { dealField: 'adverseMediaChecked' } },

    { id: 'valuation_ordered', section: 'Valuation', label: 'Valuation ordered', defaultStatus: 'PENDING', rules: { dealField: 'valuationOrdered' } },
    { id: 'valuation_received', section: 'Valuation', label: 'Valuation received', defaultStatus: '', rules: { dealField: 'valuationReceived' } },

    { id: 'bank_statements_received', section: 'Financials', label: 'Bank statements received', defaultStatus: 'PENDING', rules: { dealField: 'bankStatementsReceived' } },
    { id: 'income_verification_complete', section: 'Financials', label: 'Income verification complete', defaultStatus: 'PENDING', rules: { dealField: 'incomeVerificationComplete' } }
  ]
};

const FALLBACK_SAMPLE_DEAL = {
  dealPresentationProvided: true,
  executedApplicationForm: true,
  creditConsentSigned: true,

  driversLicenseHeld: true,
  passportHeld: false,
  medicareHeld: true,

  abnChecked: true,
  gstRegistered: false,
  companySearchDone: true,

  equifaxClear: true,
  adverseMediaChecked: true,

  valuationOrdered: true,
  valuationReceived: false,

  bankStatementsReceived: true,
  incomeVerificationComplete: false
};

const STATUS_VALUES = ['', 'PENDING', 'COMPLETE', 'N/A'];

const state = {
  checklist: null,
  rows: new Map(), // id -> { id, section, label, status, comments }
  lastDeal: null,
};

function $(id){
  return document.getElementById(id);
}

function setMeta(text){
  const el = $('sheetMeta');
  if (el) el.textContent = text;
}

async function loadJson(url, fallback){
  try{
    const res = await fetch(url, { cache: 'no-store' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return await res.json();
  } catch (_err){
    return fallback;
  }
}

function computeCheckResult(status){
  if (status === 'PENDING') return 'Incomplete';
  if (status === 'COMPLETE') return 'Complete';
  return '';
}

function statusClass(status){
  if (status === 'PENDING') return 'status-pending';
  if (status === 'COMPLETE') return 'status-complete';
  if (status === 'N/A') return 'status-na';
  return '';
}

function createSectionRow(section){
  const tr = document.createElement('tr');
  tr.className = 'tr-section';

  const td = document.createElement('td');
  td.colSpan = 4;
  td.textContent = section;
  tr.appendChild(td);
  return tr;
}

function createChecklistRow(row){
  const tr = document.createElement('tr');
  tr.dataset.rowId = row.id;

  // Item/Description
  const tdItem = document.createElement('td');
  tdItem.textContent = row.label;
  tr.appendChild(tdItem);

  // Status
  const tdStatus = document.createElement('td');
  tdStatus.className = 'status-cell';

  const select = document.createElement('select');
  select.className = 'status-select';
  select.setAttribute('aria-label', `Status for ${row.label}`);
  for (const v of STATUS_VALUES){
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    select.appendChild(opt);
  }
  select.value = row.status || '';

  select.addEventListener('change', () => {
    updateRowState(row.id, { status: select.value });
    renderRowCells(row.id);
    setMeta(`Edited: ${row.label}`);
  });

  tdStatus.appendChild(select);
  tr.appendChild(tdStatus);

  // CHECK (computed)
  const tdCheck = document.createElement('td');
  tdCheck.className = 'check-cell';
  tdCheck.dataset.cell = 'check';
  tdCheck.textContent = computeCheckResult(row.status || '');
  tr.appendChild(tdCheck);

  // COMMENTS
  const tdComments = document.createElement('td');
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'comment-input';
  input.placeholder = '';
  input.value = row.comments || '';
  input.setAttribute('aria-label', `Comments for ${row.label}`);
  input.addEventListener('input', () => {
    updateRowState(row.id, { comments: input.value });
  });
  input.addEventListener('blur', () => {
    setMeta('Ready');
  });
  input.addEventListener('focus', () => {
    setMeta(`Editing comments: ${row.label}`);
  });

  tdComments.appendChild(input);
  tr.appendChild(tdComments);

  renderRowCells(row.id, tr);
  return tr;
}

function updateRowState(rowId, patch){
  const current = state.rows.get(rowId);
  if (!current) return;
  state.rows.set(rowId, { ...current, ...patch });
}

function renderRowCells(rowId, trEl){
  const row = state.rows.get(rowId);
  if (!row) return;
  const tr = trEl || document.querySelector(`tr[data-row-id="${cssEscape(rowId)}"]`);
  if (!tr) return;

  const statusTd = tr.children[1];
  const checkTd = tr.querySelector('[data-cell="check"]');

  // conditional formatting on Status cell
  statusTd.classList.remove('status-pending','status-complete','status-na');
  const cls = statusClass(row.status);
  if (cls) statusTd.classList.add(cls);

  // computed check
  if (checkTd) checkTd.textContent = computeCheckResult(row.status);
}

function cssEscape(value){
  // minimal escape for attribute selector
  return String(value).replace(/"/g, '\\"');
}

function initRowsFromChecklist(checklist){
  state.rows.clear();
  for (const r of checklist.rows){
    state.rows.set(r.id, {
      id: r.id,
      section: r.section,
      label: r.label,
      status: r.defaultStatus || '',
      comments: ''
    });
  }
}

function renderTable(checklist){
  const body = $('gridBody');
  body.innerHTML = '';

  let currentSection = null;
  for (const r of checklist.rows){
    if (r.section !== currentSection){
      currentSection = r.section;
      body.appendChild(createSectionRow(currentSection));
    }
    const rowState = state.rows.get(r.id);
    body.appendChild(createChecklistRow(rowState));
  }

  setMeta('Ready');
}

function applyDealToChecklist(deal){
  if (!state.checklist) return;

  state.lastDeal = deal;

  for (const r of state.checklist.rows){
    const dealField = r.rules?.dealField;
    if (!dealField) continue;

    if (!(dealField in deal)){
      // leave as-is if missing
      continue;
    }

    const v = deal[dealField];
    if (v === true) updateRowState(r.id, { status: 'COMPLETE' });
    else if (v === false) updateRowState(r.id, { status: 'PENDING' });
    else updateRowState(r.id, { status: '' });

    renderRowCells(r.id);
    // keep the select in sync
    const tr = document.querySelector(`tr[data-row-id="${cssEscape(r.id)}"]`);
    if (tr){
      const sel = tr.querySelector('select.status-select');
      if (sel) sel.value = state.rows.get(r.id).status || '';
    }
  }

  setMeta('Autofill applied');
}

function exportChecklist(){
  const exportedRows = [];

  for (const r of state.checklist.rows){
    const s = state.rows.get(r.id);
    exportedRows.push({
      id: s.id,
      section: s.section,
      label: s.label,
      status: s.status || '',
      checkResult: computeCheckResult(s.status || ''),
      comments: s.comments || ''
    });
  }

  const payload = {
    sheet: state.checklist.sheet,
    version: state.checklist.version,
    exportedAt: new Date().toISOString(),
    lastDeal: state.lastDeal || null,
    rows: exportedRows
  };

  downloadJson('credit-checklist.export.json', payload);
  setMeta('Exported');
}

function downloadJson(filename, obj){
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

async function handleDealUpload(file){
  if (!file) return;
  const text = await file.text();
  const deal = JSON.parse(text);
  applyDealToChecklist(deal);
}

function initTabs(){
  document.querySelectorAll('.tab').forEach((btn) => {
    btn.addEventListener('click', () => {
      // MVP: only Credit Checklist is implemented.
      const name = btn.getAttribute('data-sheet') || '';
      if (name !== 'Credit Checklist'){
        alert(`"${name}" is a placeholder tab in this MVP. Only "Credit Checklist" is implemented.`);
        return;
      }
    });
  });
}

async function main(){
  initTabs();

  // wire buttons
  $('btnUploadDeal').addEventListener('click', () => $('dealFileInput').click());
  $('dealFileInput').addEventListener('change', async (e) => {
    const file = e.target.files && e.target.files[0];
    try{
      await handleDealUpload(file);
    } catch (err){
      alert('Could not read deal JSON file.');
      console.error(err);
    } finally {
      e.target.value = '';
    }
  });

  $('btnLoadSample').addEventListener('click', async () => {
    const deal = await loadJson(SAMPLE_DEAL_URL, FALLBACK_SAMPLE_DEAL);
    applyDealToChecklist(deal);
  });

  $('btnExportChecklist').addEventListener('click', exportChecklist);

  // load checklist template
  setMeta('Loading checklist…');
  const checklist = await loadJson(CHECKLIST_URL, FALLBACK_CHECKLIST);
  state.checklist = checklist;
  initRowsFromChecklist(checklist);
  renderTable(checklist);
  setMeta('Ready');
}

main().catch((err) => {
  console.error(err);
  alert('Failed to load the checklist app.');
});
