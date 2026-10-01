const $ = (selector, parent = document) => parent.querySelector(selector);
const app = { state: { rows: [] }, settings: {}, filtered: [], templates: [], templatesLoaded: false,
  editor: null, page: 1, savedFilesPage: 1, previewRevision: 0, detailsBusy: false, pendingConfirmation: null, toastTimer: null };
const templateNames = { standard: 'Standard', sita: 'Sita International', siemens: 'Siemens Medical' };
const commonTemplateKeys = new Set(['supplier', 'location', 'inward_date', 'certificate_weight', 'form6', 'reference', 'cod_no', 'cod_date']);
const revisionEndpoints = new Set(['/api/create', '/api/generate', '/api/manual/save', '/api/manual/generate', '/api/manual/complete', '/api/complete', '/api/backup/restore']);
const workingDraftKey = 'ecoreco.certificate.working-draft.v1';
const fieldLabels = { plant: 'Plant', source_kind: 'Use details from', certificate_type: 'Certificate type', case_id: 'Client ID / SN', supplier: 'Company name', location: 'City', inward_date: 'Received date', certificate_weight: 'Weight to print (kg)', item: 'Item description', quantity: 'Quantity', per_piece_weight: 'Per piece weight (kg)', weight: 'Inward weight (kg)', form6: 'Form 6 number', reference: 'Reference number', cod_no: 'Certificate number', cod_date: 'Certificate date' };
function isManual(row) { return row.source_kind === 'manual'; }
function templateId(row) { return row.certificate_type || 'standard'; }
function templateLabel(row) {
  return app.templates.find(template => template.id === templateId(row))?.name || templateNames[templateId(row)] || 'Standard';
}
function normalizeDate(value) {
  const text = String(value ?? '').trim();
  if (!text) return '';
  let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (match) return validDate(Number(match[1]), Number(match[2]), Number(match[3]));
  match = text.match(/^(\d{1,2})([-/.])(\d{1,2})\2(\d{4})$/);
  if (match) return validDate(Number(match[4]), Number(match[3]), Number(match[1]));
  match = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2})$/);
  if (match) return validDate(Number(match[3]) + (Number(match[3]) >= 69 ? 1900 : 2000), Number(match[2]), Number(match[1]));
  return '';
}
function validDate(year, month, day) {
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() + 1 !== month || date.getDate() !== day) return '';
  return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function element(tag, className, value) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (value !== undefined && value !== null) node.textContent = String(value);
  return node;
}

function labelPlant(value) { return value === 'plant2' ? 'Plant 2' : value === 'plant1' ? 'Plant 1' : 'Setup'; }
function stateOf(row) {
  const raw = String(row.certificate_status || 'blank').trim().toLowerCase().replace(/[\s-]+/g, '_');
  return ({ created: 'generated', drafted: 'generated', sent: 'completed',
    check_details: 'alert', working_draft: 'alert', 'generated_manual': 'generated_manual' })[raw] || raw;
}
function hasWarnings(row) {
  return (Array.isArray(row.warnings) && row.warnings.length > 0)
    || (stateOf(row) === 'ready' && row.outcome === 'needs_review')
    || Boolean(row.system_alert && row.outcome === 'needs_review');
}
function needsAttention(row) {
  const state = stateOf(row);
  return ['alert', 'blank', 'unknown'].includes(state)
    || (state === 'ready' && hasWarnings(row))
    || Boolean(row.system_alert && row.outcome === 'needs_review');
}
function stateLabel(row) {
  const state = stateOf(row);
  if (state === 'ready' && hasWarnings(row)) return 'Alert';
  return ({ ready: 'Ready', alert: 'Alert', draft: 'Draft', generated: 'Generated', generated_manual: 'Generated Manual', completed: 'Completed', na: 'N/A', blank: 'Blank', unknown: 'Alert' })[state] || 'Alert';
}
function statusClass(row) {
  const state = stateOf(row);
  if (state === 'ready' && hasWarnings(row)) return 'status-needs_attention';
  return `status-${state === 'alert' || state === 'blank' || state === 'unknown' ? 'needs_attention' : state}`;
}
function dateLabel(value) {
  if (!value) return 'Not entered';
  const normalized = normalizeDate(value);
  if (!normalized) return value;
  const parts = normalized.split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2])).toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' });
}
function escapeText(value) { return String(value ?? '').trim(); }

function readableNumber(value) {
  if (value === '' || value == null) return '';
  const number = Number(value);
  return Number.isFinite(number) ? String(Number(number.toPrecision(12))) : String(value);
}

function simpleWarning(value) {
  const text = String(value || '').trim();
  if (/^(missing company\/supplier|Company name is missing;)/i.test(text)) return 'The company name is missing. The company line will be blank.';
  if (/^(missing or invalid inward date|Material receipt date is missing or invalid;)/i.test(text)) return 'The received date is missing or invalid. It will be blank on the PDF.';
  if (/^(missing Form 6 number \/ manifest number|Form 6 \/ manifest number is missing;)/i.test(text)) return 'The Form 6 number is missing. It will be blank on the PDF.';
  if (/^total weight must be a positive number$/i.test(text)) return 'The weight is missing or zero. Check the weight before sending the PDF.';
  if (/^one or more item weights cannot be calculated$/i.test(text)) return 'Some item weights could not be added up. Check the certificate weight.';
  if (/^missing item description on row \d+$/i.test(text)) return 'An item has no description. Add one under “Change material or Excel weight”.';
  if (/^invalid or negative quantity on row \d+$/i.test(text)) return 'An item quantity is invalid. Check it under “Change material or Excel weight”.';
  return text; // Keep unfamiliar errors and their consequences, never hide them.
}

function nextStep(row) {
  switch (stateOf(row)) {
    case 'generated': case 'generated_manual': return 'Your PDF is saved. Send it, then mark it Completed.';
    case 'completed': return 'This certificate is marked Completed.';
    case 'na': return 'N/A — no certificate is needed. This row is skipped.';
    case 'blank': return 'Blank in Excel. Automatic creation skips this row. Open Edit details if you want to create it yourself.';
    case 'draft': return 'Check your saved details, then create the PDF.';
    case 'ready': return hasWarnings(row) ? 'Check the warnings below. You can fix them or confirm and create anyway.' : 'Check the company and weight, then create the PDF.';
    default: return 'Check the Excel status before creating this certificate.';
  }
}

function toast(message, isError = false) {
  const node = $('#toast');
  node.textContent = message;
  node.classList.toggle('error', isError);
  node.classList.add('visible');
  clearTimeout(app.toastTimer);
  app.toastTimer = setTimeout(() => node.classList.remove('visible'), 3300);
}

function setBusy(on, title = 'Checking workbooks…', description = 'This usually takes a few seconds.') {
  const overlay = $('#busy-overlay');
  overlay.hidden = !on;
  $('#busy-title').textContent = title;
  $('#busy-description').textContent = description;
  for (const control of document.querySelectorAll('button, input, select, textarea')) {
    if (control.closest('dialog[open]') && !on) continue;
    control.toggleAttribute('aria-busy', on);
  }
}

async function api(path, payload) {
  if (payload !== undefined && revisionEndpoints.has(path)) payload = { ...payload, expected_revision: payload.expected_revision ?? app.workspaceRevision ?? app.state.workspace_revision ?? '' };
  const options = payload === undefined ? { method: 'GET' } : {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload),
  };
  let response;
  try { response = await fetch(path, options); }
  catch (error) {
    $('#offline-screen').hidden = false;
    throw new Error('The local dashboard server is not responding. Reopen Prepare Certificates.bat.');
  }
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error((data.error || `Request failed (${response.status}).`) + (response.status === 409 ? ' Refresh Excel and reopen the details to review the latest data. Your working draft is kept.' : ''));
  const revision = data.workspace_revision ?? data.state?.workspace_revision;
  if (typeof revision === 'string') app.workspaceRevision = revision;
  return data;
}

function updateSettingsDisplay() {
  const paths = app.settings.workbook_paths || [];
  for (let index = 0; index < 2; index++) {
    const path = paths[index] || '';
    const file = path.split(/[\\/]/).filter(Boolean).pop() || 'Not connected';
    $(`#plant${index + 1}-name`).textContent = file;
    $(`#plant${index + 1}-dot`).classList.toggle('connected', Boolean(path));
  }
  $('#settings-form [name="plant1_path"]').value = paths[0] || '';
  $('#settings-form [name="plant2_path"]').value = paths[1] || '';
  $('#settings-form [name="reminders_enabled"]').checked = Boolean(app.settings.reminders_enabled);
  $('#settings-form [name="reminder_time"]').value = app.settings.reminder_time || '11:00';
}

function setState(state) {
  app.state = state || { rows: [] };
  if (typeof app.state.workspace_revision === 'string') app.workspaceRevision = app.state.workspace_revision;
  app.state.rows ||= [];
  app.page = 1;
  app.expanded = null;
  updateSettingsDisplay();
  render();
  renderFolderInventory();
}

function certificateFileName(entry) {
  return String(entry.path || '').split(/[\\/]/).pop() || 'Unnamed PDF';
}

function certificateFileTitle(entry) {
  const stem = certificateFileName(entry).replace(/\.pdf$/i, '');
  const first = stem.split('_')[0];
  const numbered = /^\d+$/.test(first) || /^cod$/i.test(first)
    || (entry.cods || []).some(cod => String(cod).toLowerCase() === first.toLowerCase());
  const name = (numbered && stem.includes('_') ? stem.slice(first.length + 1) : stem)
    .replace(/_[12]$/, '').replace(/_/g, ' ').trim();
  return name || 'Certificate PDF';
}

function currentFileNumbers(entry) {
  return Array.isArray(entry.current_cods) ? entry.current_cods : entry.cods || [];
}

function pastCertificateNumbers(entries) {
  const identity = value => String(value).trim().toUpperCase().replace(/^0+(?=\d+$)/, '');
  const active = new Set(entries.filter(entry => entry.present).flatMap(currentFileNumbers).map(identity));
  const used = new Map();
  for (const entry of entries) {
    for (const cod of entry.cods || []) {
      const key = identity(cod);
      if (key && !active.has(key) && !used.has(key)) used.set(key, String(cod));
    }
  }
  return [...used.values()].sort((a, b) => a.localeCompare(b, undefined, {numeric: true}));
}

function renderPastNumbers() {
  const numbers = pastCertificateNumbers(app.state.certificate_inventory?.entries || []);
  const search = $('#past-number-search').value.trim().toLowerCase();
  const matches = numbers.filter(number => number.toLowerCase().includes(search));
  $('#past-number-list').replaceChildren(...matches.slice(0, 50).map(number => element('li', 'past-number', number)));
  $('#past-number-count').textContent = matches.length > 50
    ? `Showing 50 of ${matches.length} numbers. Search to find a specific number.`
    : matches.length ? `${matches.length} past number${matches.length === 1 ? '' : 's'}` : 'No matching numbers.';
}

function showSavedFile(entry) {
  $('#saved-file-name').textContent = certificateFileTitle(entry);
  const facts = $('#saved-file-facts');
  facts.replaceChildren();
  for (const [label, value] of [['Certificate number', currentFileNumbers(entry).join(', ') || 'Could not be read'],
    ['Reference number', entry.references?.join(', ') || 'Not entered (optional)'], ['PDF file', certificateFileName(entry)]]) {
    const fact = element('div'); fact.append(element('dt', '', label), element('dd', '', value)); facts.append(fact);
  }
  for (const warning of entry.warnings || []) facts.append(element('p', 'inventory-warning', warning));
  $('#saved-file-modal').showModal();
}

function renderFolderInventory() {
  const inventory = app.state.certificate_inventory;
  const section = $('#folder-inventory');
  section.hidden = !inventory;
  if (!inventory) return;
  const entries = inventory.entries || [];
  const present = entries.filter(entry => entry.present);
  const warnings = inventory.warnings || [];
  $('#folder-inventory-summary').textContent = warnings.length
    ? `Saved certificates (${present.length}) · ${warnings.length} file warning${warnings.length === 1 ? '' : 's'}`
    : `Saved certificates (${present.length})`;
  $('#folder-inventory-note').textContent = present.length
    ? 'These PDFs are in your certificate folder.' : 'No saved PDFs yet. Generated certificates will appear here.';
  const notes = $('#folder-inventory-warnings');
  notes.replaceChildren();
  if (warnings.length) {
    notes.append(element('strong', '', 'Some files need checking before you create another certificate.'));
    const issues = element('ul');
    for (const warning of warnings) issues.append(element('li', 'inventory-warning', warning));
    notes.append(issues);
  }
  const list = $('#folder-inventory-files');
  list.replaceChildren();
  const page = visiblePage(present, app.savedFilesPage);
  app.savedFilesPage = page.page;
  for (const entry of page.rows) {
    const item = element('li', 'saved-file');
    const copy = element('div', 'saved-file-copy');
    const title = element('strong', '', certificateFileTitle(entry));
    title.title = certificateFileTitle(entry);
    copy.append(title, element('span', '', `Certificate ${currentFileNumbers(entry).join(', ') || 'number not read'}`));
    const button = element('button', 'button button-quiet', 'File details');
    button.type = 'button';
    button.setAttribute('aria-label', `File details for ${certificateFileTitle(entry)}`);
    button.addEventListener('click', () => showSavedFile(entry));
    item.append(copy, button);
    list.append(item);
  }
  $('#saved-file-pages').hidden = page.pages === 1;
  $('#saved-file-page-count').textContent = `Page ${page.page} of ${page.pages}`;
  $('#saved-files-previous').disabled = page.page === 1;
  $('#saved-files-next').disabled = page.page === page.pages;
  const past = pastCertificateNumbers(entries);
  $('#show-past-numbers').hidden = past.length === 0;
  $('#show-past-numbers').textContent = `Past numbers (${past.length})`;
  if ($('#past-numbers-modal').open) renderPastNumbers();
}

function summarize() {
  const rows = app.state.rows;
  const ready = rows.filter(row => stateOf(row) === 'ready' && !needsAttention(row)).length;
  const selectedPlant = $('#plant-filter').value;
  const workbookReady = rows.filter(row => !isManual(row) && stateOf(row) === 'ready'
    && (selectedPlant === 'all' || row.plant === selectedPlant)).length;
  const attention = rows.filter(needsAttention).length;
  const generated = rows.filter(row => stateOf(row) === 'generated').length;
  const generatedManual = rows.filter(row => stateOf(row) === 'generated_manual').length;
  const completed = rows.filter(row => stateOf(row) === 'completed').length;
  const overdue = rows.filter(row => Boolean(row.overdue)).length;
  $('#count-ready').textContent = ready;
  $('#count-attention').textContent = attention;
  $('#count-generated').textContent = generated;
  $('#count-generated-manual').textContent = generatedManual;
  $('#count-completed').textContent = completed;
  $('#count-overdue').textContent = overdue;
  $('#side-ready').textContent = rows.length;
  $('#side-attention').textContent = attention;
  $('#generate-count').textContent = workbookReady ? `(${workbookReady})` : '';
  $('#generate-all').disabled = workbookReady === 0;
}

function bulkGenerationPayload() {
  const selectedPlant = $('#plant-filter').value;
  return { plants: selectedPlant === 'all' ? ['plant1', 'plant2'] : [selectedPlant] };
}

function compareRows(a, b) {
  const rank = row => row.overdue ? 0 : (stateOf(row) === 'ready' && hasWarnings(row)) ? 1
    : stateOf(row) === 'ready' ? 2 : ['generated', 'generated_manual'].includes(stateOf(row)) ? 3
      : stateOf(row) === 'blank' || stateOf(row) === 'unknown' ? 4
        : stateOf(row) === 'completed' ? 5 : 6;
  return rank(a) - rank(b) || String(a.supplier || '').localeCompare(String(b.supplier || ''));
}

function matches(row) {
  const plant = $('#plant-filter').value;
  const status = $('#status-filter').value;
  const template = $('#template-filter').value;
  const search = $('#search').value.trim().toLowerCase();
  if (plant !== 'all' && row.plant !== plant) return false;
  if (template !== 'all' && templateId(row) !== template) return false;
  if (status === 'ready' && (stateOf(row) !== 'ready' || hasWarnings(row))) return false;
  if (status === 'needs_attention' && !needsAttention(row)) return false;
  if (status !== 'all' && !['ready', 'needs_attention'].includes(status) && stateOf(row) !== status) return false;
  if (search) {
    const searchable = [row.supplier, row.sn, row.case_id, row.form6, row.reference, row.certificate_no, row.plant, row.location, templateLabel(row), isManual(row) ? 'local register' : 'plant workbook']
      .map(value => String(value || '').toLowerCase()).join(' ');
    if (!searchable.includes(search)) return false;
  }
  return true;
}

function cell(value, className = '') {
  const td = element('td');
  const content = element('span', value ? className : `muted-value ${className}`, value || '—');
  td.append(content);
  return td;
}

function makeClientCell(row, index) {
  const td = element('td');
  const wrap = element('div', 'client-cell');
  const copy = element('span', 'client-copy');
  copy.append(element('strong', '', row.supplier || 'Company name missing'));
  copy.append(element('small', '', `${row.certificate_no ? `Certificate #${row.certificate_no}` : 'Number chosen automatically'} · ${templateLabel(row)}`));
  wrap.append(copy);
  td.append(wrap);
  return td;
}

function makeStatus(row) {
  const td = element('td');
  td.append(element('span', `status-pill ${statusClass(row)}`, stateLabel(row)));
  if (row.overdue) td.append(element('small', 'age-note', 'Over 30 days since received'));
  return td;
}

function actionButton(title, action, rowIndex, primary = false) {
  const button = element('button', `row-action${primary ? ' primary-row-action' : ''}`, title);
  button.type = 'button';
  button.dataset.action = action;
  button.dataset.index = String(rowIndex);
  return button;
}

function makeActions(row, index) {
  const td = element('td');
  const wrap = element('div', 'row-actions');
  if (stateOf(row) === 'ready' || (isManual(row) && ['blank', 'draft'].includes(stateOf(row)))) wrap.append(actionButton('Review & create', 'details', index, true));
  else if (['generated', 'generated_manual'].includes(stateOf(row))) wrap.append(actionButton('Mark Completed', 'complete', index, true));
  else wrap.append(actionButton('Edit details', 'details', index));
  const more = element('button', 'row-action', app.expanded === index ? 'Less' : 'More');
  more.type = 'button';
  more.dataset.expand = String(index);
  more.setAttribute('aria-expanded', String(app.expanded === index));
  more.setAttribute('aria-label', `${app.expanded === index ? 'Less' : 'More'} about ${row.supplier || `client ${row.sn}`}`);
  wrap.append(more);
  td.append(wrap);
  return td;
}

function makeDetailRow(row, index) {
  const tr = element('tr', 'detail-row');
  const td = element('td');
  td.colSpan = 5;
  const content = element('div', 'detail-content');
  content.append(element('p', 'next-step', nextStep(row)));
  const facts = element('dl', 'certificate-facts');
  for (const [label, value] of [['Weight to print', row.display_weight != null ? `${readableNumber(row.display_weight)} kg` : 'Not entered'], ['Form 6', row.form6 || 'Not entered'], ['Reference', row.reference || 'Not entered (optional)'], ['Excel row ID', row.sn || 'Manual certificate']]) {
    const fact = element('div'); fact.append(element('dt', '', label), element('dd', '', value)); facts.append(fact);
  }
  content.append(facts);
  const issues = rowWarningSummary(row);
  if (issues.length) {
    const warning = element('div', 'detail-warnings');
    warning.append(element('strong', '', ['generated', 'generated_manual', 'completed'].includes(stateOf(row)) ? 'Warnings accepted when generated' : 'Check before creating'));
    const list = element('ul');
    for (const issue of issues) list.append(element('li', '', simpleWarning(issue)));
    warning.append(list);
    content.append(warning);
  }
  const items = Array.isArray(row.items) ? row.items : [];
  if (items.length) {
    const materials = element('details', 'technical-details');
    materials.append(element('summary', '', `Material list (${items.length} item${items.length === 1 ? '' : 's'})`));
    const list = element('ul');
    for (const item of items) list.append(element('li', '', `${item.item || 'No description'}${item.weight != null ? ` · ${readableNumber(item.weight)} kg` : ''}`));
    materials.append(list); content.append(materials);
  }
  const actions = element('div', 'row-actions');
  actions.style.justifyContent = 'flex-start';
  actions.style.marginTop = '7px';
  actions.append(actionButton('Edit details', 'details', index));
  if (!isManual(row)) actions.append(actionButton('Open plant workbook', 'open-workbook', index));
  else actions.append(actionButton('Copy as new', 'copy', index));
  content.append(actions);
  const technical = element('details', 'technical-details');
  technical.append(element('summary', '', 'Technical details'));
  technical.append(element('p', '', row.detail || 'No additional notes.'));
  if (issues.length) technical.append(element('p', '', issues.join(' · ')));
  content.append(technical);
  td.append(content);
  tr.append(td);
  return tr;
}

function visiblePage(items, requested = 1) {
  const pages = Math.max(1, Math.ceil(items.length / 10));
  const page = Math.max(1, Math.min(requested, pages));
  const start = (page - 1) * 10;
  return { rows: items.slice(start, start + 10), page, pages, start };
}

function render() {
  summarize();
  const view = $('#status-filter').value === 'needs_attention' ? 'attention' : 'queue';
  $('#page-title').textContent = view === 'attention' ? 'Alerts' : 'Your certificates';
  $('#page-description').textContent = view === 'attention'
    ? 'Open a company to fix missing details or create with a warning.'
    : 'Find a company, check its details, then create the certificate.';
  $('#breadcrumb-view').textContent = view === 'attention' ? 'Alerts' : 'Certificates';
  $('#queue-title').textContent = $('#status-filter').selectedOptions[0]?.textContent || 'All companies';
  $('#queue-subtitle').textContent = ({
    generated: 'Open the certificate folder to find your PDFs. Send them, then mark them Completed.',
    generated_manual: 'These PDFs were made manually. Send them, then mark them Completed.',
    completed: 'These certificates are marked Completed.',
    na: 'No certificate is needed for these rows. Automatic creation skips them.',
    needs_attention: 'Review the alert. You can fix the details or confirm creation with a warning.',
  })[$('#status-filter').value] || 'Choose “Review & create” to start. Nothing is emailed by this app.';
  document.querySelectorAll('[data-status]').forEach(button => button.setAttribute('aria-pressed', String(button.dataset.status === $('#status-filter').value)));

  const filtered = app.state.rows.map((row, index) => ({ row, index })).filter(item => matches(item.row));
    filtered.sort((a, b) => compareRows(a.row, b.row));
  app.filtered = filtered;
  const page = visiblePage(filtered, app.page);
  app.page = page.page;
  const tbody = $('#queue-body');
  tbody.replaceChildren();
  if (!filtered.length) {
    const tr = element('tr');
    const td = element('td'); td.colSpan = 5;
    const empty = element('div', 'empty-state');
    empty.append(element('span', 'empty-icon', '✓'));
    empty.append(element('strong', '', $('#search').value.trim() ? 'No matching certificates' : 'You’re all caught up'));
    empty.append(element('span', '', $('#search').value.trim() ? 'Try another client name, ID, or reference.' : 'Refresh the workbooks or add a certificate to get started.'));
    td.append(empty); tr.append(td); tbody.append(tr);
  } else {
    for (const { row, index } of page.rows) {
      const tr = element('tr', 'data-row');
      tr.dataset.row = String(index);
      tr.append(makeClientCell(row, index));
      const plantCell = element('td');
      plantCell.append(element('span', `plant-badge ${row.plant === 'plant2' ? 'plant2' : ''}`, labelPlant(row.plant)));
      tr.append(plantCell);
      tr.append(cell(dateLabel(row.inward_date), 'field-value'));
      tr.append(makeStatus(row));
      tr.append(makeActions(row, index));
      tbody.append(tr);
      if (app.expanded === index) tbody.append(makeDetailRow(row, index));
    }
  }
  $('#visible-count').textContent = filtered.length
    ? `${page.start + 1}–${page.start + page.rows.length} of ${filtered.length} companies`
    : 'No matching companies';
  $('#page-controls').hidden = page.pages === 1;
  $('#page-count').textContent = `Page ${page.page} of ${page.pages}`;
  $('#previous-page').disabled = page.page === 1;
  $('#next-page').disabled = page.page === page.pages;
  $('#side-link-queue').classList.toggle('active', view === 'queue');
  $('#side-link-attention').classList.toggle('active', view === 'attention');
}

function setView(view) {
  app.page = 1;
  app.expanded = null;
  $('#search').value = '';
  $('#plant-filter').value = 'all';
  $('#template-filter').value = 'all';
  if (view === 'attention') $('#status-filter').value = 'needs_attention';
  else $('#status-filter').value = 'all';
  render();
}

function showBusyForApi(label, fn) {
  return async (...args) => {
    setBusy(true, label);
    try { return await fn(...args); }
    finally { setBusy(false); }
  };
}

async function refreshState() {
  setBusy(true, 'Checking plant workbooks…', 'Reading rows and rebuilding your certificate queue.');
  try {
    const response = await api('/api/check', {});
    setState(response.state);
    toast('Workbooks refreshed.');
  } catch (error) { toast(error.message, true); }
  finally { setBusy(false); }
}

function rowWarningSummary(row) {
  const detail = String(row.detail || '').trim();
  if (Array.isArray(row.warnings) && row.warnings.length) return row.warnings;
  if (row.outcome === 'needs_review' && stateOf(row) === 'ready') return [detail || 'Review this row before issuing.'];
  if (row.system_alert && detail && stateOf(row) !== 'generated') return [detail];
  return [];
}

function showWarnings(warnings, continuation) {
  const entries = warnings.map(row => typeof row === 'string' ? { supplier: continuation.payload?.details?.supplier || 'Certificate', warnings: [row] } : row);
  app.pendingConfirmation = { ...continuation, hasWarnings: entries.length > 0 };
  $('#override-reason').value = '';
  $('#override-reason').required = entries.length > 0;
  $('#warning-reason-hint').textContent = entries.length ? 'Required when accepting warnings' : 'Optional — no warnings';
  $('#warning-reason-error').hidden = true;
  $('#warning-reason-error').textContent = 'Enter a reason before continuing.';
  const list = $('#warning-list');
  list.replaceChildren();
  for (const row of entries) {
    const item = element('div', 'warning-item');
    item.append(element('strong', '', `${row.plant ? `${labelPlant(row.plant)} · ` : ''}${row.supplier || row.company || `Client ${row.sn ?? ''}`}`));
    const issues = rowWarningSummary(row);
    const details = issues.length ? issues : [row.detail || 'Please review this row.'];
    const bullets = element('ul');
    for (const issue of details) bullets.append(element('li', '', simpleWarning(issue)));
    item.append(bullets);
    list.append(item);
  }
  if (!warnings.length) list.append(element('div', 'warning-item', 'No warnings were found.'));
  $('#warning-modal').showModal();
}

async function submitGeneration(payload, confirmed = false, endpoint = '/api/generate', fromDetails = false) {
  if (app.generationBusy) return;
  app.generationBusy = true;
  payload = { ...payload, expected_revision: payload.expected_revision ?? app.workspaceRevision ?? app.state.workspace_revision ?? '' };
  if (fromDetails) setDetailsBusy(true);
  setBusy(true, confirmed ? 'Creating certificates…' : 'Reviewing certificate details…', confirmed
    ? 'Saving the PDFs and updating your certificates.'
    : 'Checking for missing or unusual details first.');
  try {
    const response = await api(endpoint, { ...payload, confirmed });
    if (response.state) setState(response.state);
    if (fromDetails && response.record_id) app.editor.recordId = response.record_id;
    if (response.requires_confirmation) {
      const continuationPayload = { ...payload, expected_revision: response.workspace_revision ?? payload.expected_revision,
        ...(endpoint === '/api/manual/generate' && response.record_id ? { record_id: response.record_id } : {}) };
      showWarnings(response.warnings || [], { endpoint, payload: continuationPayload, fromDetails, revision: app.previewRevision, reviewToken: response.review_token });
      return;
    }
    const created = response.created || response.certificates?.length || 0;
    if (created) {
      if (fromDetails) { clearWorkingDraft(); app.editor.baseline = detailsBody($('#details-form')); $('#details-modal').close(); }
      showGenerationSuccess(response);
    } else toast('No certificates were generated. Check the selected rows.');
  } catch (error) {
    if (fromDetails) { $('#details-error').textContent = error.message; $('#details-error').hidden = false; }
    toast(error.message, true);
  }
  finally { app.generationBusy = false; if (fromDetails) setDetailsBusy(false); setBusy(false); }
}

async function loadTemplates() {
  if (app.templatesLoaded) return true;
  if (app.templatesRequest) return app.templatesRequest;
  app.templatesRequest = (async () => {
    try {
      const response = await api('/api/templates');
      if (!Array.isArray(response.templates) || !Object.keys(templateNames).every(id => response.templates.some(template => template.id === id && Array.isArray(template.fields)))) {
        throw new Error('Certificate templates are not available yet. Try again shortly.');
      }
      app.templates = response.templates;
      app.templatesLoaded = true;
      for (const select of [$('#template-filter'), $('#details-form').elements.namedItem('certificate_type')]) {
        for (const option of select.options) {
          const spec = app.templates.find(template => template.id === option.value);
          if (spec) option.textContent = spec.name;
        }
      }
      render();
      return true;
    } catch (error) { toast(error.message, true); return false; }
    finally { app.templatesRequest = null; }
  })();
  return app.templatesRequest;
}

function captureTemplateFields() {
  if (!app.editor) return;
  const values = app.editor.fields[app.editor.template] ||= {};
  for (const input of $('#template-editor').querySelectorAll('[data-certificate-field]')) {
    values[input.dataset.certificateField] = input.value;
  }
}

function updateAnnexureDisplay() {
  const values = app.editor.fields[app.editor.template] || {};
  const attachment = app.editor.annexures[values._annexure];
  const name = $('#annexure-name');
  name.hidden = !values._annexure;
  name.textContent = attachment ? `${attachment.name}${attachment.pages != null ? ` · ${attachment.pages} page${attachment.pages === 1 ? '' : 's'}` : ''}` : 'A PDF list is attached.';
  $('#remove-annexure').hidden = !values._annexure;
}

async function pickAnnexure() {
  if (app.detailsBusy) return;
  captureTemplateFields();
  setDetailsBusy(true);
  setBusy(true, 'Choose a PDF list…', 'The Windows file picker will open.');
  try {
    // The picker must validate this snapshot before saving the attachment.
    // Its returned revision covers that attachment as well as the source data.
    const response = await api('/api/pick-annexure', { expected_revision: app.editor.workspaceRevision });
    if (!response.annexure) return;
    if (typeof response.workspace_revision === 'string') app.editor.workspaceRevision = response.workspace_revision;
    const attachment = response.annexure;
    app.editor.annexures[attachment.id] = attachment;
    app.editor.annexureId = attachment.id;
    app.editor.fields[app.editor.template]._annexure = attachment.id;
    updateAnnexureDisplay();
    invalidatePreview();
  } catch (error) { toast(error.message, true); }
  finally { setDetailsBusy(false); setBusy(false); }
}

function renderTemplateEditor() {
  const form = $('#details-form');
  const id = form.elements.namedItem('certificate_type').value;
  const spec = app.templates.find(template => template.id === id);
  const supplier = form.elements.namedItem('supplier');
  if (!app.editor.existing && !app.editor.touched.has('supplier') && (!supplier.value || supplier.value === app.editor.autoCompany)) {
    supplier.value = spec?.default_company || '';
    app.editor.autoCompany = supplier.value;
  }
  const container = $('#template-editor');
  container.replaceChildren();
  let printOptions;
  app.editor.template = id;
  $('#form6-hint').textContent = id === 'standard' ? 'Expected' : 'Optional';
  const values = app.editor.fields[id] ||= {};
  // Do not carry old text/layout overrides into a fixed-design certificate.
  const editableKeys = new Set((spec?.fields || []).map(field => field.key));
  for (const key of Object.keys(values)) if (key !== '_annexure' && !editableKeys.has(key)) delete values[key];
  if (app.editor.annexureId) values._annexure = app.editor.annexureId;
  else delete values._annexure;
  const status = $('#template-status');
  status.hidden = Boolean(spec);
  status.textContent = 'Template fields could not be loaded. Reopen the form to try again.';
  for (const group of ['details', 'wording']) {
    let headingAdded = false;
    for (const field of spec?.fields || []) {
      if (field.group !== group) continue;
      if (group === 'wording' && field.key !== 'company_display') continue;
      if (group === 'details' && commonTemplateKeys.has(field.key)) {
        const input = form.elements.namedItem(field.key);
        if (input && !app.editor.existing && !input.value && !app.editor.touched.has(field.key)) {
          input.value = input.type === 'date' ? normalizeDate(field.default) : String(field.default ?? '');
        }
        continue;
      }
      const optionalPrinted = ['company_display', 'weight_display'].includes(field.key);
      if (!headingAdded && !optionalPrinted) {
        if (group === 'details') container.append(element('div', 'form-section-title', 'Certificate details'));
        headingAdded = true;
      }
      const label = element('label', 'field');
      label.append(element('span', '', field.label || field.key));
      const input = element(field.multiline ? 'textarea' : 'input');
      if (field.multiline) input.rows = 3;
      else if (group === 'details' && (field.type === 'date' || field.key.endsWith('_date'))) input.type = 'date';
      input.dataset.certificateField = field.key;
      const value = Object.hasOwn(values, field.key) ? values[field.key] : field.key === 'company_display' ? '' : field.default;
      input.value = input.type === 'date' ? normalizeDate(value) : String(value ?? '');
      if (field.key === 'company_display') input.placeholder = 'Leave blank to use Company / supplier and City';
      label.append(input);
      if (optionalPrinted) {
        if (!printOptions) {
          const disclosure = element('details', 'form-disclosure');
          disclosure.append(element('summary', '', 'Use a different printed name or weight label'));
          printOptions = element('div', 'form-grid');
          disclosure.append(printOptions);
          container.append(disclosure);
        }
        printOptions.append(label);
      } else container.append(label);
    }
  }
  updateAnnexureDisplay();
  updateDetailsButtons();
}

function updateSourceControls() {
  const form = $('#details-form');
  const manual = form.elements.namedItem('source_kind').value === 'manual';
  const kind = form.elements.namedItem('certificate_type').value;
  const existingWorkbook = app.editor?.existing && app.editor?.sourceSn != null && !manual;
  $('#edit-scope-field').hidden = manual;
  form.elements.namedItem('edit_scope').querySelector('[value="certificate_only"]').disabled = !existingWorkbook;
  if (!manual && !existingWorkbook) form.elements.namedItem('edit_scope').value = 'update_excel';
  form.elements.namedItem('case_id').readOnly = Boolean(existingWorkbook);
  form.elements.namedItem('plant').disabled = Boolean(existingWorkbook);
  if (existingWorkbook) {
    form.elements.namedItem('case_id').value = app.editor.sourceSn;
    form.elements.namedItem('plant').value = app.editor.sourcePlant;
  }
  $('#edit-scope-hint').textContent = existingWorkbook
    ? 'Excel still records the certificate number, date and Generated status. Cancel does not save changes.'
    : 'A new Excel row is saved only when the PDF is created successfully.';
  form.elements.namedItem('case_id').required = !manual;
  $('#case-required').hidden = manual;
  $('#case-optional').hidden = !manual;
  $('#pick-workbook').hidden = manual || Boolean(existingWorkbook);
  $('#details-plant-label').textContent = manual ? 'Plant' : 'Plant *';
  $('#material-options').hidden = manual;
  $('#save-draft').hidden = !manual;
  for (const name of ['case_id', 'item', 'quantity', 'per_piece_weight', 'weight']) form.elements.namedItem(name).closest('.field').hidden = manual;
  form.elements.namedItem('inward_date').closest('.field').hidden = manual && kind === 'siemens';
  form.elements.namedItem('form6').closest('.field').hidden = kind !== 'standard';
  form.elements.namedItem('reference').closest('.field').hidden = manual && kind === 'siemens';
  form.elements.namedItem('certificate_weight').closest('.field').querySelector('span').textContent = manual ? 'Weight to print (kg)' : 'Weight to print (kg) · Optional';
  form.elements.namedItem('certificate_weight').placeholder = manual ? 'Enter the weight in kg' : app.editor.inwardPlaceholder;
  $('#edit-items-hint').hidden = manual || !app.editor.itemHint;
  $('#details-description').textContent = manual
    ? 'Enter the company details below. You do not need an Excel row.'
    : existingWorkbook ? 'These details came from Excel. Change only what you need, then check the preview.'
    : 'Enter the company details. The app will save a new Excel row with the PDF.';
  $('#save-details').textContent = manual ? 'Create PDF' : form.elements.namedItem('edit_scope').value === 'certificate_only'
    ? 'Create PDF only' : 'Create PDF & update Excel';
}

function updateDetailsButtons() {
  const unavailable = app.detailsBusy || !app.templatesLoaded || !app.templates.some(template => template.id === $('#details-form').elements.namedItem('certificate_type').value);
  for (const id of ['save-details', 'save-draft', 'preview-details', 'pick-annexure', 'remove-annexure', 'keep-working-draft']) $(`#${id}`).disabled = unavailable;
}

function setDetailsBusy(on) {
  app.detailsBusy = on;
  $('#details-form').inert = on;
  updateDetailsButtons();
}

function invalidatePreview() {
  app.previewRevision++;
  $('#preview-frame').removeAttribute('src');
  $('#preview-frame').hidden = true;
  $('#preview-image').removeAttribute('src');
  $('#preview-image').hidden = true;
  $('#preview-image-note').hidden = true;
  $('#preview-full-pdf').textContent = 'View full PDF here';
  $('#preview-full-pdf').setAttribute('aria-expanded', 'false');
  $('#preview-link').removeAttribute('href');
  $('#preview-document').hidden = true;
  $('#preview-number').hidden = true;
  $('#preview-warnings').hidden = true;
  $('#preview-warnings').replaceChildren();
  $('#preview-status').textContent = 'Preview your current details before creating the final PDF.';
  if (app.editor?.baseline && $('#details-modal').open) {
    renderChangeSummary();
    if (editorDirty()) saveWorkingDraft();
    else if (app.editor.workingDraftSaved && clearWorkingDraft()) {
      app.editor.workingDraftSaved = false;
      $('#working-draft-status').textContent = 'No unsaved field changes. Excel is unchanged.';
    }
  }
}

async function openDetails(row = null, copyAsNew = false, skipRecovery = false) {
  if ($('#details-modal').open || app.detailsBusy) return;
  if (!skipRecovery && offerDraftRecovery(row, copyAsNew, true)) return;
  const form = $('#details-form');
  form.reset();
  for (const disclosure of form.querySelectorAll('details')) disclosure.open = false;
  const id = templateId(row || {});
  app.editor = { existing: Boolean(row) && !copyAsNew, recordId: copyAsNew ? null : row?.record_id || null, template: id,
    sourceSn: row && !isManual(row) && !copyAsNew ? String(row.sn) : null, sourcePlant: row?.plant || null,
    fields: { [id]: { ...(row?.certificate_fields || {}) } }, annexures: {}, annexureId: row?.certificate_fields?._annexure || null,
    touched: new Set(), sourceExplicit: false, rawNumeric: {}, rawDetails: isManual(row || {}) ? { ...(row.raw_details || {}) } : {}, workspaceRevision: app.workspaceRevision ?? app.state.workspace_revision ?? '' };
  form.elements.namedItem('certificate_type').value = id;
  form.elements.namedItem('source_kind').value = row ? (isManual(row) ? 'manual' : 'workbook') : 'workbook';
  form.elements.namedItem('edit_scope').value = app.editor.sourceSn != null ? 'certificate_only' : 'update_excel';
  invalidatePreview();
  $('#details-title').textContent = copyAsNew ? 'Copy into a new certificate' : row ? `Check ${row.supplier || `Client ${row.sn}`}` : 'Create a certificate';
  $('#save-details').innerHTML = row ? 'Save changes and create PDF <span>→</span>' : 'Save and create PDF <span>→</span>';
  const items = row && Array.isArray(row.items) ? row.items : [];
  const hint = $('#edit-items-hint');
  if (items.length > 1) {
    hint.hidden = false;
    hint.textContent = `${items.length} items from Excel are included. “Change material or Excel weight” edits the first item; the others stay unchanged.`;
  } else hint.hidden = true;
  if (row) {
    const values = {
      plant: row.plant, case_id: isManual(row) ? row.case_id || '' : row.sn ?? row.case_id, supplier: row.supplier, location: row.location,
      inward_date: row.inward_date, item: items[0]?.item ?? row.item ?? '', quantity: items[0]?.quantity ?? row.quantity,
      per_piece_weight: items[0]?.per_piece_weight ?? row.per_piece_weight, weight: items[0]?.weight ?? row.weight,
      certificate_weight: row.certificate_weight, form6: row.form6, reference: row.reference,
      cod_no: row.cod_no ?? row.certificate_no, cod_date: row.cod_date ?? row.certificate_date,
    };
    if (isManual(row)) Object.assign(values, row.raw_details || {});
    for (const [name, value] of Object.entries(values)) {
      const input = form.elements.namedItem(name);
      if (input) {
        let normalized = value == null ? '' : String(value);
        if (input.type === 'date') normalized = normalizeDate(value);
        if (['quantity', 'per_piece_weight', 'weight', 'certificate_weight'].includes(name) && normalized !== '' && Number.isFinite(Number(normalized))) {
          app.editor.rawNumeric[name] = normalized;
          normalized = readableNumber(normalized);
        }
        input.value = normalized;
      }
    }
    if (copyAsNew) {
      for (const name of ['case_id', 'cod_no', 'cod_date', 'inward_date']) {
        form.elements.namedItem(name).value = '';
        delete app.editor.rawDetails[name];
      }
      hint.hidden = false;
      hint.textContent = 'Copied the previous details. Review dates, weights and serial numbers for this new certificate. A new COD will be assigned.';
    }
  } else {
    const localDate = new Date();
    localDate.setMinutes(localDate.getMinutes() - localDate.getTimezoneOffset());
    form.elements.namedItem('inward_date').value = localDate.toISOString().slice(0, 10);
  }
  const inwardTotal = items.length && items.every(item => item.weight != null && Number.isFinite(Number(item.weight)))
    ? items.reduce((sum, item) => sum + Number(item.weight), 0) : null;
  app.editor.itemHint = !hint.hidden ? hint.textContent : '';
  app.editor.inwardPlaceholder = inwardTotal != null
    ? `Leave blank to use ${readableNumber(inwardTotal)} kg from Excel` : 'Enter weight, or leave blank to use Excel';
  updateSourceControls();
  renderTemplateEditor();
  app.editor.baseline = detailsBody(form);
  renderChangeSummary();
  $('#working-draft-status').textContent = 'Changes stay in this form until you create the PDF.';
  $('#details-error').hidden = true;
  $('#details-modal').showModal();
  const editor = app.editor;
  await loadTemplates();
  if (app.editor === editor && $('#details-modal').open) {
    const changed = editorDirty();
    captureTemplateFields(); renderTemplateEditor();
    if (!changed) editor.baseline = detailsBody(form);
    renderChangeSummary();
  }
}

function comparisonFields(payload) {
  const fields = { plant: payload.plant, source_kind: payload.source_kind, ...(payload.details || {}) };
  delete fields.certificate_fields;
  for (const [key, value] of Object.entries(payload.details?.certificate_fields || {})) fields[`certificate_fields.${key}`] = value;
  return fields;
}

function fieldDifferences(before, after) {
  const original = comparisonFields(before || {}), current = comparisonFields(after || {});
  return [...new Set([...Object.keys(original), ...Object.keys(current)])].filter(key => String(original[key] ?? '') !== String(current[key] ?? ''))
    .map(key => ({ key, before: String(original[key] ?? ''), after: String(current[key] ?? '') }));
}

function editorDirty() {
  if (!app.editor?.baseline || !$('#details-modal').open) return false;
  const current = detailsBody($('#details-form'));
  return fieldDifferences(app.editor.baseline, current).length > 0 || current.edit_scope !== app.editor.baseline.edit_scope;
}

function scopeDescription(payload) {
  if (payload.source_kind === 'manual') return 'PDF + local certificate tracking. No Excel row is changed.';
  if (payload.edit_scope === 'certificate_only') return 'Edited details go to the PDF only. Excel still records the certificate number, date and Generated status.';
  return app.editor?.sourceSn != null
    ? 'Edited details go to the PDF and Excel. Excel also records the certificate number, date and Generated status.'
    : 'A new Excel row and PDF are created together. Excel records the certificate number, date and Generated status.';
}

function renderChangeSummary() {
  if (!app.editor?.baseline) return;
  const payload = detailsBody($('#details-form'));
  $('#change-scope').textContent = scopeDescription(payload);
  const changes = fieldDifferences(app.editor.baseline, payload);
  const list = $('#change-list'); list.replaceChildren();
  for (const change of changes) {
    const key = change.key.replace(/^certificate_fields\./, '');
    const spec = app.templates.find(template => template.id === app.editor.template)?.fields.find(field => field.key === key);
    const label = key === '_annexure' ? 'Attached PDF list' : fieldLabels[key] || spec?.label || key.replace(/_/g, ' ');
    const display = value => key === '_annexure' && value ? app.editor.annexures[value]?.name || value : value || '(blank)';
    const entry = element('div', 'field-change');
    const values = element('dd');
    values.append(element('span', 'change-before', `Before: ${display(change.before)}`), element('span', 'change-after', `After: ${display(change.after)}`));
    entry.append(element('dt', '', label), values); list.append(entry);
  }
  $('#change-empty').hidden = changes.length > 0;
}

function readWorkingDraft() {
  try {
    const saved = JSON.parse(localStorage.getItem(workingDraftKey) || 'null');
    return saved?.version === 1 && saved.values && saved.editor?.fields && saved.editor?.baseline && Array.isArray(saved.editor.touched) ? saved : null;
  } catch { return null; }
}

function clearWorkingDraft() {
  try { localStorage.removeItem(workingDraftKey); return true; }
  catch { toast('The browser could not remove the saved working draft.', true); return false; }
}

function saveWorkingDraft() {
  if (!app.editor?.baseline) return false;
  captureTemplateFields();
  const form = $('#details-form');
  const values = Object.fromEntries(new FormData(form).entries());
  values.plant = form.elements.namedItem('plant').value;
  values.source_kind = form.elements.namedItem('source_kind').value;
  try {
    localStorage.setItem(workingDraftKey, JSON.stringify({ version: 1, savedAt: new Date().toISOString(), values,
      editor: { ...app.editor, touched: [...app.editor.touched] } }));
    app.editor.workingDraftSaved = true;
    $('#working-draft-status').textContent = 'Working draft saved on this computer. Excel is unchanged.';
    return true;
  } catch {
    $('#working-draft-status').textContent = 'The browser could not save your working draft. Keep this form open to avoid losing edits.';
    return false;
  }
}

function offerDraftRecovery(row = null, copyAsNew = false, resumeEditor = false) {
  const draft = readWorkingDraft();
  if (!draft) return false;
  app.draftRecovery = { draft, row, copyAsNew, resumeEditor };
  const savedDate = new Date(draft.savedAt);
  $('#draft-recovery-description').textContent = `A draft for ${draft.values.supplier || 'a certificate'}${Number.isNaN(savedDate.getTime()) ? '' : ` was saved ${savedDate.toLocaleString()}`}. ${resumeEditor ? 'Restore it or discard it to continue with the certificate you selected.' : 'Restore it to continue editing, or discard it to stay on the dashboard.'}`;
  if (!$('#draft-recovery-modal').open) $('#draft-recovery-modal').showModal();
  return true;
}

async function restoreWorkingDraft() {
  const saved = app.draftRecovery?.draft;
  if (!saved || app.restoringDraft) return;
  app.restoringDraft = true;
  try {
    if (!await loadTemplates()) return;
    const form = $('#details-form'); form.reset();
    app.editor = { ...saved.editor, touched: new Set(saved.editor.touched) };
    app.editor.workingDraftSaved = true;
    for (const [name, value] of Object.entries(saved.values)) {
      const input = form.elements.namedItem(name); if (input) input.value = value;
    }
    for (const disclosure of form.querySelectorAll('details')) disclosure.open = false;
    $('#details-title').textContent = 'Continue your working draft';
    $('#edit-items-hint').textContent = app.editor.itemHint || '';
    updateSourceControls(); renderTemplateEditor(); invalidatePreview(); renderChangeSummary();
    $('#working-draft-status').textContent = 'Working draft restored. Review the changes and preview before creating.';
    $('#details-error').hidden = true;
    $('#draft-recovery-modal').close(); $('#details-modal').showModal();
    app.draftRecovery = null;
  } finally { app.restoringDraft = false; }
}

function requestDetailsClose(afterClose = null) {
  if (app.detailsBusy || app.generationBusy) return false;
  if ($('#details-modal').open && editorDirty()) {
    app.afterDraftClose = afterClose;
    if (!$('#unsaved-modal').open) $('#unsaved-modal').showModal();
    return false;
  }
  $('#details-modal').close();
  if (afterClose) afterClose();
  return true;
}

function finishDraftClose(keep) {
  if (keep ? !saveWorkingDraft() : !clearWorkingDraft()) return;
  const continuation = app.afterDraftClose;
  app.afterDraftClose = null;
  $('#unsaved-modal').close(); $('#details-modal').close();
  if (keep) toast('Working draft saved on this computer. Excel is unchanged.');
  if (continuation) continuation();
}

function certificateUrl(certificate) {
  const filename = String(certificate.filename || '');
  if (!filename || /[\\/]/.test(filename) || !/\.pdf$/i.test(filename) || filename === '..') return null;
  const expectedPath = `/api/certificates/${encodeURIComponent(filename)}`;
  const url = new URL(certificate.pdf_url || expectedPath, window.location.href);
  return url.origin === window.location.origin && url.pathname === expectedPath && !url.search && !url.hash ? url.href : null;
}

function showGenerationSuccess(response) {
  const certificates = Array.isArray(response.certificates) ? response.certificates : [];
  $('#success-title').textContent = certificates.length === 1 ? 'Certificate generated' : 'Certificates generated';
  $('#success-description').textContent = certificates.length
    ? `${certificates.length} PDF${certificates.length === 1 ? '' : 's'} saved. Check the actual files below before sending.`
    : `${response.created} PDF${response.created === 1 ? '' : 's'} generated. Open the folder to check the saved files.`;
  const list = $('#created-certificates'); list.replaceChildren();
  for (const certificate of certificates) {
    const entry = element('li', 'saved-file');
    const copy = element('div', 'saved-file-copy');
    copy.append(element('strong', '', certificate.supplier || 'Certificate'),
      element('span', '', `${labelPlant(certificate.plant)} · Certificate ${certificate.certificate_no || 'number not provided'}`),
      element('span', 'saved-pdf-name', certificate.filename));
    entry.append(copy);
    let url = null;
    try { url = certificateUrl(certificate); } catch { /* Keep file details visible even if a URL is invalid. */ }
    if (url) {
      const link = element('a', 'button button-quiet', 'Open PDF');
      link.href = url; link.target = '_blank'; link.rel = 'noopener';
      link.setAttribute('aria-label', `Open PDF for ${certificate.supplier || certificate.filename}`); entry.append(link);
    }
    list.append(entry);
  }
  $('#success-modal').showModal();
}

function detailsBody(form) {
  const data = new FormData(form);
  const details = {};
  for (const [key, value] of data.entries()) {
    if (key === 'plant' || key === 'source_kind' || key === 'edit_scope') continue;
    details[key] = String(value).trim();
  }
  // Clean up floating-point tails on screen, but preserve the original numeric
  // value unless the user actually edits it.
  for (const [key, value] of Object.entries(app.editor.rawNumeric || {})) {
    if (!app.editor.touched.has(key)) details[key] = value;
  }
  // Date pickers cannot display malformed saved dates. Preserve them until
  // explicitly edited so changing an unrelated field cannot erase a warning.
  for (const [key, value] of Object.entries(app.editor.rawDetails || {})) {
    const input = form.elements.namedItem(key);
    if (input?.type === 'date' && value && !normalizeDate(value) && !app.editor.touched.has(key)) details[key] = value;
  }
  captureTemplateFields();
  details.certificate_fields = { ...app.editor.fields[app.editor.template] };
  const spec = app.templates.find(template => template.id === app.editor.template);
  for (const field of spec?.fields || []) {
    if (field.group === 'details' && commonTemplateKeys.has(field.key)) delete details.certificate_fields[field.key];
  }
  const source = form.elements.namedItem('source_kind').value;
  return { plant: data.get('plant') || form.elements.namedItem('plant').value, details, source_kind: source,
    edit_scope: data.get('edit_scope') || 'update_excel',
    ...(app.editor.workspaceRevision != null ? { expected_revision: app.editor.workspaceRevision } : {}),
    ...(source === 'workbook' && app.editor.sourceSn != null ? { source_sn: app.editor.sourceSn, source_plant: app.editor.sourcePlant } : {}),
    record_id: source === 'manual' ? app.editor.recordId : null };
}

async function saveDetails(event) {
  event.preventDefault();
  if (app.detailsBusy || !app.templatesLoaded) return;
  const payload = detailsBody(event.currentTarget);
  await submitGeneration(payload, false, payload.source_kind === 'manual' ? '/api/manual/generate' : '/api/create', true);
}

async function saveDraft() {
  const form = $('#details-form');
  if (app.detailsBusy || !app.templatesLoaded || form.elements.namedItem('source_kind').value !== 'manual' || !form.reportValidity()) return;
  const payload = detailsBody(form);
  setDetailsBusy(true);
  setBusy(true, 'Saving your draft…', 'Saving the details in your local register.');
  try {
    const response = await api('/api/manual/save', payload);
    app.editor.recordId = response.record_id || app.editor.recordId;
    if (response.state) setState(response.state);
    clearWorkingDraft(); app.editor.baseline = detailsBody(form);
    $('#details-modal').close();
    toast('Draft saved in your local register.');
  } catch (error) { toast(error.message, true); }
  finally { setDetailsBusy(false); setBusy(false); }
}

async function previewDetails() {
  const form = $('#details-form');
  if (app.detailsBusy || !app.templatesLoaded || !form.reportValidity()) return;
  const payload = detailsBody(form);
  invalidatePreview();
  const revision = app.previewRevision;
  $('#preview-details').disabled = true;
  $('#preview-status').textContent = 'Preparing your PDF preview…';
  try {
    const response = await api('/api/preview', payload);
    if (revision !== app.previewRevision || !$('#details-modal').open) return;
    if (!response.preview_url) throw new Error('The PDF preview is not available. Please try again.');
    const url = new URL(response.preview_url, window.location.href);
    if (url.origin !== window.location.origin || !['http:', 'https:'].includes(url.protocol)) throw new Error('The PDF preview address is not available in this workspace.');
    $('#preview-link').href = url.href;
    $('#preview-document').hidden = false;
    let imageUrl = null;
    if (response.preview_image_url) {
      try {
        const candidate = new URL(response.preview_image_url, window.location.href);
        if (candidate.origin === window.location.origin && /^\/previews\/[a-f0-9]{32}\.png$/.test(candidate.pathname) && !candidate.search && !candidate.hash) imageUrl = candidate.href;
      } catch { /* The optional image must not prevent opening the full PDF. */ }
    }
    $('#preview-image').hidden = !imageUrl;
    $('#preview-image-note').hidden = Boolean(imageUrl);
    if (imageUrl) $('#preview-image').src = imageUrl;
    $('#preview-status').textContent = imageUrl
      ? 'First-page preview of your current details. Open the full PDF to check any attached list. Editing a field clears this preview.'
      : 'PDF preview ready. Open the full PDF to check every page; the first-page image is unavailable.';
    $('#preview-number').textContent = `Suggested COD: ${response.certificate_no || 'Not available'}. Previewing does not reserve this number.`;
    $('#preview-number').hidden = false;
    const warnings = Array.isArray(response.warnings) ? response.warnings : [];
    $('#preview-warnings').replaceChildren(...warnings.map(warning => element('li', '', simpleWarning(warning))));
    $('#preview-warnings').hidden = warnings.length === 0;
  } catch (error) {
    if (revision === app.previewRevision && $('#details-modal').open) {
      $('#preview-status').textContent = error.message;
      toast(error.message, true);
    }
  } finally { updateDetailsButtons(); }
}

function toggleFullPdfPreview() {
  const frame = $('#preview-frame');
  if ($('#preview-document').hidden || !$('#preview-link').href) return;
  frame.hidden = !frame.hidden;
  if (!frame.hidden) frame.src = $('#preview-link').href;
  else frame.removeAttribute('src');
  $('#preview-full-pdf').textContent = frame.hidden ? 'View full PDF here' : 'Hide embedded PDF';
  $('#preview-full-pdf').setAttribute('aria-expanded', String(!frame.hidden));
}

async function confirmWarnings() {
  const next = app.pendingConfirmation;
  if (!next) return;
  const reason = $('#override-reason').value.trim();
  if (next.hasWarnings && !reason) {
    $('#warning-reason-error').hidden = false; $('#override-reason').focus(); return;
  }
  if (reason.length > 2000) { $('#warning-reason-error').textContent = 'Keep the reason to 2,000 characters or fewer.'; $('#warning-reason-error').hidden = false; return; }
  $('#warning-modal').close();
  app.pendingConfirmation = null;
  if (!next) return;
  if (next.fromDetails && next.revision !== app.previewRevision) {
    toast('Details have changed. Create the PDF again to review the current details.');
    return;
  }
  await submitGeneration({ ...next.payload, review_token: next.reviewToken, override_reason: reason || 'No warnings were reported; confirmed creation.' }, true, next.endpoint, next.fromDetails);
}

function confirmAction({ title, message, confirmLabel, cancelLabel = 'Cancel' }) {
  if (app.actionConfirmation) return Promise.resolve(false);
  $('#action-confirmation-title').textContent = title;
  $('#action-confirmation-message').textContent = message;
  $('#confirm-action').textContent = confirmLabel;
  $('#cancel-action').textContent = cancelLabel;
  return new Promise(resolve => {
    app.actionConfirmation = resolve;
    $('#action-confirmation-modal').showModal();
    $('#cancel-action').focus();
  });
}

function settleActionConfirmation(approved = false) {
  const resolve = app.actionConfirmation;
  app.actionConfirmation = null;
  $('#action-confirmation-modal').close();
  if (resolve) resolve(approved);
}

async function handleAction(button) {
  const row = app.state.rows[Number(button.dataset.index)];
  if (!row) return;
  const action = button.dataset.action;
  if (action === 'copy') { openDetails(row, true); return; }
  if (action === 'details') { openDetails(row); return; }
  if (action === 'generate') {
    await submitGeneration(isManual(row) ? { plant: row.plant, record_id: row.record_id } : { plant: row.plant, sn: row.sn }, false,
      isManual(row) ? '/api/manual/generate' : '/api/generate');
    return;
  }
  if (action === 'complete') {
    if (app.completionBusy || app.generationBusy || app.backupBusy) return;
    const expectedRevision = app.workspaceRevision ?? app.state.workspace_revision ?? '';
    const approved = await confirmAction({ title: 'Mark certificate Completed?',
      message: `Confirm that you have sent the certificate for ${row.supplier || `Client ${row.sn}`}. This marks it Completed in ${isManual(row) ? 'the local certificate register' : 'Excel'} and stops its reminder.`,
      confirmLabel: 'Mark Completed', cancelLabel: 'Go back' });
    if (!approved || app.completionBusy || app.generationBusy || app.backupBusy) return;
    app.completionBusy = true;
    setBusy(true, isManual(row) ? 'Updating your local register…' : 'Updating the workbook…');
    try { const response = await api(isManual(row) ? '/api/manual/complete' : '/api/complete', { expected_revision: expectedRevision,
      ...(isManual(row) ? { record_id: row.record_id } : { plant: row.plant, sn: row.sn }) }); setState(response.state); toast('Marked Completed.'); }
    catch (error) { toast(error.message, true); }
    finally { app.completionBusy = false; setBusy(false); }
    return;
  }
  if (action === 'open-workbook' && !isManual(row)) {
    try { await api('/api/open-workbook', { plant: row.plant, sn: row.sn }); toast('Opened the plant workbook.'); }
    catch (error) { toast(error.message, true); }
  }
}

function openSettings() {
  updateSettingsDisplay();
  $('#settings-status').textContent = '';
  $('#settings-modal').showModal();
  loadBackupStatus();
}

function displayBackupStatus(response) {
  const destination = response.destination || response.backup_destination || app.settings.backup_destination;
  $('#backup-destination').textContent = destination ? `Backup destination: ${destination}` : 'Choose a backup destination on this computer.';
  if (response.health) $('#backup-health').textContent = typeof response.health === 'string' ? response.health : JSON.stringify(response.health, null, 2);
}

async function loadBackupStatus() {
  $('#backup-status').textContent = '';
  try {
    const response = await api('/api/backup/list', {});
    displayBackupStatus(response);
  } catch (error) { $('#backup-status').textContent = error.message; }
}

async function backupAction(kind) {
  if (app.backupBusy) return;
  app.backupBusy = true;
  for (const id of ['pick-backup-destination', 'create-backup', 'restore-backup']) $(`#${id}`).disabled = true;
  setBusy(true, kind === 'pick-destination' ? 'Choose a backup destination…' : 'Creating a backup…');
  try {
    const response = await api(`/api/backup/${kind}`, {});
    if (response.cancelled) return;
    if (response.settings) app.settings = { ...app.settings, ...response.settings };
    displayBackupStatus(response);
    await loadBackupStatus();
    $('#backup-status').textContent = response.message || (kind === 'create' ? 'Backup created.' : 'Backup destination saved.');
  } catch (error) { $('#backup-status').textContent = error.message; }
  finally {
    app.backupBusy = false; setBusy(false);
    for (const id of ['pick-backup-destination', 'create-backup', 'restore-backup']) $(`#${id}`).disabled = false;
  }
}

function backupOptionLabel(backup) {
  const timestamp = backup.created_at || backup.created;
  const date = timestamp ? new Date(timestamp) : null;
  const parts = [date && !Number.isNaN(date.getTime())
    ? date.toLocaleString(undefined, {day:'2-digit', month:'short', year:'numeric', hour:'numeric', minute:'2-digit'})
    : 'Date not recorded'];
  if (Number.isInteger(backup.file_count) && backup.file_count >= 0) parts.push(`${backup.file_count.toLocaleString()} file${backup.file_count === 1 ? '' : 's'}`);
  if (['pre-restore', 'pre_restore'].includes(backup.purpose)) parts.push('Pre-restore recovery copy');
  else if (backup.purpose === 'manual') parts.push('Manual backup');
  else if (backup.purpose) parts.push(String(backup.purpose).replace(/[_-]/g, ' ').replace(/^\w/, letter => letter.toUpperCase()));
  if (backup.name) parts.push(backup.name);
  return parts.join(' · ');
}

async function openBackupRestore() {
  if (app.backupBusy || $('#details-modal').open) return;
  setBusy(true, 'Finding available backups…');
  app.backupBusy = true;
  try {
    const response = await api('/api/backup/list', {});
    displayBackupStatus(response);
    const backups = response.backups || [];
    const select = $('#backup-selection'); select.replaceChildren();
    for (const backup of backups) {
      const option = element('option', '', backupOptionLabel(backup));
      option.value = backup.backup_id || backup.id;
      option.disabled = backup.valid === false;
      if (option.disabled) option.textContent += ' · Unavailable';
      select.append(option);
    }
    const available = backups.filter(backup => backup.valid !== false);
    if (!available.length) { $('#backup-status').textContent = backups.length ? 'The registered backups are missing or damaged. Choose another destination or create a new backup.' : 'No backups found in the selected destination. Create a backup first.'; return; }
    select.value = available[0].backup_id || available[0].id;
    $('#backup-restore-status').textContent = '';
    $('#confirm-backup-restore').disabled = false;
    $('#backup-restore-modal').showModal();
  } catch (error) { $('#backup-status').textContent = error.message; }
  finally { app.backupBusy = false; setBusy(false); }
}

async function restoreBackup() {
  if (app.backupBusy) return;
  const backupId = $('#backup-selection').value;
  if (!backupId) return;
  app.backupBusy = true; $('#confirm-backup-restore').disabled = true;
  setBusy(true, 'Restoring the selected backup…', 'Saving a recovery snapshot of current data before replacement.');
  try {
    const response = await api('/api/backup/restore', { backup_id: backupId, confirmed: true });
    if (response.settings) app.settings = response.settings;
    if (response.state) setState(response.state);
    else {
      const current = await api('/api/state'); app.settings = current.settings || app.settings; setState(current.state);
    }
    $('#backup-restore-modal').close();
    updateSettingsDisplay(); await loadBackupStatus();
    $('#backup-status').textContent = [response.message || 'Backup restored. A recovery snapshot of the previous data was saved.', ...(response.restored?.warnings || [])].join(' ');
  } catch (error) { $('#backup-restore-status').textContent = error.message; }
  finally { app.backupBusy = false; $('#confirm-backup-restore').disabled = false; setBusy(false); }
}

function historyFields(value, prefix = '', result = {}) {
  if (value && typeof value === 'object') {
    for (const [key, item] of Object.entries(value)) historyFields(item, prefix ? `${prefix}.${key}` : key, result);
  } else if (prefix) result[prefix] = String(value ?? '');
  return result;
}

function historyDifferences(record) {
  const before = historyFields(record.before), after = historyFields(record.after);
  const visibleValue = value => value == null || !String(value).trim() ? '(blank)' : String(value);
  return [...new Set([...Object.keys(before), ...Object.keys(after)])]
    .map(key => ({key, before: visibleValue(before[key]), after: visibleValue(after[key])}))
    .filter(change => change.before !== change.after);
}

function historyFieldLabel(key, record) {
  const printedLabels = {company_display:'Company line',weight_display:'Weight label',_annexure:'Attached PDF list',serial_no:'Serial number'};
  const id = record.after?.certificate_type || record.after?.details?.certificate_type || record.before?.certificate_type || record.before?.details?.certificate_type || 'standard';
  const fields = app.templates.find(template => template.id === id)?.fields || [];
  return key.split('.').filter(part => !['details', 'certificate_fields', 'items'].includes(part))
    .map(part => /^\d+$/.test(part) ? `Item ${Number(part) + 1}`
      : printedLabels[part] || fieldLabels[part] || fields.find(field => field.key === part)?.label
        || part.replace(/_/g, ' ').replace(/^\w/, letter => letter.toUpperCase()))
    .join(' · ');
}

function activityState(state) {
  return ({complete:'Generated',rolled_back:'Recovered after an error',prepared:'Creation started',published:'PDF saved; tracking pending',source_committed:'Tracking saved; completion pending'})[state] || String(state || 'Unknown state').replace(/_/g, ' ');
}

function activityTime(record) {
  const value = record.completed_at || record.started_at;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? String(value || 'Time not recorded') : date.toLocaleString();
}

async function loadHistory() {
  if (app.historyBusy) return;
  app.historyBusy = true;
  $('#history-status').textContent = 'Loading recent activity…';
  try {
    const response = await api('/api/history');
    const records = (response.history || []).slice(-50).reverse();
    const list = $('#history-list'); list.replaceChildren();
    for (const record of records) {
      const item = element('li', 'history-entry'), copy = element('div');
      copy.append(element('strong', '', `Certificate ${record.cod || 'number not recorded'} · ${record.after?.supplier || record.before?.supplier || record.key || 'Company'}`), element('span', '', `${activityState(record.state)} · ${activityTime(record)}`));
      const button = element('button', 'button button-quiet', 'View changes'); button.type = 'button';
      button.addEventListener('click', () => showHistoryDetail(record)); item.append(copy, button); list.append(item);
    }
    $('#history-status').textContent = records.length ? `${records.length} recent record${records.length === 1 ? '' : 's'}.` : 'No certificate activity recorded yet.';
  } catch (error) { $('#history-status').textContent = error.message; }
  finally { app.historyBusy = false; }
}

function showHistoryDetail(record) {
  $('#history-detail-title').textContent = `Certificate ${record.cod || 'activity'}`;
  $('#history-detail-description').textContent = `${activityState(record.state)} · ${activityTime(record)}`;
  const detail = $('#history-detail'); detail.replaceChildren();
  const facts = element('dl', 'history-facts');
  for (const [label, value] of [['Company', record.after?.supplier || record.before?.supplier || 'Not recorded'], ['Plant', labelPlant(record.plant)], ['Source', record.source_kind === 'manual' ? 'Local certificate register' : 'Excel'], ['Operator', record.operator || 'Not recorded'], ['PDF file', String(record.pdf_path || '').split(/[\\/]/).pop() || 'Not recorded'], ['Accepted warning reason', record.override_reason || 'No reason recorded']]) {
    const fact = element('div'); fact.append(element('dt', '', label), element('dd', '', value)); facts.append(fact);
  }
  detail.append(facts, element('h3', '', 'Recorded field changes'));
  const changes = historyDifferences(record), list = element('dl', 'history-changes');
  for (const change of changes) {
    const label = historyFieldLabel(change.key, record);
    const entry = element('div', 'field-change'), values = element('dd');
    values.append(element('span', 'change-before', `Before: ${change.before}`), element('span', 'change-after', `After: ${change.after}`));
    entry.append(element('dt', '', label), values); list.append(entry);
  }
  if (changes.length) detail.append(list); else detail.append(element('p', '', 'No field changes recorded.'));
  if (record.accepted_warnings?.length) {
    detail.append(element('h3', '', 'Accepted warnings'));
    const warnings = element('ul');
    for (const warning of record.accepted_warnings) warnings.append(element('li', '', simpleWarning(warning)));
    detail.append(warnings);
  }
  if (record.previous_pdfs?.length) detail.append(element('p', '', `${record.previous_pdfs.length} previous PDF${record.previous_pdfs.length === 1 ? '' : 's'} retained in the recovery archive.`));
  $('#history-detail-modal').showModal();
}

async function saveSettings(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const payload = {
    workbook_paths: [form.elements.namedItem('plant1_path').value.trim(), form.elements.namedItem('plant2_path').value.trim()],
    reminders_enabled: form.elements.namedItem('reminders_enabled').checked,
    reminder_time: form.elements.namedItem('reminder_time').value,
    update_reminder: event.submitter?.value === 'schedule',
  };
  setBusy(true, 'Saving settings…', 'Checking the workbook paths and updating your Windows reminder.');
  try {
    const response = await api('/api/settings', payload);
    app.settings = response.settings;
    updateSettingsDisplay();
    $('#settings-modal').close();
    toast(response.message || 'Settings saved.');
    await refreshState();
  } catch (error) { $('#settings-status').textContent = error.message; toast(error.message, true); }
  finally { setBusy(false); }
}

async function stopApp() {
  if (app.detailsBusy || app.generationBusy || app.backupBusy || app.completionBusy || app.stopping) return;
  if ($('#details-modal').open && editorDirty()) { requestDetailsClose(stopApp); return; }
  const approved = await confirmAction({ title: 'Close dashboard?',
    message: 'This closes the certificate dashboard on this computer. Your saved files and working drafts are kept. Open Prepare Certificates to use it again.',
    confirmLabel: 'Close dashboard now', cancelLabel: 'Keep open' });
  if (!approved) return;
  if (app.detailsBusy || app.generationBusy || app.backupBusy || app.completionBusy || app.stopping) return;
  app.stopping = true;
  try { await api('/api/stop', {}); $('#offline-screen').hidden = false; }
  catch (error) { toast(error.message, true); }
  finally { app.stopping = false; }
}

function wire() {
  $('#confirm-action').addEventListener('click', () => settleActionConfirmation(true));
  $('#action-confirmation-modal').addEventListener('close', () => {
    if (!$('#action-confirmation-modal').open && app.actionConfirmation) {
      const resolve = app.actionConfirmation; app.actionConfirmation = null; resolve(false);
    }
  });
  $('#keep-working-draft').addEventListener('click', () => { if (saveWorkingDraft()) toast('Working draft saved on this computer. Excel is unchanged.'); });
  $('#save-and-close').addEventListener('click', () => finishDraftClose(true));
  $('#discard-edits').addEventListener('click', () => finishDraftClose(false));
  $('#restore-working-draft').addEventListener('click', restoreWorkingDraft);
  $('#discard-working-draft').addEventListener('click', () => {
    const next = app.draftRecovery;
    if (!clearWorkingDraft()) return;
    app.draftRecovery = null; $('#draft-recovery-modal').close();
    if (next?.resumeEditor) openDetails(next.row || null, next.copyAsNew || false, true);
  });
  $('#create-another').addEventListener('click', () => { $('#success-modal').close(); openDetails(); });
  $('#success-open-folder').addEventListener('click', async () => {
    try { await api('/api/open-folder', {}); } catch (error) { toast(error.message, true); }
  });
  $('#create-backup').addEventListener('click', () => backupAction('create'));
  $('#pick-backup-destination').addEventListener('click', () => backupAction('pick-destination'));
  $('#restore-backup').addEventListener('click', openBackupRestore);
  $('#confirm-backup-restore').addEventListener('click', restoreBackup);
  $('#reload-history').addEventListener('click', loadHistory);
  $('#history-section').addEventListener('toggle', () => { if ($('#history-section').open) loadHistory(); });
  window.addEventListener('beforeunload', event => {
    if (!editorDirty() && !app.generationBusy && !app.backupBusy && !app.completionBusy) return;
    if (editorDirty()) saveWorkingDraft();
    event.preventDefault(); event.returnValue = '';
  });
  for (const [id, change] of [['saved-files-previous', -1], ['saved-files-next', 1]]) {
    $(`#${id}`).addEventListener('click', () => { app.savedFilesPage += change; renderFolderInventory(); });
  }
  $('#show-past-numbers').addEventListener('click', () => {
    $('#past-number-search').value = '';
    renderPastNumbers();
    $('#past-numbers-modal').showModal();
  });
  $('#past-number-search').addEventListener('input', renderPastNumbers);
  $('#refresh-button').addEventListener('click', refreshState);
  $('#new-button').addEventListener('click', () => openDetails());
  $('#generate-all').addEventListener('click', () => submitGeneration(bulkGenerationPayload()));
  const restartList = () => { app.page = 1; app.expanded = null; render(); };
  $('#status-filter').addEventListener('change', restartList);
  $('#plant-filter').addEventListener('change', restartList);
  $('#template-filter').addEventListener('change', restartList);
  $('#search').addEventListener('input', restartList);
  for (const [id, change] of [['previous-page', -1], ['next-page', 1]]) {
    $(`#${id}`).addEventListener('click', () => {
      app.page += change;
      app.expanded = null;
      render();
      $('#queue-title').scrollIntoView({block: 'start', behavior: 'smooth'});
    });
  }
  $('#queue-body').addEventListener('click', event => {
    const button = event.target.closest('button[data-action]');
    if (button) { handleAction(button); return; }
    const expandable = event.target.closest('[data-expand]');
    if (expandable) {
      const index = Number(expandable.dataset.expand);
      app.expanded = app.expanded === index ? null : index;
      render();
    }
  });
  document.querySelectorAll('[data-view]').forEach(button => button.addEventListener('click', () => setView(button.dataset.view)));
  document.querySelectorAll('[data-status]').forEach(button => button.addEventListener('click', () => {
    app.page = 1;
    app.expanded = null;
    $('#status-filter').value = button.dataset.status;
    $('#search').value = '';
    $('#plant-filter').value = 'all';
    $('#template-filter').value = 'all';
    render();
  }));
  $('#details-form').addEventListener('submit', saveDetails);
  $('#save-draft').addEventListener('click', saveDraft);
  $('#preview-details').addEventListener('click', previewDetails);
  $('#preview-full-pdf').addEventListener('click', toggleFullPdfPreview);
  $('#preview-image').addEventListener('error', () => {
    if ($('#preview-image').hidden || !$('#details-modal').open) return;
    $('#preview-image').hidden = true;
    $('#preview-image-note').hidden = false;
    $('#preview-status').textContent = 'The first-page image could not be displayed. Open the full PDF to check the certificate and attached list.';
  });
  $('#pick-annexure').addEventListener('click', pickAnnexure);
  $('#remove-annexure').addEventListener('click', () => {
    captureTemplateFields();
    app.editor.annexureId = null;
    delete app.editor.fields[app.editor.template]._annexure;
    updateAnnexureDisplay();
    invalidatePreview();
  });
  $('#details-form').addEventListener('input', event => {
    if (event.target.name) app.editor?.touched.add(event.target.name);
    invalidatePreview();
  });
  $('#details-form').addEventListener('change', event => {
    const form = event.currentTarget;
    if (event.target.name) app.editor?.touched.add(event.target.name);
    if (event.target.name === 'edit_scope') updateSourceControls();
    if (event.target.name === 'certificate_type') {
      captureTemplateFields();
      if (!app.editor.existing && !app.editor.sourceExplicit) {
        form.elements.namedItem('source_kind').value = event.target.value === 'standard' ? 'workbook' : 'manual';
      }
      renderTemplateEditor();
      updateSourceControls();
    } else if (event.target.name === 'source_kind') {
      app.editor.sourceExplicit = true;
      updateSourceControls();
    }
    invalidatePreview();
  });
  $('#details-modal').addEventListener('close', () => { if (!$('#details-modal').open) invalidatePreview(); });
  $('#warning-modal').addEventListener('close', () => { if (!$('#warning-modal').open) app.pendingConfirmation = null; });
  $('#settings-form').addEventListener('submit', saveSettings);
  $('#confirm-warnings').addEventListener('click', confirmWarnings);
  $('#open-settings').addEventListener('click', openSettings);
  $('#open-folder').addEventListener('click', async () => {
    try { await api('/api/open-folder', {}); }
    catch (error) { toast(error.message, true); }
  });
  $('#stop-app').addEventListener('click', stopApp);
  $('#pick-workbook').addEventListener('click', async () => {
    setBusy(true, 'Choose a workbook…', 'The Windows file picker will open.');
    try {
      const result = await api('/api/pick-workbook', {});
      if (result.cancelled) return;
      $('#details-form').elements.namedItem('plant').value = result.plant;
      invalidatePreview();
      toast(`${labelPlant(result.plant)} workbook selected.`);
    } catch (error) { toast(error.message, true); }
    finally { setBusy(false); }
  });
  document.querySelectorAll('[data-pick-setting]').forEach(button => button.addEventListener('click', async () => {
    setBusy(true, 'Choose an Excel workbook…', 'The Windows file picker will open.');
    try {
      const result = await api('/api/pick-workbook', { for_settings: true });
      if (!result.cancelled) $('#settings-form').elements.namedItem(button.dataset.pickSetting).value = result.path;
    } catch (error) { toast(error.message, true); }
    finally { setBusy(false); }
  }));
  const closeDialog = dialog => {
    if (dialog.id === 'details-modal') requestDetailsClose();
    else if (!app.generationBusy && !app.backupBusy && !app.restoringDraft) dialog.close();
  };
  document.querySelectorAll('[data-close]').forEach(button => button.addEventListener('click', () => closeDialog($(`#${button.dataset.close}`))));
  for (const dialog of document.querySelectorAll('dialog')) {
    dialog.addEventListener('click', event => {
      if (event.target !== dialog || dialog.id === 'success-modal') return;
      const bounds = dialog.getBoundingClientRect();
      if (event.clientX < bounds.left || event.clientX > bounds.right || event.clientY < bounds.top || event.clientY > bounds.bottom) closeDialog(dialog);
    });
    dialog.addEventListener('cancel', event => { event.preventDefault(); closeDialog(dialog); });
  }
  window.addEventListener('keydown', event => {
    if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'k') {
      event.preventDefault(); $('#search').focus();
    }
  });
}

async function init() {
  const now = new Date();
  $('#today-label').textContent = now.toLocaleDateString(undefined, { weekday: 'short', day: 'numeric', month: 'short' });
  wire();
  const templatesReady = loadTemplates();
  try {
    const response = await api('/api/state');
    app.settings = response.settings || {};
    setState(response.state);
    await templatesReady;
    await refreshState();
    offerDraftRecovery();
  } catch (error) {
    toast(error.message, true);
    $('#queue-body').replaceChildren();
    const tr = element('tr'); const td = element('td'); td.colSpan = 5;
    const empty = element('div', 'empty-state');
    empty.append(element('span', 'empty-icon', '!'));
    empty.append(element('strong', '', 'Could not load certificate data'));
    empty.append(element('span', '', error.message));
    td.append(empty); tr.append(td); $('#queue-body').append(tr);
  }
}

init();
