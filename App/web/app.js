const $ = (selector, parent = document) => parent.querySelector(selector);
const app = { state: { rows: [] }, settings: {}, filtered: [], templates: [], templatesLoaded: false,
  editor: null, materialRows: [], page: 1, savedFilesPage: 1, workspaceView: 'records', previewRevision: 0, detailsBusy: false, pendingConfirmation: null, toastTimer: null, savedFileEntry: null };
const templateNames = { standard: 'Recycling', sita: 'International', siemens: 'Medical', destruction: 'Data Destruction' };
const commonTemplateKeys = new Set(['supplier', 'location', 'inward_date', 'certificate_weight', 'form6', 'reference', 'cod_no', 'cod_date']);
const revisionEndpoints = new Set(['/api/create', '/api/generate', '/api/manual/save', '/api/manual/generate', '/api/manual/complete', '/api/complete', '/api/backup/restore', '/api/clients/add', '/api/clients/update', '/api/clients/followup', '/api/clients/remove-preview', '/api/clients/remove', '/api/destruction/request']);
const followupFields = {comments:'Comments', transporter_details:'Transporter / vehicle details', docket_vehicle_no:'Docket no. / vehicle no.', weight_slip_scale:'Weight slip / scale weight'};
revisionEndpoints.add('/api/entries/delete-preview'); revisionEndpoints.add('/api/entries/delete');
revisionEndpoints.add('/api/entries/undo-delete');
const deletionUndoKey = 'ecoreco-last-deletion';
revisionEndpoints.add('/api/clients/status'); revisionEndpoints.add('/api/certificates/link-external');
revisionEndpoints.add('/api/certificates/save-upload');
function hasInward(row) { return row && !isManual(row) && ['plant1','plant2'].includes(row.plant) && String(row.sn || '').trim(); }
function missingFollowup(row) { return hasInward(row) ? Object.keys(followupFields).filter(key => !String(row[key] || '').trim()) : []; }
const workingDraftKey = 'ecoreco.certificate.working-draft.v1';
const showWorkingDraftRecovery = false;
const fieldLabels = { plant: 'Plant', source_kind: 'Use details from', certificate_type: 'Certificate type', case_id: 'SN', supplier: 'Company name', location: 'City', inward_date: 'Received date', certificate_weight: 'Certificate Weight (kg)', item: 'Item description', items: 'Materials and weights', quantity: 'Quantity', per_piece_weight: 'Per piece weight (kg)', weight: 'Inward weight (kg)', form6: 'Form 6 number', reference: 'Reference number', cod_no: 'Certificate number', cod_date: 'Certificate date', starting_cod:'Starting COD number' };
Object.assign(fieldLabels, followupFields);
function isManual(row) { return row.source_kind === 'manual'; }

function excelTarget(entry) {
  if (entry.external || entry.linkable_external) {
    const linked = savedFileRecord(entry);
    if (linked && !isManual(linked)) return {plant:linked.plant, sn:String(linked.sn), sheet:linked.sheet, row:linked.row};
  }
  const source = entry.managed_owner || entry;
  const record = source.source_kind === 'manual'
    ? app.state.rows.find(row => isManual(row) && row.record_id === (source.key || source.record_id)) || entry : entry;
  if (record.detached_inward) return null;
  const link = record.destruction_source || record.raw_details?.destruction_source;
  if (link?.record_id) return null;
  if (link && ['plant1','plant2'].includes(link.plant) && String(link.case_id || '').trim()) {
    const inward = app.state.rows.filter(row => !isManual(row) && row.plant === link.plant
      && String(row.sn ?? row.case_id).trim() === String(link.case_id).trim());
    if (inward.length === 1 && !recordIdentityIssue(inward[0])) return {plant:link.plant,sn:String(link.case_id),sheet:inward[0].sheet,row:inward[0].row};
  }
  if (isManual(record) && templateId(record) === 'destruction' && record.supplier && record.location) {
    const same = app.state.rows.filter(row => !isManual(row) && row.plant === record.plant
      && String(row.supplier || '').trim().toLowerCase() === record.supplier.trim().toLowerCase()
      && String(row.location || '').trim().toLowerCase() === record.location.trim().toLowerCase());
    if (same.length === 1 && !recordIdentityIssue(same[0])) return {plant:same[0].plant,sn:String(same[0].sn),sheet:same[0].sheet,row:same[0].row};
  }
  const sn = String(source.key ?? source.sn ?? '').trim();
  return source.source_kind === 'workbook' && ['plant1','plant2'].includes(source.plant) && sn
    ? {plant:source.plant, sn,
       sheet:source.source_sheet ?? source.sheet,
       row:source.source_row ?? source.row} : null;
}

async function openClientInExcel(plant, sn, sheet, row) {
  if (app.openingExcel) return;
  app.openingExcel = true;
  setBusy(true, `Opening ${labelPlant(plant)} Excel…`, 'Finding and selecting this client’s rows.');
  try {
    const response = await api('/api/open-workbook', {plant, sn, sheet, row});
    if (response.row_selected === false) throw new Error(response.warning || 'Excel opened, but this row was not selected. Refresh and try again.');
    if (!response.selected) throw new Error('Excel could not confirm the selected client. Refresh and try again.');
    if (response.warning) toast('Excel opened. Check the SN, then save and refresh.');
    else toast(`${labelPlant(plant)} Excel opened with this client’s rows selected.`);
  } catch (error) { toast(error.message, true); }
  finally { app.openingExcel = false; setBusy(false); }
}

function excelOpenButton(target, supplier, className = 'button button-quiet') {
  const button = element('button', className, 'Open in Excel'); button.type = 'button';
  button.title = `Open ${labelPlant(target.plant)} Excel and select this client`;
  button.setAttribute('aria-label', `Open ${labelPlant(target.plant)} Excel for ${supplier || 'this client'}`);
  button.addEventListener('click', () => openClientInExcel(target.plant, target.sn, target.sheet, target.row));
  return button;
}
function templateId(row) { return (!recordIdentityIssue(row) && row.last_certificate_details?.certificate_type) || row.certificate_type || 'standard'; }
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
  match = text.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2})$/);
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

const materialFields = [
  ['item', 'Material or asset', 'text', 'e.g. Recycled computer equipment'],
  ['weight', 'Weight (KG)', 'text', 'Enter this material’s weight in kg'],
  ['quantity', 'Quantity', 'text', 'Optional'],
  ['per_piece_weight', 'Per Pcs Weight in (KG)', 'text', 'Optional'],
];
const clientMaterialFields = [
  ['item', 'Material or asset', 'text', 'e.g. Recycled computer equipment'],
  ['quantity', 'Quantity in Pcs', 'text', 'e.g. 2 or 1 lot'],
  ['per_piece_weight', 'Per Pcs Weight in (KG)', 'text', 'Optional'],
  ['weight', 'Weight (KG)', 'text', 'Auto-calculated or enter this material’s weight'],
];

function materialInputName(index, key) { return index === 0 ? key : `material_${index}_${key}`; }

function clientMaterialValues() {
  return (app.clientMaterialRows || []).map(inputs => Object.fromEntries(
    Object.entries(inputs).map(([key, input]) => [key,
      key === 'weight' && input.dataset.autoCalculated === 'true' ? '' : input.value || ''])));
}

function materialWeightEditor(item, index, onChange) {
  const grid = element('div', 'client-material-grid');
  const inputs = {};
  const addField = ([key, label, type, placeholder], className = '') => {
    const field = element('label', `field ${className}`.trim()); field.append(element('span', '', label));
    const input = element(key === 'item' ? 'textarea' : 'input');
    if (key === 'item') input.rows = 1;
    else input.type = type;
    input.value = String(item[key] ?? '');
    input.maxLength = key === 'item' ? 500 : 100;
    input.inputMode = key === 'item' || key === 'quantity' ? 'text' : 'decimal'; input.placeholder = placeholder;
    input.setAttribute('aria-label', `Material ${index + 1}: ${label}`);
    field.append(input); inputs[key] = input; return field;
  };
  grid.append(addField(clientMaterialFields[0], 'client-material-name'));
  const calculation = element('div', 'material-calculation client-material-calculation');
  calculation.append(element('strong', 'client-material-calc-title', 'Calculate weight'));
  const equation = element('div', 'client-material-equation');
  equation.append(addField(clientMaterialFields[2]), element('span', 'client-material-operator', '×'),
    addField(clientMaterialFields[1]), element('span', 'client-material-operator', '='), addField(clientMaterialFields[3]));
  calculation.append(equation);
  grid.append(calculation);
  const hint = element('small', 'client-material-hint');
  hint.hidden = true;
  grid.append(hint); Object.defineProperty(inputs, 'hint', {value:hint});
  const updateCalculatedWeight = (recalculate = true, changedKey = null) => {
    const parse = value => {
      const raw = String(value ?? '').trim();
      if (!raw) return null;
      const grouping = '(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+|[0-9]{1,2}(?:,[0-9]{2})+,[0-9]{3})';
      if (!new RegExp(`^[+]?(?:${grouping}(?:\\.[0-9]*)?|\\.[0-9]+)(?:[eE][+-]?[0-9]+)?$`).test(raw)) return null;
      const number = Number(raw.replaceAll(',', ''));
      return Number.isFinite(number) && number >= 0 ? number : null;
    };
    const quantity = parse(inputs.quantity.value), perPiece = parse(inputs.per_piece_weight.value);
    if (quantity !== null && perPiece !== null) {
      if (recalculate || !String(inputs.weight.value ?? '').trim()) {
        inputs.weight.value = String(Number((quantity * perPiece).toFixed(6)));
        inputs.weight.dataset.autoCalculated = 'true';
      }
      hint.textContent = '';
      hint.hidden = true;
    } else {
      if (inputs.weight.dataset.autoCalculated === 'true') inputs.weight.value = '';
      delete inputs.weight.dataset.autoCalculated;
      hint.textContent = String(inputs.quantity.value).trim() && parse(inputs.quantity.value) === null
        ? 'Enter weight directly.' : '';
      hint.hidden = !hint.textContent;
    }
    onChange(changedKey, inputs.weight.dataset.autoCalculated === 'true');
  };
  for (const key of ['quantity', 'per_piece_weight']) inputs[key].addEventListener('input', () => updateCalculatedWeight(true, key));
  inputs.weight.addEventListener('input', () => {
    delete inputs.weight.dataset.autoCalculated;
    onChange('weight', false);
  });
  inputs.item.addEventListener('input', () => onChange('item', false));
  updateCalculatedWeight(false);
  return {grid, inputs};
}

function renderClientMaterials(items = [{}]) {
  const container = $('#client-material-rows');
  $('#add-client-error').hidden = true;
  const source = items.length ? items : [{}];
  app.clientMaterialRows = [];
  container.replaceChildren();
  source.forEach((item, index) => {
    const row = element('section', 'material-row');
    const heading = element('div', 'material-row-heading');
    heading.append(element('strong', '', `Material ${index + 1}`));
    if (source.length > 1) {
      const remove = element('button', 'text-button', 'Remove'); remove.type = 'button';
      remove.setAttribute('aria-label', `Remove client material ${index + 1}`);
      remove.addEventListener('click', () => {
        if (app.clientBusy) return;
        const next = clientMaterialValues(); next.splice(index, 1); renderClientMaterials(next);
      });
      heading.append(remove);
    }
    row.append(heading);
    const {grid, inputs} = materialWeightEditor(item, index, updateClientMaterialWeightTotal);
    row.append(grid); container.append(row); app.clientMaterialRows.push(inputs);
  });
  $('#add-client-material').disabled = source.length >= 100;
  updateClientMaterialWeightTotal();
  if (app.clientPendingAttempt && !app.clientBusy) validateClientSerialNumber();
}

function updateClientMaterialWeightTotal() {
  const summary = $('#client-material-weight-total');
  if (!summary) return;
  const result = materialWeightTotal(clientMaterialValues());
  summary.hidden = result.state === 'empty';
  summary.textContent = result.state === 'complete' ? `Total weight: ${readableNumber(result.total)} kg`
    : result.state === 'invalid' ? 'Check the weights.'
      : result.state === 'incomplete' ? 'Enter each asset’s weight.' : '';
}

function duplicateSerialNumber(plant, value) {
  const serial = String(value ?? '').trim().toLowerCase();
  if (!serial) return null;
  return app.state.rows.find(row => !isManual(row) && row.plant === plant
    && String(row.sn ?? row.case_id ?? '').trim().toLowerCase() === serial) || null;
}

function clientEntryPayload(form = $('#add-client-form')) {
  const values = Object.fromEntries(new FormData(form).entries());
  const details = Object.fromEntries(['sn','supplier','location','inward_date','form6','reference',
    'comments','transporter_details','docket_vehicle_no','weight_slip_scale']
    .map(key => [key, String(values[key] ?? '').trim()]));
  details.inward_date = normalizeDate(details.inward_date) || details.inward_date;
  details.items = clientMaterialValues().map(item => Object.fromEntries(
    ['item','quantity','per_piece_weight','weight'].map(key => [key, String(item[key] ?? '').trim()])));
  details.sn = ''; // Allocate from the latest workbook under the server lock.
  return {plant:String(values.plant ?? '').trim(), details};
}

function validateClientSerialNumber() {
  const input = $('#add-client-form').elements.namedItem('sn');
  input.value = ''; input.setCustomValidity(''); input.setAttribute('aria-invalid', 'false');
  const hint = $('#client-serial-number-hint');
  hint.textContent = ''; hint.hidden = true; hint.classList.remove('field-error');
  return true;
}

function openAddClient() {
  if (app.loadingWorkspace || app.clientBusy || app.generationBusy || app.backupBusy
      || app.completionBusy || app.settingsBusy || app.stopping) return;
  const form = $('#add-client-form');
  if (!app.clientEntryInitialized) {
    form.reset();
    const selected = $('#plant-filter').value;
    form.elements.namedItem('plant').value = selected === 'plant2' ? 'plant2' : 'plant1';
    renderClientMaterials();
    app.clientEntryInitialized = true;
  }
  // A stable request ID makes a retry safe even if the previous response was lost.
  app.clientRevision = app.workspaceRevision ?? app.state.workspace_revision ?? '';
  app.clientRequestId ||= Array.from({length:32}, () => Math.floor(Math.random() * 16).toString(16)).join('');
  $('#add-client-error').hidden = true;
  validateClientSerialNumber();
  formatFormDates(form);
  $('#add-client-modal').showModal();
  form.elements.namedItem('supplier').focus();
}

async function saveClient() {
  const form = $('#add-client-form');
  if (app.clientBusy) return;
  const serialValid = validateClientSerialNumber();
  if (!serialValid) { form.elements.namedItem('sn').focus(); return; }
  if (!form.reportValidity()) return;
  app.clientRequestId ||= Array.from({length:32}, () => Math.floor(Math.random() * 16).toString(16)).join('');
  const payload = clientEntryPayload(form);
  app.clientBusy = true;
  $('#add-client-error').hidden = true;
  const controls = [...form.querySelectorAll('button, input, select, textarea')];
  controls.forEach(control => { control.disabled = true; });
  $('#save-client').disabled = true; $('#save-client').textContent = 'Adding to Excel…';
  try {
    app.clientPendingAttempt = {requestId:app.clientRequestId, signature:JSON.stringify(payload)};
    const response = await api('/api/clients/add', {...payload, request_id:app.clientRequestId, expected_revision:app.clientRevision});
    if (!response.client_added || !response.client) throw new Error('The client could not be confirmed. Check Excel before adding it again.');
    app.clientEntryInitialized = false; app.clientRevision = null; app.clientRequestId = null; app.clientPendingAttempt = null;
    $('#add-client-modal').close();
    $('#status-filter').value = 'all'; $('#template-filter').value = 'all';
    $('#serial-filter').value = ''; $('#received-date-filter').value = '';
    $('#plant-filter').value = response.client.plant; $('#search').value = String(response.client.sn);
    setWorkspaceView('records'); setState(response.state);
    const row = app.state.rows.find(item => !isManual(item) && item.plant === response.client.plant
      && String(item.sn) === String(response.client.sn));
    const next = response.client.status === 'na' ? 'No certificate needed.'
      : row && readyForBatch(row) ? 'Ready to create a certificate.' : 'Open Review & edit to finish the certificate details.';
    toast(`${response.client.supplier} added to ${labelPlant(response.client.plant)}. ${next}`);
  } catch (error) {
    if (error.clientRequestEnded) {
      app.clientRequestId = null;
      app.clientPendingAttempt = null;
      app.clientRevision = null;
    }
    if (error.workspaceRecovered) {
      app.clientRevision = app.workspaceRevision;
      $('#add-client-error').textContent = 'The Excel list was refreshed. Your entries are kept. Review them, then choose Add to Excel again.';
    } else {
      $('#add-client-error').textContent = noticeText(error.message);
    }
    $('#add-client-error').hidden = false;
    $('#add-client-error').focus();
    $('#add-client-error').scrollIntoView?.({block:'nearest'});
  } finally {
    app.clientBusy = false;
    controls.forEach(control => { control.disabled = false; });
    $('#add-client-material').disabled = (app.clientMaterialRows || []).length >= 100;
    $('#save-client').disabled = false; $('#save-client').textContent = 'Add to Excel';
  }
}

function materialValues() {
  return (app.materialRows || []).map(row => Object.fromEntries(
    Object.entries(row.inputs).map(([key, input]) => [key, input.value || ''])));
}

function materialWeightTotal(items = materialValues()) {
  const active = items.filter(item => Object.values(item || {}).some(value => String(value ?? '').trim()));
  if (!active.length) return { state: 'empty', total: null };
  const numeric = value => {
    const raw = String(value ?? '').trim();
    if (!raw) return null;
    const grouping = '(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+|[0-9]{1,2}(?:,[0-9]{2})+,[0-9]{3})';
    if (!new RegExp(`^[+-]?(?:${grouping}(?:\\.[0-9]*)?|\\.[0-9]+)(?:[eE][+-]?[0-9]+)?$`).test(raw)) return NaN;
    const parsed = Number(raw.replaceAll(',', ''));
    return Number.isFinite(parsed) ? parsed : NaN;
  };
  let total = 0;
  for (const item of active) {
    const weight = numeric(item.weight);
    const quantity = numeric(item.quantity);
    const perPiece = numeric(item.per_piece_weight);
    if (Number.isNaN(weight) || (weight != null && weight < 0)) return { state: 'invalid', total: null };
    if ((quantity != null && !Number.isNaN(quantity) && quantity < 0)
        || (perPiece != null && !Number.isNaN(perPiece) && perPiece < 0)) {
      return { state: 'invalid', total: null };
    }
    // A provided line weight is authoritative. In particular, a text quantity
    // such as "1 Lot" cannot be multiplied and must not invalidate that weight.
    if (weight != null) { total += weight; continue; }
    if (Number.isNaN(perPiece)) {
      return { state: 'invalid', total: null };
    }
    if (Number.isNaN(quantity)) return { state: 'incomplete', total: null };
    const lineWeight = quantity != null && perPiece != null ? quantity * perPiece : null;
    if (lineWeight == null) return { state: 'incomplete', total: null };
    total += lineWeight;
  }
  return Number.isFinite(total) ? { state: 'complete', total } : { state: 'invalid', total: null };
}

function updateMaterialWeightTotal(items = materialValues()) {
  const summary = $('#material-weight-total');
  if (!summary) return;
  const result = materialWeightTotal(items);
  summary.hidden = result.state === 'empty';
  summary.textContent = result.state === 'complete' ? `Total weight: ${readableNumber(result.total)} kg`
    : result.state === 'invalid' ? 'Check the weights.'
      : result.state === 'incomplete' ? 'Enter each asset’s weight.' : '';
  const active = items.filter(item => Object.values(item || {}).some(value => String(value ?? '').trim())).length;
  $('#material-options-title').textContent = `Assets & weights${active ? ` · ${active} item${active === 1 ? '' : 's'}` : ''}`;
}

function materialValuesFromFormValues(values = {}) {
  const rows = [{ item: values.item || '', quantity: values.quantity || '',
    per_piece_weight: values.per_piece_weight || '', weight: values.weight || '' }];
  for (const [name, value] of Object.entries(values)) {
    const match = name.match(/^material_(\d+)_(item|quantity|per_piece_weight|weight)$/);
    if (!match) continue;
    const index = Number(match[1]);
    rows[index] ||= { item: '', quantity: '', per_piece_weight: '', weight: '' };
    rows[index][match[2]] = value;
  }
  return rows;
}

function renderMaterialRows(items = [{}]) {
  const container = $('#material-rows');
  const source = Array.isArray(items) && items.length ? items : [{}];
  app.materialRows = [];
  container.replaceChildren();
  source.forEach((item, index) => {
    const row = element('section', 'material-row');
    row.dataset.materialRow = 'true';
    const heading = element('div', 'material-row-heading');
    heading.append(element('strong', '', `Material ${index + 1}`));
    const remove = element('button', 'button button-quiet material-remove', 'Remove');
    remove.type = 'button';
    remove.hidden = source.length < 2;
    remove.setAttribute('aria-label', `Remove material ${index + 1}`);
    remove.addEventListener('click', () => {
      const next = materialValues();
      next.splice(index, 1);
      if (index === 0 && app.editor) {
        for (const key of ['item', 'quantity', 'per_piece_weight', 'weight']) {
          app.editor.touched.add(key);
          if (app.editor.rawNumeric) delete app.editor.rawNumeric[key];
        }
      }
      renderMaterialRows(next);
      if (app.editor) { app.editor.touched.add('items'); invalidatePreview(); }
      schedulePreview();
    });
    heading.append(remove);
    const {grid, inputs} = materialWeightEditor(item, index, (key, calculated) => {
      if (!key) return;
      if (app.editor) {
        app.editor.touched.add('items');
        app.editor.touched.add(materialInputName(index, key));
        if (calculated || ['quantity', 'per_piece_weight'].includes(key)) app.editor.touched.add(materialInputName(index, 'weight'));
      }
      updateMaterialWeightTotal(); invalidatePreview(); schedulePreview();
    });
    for (const [key, input] of Object.entries(inputs)) input.name = materialInputName(index, key);
    row.append(heading, grid);
    container.append(row);
    app.materialRows.push({ element: row, inputs });
  });
  $('#add-material').disabled = source.length >= 100;
  updateMaterialWeightTotal();
}

function addMaterialRow() {
  if ((app.materialRows || []).length >= 100) return;
  renderMaterialRows([...materialValues(), {}]);
  if (app.editor) { app.editor.touched.add('items'); invalidatePreview(); }
  const input = app.materialRows.at(-1)?.inputs.item;
  input?.focus();
  schedulePreview();
}

function labelPlant(value) { return value === 'plant2' ? 'Plant 2' : value === 'plant1' ? 'Plant 1' : value === null || value === 'local' ? 'App only' : 'Setup'; }
function stateOf(row) {
  const raw = String(row.certificate_status || 'blank').trim().toLowerCase().replace(/[\s-]+/g, '_');
  return ({ created: 'generated', drafted: 'generated', sent: 'completed',
    check_details: 'alert', working_draft: 'alert', generated_externally: 'generated_manual',
    generated_manual: 'generated_manual' })[raw] || raw;
}
function hasWarnings(row) {
  return (Array.isArray(row.warnings) && row.warnings.length > 0)
    && row.warnings.some(warning => !isRoutineNotice(warning))
    || (stateOf(row) === 'ready' && row.outcome === 'needs_review')
    || Boolean(row.system_alert);
}
function needsAttention(row) {
  const state = stateOf(row);
  return Boolean(row.destruction_pending) || ['alert', 'blank', 'unknown'].includes(state)
    || (state === 'ready' && hasWarnings(row))
    || Boolean(row.system_alert) || Boolean(row.reminder_due);
}
function stateLabel(row) {
  const state = stateOf(row);
  if (state === 'blank') return '';
  if (state === 'ready' && hasWarnings(row)) return 'Needs attention';
  return ({ ready: 'Ready', alert: 'Needs attention', draft: 'Draft', generated: 'Generated', generated_manual: 'Generated Externally', completed: 'Completed', na: 'N/A', unknown: 'Needs attention' })[state] || 'Needs attention';
}
function statusClass(row) {
  const state = stateOf(row);
  if (state === 'ready' && hasWarnings(row)) return 'status-needs_attention';
  return `status-${state === 'alert' || state === 'unknown' ? 'needs_attention' : state}`;
}
function dateLabel(value) {
  if (!value) return 'Not entered';
  const normalized = normalizeDate(value);
  if (!normalized) return value;
  const parts = normalized.split('-');
  return `${parts[2]}/${parts[1]}/${parts[0].slice(-2)}`;
}
function isDateInput(input) { return input?.type === 'date' || input?.dataset?.dateInput === 'true'; }
function dateDigits(value) { return String(value ?? '').replace(/\D/g, '').slice(0, 8); }
function formatTypedDate(value) {
  const digits = dateDigits(value);
  if (digits.length <= 2) return digits;
  if (digits.length <= 4) return `${digits.slice(0, 2)}/${digits.slice(2)}`;
  return `${digits.slice(0, 2)}/${digits.slice(2, 4)}/${digits.slice(4)}`;
}
function configureDateInput(input, value = input.value) {
  input.type = 'text'; input.dataset.dateInput = 'true'; input.inputMode = 'numeric';
  input.placeholder = 'DD/MM/YY'; input.maxLength = 10;
  input.value = value ? dateLabel(value) : '';
  if (!input.dataset.dateWired) {
    input.dataset.dateWired = 'true';
    input.addEventListener('input', () => {
      const start = input.selectionStart ?? input.value.length;
      const digitsBefore = dateDigits(input.value.slice(0, start)).length;
      const formatted = formatTypedDate(input.value);
      if (formatted !== input.value) {
        input.value = formatted;
        let caret = digitsBefore;
        if (digitsBefore > 2) caret++;
        if (digitsBefore > 4) caret++;
        input.setSelectionRange(caret, caret);
      }
    });
  }
  const validate = () => input.setCustomValidity(input.value.trim() && !normalizeDate(input.value) ? 'Enter a valid date as DD/MM/YY.' : '');
  if (!input.dataset.dateValidationWired) {
    input.dataset.dateValidationWired = 'true';
    input.addEventListener('input', validate);
    input.addEventListener('blur', () => { validate(); if (normalizeDate(input.value)) input.value = dateLabel(input.value); });
  }
  if (!input.dataset.datePickerWired && input.parentElement) {
    input.dataset.datePickerWired = 'true';
    const label = input.getAttribute('aria-label') || input.closest('label')?.querySelector('span')?.textContent.trim() || 'date';
    const control = element('span', 'date-input-control');
    input.parentElement.insertBefore(control, input);
    control.append(input);
    const button = element('button', 'date-picker-button');
    button.type = 'button'; button.title = 'Choose a date';
    button.setAttribute('aria-label', `Open calendar for ${label}`);
    button.innerHTML = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 10h18"/></svg>';
    const picker = element('input', 'date-picker-native');
    picker.type = 'date'; picker.tabIndex = -1; picker.setAttribute('aria-label', 'Choose date from calendar');
    picker.value = normalizeDate(input.value);
    picker.addEventListener('input', event => event.stopPropagation());
    picker.addEventListener('change', event => {
      event.stopPropagation();
      if (!picker.value) return;
      input.value = dateLabel(picker.value);
      input.dispatchEvent(new Event('input', {bubbles:true}));
      input.dispatchEvent(new Event('change', {bubbles:true}));
    });
    button.addEventListener('click', event => {
      event.preventDefault();
      if (input.readOnly || input.disabled) return;
      picker.value = normalizeDate(input.value);
      if (typeof picker.showPicker === 'function') picker.showPicker();
      else { picker.focus(); picker.click(); }
    });
    control.append(button, picker);
  }
}
function formatFormDates(form) {
  if (!form.elements) return;
  for (const name of ['inward_date', 'cod_date']) {
    const input = form.elements.namedItem(name); if (input) configureDateInput(input);
  }
}
function dateTimeLabel(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value || 'Time not recorded');
  return `${dateLabel(validDate(date.getFullYear(), date.getMonth() + 1, date.getDate()))} ${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
}

function escapeText(value) { return String(value ?? '').trim(); }

function readableNumber(value) {
  if (value === '' || value == null) return '';
  const number = Number(value);
  return Number.isFinite(number) ? String(Number(number.toPrecision(12))) : String(value);
}

function workbookInUse(text) {
  return /Windows error (?:32|33)|(?:workbook|Excel).*\b(?:in use|busy)\b|Windows still reports .* in use/i.test(text);
}

function warningInfo(value) {
  const text = String(value || '').trim();
  const info = (title, message, check = '') => ({title, message, check, original: text});
  const missingLocalPdf = text.match(/^PDF missing\. Ready to recreate with COD (.+)\.$/i);
  if (missingLocalPdf) return {...info('PDF missing', '', `Create it again with COD ${missingLocalPdf[1]}.`), compact:true};
  if (/^PDF missing from Certificates:/i.test(text)) {
    const cod = text.match(/Keep COD (\S+) when recreating it\./i)?.[1];
    const newCod = /A new COD will be assigned when you create it\./i.test(text);
    const remaining = text.replace(/^PDF missing from Certificates:\s*/i, '')
      .replace(/^(?:Excel was changed to Ready\.|Ready to recreate; Excel status update is pending\.|Still Ready to recreate\.)\s*/i, '')
      .replace(/^(?:Keep COD \S+ when recreating it\.|A new COD will be assigned when you create it\.)\s*/i, '')
      .replace(/^Ready to generate COD [^\s.]+\.\s*/i, '');
    const check = cod ? `Create it again with COD ${cod}.`
      : newCod ? 'Create the PDF again to assign a new COD.' : 'Create the PDF again.';
    return {...info('PDF missing', remaining, check), compact:true};
  }
  if (/^Enter the actual date the devices were destroyed\.?$/i.test(text)) return info(
    'Destruction date missing', 'Enter the date the devices were destroyed.');
  let receiptMatch = text.match(/^Device receipt (\d+): enter a valid received date and a positive whole device count\.?$/i);
  if (receiptMatch) return info(`Device receipt ${receiptMatch[1]} needs checking`,
    'Enter when the devices arrived and how many there were. Use a whole number greater than zero.');
  if (/^Add at least one device receipt with a received date and device count\.?$/i.test(text)) return info(
    'Device receipt details missing', 'Add at least one received date and device count.');
  const serial = text.match(/Serial Number\s+["“]?(.+?)["”]?\s+(?:is repeated|appears more than once|is already used)/i);
  if (/Serial Number.*(?:is repeated|appears more than once)/i.test(text)) {
    const rows = text.match(/rows?\s+(\d+(?:(?:\s*,\s*|\s+and\s+)\d+)*)/i)?.[1]?.split(/\s*(?:,|and)\s*/);
    return {...info(`Serial Number${serial ? ` ${serial[1]}` : ''} is repeated.`,
      rows ? `Excel rows ${rows.join(' and ')}.` : '', 'Change the repeated number in Excel, then refresh.'), compact:true};
  }
  if (/Serial Number.*(?:already used|already exists)/i.test(text)) return {...info(
    `Serial Number${serial ? ` ${serial[1]}` : ''} is already used.`, '', 'Use another number or leave it blank.'), compact:true};
  if (/Serial Number.*(?:linked to|different certificate history)/i.test(text)) return {...info(
    'Check this Serial Number.', '', 'Check its Serial Number and COD in Excel, then refresh.'), compact:true};
  if (/files changed since|Excel changed since|reviewed.*changed/i.test(text)) return {...info(
    'Excel has changed.', '', 'Refresh and review the latest details.'), compact:true};
  let sourceMatch = text.match(/^(.+?\.xls[mx]) is unavailable: (.+)$/i);
  if (sourceMatch && (workbookInUse(sourceMatch[2]) || /(?:close.*Excel|already locked)/i.test(sourceMatch[2]))) {
    return {...info(`${sourceMatch[1]} is in use.`, '',
      'Save and close it in Excel. If Excel is already closed, close any file preview, then retry.'), compact:true};
  }
  if (workbookInUse(text)) return {...info(
    'The Excel file is in use.', '', 'Save and close it in Excel. If Excel is already closed, close any file preview, then retry.'), compact:true};
  if (sourceMatch) return info(`${sourceMatch[1]} is unavailable`, sourceMatch[2], 'Check this file in Settings, then click Refresh.');
  sourceMatch = text.match(/^(.+?\.xls[mx]) could not be read: (.+)$/i);
  if (sourceMatch) return info(`${sourceMatch[1]} could not be read`, sourceMatch[2], 'Correct the workbook issue shown here, then click Refresh.');
  if (/^Excel row \d+: total row needs checking\./i.test(text)) return info('Total row needs checking', text,
    'Keep the material rows. Correct or remove the separate total row in Excel, then refresh.');
  if (/^Merged header .+ hides a populated column\./i.test(text)) return info('Excel headings are merged', text,
    'Unmerge the heading cells and give each populated column its own heading.');
  if (/^Column .+ has data but is not recognised\./i.test(text)) return info('An Excel column could not be matched', text,
    'Use one of the heading names listed here, then refresh.');
  if (/^formula in (\S+) on row (\d+);/i.test(text)) {
    const parts = text.match(/^formula in (\S+) on row (\d+);/i);
    return info(`Excel row ${parts[2]} uses a formula`, `The ${fieldLabels[parts[1]] || parts[1]} must be a confirmed value.`,
      'Copy this cell and use Paste Special → Values in Excel, then refresh.');
  }
  if (/^COD Date is in the future;/i.test(text)) return info('COD date is in the future', 'The certificate will use a date later than today.',
    'Correct the date, or confirm this future date deliberately before creating.');
  const fileWarning = text.match(/^(.+\.pdf): (.+)$/i);
  if (fileWarning) {
    const warning = warningInfo(fileWarning[2]);
    if (warning.original) return {...warning, file: certificateFileTitle({path:fileWarning[1]}), original:text};
  }
  let match = text.match(/^COD (\S+) (?:is already recorded in certificate PDF history|belongs to another certificate)/i);
  const replacement = text.match(/(?:new )?automatic COD ([^\s;.,]+)/i);
  if (match && replacement && !/was used after confirmation/i.test(text)) return info(`COD ${match[1]} is already used`,
    `This certificate will use COD ${replacement[1]}. Existing PDFs stay unchanged.`,
    'Check whether this inward already has a certificate before continuing.');
  if (/^(missing company\/supplier|missing supplier name|Company name is missing;)/i.test(text)) return info('Company name missing', 'The company line will be blank on the PDF.', 'Add the company name before creating.');
  if (/^missing inward date/i.test(text)) return info('Received date missing', 'The received date will be blank on the PDF.', 'Check the date in Excel.');
  if (/^invalid inward date/i.test(text)) return info('Received date needs fixing', 'The received date is not valid.', 'Check the date in Excel.');
  if (/^(inward date is in the future|Material receipt date is in the future)/i.test(text)) return info('Received date is in the future', 'The received date is later than today.', 'Check the date before continuing.');
  if (/^(missing or invalid inward date|Material receipt date is missing or invalid;)/i.test(text)) return info('Received date missing or invalid', 'It will be blank on the PDF.', 'Check the received date before creating.');
  if (/^(missing Form 6 number \/ manifest number|Form 6 \/ manifest number is missing;|Form 6 number is required to create the certificate\.)/i.test(text)) return info('Form 6 number missing', 'Enter it when creating the certificate.', 'Open the certificate details and add the Form 6 number.');
  if (/^total weight must be a positive number$/i.test(text)) return info('Weight missing or zero', 'The weight must be greater than zero.', 'Check the weight before sending the PDF.');
  if (/^Weight is missing, invalid or not positive;/i.test(text)) return info('Weight needs checking', 'The source weight is missing, zero or invalid.', 'Check the weight shown in the preview before continuing.');
  if (/^one or more item weights cannot be calculated$/i.test(text)) return info('Material weight missing', 'Some material weights could not be added up.', 'Check the material weights and certificate total.');
  if (/^(Certificate\/net weight is invalid|Certificate Weight .*invalid)/i.test(text)) return info('Certificate weight needs fixing', 'The material total will be used if available.', 'Check the printed weight in the preview.');
  if (/^no item rows$/i.test(text)) return info('Material list missing', 'No materials were entered.', 'Check the material list and certificate weight.');
  match = text.match(/^missing item description on row (\d+)$/i);
  if (match) return info(`Material name missing on Excel row ${match[1]}`, 'This material has no name.', 'Add a material name in Excel.');
  match = text.match(/^invalid or negative (quantity|weight|per_piece_weight) on row (\d+)$/i);
  if (match) return info(`Excel row ${match[2]} needs checking`, `The ${match[1] === 'per_piece_weight' ? 'weight per piece' : match[1]} is invalid or below zero.`, 'Correct this material in Excel and check the total weight.');
  match = text.match(/^invalid weight or unsupported weight formula on row (\d+);/i);
  if (match) return info(`Weight on Excel row ${match[1]} needs fixing`, 'The app could not read or calculate this weight.', 'Enter a confirmed weight, or multiply quantity by weight per piece from the same row.');
  match = text.match(/^invalid or negative certificate\/net weight on row (\d+)$/i);
  if (match) return info('Certificate weight needs fixing', `The certificate weight on Excel row ${match[1]} is invalid. The material total will be used instead.`, 'Check the printed weight in the preview.');
  match = text.match(/^conflicting certificate\/net weight on continuation row (\d+)$/i);
  if (match) return info('Certificate weights do not match', `Excel row ${match[1]} has a different certificate weight for this inward.`, 'Use the same certificate weight throughout this inward.');
  if (/^(?:SN|Serial Number) is repeated on more than one batch in this plant workbook$/i.test(text)) return info('Serial Number is repeated', 'More than one inward has the same Serial Number in this plant’s Excel file.', 'Open Excel and give each inward a unique Serial Number.');
  if (/^Form 6 number is repeated on more than one batch in this plant workbook$/i.test(text)) return info('Form 6 number is repeated', 'More than one inward uses this Form 6 number.', 'Check whether these inwards should share the same Form 6.');
  match = text.match(/^material (\d+) (?:has an? (invalid|negative) (.+?)|is missing its description|weight is missing or cannot be calculated)(?:;.*)?$/i);
  if (match) return info(`Material ${match[1]} needs checking`, match[2] ? `The ${match[3]} is ${match[2]}.`
    : /description/i.test(text) ? 'The material name is missing.' : 'The material weight is missing or cannot be calculated.', 'Check this material and the certificate total.');
  if (/^an item has an invalid or negative weight$/i.test(text)) return info('Material weight needs fixing', 'A material has an invalid or negative weight.', 'Check each material weight.');
  if (/^(Certificate date is invalid;|COD Date is not a valid date;)/i.test(text)) return info('COD date needs fixing', 'The certificate date is not valid.', /today.*will be used/i.test(text) ? 'If you continue, today’s date will be used.' : 'Correct the date in Excel, or clear it to use today.');
  if (/^COD No contains a formula;/i.test(text)) return info('COD number uses an Excel formula', 'The COD number must be a fixed value.', 'Enter a fixed COD number in Excel, or clear it to assign one automatically.');
  match = text.match(/^Plant ([12]) workbook is not connected;/i);
  if (match) return info(`Plant ${match[1]} Excel file is not connected`, 'The app could not check that plant for used COD numbers.', 'Connect the file in Settings, or check the COD yourself before sending.');
  match = text.match(/^Plant ([12]) workbook could not be checked for existing CODs:/i);
  if (match) return info(`Plant ${match[1]} COD numbers could not be checked`, 'The app could not read that plant’s Excel file.', 'Check the file and COD before sending.');
  if (/couldn't find a COD in .*rename the PDF so the COD comes first/i.test(text)) return info('A saved PDF needs its COD in the filename', 'The app could not identify a COD in this PDF, so it cannot safely choose the next number.', 'Open the PDF, read its COD, put that number at the start of the filename, then click Refresh. Move unrelated PDFs outside Certificates.');
  match = text.match(/^(?:Missing (.+?); enter N\/A if it does not apply|(.+?) is missing; enter N\/A if it does not apply, or continue with a blank value\.)$/i);
  if (match) return info(`${match[1] || match[2]} missing`, 'This value will be blank if you continue.', 'Enter the value, or N/A if it does not apply.');
  match = text.match(/^(.+?) is long for its space; the PDF may use small text or clip it\./i);
  if (match) return info(`${match[1]} is too long`, 'The PDF may show small text or cut off part of this value.', 'Shorten it and check the preview.');
  match = text.match(/^(.+?) extends beyond the page; move it inside/i);
  if (match) return info(`${match[1]} is outside the page`, 'Part of this value may not appear on the PDF.', 'Move it inside the page and check the preview.');
  match = text.match(/^(.+?) may overlap (.+?); check the preview\.$/i);
  if (match) return info('Certificate text may overlap', `${match[1]} may cover ${match[2]}.`, 'Check that both can be read in the preview.');
  if (/^Filename COD and printed COD differ;/i.test(text)) return info('PDF has two different COD numbers', 'The filename and certificate show different numbers. Both are treated as used.', 'Open the PDF and check its COD.');
  if (/^More than one printed COD was found;/i.test(text)) return info('PDF shows several COD numbers', 'All of these numbers are treated as used.', 'Open the PDF and check each certificate.');
  if (/^PDF has no readable text/i.test(text)) return info('Scanned PDF needs a check', 'Only the COD in its filename could be checked.', 'Open the PDF and check its COD and reference yourself.');
  if (/^No COD could be identified\./i.test(text)) return info('PDF has no readable COD', 'The app could not find this certificate’s number.', 'Rename the PDF to include its COD before creating another certificate.');
  match = text.match(/^Only the first (\d+) PDF pages were inspected;/i);
  if (match) return info('Some PDF pages need a check', `The app checked only the first ${match[1]} pages.`, 'Check the CODs on any later pages yourself.');
  if (/^Could not read PDF details:/i.test(text)) return info('Saved PDF could not be read', 'The COD in its filename is still treated as used.', 'Open the PDF and check its COD and reference yourself.');
  match = text.match(/^COD (\S+) appears in multiple PDFs \(.+\)\. Verify which certificate is current before sending\.$/i);
  if (match) return info(`COD ${match[1]} appears on several PDFs`, 'Several saved PDFs use this number.', 'Check which certificate is current before sending it.');
  return {title: 'Check this issue', message: text, check: '', original: ''}; // Unfamiliar consequences remain visible in full.
}

function simpleWarning(value) {
  const warning = warningInfo(value);
  return warning.original ? [warning.title, warning.message, warning.check].filter(Boolean).join('. ').replace(/\.{2}/g, '.') : warning.message;
}

function isRoutineNotice(value) {
  const text = String(value || '').trim();
  return /^Excel is open or in use\. Showing its saved contents/i.test(text)
    || /^Reference .+ (?:already appears in|is also used by) /i.test(text);
}

function noticeText(value) {
  const warning = warningInfo(value);
  if (warning.compact) return [warning.title, warning.message, warning.check].filter(Boolean).join(' ');
  return simpleWarning(value);
}

function showDetailsError(message, {previewNotice = false} = {}) {
  const node = $('#details-error');
  node.classList.toggle('preview-notice', previewNotice);
  node.textContent = noticeText(message);
  node.hidden = false;
  const form = $('#details-form');
  const existingWorkbook = form.elements.namedItem('source_kind').value === 'workbook' && app.editor?.sourceSn != null;
  const plant = existingWorkbook ? app.editor.sourcePlant : form.elements.namedItem('plant').value;
  const sn = existingWorkbook ? app.editor.sourceSn : form.elements.namedItem('case_id').value;
  if (['plant1', 'plant2'].includes(plant) && sn && /Serial Number|Excel|workbook/i.test(String(message))
      && !workbookInUse(String(message))) {
    node.append(excelOpenButton({plant, sn:String(sn)}, '', 'button button-quiet'));
  }
}

function warningCard(value) {
  const warning = warningInfo(value);
  const card = element('li', 'issue-card');
  card.append(element('strong', 'issue-title', warning.title));
  if (warning.message) card.append(element('p', 'issue-message', warning.message));
  if (warning.file) card.append(element('p', 'issue-file', `Certificate: ${warning.file}`));
  if (warning.check) {
    const check = element('p', 'issue-check');
    check.append(element('span', '', warning.check)); card.append(check);
  }
  return card;
}

function compactWarningCard(value) {
  const warning = warningInfo(value);
  const card = element('li', 'issue-card');
  card.classList.add('compact-issue');
  card.append(element('strong', 'issue-title', warning.title));
  if (warning.message) card.append(element('p', 'issue-message', warning.message));
  if (warning.file) card.append(element('p', 'issue-file', `Certificate: ${warning.file}`));
  if (warning.check) card.append(element('p', 'issue-check', warning.check));
  return card;
}

function uniqueWarnings(warnings) {
  const unique = new Map();
  for (const value of warnings) {
    if (isRoutineNotice(value)) continue;
    const warning = warningInfo(value);
    const key = JSON.stringify([warning.title, warning.message, warning.check, warning.file]);
    if (!unique.has(key)) unique.set(key, value);
  }
  return [...unique.values()];
}

function renderWarningList(list, warnings, compact = false) {
  list.classList.add('issue-list');
  list.replaceChildren(...uniqueWarnings(warnings).map(compact ? compactWarningCard : warningCard));
}

function nextStep(row) {
  if (recordIdentityIssue(row)) return recordIdentityIssue(row);
  const destructionAction = destructionFollowupAction(row);
  if (destructionAction) return destructionAction;
  switch (stateOf(row)) {
    case 'generated': case 'generated_manual': return 'Your PDF is saved. Send it, then mark it Completed.';
    case 'completed': return 'This certificate is marked Completed.';
    case 'na': return 'N/A — no certificate is needed. This row is skipped.';
    case 'blank': return 'Fill the inward details in Excel. Ready appears automatically when they are complete.';
    case 'draft': return 'Check your saved details, then create the PDF.';
    case 'ready': return hasWarnings(row) ? 'Check the warnings below. You can fix them or confirm and create anyway.' : 'Check the company and weight, then create the PDF.';
    default: return 'Check the Excel status before creating this certificate.';
  }
}

function syncToastLayer() {
  const node = $('#toast');
  const dialogs = (app.toastDialogs || [...document.querySelectorAll('dialog')]).filter(dialog => dialog.open);
  const host = dialogs.at(-1) || document.body;
  // A modal makes elements outside it inert. Keep the dismiss button within
  // the active modal, and promote the notification above its blurred backdrop.
  node.hidePopover?.();
  if (node.parentElement !== host) host.append(node);
  if (node.classList.contains('visible')) node.showPopover?.();
}

function dismissToast() {
  const node = $('#toast');
  node.classList.remove('visible');
  node.hidePopover?.();
  clearTimeout(app.toastTimer);
}

function toast(message, isError = false, canUndoDeletion = false) {
  const node = $('#toast');
  $('#toast-message').textContent = isError ? noticeText(message) : message;
  $('#dismiss-toast').hidden = !isError && !canUndoDeletion;
  $('#undo-deletion-toast').hidden = !canUndoDeletion;
  node.classList.toggle('error', isError);
  node.classList.add('visible');
  syncToastLayer();
  clearTimeout(app.toastTimer);
  if (!isError && !canUndoDeletion) app.toastTimer = setTimeout(dismissToast, 4500);
}

function setBusy(on, title = 'Checking workbooks…', description = 'This usually takes a few seconds.') {
  const overlay = $('#busy-overlay');
  overlay.hidden = !on;
  $('#busy-title').textContent = title;
  $('#busy-description').textContent = description;
  for (const control of document.querySelectorAll('button, input, select, textarea')) {
    if (on) control.setAttribute('aria-busy', 'true');
    else control.removeAttribute('aria-busy');
  }
}

async function api(path, payload, {background = false} = {}) {
  const stateVersion = app.stateVersion || 0;
  if (payload !== undefined && revisionEndpoints.has(path)) payload = { ...payload, expected_revision: payload.expected_revision ?? app.workspaceRevision ?? app.state.workspace_revision ?? '' };
  const options = payload === undefined ? { method: 'GET' } : {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload),
  };
  let response;
  try { response = await fetch(path, options); }
  catch (error) {
    if (background) throw new Error('The local dashboard could not be reached. Click Refresh to check the connection.');
    $('#offline-title').textContent = 'Connection lost';
    $('#offline-message').textContent = 'The app is not responding. Your saved files are safe. Check the connection, or reopen Prepare Certificates. Do not repeat creation until you have checked Saved certificates.';
    $('#reconnect').hidden = false;
    $('#offline-screen').hidden = false;
    if (document.querySelector('dialog[open]') && !$('#connection-modal').open) $('#connection-modal').showModal();
    throw new Error('The local dashboard server is not responding. Reopen Prepare Certificates.bat.');
  }
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error((data.error || `Request failed (${response.status}).`) + (response.status === 409 ? ' Refresh Excel and reopen the details to review the latest data. Your working draft is kept.' : ''));
    error.clientRequestEnded = path === '/api/clients/add' && data.client_request_ended === true;
    error.workspaceChanged = /files changed|Excel has changed|files have changed/i.test(error.message);
    if (error.workspaceChanged && path !== '/api/check') {
      // Reload the queue, never replay the rejected write. An open review
      // pauses normal polling, so otherwise every later action stays stale.
      try {
        const current = await api('/api/check', {}, {background});
        if (current.state && typeof current.workspace_revision === 'string') {
          setState(current.state);
          error.workspaceRecovered = true;
          error.message += ' The company list has now been refreshed. Reopen the entry to review the latest Excel details. Your entered details are kept.';
        }
      } catch { /* Keep the original error if Excel cannot be refreshed yet. */ }
    }
    throw error;
  }
  // A revision belongs to the rows that were actually read. Preview, history,
  // picker and cached startup responses must not bless stale Excel rows.
  const freshRows = data.state && (path === '/api/check' || path === '/api/workbook-changes' || revisionEndpoints.has(path));
  const revision = freshRows ? data.workspace_revision ?? data.state?.workspace_revision
    : data.state?.workspace_revision ?? data.workspace_revision;
  const oldBackgroundRead = path === '/api/workbook-changes'
    && (stateVersion !== (app.stateVersion || 0) || app.completionBusy
      || app.removalBusy || app.entryDeletionBusy || document.querySelector('dialog[open]'));
  if (freshRows && !oldBackgroundRead && typeof revision === 'string') {
    data.state.workspace_revision = revision;
    app.workspaceRevision = revision;
  }
  return data;
}

function updateSettingsDisplay() {
  const paths = app.settings.workbook_paths || [];
  $('#settings-form [name="plant1_path"]').value = paths[0] || '';
  $('#settings-form [name="plant2_path"]').value = paths[1] || '';
  $('#settings-form [name="reminders_enabled"]').checked = Boolean(app.settings.reminders_enabled);
  $('#settings-form [name="reminder_time"]').value = app.settings.reminder_time || '11:00';
  const cod = $('#settings-form').elements.namedItem('next_cod');
  if (cod) {
    cod.value = app.settings.next_cod == null ? '' : String(app.settings.next_cod).padStart(app.state.numbering?.width || 4, '0');
    $('#settings-cod-hint').textContent = app.state.numbering?.next_cod
      ? `Next available COD: ${app.state.numbering.next_cod}. Leave blank for automatic numbering.`
      : 'Leave blank to continue automatic numbering.';
  }
  updateSettingsFiles();
  updateReminderControls();
}

function updateSettingsFiles() {
  const form = $('#settings-form');
  for (const plant of ['plant1', 'plant2']) {
    const path = form.elements.namedItem(`${plant}_path`).value.trim();
    const label = $(`#${plant}-file-name`);
    label.textContent = path.split(/[\\/]/).pop() || 'Not connected';
    label.title = path;
    $(`#${plant}-disconnect`).hidden = !path;
  }
}

function updateReminderControls() {
  const form = $('#settings-form');
  const enabled = Boolean(form.elements.namedItem('reminders_enabled').checked);
  const time = form.elements.namedItem('reminder_time');
  if (!enabled && !/^([01]\d|2[0-3]):[0-5]\d$/.test(time.value)) time.value = app.settings.reminder_time || '11:00';
  form.elements.namedItem('reminder_time').closest('.field').hidden = !enabled;
  // Keep its value in the payload so turning the reminder back on retains the time.
}

function setState(state, {preserveSettings = true} = {}) {
  app.stateVersion = (app.stateVersion || 0) + 1;
  const pending = preserveSettings && $('#settings-modal').open ? settingsValues() : null;
  app.state = state || { rows: [] };
  if (typeof app.state.workspace_revision === 'string') app.workspaceRevision = app.state.workspace_revision;
  app.state.rows ||= [];
  const systemRow = row => row.source_kind === 'system' || (!row.sn && !row.case_id && !row.record_id && !isManual(row));
  app.state.workbook_issues = [...(app.state.workbook_issues || []), ...app.state.rows.filter(systemRow)];
  app.state.rows = app.state.rows.filter(row => !systemRow(row));
  app.page = 1;
  app.expanded = null;
  updateSettingsDisplay();
  if (pending) {
    const form = $('#settings-form');
    form.elements.namedItem('plant1_path').value = pending.workbook_paths[0];
    form.elements.namedItem('plant2_path').value = pending.workbook_paths[1];
    form.elements.namedItem('reminders_enabled').checked = pending.reminders_enabled;
    form.elements.namedItem('reminder_time').value = pending.reminder_time;
    if (form.elements.namedItem('next_cod')) form.elements.namedItem('next_cod').value = pending.next_cod;
    updateSettingsFiles(); updateReminderControls();
  }
  render();
  renderFolderInventory();
  if (app.workspaceView === 'analytics' && !app.loadingWorkspace) window.CertificateAnalytics?.load();
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

function savedFilePlant(entry) {
  const record = savedFileRecord(entry);
  if (record && ['plant1','plant2'].includes(record.plant)) return record.plant;
  if (record && isManual(record) && !record.plant) return 'local';
  if (entry.managed_owner?.source_kind === 'manual' && !entry.managed_owner.plant) return 'local';
  const owner = entry.managed_owner?.plant;
  if (['plant1', 'plant2'].includes(owner)) return owner;
  const match = certificateFileName(entry).match(/_([12])\.pdf$/i);
  return match ? `plant${match[1]}` : 'unknown';
}

function savedFileUrl(entry) {
  const match = String(entry.path || '').replaceAll('\\', '/').match(/^Certificates\/(.+\.pdf)$/i);
  if (!entry.present || !match) return null;
  const relativePath = match[1];
  try { return certificateUrl({filename: relativePath.split('/').at(-1), relative_path: relativePath}); } catch { return null; }
}

function savedFileLabel(entry) {
  const filename = certificateFileName(entry);
  const relativePath = String(entry.path || '').replaceAll('\\', '/').replace(/^Certificates\//i, '');
  const row = savedFileUrl(entry) && app.state.rows.find(row => row.pdf_filename === filename
    && row.pdf_relative_path === relativePath);
  return (row && savedCertificateDetails(row).supplier) || certificateFileTitle(entry);
}

function savedCertificateDetails(row) {
  return row.uploaded_pdf_sha256 && row.last_certificate_details && !recordIdentityIssue(row)
    ? row.last_certificate_details : row;
}

function savedFileRecord(entry) {
  const filename = certificateFileName(entry);
  const relativePath = String(entry.relative_path || entry.path || '').replaceAll('\\', '/').replace(/^Certificates\//i, '');
  const exact = app.state.rows.find(row => row.pdf_filename === filename && row.pdf_relative_path === relativePath);
  if (exact) return exact;
  const owner = entry.managed_owner;
  if (owner?.source_kind === 'manual') {
    return app.state.rows.find(row => isManual(row) && row.record_id === String(owner.key ?? owner.record_id ?? '')) || null;
  }
  if (owner?.source_kind === 'workbook') {
    return app.state.rows.find(row => !isManual(row) && row.plant === owner.plant
      && String(row.sn ?? row.case_id).trim() === String(owner.key ?? '').trim()) || null;
  }
  return null;
}

function savedFileTypeId(entry) {
  const row = savedFileRecord(entry);
  return row ? templateId(row) : entry.certificate_type
    || (/_hdd\.pdf$/i.test(certificateFileName(entry)) ? 'destruction' : '');
}
function savedFileType(entry) {
  const kind = savedFileTypeId(entry);
  return templateNames[kind] || (kind ? app.templates.find(template => template.id === kind)?.name : '')
    || 'Type not identified';
}

function savedFileMatches(entry) {
  const plant = $('#saved-plant-filter').value || 'all';
  if (plant !== 'all' && savedFilePlant(entry) !== plant) return false;
  const type = $('#saved-template-filter').value || 'all';
  if (type !== 'all' && savedFileTypeId(entry) !== type) return false;
  const words = $('#saved-file-search').value.toLowerCase().trim().split(/\s+/).filter(Boolean);
  const record = savedFileRecord(entry);
  const reference = record && savedCertificateDetails(record).reference;
  const haystack = [savedFileLabel(entry), reference, entry.path, ...currentFileNumbers(entry), ...(entry.references || [])]
    .join(' ').replaceAll('_', ' ').toLowerCase();
  return words.every(word => haystack.includes(word));
}

function setWorkspaceView(view) {
  app.workspaceView = ['files', 'analytics'].includes(view) ? view : 'records';
  const files = app.workspaceView === 'files';
  const analytics = app.workspaceView === 'analytics';
  document.body.classList.toggle('analytics-view', analytics);
  $('#analytics-workspace').hidden = !analytics;
  $('#view-analytics').setAttribute('aria-pressed', String(analytics));
  $('#records-workspace').hidden = files || analytics;
  $('#folder-inventory').hidden = !files;
  $('#view-records').setAttribute('aria-pressed', String(!files && !analytics));
  $('#view-files').setAttribute('aria-pressed', String(files));
  if (files) renderFolderInventory();
  $('#page-title').textContent = files ? 'Saved certificates' : 'Your certificates';
  $('#page-description').textContent = files ? 'Find a PDF and open it. Files added manually appear here too.'
    : 'Find a company, check its certificate, and create the PDF.';
  if (analytics) {
    $('#page-title').textContent = 'Dashboard';
    $('#page-description').textContent = 'Your certificates and pending work.';
    if (!app.loadingWorkspace) window.CertificateAnalytics?.load();
  }
}

const recordFilterDefaults = {
  search:'', 'plant-filter':'all', 'status-filter':'all', 'template-filter':'all',
  'serial-filter':'', 'received-date-filter':'', 'received-to-filter':'', 'record-followup-filter':'all',
};
const recordStatusTitles = {
  all:'All companies', ready:'Ready', needs_attention:'Needs attention', details_to_finish:'Details to finish',
  generated:'Generated', generated_manual:'Generated Externally', completed:'Completed', blank:'No status selected',
  na:'N/A', pending_review:'Not ready yet', awaiting_send:'Created, not completed',
  destruction_pending:'Data Destruction follow-up', overdue_unissued:'No certificate after 30 days',
  needs_inward_date:'Inwards that need a received date', pdf_missing:'Missing or unavailable PDF',
  missing_paperwork:'Missing Form 6 or reference',
};
const savedFilterDefaults = {'saved-file-search':'', 'saved-plant-filter':'all', 'saved-template-filter':'all'};

function clearSavedFilter(id) {
  $(`#${id}`).value = savedFilterDefaults[id];
}

function restartSavedList() {
  app.savedFilesPage = 1;
  renderFolderInventory();
}

function updateSavedFilters() {
  const search = $('#saved-file-search').value.trim();
  const plant = $('#saved-plant-filter').value || 'all';
  const type = $('#saved-template-filter').value || 'all';
  const chips = [];
  if (plant !== 'all') chips.push([plant === 'unknown' ? 'Plant not identified' : labelPlant(plant), ['saved-plant-filter']]);
  if (type !== 'all') chips.push([templateNames[type] || type, ['saved-template-filter']]);
  if (search) chips.push([`Search: ${search}`, ['saved-file-search']]);
  $('#clear-file-search').hidden = chips.length === 0;
  $('#saved-filter-count').textContent = String(chips.length);
  renderFilterChips($('#saved-filter-chips'), chips, clearSavedFilter, restartSavedList, $('#saved-file-search'));
}

function clearRecordFilter(id) {
  const control = $(`#${id}`);
  if (id === 'attention-only') control.checked = false;
  else control.value = recordFilterDefaults[id];
  if (id === 'received-date-filter' || id === 'received-to-filter') control.setCustomValidity('');
}

function clearRecordFilterValues() {
  for (const id of [...Object.keys(recordFilterDefaults), 'attention-only']) clearRecordFilter(id);
  app.analyticsAsOf = null;
}

function restartRecordList() {
  app.page = 1; app.expanded = null;
  render();
}

function resetRecordFilters() {
  clearRecordFilterValues();
  restartRecordList();
  $('#search').focus();
}

function readRecordFilters() {
  const plant = $('#plant-filter').value || 'all', status = $('#status-filter').value || 'all';
  const template = $('#template-filter').value || 'all', followup = $('#record-followup-filter').value || 'all';
  const search = $('#search').value.trim(), serial = $('#serial-filter').value.trim();
  const fromText = $('#received-date-filter').value.trim(), toText = $('#received-to-filter').value.trim();
  const from = normalizeDate(fromText), to = normalizeDate(toText);
  const attentionOnly = status === 'needs_attention' && Boolean($('#attention-only').checked);
  return {plant, status, template, followup, search, serial, fromText, toText, from, to, attentionOnly};
}

function updateRecordFilterButton(filters = readRecordFilters()) {
  const {plant, status, template, followup, search, serial, fromText, toText, from, to, attentionOnly} = filters;
  const chips = [];
  if (plant !== 'all') chips.push([labelPlant(plant), ['plant-filter']]);
  if (status !== 'all') chips.push([recordStatusTitles[status] || status, ['status-filter']]);
  if (attentionOnly) chips.push(['Needs attention status only', ['attention-only']]);
  if (template !== 'all') chips.push([templateNames[template] || template, ['template-filter']]);
  if (followup !== 'all') chips.push([$('#record-followup-filter').selectedOptions?.[0]?.textContent || recordStatusTitles[followup], ['record-followup-filter']]);
  if (search) chips.push([`Search: ${search}`, ['search']]);
  if (serial) chips.push([`SN: ${serial}`, ['serial-filter']]);
  if (fromText || toText) {
    const label = fromText && toText ? `Received: ${dateLabel(fromText)}–${dateLabel(toText)}`
      : fromText ? `Received from ${dateLabel(fromText)}` : `Received through ${dateLabel(toText)}`;
    chips.push([label, ['received-date-filter', 'received-to-filter']]);
  }
  const count = chips.length;
  const badge = $('#record-filter-count');
  badge.textContent = String(count);
  badge.hidden = count === 0;
  $('#reset-record-filters').hidden = count === 0;
  $('#clear-filters').disabled = count === 0;
  $('#open-record-filters').classList.toggle('has-filters', count > 0);
  $('#open-record-filters').setAttribute('aria-label', count ? `Filters, ${count} active` : 'Filters');
  const dateError = $('#received-range-error');
  dateError.textContent = (fromText && !from) || (toText && !to) ? 'Enter dates as DD/MM/YY.'
    : from && to && from > to ? 'Received from must be on or before Received to.' : '';
  dateError.hidden = !dateError.textContent;

  renderFilterChips($('#record-filter-chips'), chips, clearRecordFilter, restartRecordList, $('#search'));
}

function renderFilterChips(strip, chips, clearFilter, restartList, searchInput) {
  strip.hidden = chips.length === 0;
  const signature = JSON.stringify(chips);
  if (strip.dataset.filters === signature) return;
  strip.dataset.filters = signature;
  strip.replaceChildren();
  chips.forEach(([label, ids], index) => {
    const button = element('button', 'record-filter-chip');
    button.type = 'button';
    button.setAttribute('aria-label', `Remove filter ${label}`);
    const cross = element('span', '', '×');
    cross.setAttribute('aria-hidden', 'true');
    button.append(element('span', '', label), cross);
    button.addEventListener('click', () => {
      ids.forEach(clearFilter);
      restartList();
      (strip.children[index] || strip.children[index - 1] || searchInput).focus();
    });
    strip.append(button);
  });
}

function positionRecordFilters() {
  const panel = $('#more-filters');
  const rect = $('#open-record-filters').getBoundingClientRect();
  const viewportHeight = window.innerHeight || 800;
  const viewportWidth = window.innerWidth || 1280;
  const width = Math.min(panel.getBoundingClientRect().width || 480, viewportWidth - 32);
  const height = Math.min((panel.scrollHeight || 418) + 2, viewportHeight - 32);
  const top = Math.max(16, Math.min(rect.bottom + 8, viewportHeight - height - 16));
  panel.style.top = `${top}px`;
  panel.style.right = `${Math.max(16, Math.min(viewportWidth - rect.right, viewportWidth - width - 16))}px`;
  panel.style.maxHeight = `${viewportHeight - top - 16}px`;
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

function savedFileSource(entry) {
  const path = String(entry.relative_path || entry.path || '').replaceAll('\\', '/').replace(/^Certificates\//, '');
  const folder = path.split('/')[0];
  const linked = entry.linkable_external ? savedFileRecord(entry) : null;
  if (entry.external || (entry.sha256 && linked?.external_pdf?.sha256 === entry.sha256)
      || ['manually added','externally created','generated externally','external certificates'].includes(folder.toLowerCase())) return 'Manually Added';
  if (entry.app_generated || ['Auto Generated', 'Created in App'].includes(folder)) return 'Auto Generated';
  if (['manual', 'workbook'].includes(entry.managed_owner?.source_kind)) return 'Auto Generated';
  return 'Existing PDF';
}

function inwardPicker(id, onSelection = () => {}) {
  const input = $(`#${id}`);
  const list = $(`#${id}-options`);
  const status = $(`#${id}-results`);
  let rows = [], matches = [], selected = null, active = -1;
  let emptyMessage = 'No available Excel inwards.';
  const key = row => row.record_id || `${row.plant}:${row.sn}`;
  const company = row => savedCertificateDetails(row).supplier || 'Company name missing';
  const identity = row => [labelPlant(row.plant), row.plant && `SN ${row.sn}`, templateNames[templateId(row)],
    row.certificate_no && `COD ${row.certificate_no}`].filter(Boolean).join(' · ');
  const label = row => `${company(row)} · ${identity(row)}`;
  const normalize = value => String(value).toLocaleLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, ' ').trim().replace(/\bplant\s+([12])\b/g, 'plant$1');

  function close() {
    list.hidden = status.hidden = true;
    input.setAttribute('aria-expanded', 'false');
    input.removeAttribute('aria-activedescendant');
    active = -1;
  }

  function render() {
    const query = selected ? '' : normalize(input.value);
    const terms = query ? query.split(' ') : [];
    const found = rows.filter(row => terms.every(term => row.search.includes(term)));
    const serialQuery = query.replace(/^sn /, '');
    if (serialQuery) found.sort((a, b) => Number(b.serial === serialQuery) - Number(a.serial === serialQuery));
    matches = found.slice(0, 12);
    active = -1;
    input.removeAttribute('aria-activedescendant');
    const nodes = matches.map(({row}, index) => {
      const option = element('div', 'inward-option');
      option.id = `${id}-option-${index}`;
      option.setAttribute('role', 'option');
      option.setAttribute('aria-selected', 'false');
      option.append(element('strong', '', company(row)), element('span', '', identity(row)));
      option.addEventListener('mousedown', event => event.preventDefault());
      option.addEventListener('click', () => choose(row));
      return option;
    });
    list.replaceChildren(...nodes);
    list.hidden = !nodes.length;
    input.setAttribute('aria-expanded', String(Boolean(nodes.length)));
    status.hidden = false;
    status.textContent = !rows.length ? emptyMessage
      : !found.length ? 'No matching inwards.'
      : found.length > nodes.length ? `Showing ${nodes.length} of ${found.length} matches. Keep typing to narrow the list.`
      : `${found.length} matching inward${found.length === 1 ? '' : 's'}.`;
  }

  function choose(row) {
    selected = row;
    input.value = label(row);
    input.setCustomValidity('');
    close();
    onSelection();
  }

  input.addEventListener('focus', () => {
    if (selected) input.select?.();
    render();
  });
  input.addEventListener('click', () => { if (list.hidden) render(); });
  input.addEventListener('blur', close);
  input.addEventListener('input', () => {
    selected = null;
    input.setCustomValidity(input.value ? 'Choose an inward from the suggestions.' : '');
    render();
    onSelection();
  });
  input.addEventListener('keydown', event => {
    if (event.isComposing) return;
    if (event.key === 'Escape' && !status.hidden) {
      event.preventDefault();
      close();
    } else if (event.key === 'Enter' && active >= 0) {
      event.preventDefault();
      choose(matches[active].row);
    } else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      event.preventDefault();
      if (list.hidden) render();
      if (!matches.length) return;
      list.children[active]?.setAttribute('aria-selected', 'false');
      active = event.key === 'ArrowDown' ? Math.min(active + 1, matches.length - 1)
        : active < 0 ? matches.length - 1 : Math.max(0, active - 1);
      const option = list.children[active];
      option.setAttribute('aria-selected', 'true');
      input.setAttribute('aria-activedescendant', option.id);
      option.scrollIntoView?.({block:'nearest'});
    }
  });

  return {
    select(row) {
      const candidate = rows.find(item => key(item.row) === key(row));
      if (candidate) choose(candidate.row);
    },
    update(candidates, {reset = false, disabled = false, empty = 'No available Excel inwards.'} = {}) {
      emptyMessage = empty;
      rows = candidates.map(row => ({row, search:normalize(label(row)), serial:normalize(row.sn)}));
      selected = !reset && !disabled && selected ? candidates.find(row => key(row) === key(selected)) || null : null;
      input.disabled = disabled;
      input.value = selected ? label(selected) : '';
      input.setCustomValidity('');
      close();
    },
    row() {
      return selected && !input.disabled ? selected : null;
    },
    selection() {
      const row = this.row();
      if (!row) return null;
      return row.record_id ? {record_id:row.record_id} : {plant:row.plant, sn:String(row.sn)};
    },
  };
}

const uploadInwardPicker = inwardPicker('upload-inward', updateUploadSelection);
const uploadDetailFields = ['supplier', 'location', 'inward_date', 'reference', 'form6', 'certificate_weight'];
let uploadDetailDraft = null;

function uploadDestructionRequest(row) {
  return linkedDestructionRequest(isManual(row) ? {record_id:row.record_id}
    : {plant:row.plant, case_id:row.sn});
}

function updateUploadTarget(reset = false) {
  const mode = $('#external-upload-target').value;
  const local = mode === 'local';
  const replace = mode === 'replace';
  if (reset) uploadDetailDraft = null;
  $('#upload-inward-field').hidden = local;
  $('#upload-inward-label').textContent = replace ? 'Existing certificate' : 'Client';
  $('#upload-inward').placeholder = replace ? 'Search company, COD or plant' : 'Search company, SN or plant';
  $('#upload-replacement-note').hidden = !replace;
  $('#save-uploaded-certificate').textContent = replace ? 'Replace PDF' : 'Save certificate';
  const form = $('#upload-certificate-form');
  const type = form.elements.namedItem('certificate_type');
  if (type.disabled && !replace) {
    type.value = '';
    app.uploadTypeExplicit = false;
    form.elements.namedItem('cod_date').value = '';
  }
  type.disabled = replace;
  const hdd = type.value === 'destruction';
  const eligible = app.state.rows.filter(row => {
    if (recordIdentityIssue(row) || row.workbook_read_only) return false;
    if (replace) {
      if (!row.certificate_no || !row.pdf_relative_path || row.pdf_missing
          || !['generated','generated_manual','completed'].includes(stateOf(row))) return false;
      if (hasInward(row)) return true;
      if (!isManual(row) || !row.record_id) return false;
      const link = row.destruction_source;
      return !link?.plant || app.state.rows.some(inward => hasInward(inward) && inward.plant === link.plant
        && String(inward.sn) === String(link.case_id) && !inward.workbook_read_only && !recordIdentityIssue(inward));
    }
    if (hdd) {
      if (!hasInward(row) && !(isManual(row) && row.record_id && templateId(row) !== 'destruction')) return false;
      const linked = uploadDestructionRequest(row);
      return !row.hdd_cod_no && !(templateId(row) === 'destruction' && row.certificate_no)
        && (!linked || (!linked.certificate_no && stateOf(linked) !== 'na'));
    }
    return hasInward(row) && !row.external_pdf && !row.pdf_relative_path && !row.pdf_missing
      && ['blank','alert','ready','generated_manual'].includes(stateOf(row));
  });
  uploadInwardPicker.update(eligible, {reset, disabled:local,
    empty:replace ? 'No issued certificates with a current PDF.' : hdd
      ? 'No matching clients without a Data Destruction certificate.'
      : 'No available Excel inwards. To add an HDD certificate, choose Data Destruction above.'});
  updateUploadSelection();
}

function updateUploadSelection() {
  const mode = $('#external-upload-target').value;
  const form = $('#upload-certificate-form');
  const type = form.elements.namedItem('certificate_type');
  const row = uploadInwardPicker.row();
  const local = mode === 'local';
  const replace = mode === 'replace';
  const company = $('#upload-company-fields');
  company.disabled = company.hidden = !local && !row;
  $('#upload-detail-scope').hidden = company.hidden;
  const rowKey = JSON.stringify([mode, uploadInwardPicker.selection()]);
  if (replace || (row && rowKey !== uploadDetailDraft?.rowKey && !app.uploadTypeExplicit && type.value !== 'destruction')) {
    type.value = row ? templateId(row) : '';
  }
  const hdd = type.value === 'destruction';
  const pending = row && mode === 'workbook' && hdd && uploadDestructionRequest(row);
  const saved = pending ? { ...pending, ...pending.raw_details }
    : row && (replace || isManual(row)) ? { ...row, ...row.raw_details, ...row.last_certificate_details }
    : row && hdd ? savedCertificateDetails(row) : row;
  const key = JSON.stringify([rowKey, mode === 'workbook' && hdd]);
  const prior = uploadDetailDraft?.key === key ? uploadDetailDraft.defaults : null;
  const defaults = {};
  for (const name of [...uploadDetailFields, 'cod_date']) {
    const field = form.elements.namedItem(name);
    const value = name === 'cod_date'
      ? (hdd && !replace ? pending?.certificate_date : row?.certificate_date)
      : saved?.[name];
    defaults[name] = escapeText(name.endsWith('_date') && value ? dateLabel(value) : value);
    if (!prior || field.value === prior[name]) {
      field.value = defaults[name];
      field.setCustomValidity('');
    }
  }
  uploadDetailDraft = { key, rowKey, defaults };
  if (local) fillUploadReference();
  const items = hdd ? [] : saved?.items?.length ? saved.items : row?.items || [];
  $('#upload-materials').hidden = !items.length;
  if (!prior) $('#upload-materials').open = false;
  $('#upload-materials-summary').textContent = `Materials (${items.length})`;
  $('#upload-materials-list').replaceChildren(...items.map(item => {
    const quantity = item.quantity === null || item.quantity === undefined ? '' : readableNumber(item.quantity);
    const weight = item.weight === null || item.weight === undefined || item.weight === '' ? '' : `${readableNumber(item.weight)} kg`;
    return element('li', '', [item.item || 'Description missing', quantity && `Qty ${quantity}`, weight].filter(Boolean).join(' · '));
  }));
  updateUploadCod();
}

function updateUploadCod() {
  const hint = $('#upload-cod-hint');
  const input = $('#upload-cod');
  const upload = app.uploadedFile;
  const replace = $('#external-upload-target').value === 'replace';
  const row = uploadInwardPicker.row();
  if (replace) {
    const cod = row?.certificate_no || '';
    input.value = cod;
    input.readOnly = true;
    hint.textContent = cod ? 'The replacement PDF must use this COD.' : 'Choose the existing certificate first.';
    return;
  }
  const form = $('#upload-certificate-form');
  const existing = row && form.elements.namedItem('certificate_type').value !== 'destruction' ? row.certificate_no : '';
  if (!upload) {
    input.readOnly = true;
    input.value = existing || '';
    hint.textContent = existing ? 'Using the inward’s existing COD.' : 'Choose a PDF to check its COD.';
    return;
  }
  const editable = !upload.ambiguous_cod && !upload.cod && !existing;
  input.readOnly = !editable;
  input.value = upload.ambiguous_cod ? '' : upload.cod || existing || upload.manualCod || upload.filename_cod || '';
  input.placeholder = editable ? app.state.numbering?.next_cod || 'Assigned automatically' : 'Assigned automatically';
  hint.textContent = upload.ambiguous_cod ? 'The first page has multiple CODs. Choose a PDF with one certificate number.'
    : upload.cod ? 'Read from the PDF.'
    : existing ? 'Using the inward’s existing COD.'
    : 'Enter the COD printed on the file, or leave blank to assign the next available number.';
}

function openUploadCertificate({entry = null, replacement = null} = {}) {
  if (app.loadingWorkspace || app.uploadingCertificate || app.generationBusy) return;
  const form = $('#upload-certificate-form');
  form.reset();
  app.uploadedFile = null;
  app.uploadTypeExplicit = false;
  const fileInput = $('#external-upload-file');
  fileInput.disabled = Boolean(entry);
  $('#upload-file-field').hidden = Boolean(entry);
  const filename = $('#upload-saved-filename');
  filename.hidden = !entry;
  filename.textContent = entry ? certificateFileName(entry) : '';
  $('#save-uploaded-certificate').disabled = Boolean(entry);
  formatFormDates(form);
  $('#upload-certificate-error').hidden = true;
  const kind = entry && savedFileTypeId(entry);
  if (kind) {
    form.elements.namedItem('certificate_type').value = kind;
    app.uploadTypeExplicit = true;
  }
  const hasClient = app.state.rows.some(row => hasInward(row) || (kind === 'destruction' && isManual(row)));
  $('#external-upload-target').value = replacement ? 'replace' : hasClient ? 'workbook' : 'local';
  updateUploadTarget(true);
  if (replacement) uploadInwardPicker.select(replacement);
  if (entry) $('#saved-file-modal').close();
  $('#upload-certificate-modal').showModal();
  if (entry) inspectSavedCertificate(entry);
}

function fillUploadReference() {
  const reference = $('#upload-certificate-form').elements.namedItem('reference');
  if (!reference.value.trim() && app.uploadedFile?.reference) {
    reference.value = app.uploadedFile.reference;
    app.uploadedFile.suggestedReference = reference.value;
  }
}

async function inspectUploadedCertificate() {
  const file = $('#external-upload-file').files[0];
  const previous = app.uploadedFile;
  app.uploadedFile = null;
  const form = $('#upload-certificate-form');
  const reference = form.elements.namedItem('reference');
  if (previous?.suggestedReference && reference.value === previous.suggestedReference) reference.value = '';
  const error = $('#upload-certificate-error');
  const submit = $('#save-uploaded-certificate');
  error.hidden = true;
  submit.disabled = Boolean(file);
  updateUploadCod();
  if (!file) return;
  try {
    const contentType = /\.png$/i.test(file.name) ? 'image/png'
      : /\.jpe?g$/i.test(file.name) ? 'image/jpeg' : 'application/pdf';
    const response = await fetch('/api/certificates/upload-file', {method:'POST',
      headers:{'Content-Type':contentType}, body:file});
    const upload = await response.json();
    if ($('#external-upload-file').files[0] !== file || !$('#upload-certificate-modal').open) return;
    if (!response.ok || !upload.token) throw new Error(upload.error || 'The certificate file could not be uploaded.');
    app.uploadedFile = {file, ...upload};
    if ($('#external-upload-target').value === 'local') fillUploadReference();
    updateUploadCod();
  } catch (exception) {
    if ($('#external-upload-file').files[0] === file) {
      error.textContent = exception.message;
      error.hidden = false;
    }
  } finally {
    if ($('#external-upload-file').files[0] === file) submit.disabled = !app.uploadedFile;
  }
}

async function inspectSavedCertificate(entry) {
  const input = $('#external-upload-file');
  try {
    const result = await api('/api/certificates/upload-saved-file', {
      relative_path:entry.path.replace(/^Certificates\//i, ''), sha256:entry.sha256,
    });
    if (!$('#upload-certificate-modal').open || !input.disabled || app.savedFileEntry !== entry) return;
    app.uploadedFile = result;
    if ($('#external-upload-target').value === 'local') fillUploadReference();
    updateUploadCod();
    $('#save-uploaded-certificate').disabled = false;
  } catch (exception) {
    if (!$('#upload-certificate-modal').open || !input.disabled || app.savedFileEntry !== entry) return;
    const error = $('#upload-certificate-error');
    error.textContent = exception.message;
    error.hidden = false;
  }
}

async function saveUploadedCertificate(event) {
  event.preventDefault();
  if (app.uploadingCertificate) return;
  const form = event.currentTarget;
  const upload = app.uploadedFile;
  const input = $('#external-upload-file');
  if (!upload || (!input.disabled && upload.file !== input.files[0])) return;
  const details = Object.fromEntries([...new FormData(form)].map(([key, value]) => [key, String(value).trim()]));
  const selected = uploadInwardPicker.selection();
  const mode = $('#external-upload-target').value;
  const replace = mode === 'replace';
  const cod = $('#upload-cod');
  if (!replace && !cod.readOnly && cod.value.trim()) details.cod_no = cod.value.trim();
  if (mode !== 'local' && !selected) return;
  const certificateEdits = {};
  if (selected) {
    for (const name of uploadDetailFields) {
      if (name in details) certificateEdits[name] = details[name];
      delete details[name];
    }
  }
  const error = $('#upload-certificate-error');
  error.hidden = true;
  app.uploadingCertificate = true;
  form.inert = true;
  $('#save-uploaded-certificate').textContent = 'Saving…';
  try {
    const saved = await api('/api/certificates/save-upload', {
      token:upload.token, details,
      inward:mode === 'workbook' ? selected : null,
      replacement:replace ? selected : null,
      ...(Object.keys(certificateEdits).length ? {certificate_edits:certificateEdits} : {}),
    });
    setState(saved.state);
    $('#upload-certificate-modal').close();
    setWorkspaceView('files');
    $('#saved-file-search').value = saved.certificate_no;
    $('#saved-template-filter').value = 'all';
    $('#saved-plant-filter').value = 'all';
    app.savedFilesPage = 1;
    renderFolderInventory();
    const action = saved.replaced ? 'PDF replaced' : 'PDF saved';
    toast(saved.state?.setup_warning || `Certificate ${saved.certificate_no}: ${action}.`, Boolean(saved.state?.setup_warning));
  } catch (exception) {
    if (exception.workspaceRecovered) updateUploadTarget();
    if (/upload has expired/i.test(exception.message)) {
      app.uploadedFile = null;
      $('#save-uploaded-certificate').disabled = true;
      if (input.disabled && app.savedFileEntry) inspectSavedCertificate(app.savedFileEntry);
      else input.value = '';
    }
    error.textContent = exception.message;
    error.hidden = false;
  } finally {
    app.uploadingCertificate = false;
    form.inert = false;
    $('#save-uploaded-certificate').textContent = replace ? 'Replace PDF' : 'Save certificate';
  }
}

function showSavedFile(entry) {
  app.savedFileEntry = entry;
  $('#saved-file-name').textContent = certificateFileTitle(entry);
  const facts = $('#saved-file-facts');
  facts.replaceChildren();
  const pathParts = String(entry.path || '').replaceAll('\\', '/').split('/').filter(Boolean);
  const folder = pathParts.length > 2 ? pathParts.slice(1, -1).map(part => part === 'Created in App' ? 'Auto Generated' : part).join(' / ') : '';
  const details = [['Source', savedFileSource(entry)], ['Certificate type', savedFileType(entry)],
    ['Certificate number', currentFileNumbers(entry).join(', ') || 'Could not be read']];
  const record = savedFileRecord(entry);
  const certificate = record && savedCertificateDetails(record);
  if (certificate && record.uploaded_pdf_sha256) {
    details.push(['Company', certificate.supplier || 'Not entered'], ['City / site', certificate.location || 'Not entered'],
      ['Received date', dateLabel(certificate.inward_date) || 'Not entered']);
  }
  if (savedFileTypeId(entry) !== 'destruction') {
    details.push(['Reference number', (certificate && record.uploaded_pdf_sha256 ? certificate.reference : entry.references?.join(', ')) || 'Not entered']);
  }
  if (folder) details.push(['Folder', folder]);
  details.push(['PDF file', certificateFileName(entry)]);
  for (const [label, value] of details) {
    const fact = element('div'); fact.append(element('dt', '', label), element('dd', '', value)); facts.append(fact);
  }
  const typeChoice = $('#saved-type-choice');
  const selectType = $('#saved-type-select');
  const identifiedByApp = Boolean(record);
  const filenameType = /_hdd\.pdf$/i.test(certificateFileName(entry)) ? 'destruction' : '';
  typeChoice.hidden = identifiedByApp;
  selectType.value = entry.certificate_type || filenameType || '';
  $('#saved-type-status').textContent = '';
  $('#use-saved-upload').hidden = !(entry.present && (entry.linkable_external ?? entry.external)
    && entry.sha256 && !identifiedByApp);
  renderWarningList($('#saved-file-warnings'), entry.warnings || []);
  $('#saved-file-warnings').hidden = !entry.warnings?.length;
  const url = savedFileUrl(entry);
  $('#saved-file-open').hidden = !url;
  $('#saved-file-open').dataset.relativePath = entry.relative_path || String(entry.path || '').replace(/^Certificates[\\/]/, '').replaceAll('\\', '/');
  const target = excelTarget(entry);
  $('#saved-file-excel').hidden = !target;
  $('#saved-file-excel').dataset.plant = target?.plant || '';
  $('#saved-file-excel').dataset.sn = target?.sn || '';
  if (!url) facts.append(element('p', '', 'Open the PDF folder to find this file.'));
  if (!$('#saved-file-modal').open) $('#saved-file-modal').showModal();
}

function renderFolderInventory() {
  const inventory = app.state.certificate_inventory;
  const section = $('#folder-inventory');
  section.hidden = app.workspaceView !== 'files';
  const entries = inventory?.entries || [];
  const present = entries.filter(entry => entry.present);
  const warnings = inventory?.warnings || [];
  $('#saved-total').textContent = present.length;
  $('#folder-inventory-summary').textContent = warnings.length
    ? `Saved certificates (${present.length}) · ${warnings.length} file warning${warnings.length === 1 ? '' : 's'}`
    : `Saved certificates (${present.length})`;
  $('#folder-inventory-note').textContent = present.length
    ? 'All certificate PDFs are saved in the Certificates folder.' : 'No saved PDFs yet. Generated certificates will appear here.';
  const notes = $('#folder-inventory-warnings');
  notes.replaceChildren();
  if (warnings.length) {
    const review = element('details', 'file-alerts');
    review.append(element('summary', '', `${warnings.length} file alert${warnings.length === 1 ? '' : 's'} · review before creating`));
    const issues = element('ul');
    renderWarningList(issues, warnings);
    review.append(issues); notes.append(review);
  }
  const list = $('#folder-inventory-files');
  list.replaceChildren();
  const filtered = present.filter(savedFileMatches);
  const sort = $('#saved-file-sort').value || 'newest';
  filtered.sort((a, b) => sort === 'company' ? savedFileLabel(a).localeCompare(savedFileLabel(b), undefined, {numeric:true})
    : sort === 'number' ? currentFileNumbers(a).join(', ').localeCompare(currentFileNumbers(b).join(', '), undefined, {numeric:true})
      : Number(b.mtime_ns || 0) - Number(a.mtime_ns || 0) || savedFileLabel(a).localeCompare(savedFileLabel(b)));
  const page = visiblePage(filtered, app.savedFilesPage);
  app.savedFilesPage = page.page;
  for (const entry of page.rows) {
    const item = element('li', 'saved-file');
    const copy = element('div', 'saved-file-copy');
    const title = element('strong', '', savedFileLabel(entry));
    title.title = certificateFileTitle(entry);
    const plant = savedFilePlant(entry);
    const fileType = savedFileType(entry);
    const numberLabel = fileType !== 'Type not identified' ? fileType : 'COD';
    const meta = element('div', 'saved-file-meta');
    meta.append(element('span', 'saved-file-type', numberLabel),
      element('span', 'saved-file-cod', `${currentFileNumbers(entry).join(', ') || 'not identified'} · ${plant === 'unknown' ? 'Plant not identified' : labelPlant(plant)}`),
      element('span', 'saved-file-source', savedFileSource(entry)));
    copy.append(title, meta);
    if (entry.warnings?.length) copy.append(element('small', 'file-warning-note', `${entry.warnings.length} file warning${entry.warnings.length === 1 ? '' : 's'} · check Details`));
    const actions = element('div', 'saved-file-actions');
    const url = savedFileUrl(entry);
    if (url) {
      const link = element('button', 'button button-primary', 'Open PDF'); link.type = 'button';
      link.addEventListener('click', () => openCertificateLocally(entry.relative_path || entry.path?.replace(/^Certificates[\\/]/, '').replaceAll('\\', '/')));
      link.setAttribute('aria-label', `Open PDF for ${savedFileLabel(entry)}`);
      actions.append(link);
    }
    const excel = excelTarget(entry);
    if (excel) actions.append(excelOpenButton(excel, savedFileLabel(entry)));
    const button = element('button', 'text-button', 'Details');
    button.type = 'button';
    button.setAttribute('aria-label', `File details for ${certificateFileTitle(entry)}`);
    button.addEventListener('click', () => showSavedFile(entry));
    actions.append(button);
    item.append(copy, actions);
    list.append(item);
  }
  $('#saved-file-pages').hidden = page.pages === 1;
  $('#saved-file-page-count').textContent = `Page ${page.page} of ${page.pages}`;
  $('#saved-files-previous').disabled = page.page === 1;
  $('#saved-files-next').disabled = page.page === page.pages;
  $('#saved-visible-count').textContent = filtered.length ? `${page.start + 1}–${page.start + page.rows.length} of ${filtered.length} PDFs` : 'No matching PDFs';
  updateSavedFilters();
  if (!filtered.length) {
    const empty = element('li', 'file-empty');
    empty.append(element('strong', '', present.length ? 'No PDFs match this search' : 'No PDFs saved yet'),
      element('span', '', present.length ? 'Try another name or number, or clear the filters.' : 'Create a certificate or put a correctly named PDF in the certificate folder, then choose Refresh.'));
    list.append(empty);
  }
  const past = pastCertificateNumbers(entries);
  $('#show-past-numbers').hidden = past.length === 0;
  $('#show-past-numbers').textContent = `Past numbers (${past.length})`;
  if ($('#past-numbers-modal').open) renderPastNumbers();
}

function readyForBatch(row) {
  return !isManual(row) && stateOf(row) === 'ready' && row.outcome === 'ready'
    && !row.issuance_blocked && !recordIdentityIssue(row);
}

function summarize(filters = readRecordFilters()) {
  const destructionPending = app.state.rows.filter(row => row.destruction_pending).length;
  $('#destruction-pending').hidden = !destructionPending;
  $('#destruction-pending').textContent = `Data Destruction · ${destructionPending} to finish →`;
  $('#destruction-pending').title = 'Open pending HDD certificates. Payment may depend on these.';
  const rows = app.state.rows.filter(row => matches(row, true, filters));
  const ready = rows.filter(row => stateOf(row) === 'ready' && !hasWarnings(row)).length;
  const selectedPlant = $('#plant-filter').value;
  const workbookReady = app.state.rows.filter(row => readyForBatch(row)
    && (selectedPlant === 'all' || row.plant === selectedPlant)).length;
  const attention = rows.filter(needsAttention).length;
  const generated = rows.filter(row => stateOf(row) === 'generated').length;
  const generatedExternally = rows.filter(row => stateOf(row) === 'generated_manual').length;
  const completed = rows.filter(row => stateOf(row) === 'completed').length;
  const overdue = rows.filter(row => Boolean(row.overdue)).length;
  $('#count-all').textContent = rows.length;
  $('#count-blank').textContent = rows.filter(row => stateOf(row) === 'blank').length;
  $('#count-na').textContent = rows.filter(row => stateOf(row) === 'na').length;
  $('#count-ready').textContent = ready;
  $('#count-attention').textContent = attention;
  $('#count-details-to-finish').textContent = rows.filter(row => missingFollowup(row).length).length;
  $('#count-generated').textContent = generated;
  $('#count-generated-externally').textContent = generatedExternally;
  $('#count-completed').textContent = completed;
  $('#count-overdue').textContent = overdue;
  $('#generate-count').textContent = workbookReady ? `(${workbookReady})` : '';
  $('#generate-all').disabled = workbookReady === 0 || app.state.scan_complete === false;
  $('#open-bulk').disabled = Boolean(app.loadingWorkspace) || workbookReady === 0 || app.state.scan_complete === false;
}

function bulkGenerationPayload() {
  const selectedPlant = $('#plant-filter').value;
  return { plants: selectedPlant === 'all' ? ['plant1', 'plant2'] : [selectedPlant] };
}

function updateBulkCodHint() {
  const input = $('#bulk-start-cod');
  input.setCustomValidity('');
  const value = input.value.trim();
  if (!/^[0-9]{1,12}$/.test(value) || Number(value) < 1) {
    $('#bulk-cod-hint').textContent = 'No previous COD found. Enter the number to start from.';
    return;
  }
  const {prefix = '', width = 4} = app.state.numbering || {};
  const cod = sequence => prefix + String(sequence).padStart(width, '0');
  $('#bulk-cod-hint').textContent = `First COD: ${cod(Number(value))}. Next: ${cod(Number(value) + 1)}.`;
}

function openBulkGeneration() {
  if (app.loadingWorkspace || app.generationBusy || app.detailsBusy || app.backupBusy
      || app.completionBusy || app.clientBusy || app.settingsBusy || app.stopping) return;
  const payload = bulkGenerationPayload();
  const rows = app.state.rows.filter(row => readyForBatch(row)
    && payload.plants.includes(row.plant)).sort(compareRows);
  const count = rows.length;
  if (!count) { toast('No Ready Excel rows in this plant.'); return; }
  app.bulkSelection = { ...payload, expected_revision: app.workspaceRevision ?? app.state.workspace_revision ?? '' };
  app.bulkRows = rows; app.bulkChecked = new Set(); app.bulkCheckboxes = []; app.bulkLabels = [];
  $('#bulk-search').value = '';
  $('#bulk-select-all').checked = false; $('#bulk-select-all').indeterminate = false;
  $('#bulk-choices').replaceChildren(...rows.map((row, index) => {
    const label = element('label', 'bulk-choice'), input = element('input'); input.type = 'checkbox';
    input.setAttribute('aria-label', `${row.supplier}, ${labelPlant(row.plant)}, SN ${row.sn}`);
    input.addEventListener('change', () => { if (input.checked) app.bulkChecked.add(index); else app.bulkChecked.delete(index); updateBulkSelection(); });
    const copy = element('span', 'bulk-client-copy'); copy.append(element('strong', '', row.supplier || 'Company name missing'),
      element('small', '', `${labelPlant(row.plant)} · SN ${row.sn} · ${row.location || 'City not entered'}`));
    const meta = element('span', 'bulk-client-meta'); meta.append(element('span', '', dateLabel(row.inward_date)));
    if (hasWarnings(row)) meta.append(element('small', 'bulk-check-note', 'Needs a check'));
    app.bulkCheckboxes.push(input); app.bulkLabels.push(label); label.append(input, copy, meta); return label;
  }));
  const numbering = app.state.numbering;
  const knownCod = app.state.rows.some(row => String(row.certificate_no || '').trim())
    || (app.state.certificate_inventory?.entries || []).some(entry => entry.cods?.length);
  const needsStart = numbering ? numbering.needs_start : !knownCod;
  app.bulkSelectionNeedsStart = needsStart;
  const startInput = $('#bulk-start-cod');
  startInput.value = '';
  startInput.placeholder = numbering?.next_cod?.replace(numbering?.prefix || '', '') || 'e.g. 0302';
  startInput.required = needsStart;
  startInput.setCustomValidity('');
  $('#bulk-start-field').hidden = !needsStart;
  $('#bulk-cod-hint').textContent = 'No previous COD found. Enter the number to start from.';
  $('#bulk-next-cod').hidden = needsStart;
  $('#bulk-next-cod').textContent = numbering?.next_cod
    ? `Next COD: ${numbering.next_cod}` : 'COD numbers are assigned automatically';
  const plants = payload.plants.map(labelPlant).join(' and ');
  $('#bulk-description').textContent = `${count} ${count === 1 ? 'client' : 'clients'} ready · ${plants}. Choose who needs a PDF.`;
  filterBulkClients();
  $('#bulk-generation-modal').showModal();
  if (needsStart) startInput.focus();
  else $('#bulk-search').focus();
}

function filterBulkClients() {
  const search = $('#bulk-search').value.trim().toLowerCase();
  const words = search.split(/\s+/).filter(Boolean);
  app.bulkVisible = [];
  (app.bulkRows || []).forEach((row, index) => {
    const text = [row.supplier, row.sn, row.reference, row.form6, row.location, labelPlant(row.plant), row.inward_date, dateLabel(row.inward_date)].join(' ').toLowerCase();
    const visible = words.every(word => text.includes(word));
    if (app.bulkLabels?.[index]) app.bulkLabels[index].hidden = !visible;
    if (visible) app.bulkVisible.push(index);
  });
  $('#bulk-empty').hidden = app.bulkVisible.length > 0;
  $('#bulk-choices').hidden = app.bulkVisible.length === 0;
  $('#bulk-clear-search').hidden = !search;
  $('#bulk-select-label').textContent = search ? 'Select shown' : 'Select all';
  $('#bulk-visible-count').textContent = search ? `${app.bulkVisible.length} of ${app.bulkRows?.length || 0} shown` : `${app.bulkRows?.length || 0} clients`;
  updateBulkSelection();
}

function updateBulkSelection() {
  const count = app.bulkChecked?.size || 0;
  const visible = app.bulkVisible || (app.bulkRows || []).map((row,index) => index);
  const checkedVisible = visible.filter(index => app.bulkChecked?.has(index)).length;
  const hiddenSelected = count - checkedVisible;
  $('#bulk-selected-count').textContent = `${count} selected`;
  $('#bulk-selection-note').textContent = count > 500 ? 'Choose up to 500 clients at a time.' : hiddenSelected
    ? `${hiddenSelected} selected ${hiddenSelected === 1 ? 'client is' : 'clients are'} hidden by this search.` : 'Only checked clients will get a PDF.';
  $('#generate-count').textContent = `(${count})`;
  $('#generate-all').disabled = count === 0 || count > 500;
  $('#bulk-select-all').disabled = visible.length === 0;
  $('#bulk-select-all').checked = visible.length > 0 && checkedVisible === visible.length;
  $('#bulk-select-all').indeterminate = checkedVisible > 0 && checkedVisible < visible.length;
  (app.bulkLabels || []).forEach((label,index) => label.classList.toggle('bulk-choice-selected', Boolean(app.bulkChecked?.has(index))));
}

async function startBulkGeneration() {
  const selection = app.bulkSelection;
  if (!selection || app.generationBusy || !app.bulkChecked?.size || app.bulkChecked.size > 500) return;
  const payload = {expected_revision:selection.expected_revision,
    selections:app.bulkRows.filter((row,index) => app.bulkChecked.has(index)).map(row => ({plant:row.plant,sn:String(row.sn)}))};
  if (app.bulkSelectionNeedsStart) {
    const input = $('#bulk-start-cod');
    const value = input.value.trim();
    if (!/^[0-9]{1,12}$/.test(value) || Number(value) < 1) {
      input.setCustomValidity('Enter the starting COD number (1 or higher).');
      input.reportValidity();
      input.focus();
      return;
    }
    input.setCustomValidity('');
    payload.starting_cod = value;
  }
  $('#bulk-generation-modal').close();
  app.bulkSelection = null;
  await submitGeneration(payload);
}

function compareRows(a, b) {
  // ISO dates sort chronologically; unknown dates follow dated inwards.
  const received = row => normalizeDate(savedCertificateDetails(row).inward_date);
  const company = row => String(savedCertificateDetails(row).supplier || '');
  const order = $('#company-sort')?.value;
  if (order === 'company') return company(a).localeCompare(company(b), undefined, {numeric:true}) || String(a.plant).localeCompare(String(b.plant));
  if (order === 'serial') {
    const first=companySerial(a),second=companySerial(b);
    return (first === '—') - (second === '—') || first.localeCompare(second, undefined, {numeric:true}) || String(a.plant).localeCompare(String(b.plant));
  }
  if (order === 'received-newest') return received(b).localeCompare(received(a)) || companySerial(a).localeCompare(companySerial(b), undefined, {numeric:true});
  return (received(a) || '9999-99-99').localeCompare(received(b) || '9999-99-99')
    || String(a.plant || '').localeCompare(String(b.plant || ''))
    || String(a.sn ?? a.case_id ?? '').localeCompare(String(b.sn ?? b.case_id ?? ''), undefined, {numeric:true})
    || company(a).localeCompare(company(b));
}

function companySerial(row) { return String(isManual(row) ? row.destruction_source?.case_id || row.case_id || '—' : row.sn ?? row.case_id ?? '—'); }

const recordFollowupModes = new Set(['needs_inward_date', 'pending_review', 'awaiting_send', 'destruction_pending', 'overdue_unissued', 'pdf_missing', 'missing_paperwork']);

function unavailableCertificatePdf(row) {
  return Boolean(row.pdf_missing) || row.issue_code === 'pdf_unavailable'
    || (row.system_alert && /^(?:The external PDF could not be matched|The externally created PDF linked to this completed inward is missing|The PDF is missing, but Excel could not be changed to Ready)/.test(row.detail || ''));
}

function matchesRecordFollowup(row, mode, state) {
  switch (mode) {
    case 'needs_inward_date':
      return hasInward(row) && !row.destruction_source && !normalizeDate(row.inward_date);
    case 'awaiting_send':
      return ['generated', 'generated_manual'].includes(state) && !unavailableCertificatePdf(row);
    case 'destruction_pending':
      return row.destruction_requested !== false && templateId(row) === 'destruction' && !['completed', 'na'].includes(state);
    case 'pdf_missing':
      return unavailableCertificatePdf(row);
    case 'missing_paperwork': {
      const type = templateId(row);
      return type !== 'destruction' && !['generated', 'generated_manual', 'completed', 'na'].includes(state)
        && (!String(row.reference || '').trim() || (type === 'standard' && !String(row.form6 || '').trim()));
    }
    case 'pending_review':
      return !['generated', 'generated_manual', 'completed', 'na'].includes(state)
        && (state !== 'ready' || hasWarnings(row) || row.issuance_blocked || Boolean(recordIdentityIssue(row)));
    case 'overdue_unissued': {
      if (['generated', 'generated_manual', 'completed', 'na'].includes(state) || row.destruction_source) return false;
      const day = normalizeDate(row.inward_date);
      const today = app.analyticsAsOf || new Date().toLocaleDateString('en-CA');
      return Boolean(day) && (Date.parse(today) - Date.parse(day)) / 86400000 > 30;
    }
    default:
      return true;
  }
}

function matches(row, ignoreStatus = false, filters = readRecordFilters()) {
  const {plant, template, followup, search, serial, from, to, attentionOnly} = filters;
  const status = ignoreStatus ? 'all' : filters.status;
  if (plant !== 'all' && (plant === 'local' ? !isManual(row) || Boolean(row.plant) : row.plant !== plant)) return false;
  if (template !== 'all' && templateId(row) !== template) return false;
  if (serial && !companySerial(row).toLowerCase().includes(serial.toLowerCase())) return false;
  if (from || to) {
    const day = normalizeDate(savedCertificateDetails(row).inward_date);
    if (!day || (from && day < from) || (to && day > to)) return false;
  }
  const state = stateOf(row);
  if (!matchesRecordFollowup(row, followup, state)) return false;
  if (status === 'ready' && (state !== 'ready' || hasWarnings(row))) return false;
  if (status === 'needs_attention' && (attentionOnly ? stateLabel(row) !== 'Needs attention' : !needsAttention(row))) return false;
  if (status === 'details_to_finish' && !missingFollowup(row).length) return false;
  if (recordFollowupModes.has(status)) {
    if (!matchesRecordFollowup(row, status, state)) return false;
  } else if (!['all', 'ready', 'needs_attention', 'details_to_finish'].includes(status) && state !== status) {
    return false;
  }
  if (search) {
    const details = savedCertificateDetails(row);
    const searchable = [details.supplier, details.form6, details.reference, details.location, row.supplier, row.sn, row.case_id, row.form6, row.reference, row.certificate_no, row.plant, row.location, templateLabel(row), isManual(row) ? 'local register' : 'plant workbook']
      .map(value => String(value || '').toLowerCase()).join(' ');
    if (!searchable.includes(search.toLowerCase())) return false;
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
  const certificate = savedCertificateDetails(row);
  const td = element('td');
  const wrap = element('div', 'client-cell');
  const copy = element('span', 'client-copy');
  copy.append(element('strong', '', certificate.supplier || 'Company name missing'));
  const meta = element('span', 'saved-file-meta');
  meta.append(element('span', 'saved-file-type', templateNames[templateId(row)] || templateLabel(row)));
  const details = [row.certificate_no || '', certificate.location || ''].filter(Boolean);
  if (details.length) meta.append(element('small', 'saved-file-cod', details.join(' · ')));
  copy.append(meta);
  const issues = rowWarningSummary(row);
  if (issues.length && $('#status-filter').value !== 'needs_attention' && !['generated', 'generated_manual', 'completed'].includes(stateOf(row))) {
    const warning = warningInfo(issues[0]);
    const message = warning.title;
    const note = element('small', 'row-alert-note', message);
    note.title = issues.map(simpleWarning).join(' · '); copy.append(note);
  }
  if ($('#status-filter').value === 'details_to_finish') {
    const shortNames = {comments:'Comments', transporter_details:'Transport details', docket_vehicle_no:'Docket / vehicle no.', weight_slip_scale:'Weight slip'};
    const missing = missingFollowup(row);
    copy.append(element('small', 'followup-missing', `Missing: ${missing.map(key => shortNames[key]).join(' · ')}`));
  }
  wrap.append(copy);
  td.append(wrap);
  return td;
}

function attentionWarning(value) {
  const text = String(value || '').trim();
  if (/^(missing company\/supplier|missing supplier name|Company name is missing;)/i.test(text)) return 'Company name missing. Add it in Review & edit.';
  if (/^(missing Form 6|Form 6 \/ manifest number is missing|Form 6 number is required)/i.test(text)) return 'Form 6 missing. Add it in Review & edit.';
  if (/^(Enter the reference number|Reference (?:number )?is (?:missing|required))/i.test(text)) return 'Reference number missing. Add it in Review & edit.';
  if (/^(missing inward date|invalid inward date|missing or invalid inward date|Material receipt date is missing or invalid)/i.test(text)) return 'Received date missing or invalid. Fix it in Review & edit.';
  if (/^(inward date is in the future|Material receipt date is in the future)/i.test(text)) return 'Received date is in the future. Check the date.';
  if (/^(Weight is missing, invalid or not positive|total weight must be a positive number)/i.test(text)) return 'Weight missing or invalid. Check it in Review & edit.';
  if (/^Enter the actual date the devices were destroyed/i.test(text)) return 'Destruction date missing. Add it in Review & edit.';
  if (/^Add at least one device receipt/i.test(text)) return 'Device details missing. Add the received date and device count.';
  const issue = warningInfo(text);
  if (!issue.original) return issue.message;
  if (/^PDF missing from Certificates:/i.test(text) && issue.message) return attentionWarning(issue.message);
  return [issue.title, issue.check || issue.message].filter(Boolean)
    .map(part => part.replace(/[.\s]+$/, '')).join('. ') + '.';
}

function missingPdfReason(row) {
  return row.certificate_no
    ? `PDF missing. Create it again with COD ${row.certificate_no}.`
    : 'PDF missing. Create it again in Review & edit.';
}

function destructionFollowup(row) {
  if (!row.destruction_pending) return null;
  const state = stateOf(row);
  if (state === 'completed' || state === 'na') return null;
  if (unavailableCertificatePdf(row)) return 'missing';
  if (state === 'generated' || state === 'generated_manual') return row.pdf_filename ? 'send' : 'missing';
  return 'create';
}

function destructionFollowupAction(row) {
  return ({
    missing:'Data Destruction PDF missing. Create it again in Review & edit.',
    send:'Send the Data Destruction PDF, then click Mark Completed.',
    create:'Add the device details in Review & edit, then create the Data Destruction PDF.',
  })[destructionFollowup(row)];
}

function attentionReason(row) {
  const state = stateOf(row);
  if (recordIdentityIssue(row)) return attentionWarning(recordIdentityIssue(row));
  const destructionAction = destructionFollowupAction(row);
  if (destructionAction) return destructionAction;
  if (Number.isFinite(row.age_days) && row.age_days < 0) return 'Received date is in the future. Check the date.';
  const issues = rowWarningSummary(row);
  if (issues.length && (row.issuance_blocked || !['generated', 'generated_manual', 'completed'].includes(state))) return attentionWarning(issues[0]);
  if (row.issuance_blocked) return 'Open Review & edit to fix the missing details.';
  if (row.pdf_missing) return missingPdfReason(row);
  if (['generated', 'generated_manual'].includes(state)) return row.pdf_filename
    ? 'PDF ready. Send it, then click Mark Completed.' : missingPdfReason(row);
  if (state === 'blank') return 'Add the missing details in Review & edit.';
  if (row.reminder_due || row.overdue) return Number.isFinite(row.age_days) && row.age_days >= 0
    ? `No certificate yet. Received ${row.age_days} days ago.` : 'No certificate yet. Open Review & edit to create it.';
  return 'Open Review & edit to check the details.';
}

function makeStatus(row) {
  const td = element('td', 'queue-status');
  const state = stateOf(row);
  const status = element('span', `status-pill ${statusClass(row)}`, stateLabel(row));
  if (state === 'blank') {
    status.setAttribute('aria-label', 'No status selected');
    status.title = 'No status selected';
  }
  td.append(status);
  if ($('#status-filter').value === 'needs_attention') {
    td.append(element('small', 'age-note attention-reason', attentionReason(row)));
    return td;
  }
  const destruction = destructionFollowup(row);
  if (destruction) {
    const note = {missing:'PDF missing', send:'PDF ready to send', create:'PDF to create'}[destruction];
    td.append(element('small', destruction === 'send' ? 'age-note' : 'row-alert-note', note));
    return td;
  }
  if (['generated', 'generated_manual', 'completed', 'na'].includes(state)) return td;
  if (Number.isFinite(row.age_days) && row.age_days < 0 && needsAttention(row)) td.append(element('small', 'age-note', 'Received date is in the future'));
  else if (row.reminder_due) td.append(element('small', 'age-note', `${row.age_days ?? '25+'} days since received · Pending`));
  else if (row.overdue) td.append(element('small', 'age-note', 'Over 30 days since received'));
  return td;
}

function actionButton(title, action, rowIndex, primary = false) {
  const button = element('button', `row-action${primary ? ' primary-row-action' : ''}`, title);
  button.type = 'button';
  button.dataset.action = action;
  button.dataset.index = String(rowIndex);
  return button;
}

function recordIdentityIssue(row) {
  if (row?.duplicate_serial_number) return noticeText((row.warnings || []).find(value => /Serial Number.*(?:repeated|more than once)/i.test(value))
    || `Serial Number ${row.sn || ''} is repeated.`);
  if (row?.record_identity_conflict) return typeof row.record_identity_conflict === 'string'
    ? noticeText(row.record_identity_conflict) : 'Check the Serial Number and COD in Excel, then refresh.';
  return '';
}

function certificateWasGenerated(row) {
  return Boolean(row.has_issued_certificate || row.pdf_filename || row.hdd_cod_no
    || ['generated', 'generated_manual', 'completed'].includes(stateOf(row)));
}
function canSetPrimaryCertificateStatus(row) {
  return hasInward(row) && !recordIdentityIssue(row) && !row.has_issued_certificate && !row.pdf_filename
    && !row.external_pdf && !row.pdf_missing && ['blank','ready','alert','na'].includes(stateOf(row));
}
function canDeleteEntry(row) {
  if (recordIdentityIssue(row) || certificateWasGenerated(row)) return false;
  const linked = app.state.rows.filter(item => isManual(item) && (item.destruction_source || item.raw_details?.destruction_source)
    && ((item.destruction_source || item.raw_details?.destruction_source).record_id === row.record_id && row.record_id
      || !isManual(row) && (item.destruction_source || item.raw_details?.destruction_source).plant === row.plant
        && String((item.destruction_source || item.raw_details?.destruction_source).case_id) === String(row.sn)));
  return !linked.some(certificateWasGenerated);
}

function deletionAction(row, index, action) {
  const remove=actionButton('',action,index);
  const label=action==='remove-inward' ? 'Delete inward' : 'Delete entry';
  remove.classList.add('remove-inward-action');remove.title=label;
  remove.setAttribute('aria-label',`${label} for ${row.supplier || 'this company'}`);
  remove.innerHTML = '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true"><path d="M3 6h18M9 6V4h6v2M5 6l1 14h12l1-14M10 10v6M14 10v6"/></svg>';
  return remove;
}

function makeActions(row, index) {
  const td = element('td', 'queue-actions');
  const wrap = element('div', 'row-actions');
  if ($('#status-filter').value === 'details_to_finish' && hasInward(row) && !recordIdentityIssue(row) && certificateWasGenerated(row)) {
    wrap.append(actionButton('Review & edit', 'edit-details', index, true));
    td.append(wrap); return td;
  }
  if (recordIdentityIssue(row)) {
    wrap.append(actionButton('View details', 'details', index));
  }
  else if (stateOf(row) === 'ready' || (isManual(row) && ['blank', 'draft'].includes(stateOf(row)))) {
    wrap.append(actionButton('Review & edit', 'edit-details', index, true));
  }
  else if (['generated', 'generated_manual', 'completed'].includes(stateOf(row)) && row.pdf_filename) {
    let url = null;
    try { url = certificateUrl({filename: row.pdf_filename, relative_path: row.pdf_relative_path}); } catch { /* Keep invalid paths out of links. */ }
    if (url) {
      const link = element('button', 'row-action primary-row-action', 'Open PDF'); link.type = 'button';
      link.addEventListener('click', () => openCertificateLocally(row.pdf_relative_path || row.pdf_filename));
      link.setAttribute('aria-label', `Open PDF for ${row.supplier || 'this company'}`);
      wrap.append(link);
      if (stateOf(row) !== 'completed') wrap.append(actionButton('Mark Completed', 'complete', index));
    } else wrap.append(actionButton('Review & edit', 'edit-details', index));
  }
  else if (['generated', 'generated_manual'].includes(stateOf(row))) wrap.append(actionButton('Review & edit', 'edit-details', index, true));
  else {
    wrap.append(actionButton('Review & edit', 'edit-details', index, true));
  }
  if (!recordIdentityIssue(row) && (row.has_issued_certificate || ['generated', 'generated_manual', 'completed'].includes(stateOf(row)))) {
    if (!wrap.children.length || ![...wrap.children].some(button => button.dataset.action === 'edit-details')) {
      wrap.append(actionButton('Review & edit', 'edit-details', index));
    }
    if (stateOf(row) === 'completed') wrap.append(actionButton('Undo Completed', 'undo-completed', index));
  }
  if (canSetPrimaryCertificateStatus(row) || (isManual(row) && row.record_id && !recordIdentityIssue(row) && !certificateWasGenerated(row))) {
    const state = stateOf(row);
    if (state === 'na') wrap.append(actionButton('Undo N/A', 'undo-na', index));
    else if (['blank','draft','ready','alert'].includes(state)) wrap.append(actionButton('Mark N/A', 'mark-na', index));
  }
  if (isManual(row)) {
    const target = excelTarget(row); if (target) wrap.append(excelOpenButton(target, row.supplier, 'row-action'));
    if (row.record_id && canDeleteEntry(row)) {
      wrap.append(deletionAction(row,index,'delete-entry'));
    }
  }
  else if (['plant1','plant2'].includes(row.plant) && String(row.sn ?? '').trim()) {
    wrap.append(excelOpenButton({plant:row.plant,sn:String(row.sn),sheet:row.sheet,row:row.row}, row.supplier, 'row-action'));
    if (canDeleteEntry(row)) {
      wrap.append(deletionAction(row,index,'remove-inward'));
    }
  }
  const main = element('div', 'row-actions-main');
  const extra = element('div', 'row-actions-extra');
  for (const control of [...wrap.children]) {
    (main.children.length < 2 && control.tagName?.toLowerCase() !== 'details' && control.tag !== 'details' ? main : extra).append(control);
  }
  if (main.children.length) wrap.append(main);
  if (extra.children.length) wrap.append(extra);
  td.append(wrap);
  return td;
}

function makeDetailRow(row, index) {
  const certificate = savedCertificateDetails(row);
  const tr = element('tr', 'detail-row');
  const td = element('td');
  td.colSpan = 6;
  const content = element('div', 'detail-content');
  content.append(element('p', 'next-step', nextStep(row)));
  const facts = element('dl', 'certificate-facts');
  const weight = certificate.certificate_weight || row.display_weight;
  for (const [label, value] of [['Certificate Weight', weight != null ? `${readableNumber(weight)} kg` : 'Not entered'], ['Form 6', certificate.form6 || 'Not entered'], ['Reference', certificate.reference || 'Not entered'], ['SN', row.sn || 'Manual certificate']]) {
    if (label === 'Reference' && templateId(row) === 'destruction') continue;
    const fact = element('div'); fact.append(element('dt', '', label), element('dd', '', value)); facts.append(fact);
  }
  content.append(facts);
  const issues = rowWarningSummary(row);
  if (issues.length) {
    const warning = element('div', 'detail-warnings');
    warning.append(element('strong', '', ['generated', 'generated_manual', 'completed'].includes(stateOf(row)) ? 'Warnings accepted when generated' : 'Check before creating'));
    const list = element('ul');
    renderWarningList(list, issues, true);
    warning.append(list);
    content.append(warning);
  }
  const items = Array.isArray(certificate.items) ? certificate.items : [];
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
  actions.append(actionButton(recordIdentityIssue(row) ? 'View details' : 'Edit details', recordIdentityIssue(row) ? 'details' : 'edit-details', index));
  if (!isManual(row)) actions.append(actionButton('Open plant workbook', 'open-workbook', index));
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
  const filters = readRecordFilters();
  summarize(filters);
  renderWorkbookIssues();
  const view = $('#status-filter').value === 'needs_attention' ? 'attention' : 'queue';
  $('#page-title').textContent = app.workspaceView === 'files' ? 'Saved certificates' : view === 'attention' ? 'Needs attention' : 'Your certificates';
  $('#page-description').textContent = view === 'attention'
    ? 'Check missing details and pending certificates.'
    : 'Find a company, check its certificate, and create the PDF.';
  if (app.workspaceView === 'files') $('#page-description').textContent = 'Find a PDF and open it. Files added manually appear here too.';
  if (app.workspaceView === 'analytics') {
    $('#page-title').textContent = 'Dashboard';
    $('#page-description').textContent = 'Your certificates and pending work.';
  }
  $('#queue-title').textContent = recordStatusTitles[filters.status] || recordStatusTitles.all;
  $('#queue-subtitle').textContent = ({
    generated: 'Open a PDF to check it. After sending, mark it Completed.',
    generated_manual: 'These PDFs were prepared outside the app. Send them, then mark them Completed.',
    completed: 'These certificates are marked Completed.',
    na: 'No certificate is needed for these rows. Automatic creation skips them.',
    needs_attention: 'Read the issue beside each company, then choose its next action. Completed and N/A stop date reminders.',
    needs_inward_date: 'Use Review & edit or Open in Excel to check the missing or invalid received date.',
    details_to_finish: 'These clients have missing comments, transport, docket or weight slip details. Open Review & edit to fill them.',
  })[$('#status-filter').value] || 'Choose a company to check its certificate.';
  document.querySelectorAll('[data-quick-status]').forEach(button => button.setAttribute('aria-pressed', String(button.dataset.quickStatus === $('#status-filter').value)));
  document.querySelectorAll('[data-filter-status]').forEach(button => button.setAttribute('aria-pressed', String(button.dataset.filterStatus === $('#status-filter').value)));
  $('#attention-status-filter').hidden = view !== 'attention';
  if (view !== 'attention') $('#attention-only').checked = false;
  updateRecordFilterButton(filters);

  const filtered = app.state.rows.map((row, index) => ({ row, index })).filter(item => matches(item.row, false, filters));
  filtered.sort((a, b) => compareRows(a.row, b.row));
  app.filtered = filtered;
  const page = visiblePage(filtered, app.page);
  app.page = page.page;
  const tbody = $('#queue-body');
  tbody.replaceChildren();
  if (!filtered.length) {
    const tr = element('tr');
    const td = element('td'); td.colSpan = 6;
    const empty = element('div', 'empty-state');
    empty.append(element('span', 'empty-icon', '✓'));
    const hasRecords = app.state.rows.length > 0;
    empty.append(element('strong', '', hasRecords ? 'No companies match these filters' : 'No companies loaded yet'));
    empty.append(element('span', '', hasRecords ? 'Try another name or number, or clear the filters.' : 'Choose your plant Excel files in Settings, then select Refresh.'));
    const add = element('button', 'button button-quiet', 'Company not listed? New certificate');
    add.type = 'button'; add.addEventListener('click', () => openDetails()); empty.append(add);
    td.append(empty); tr.append(td); tbody.append(tr);
  } else {
    for (const { row, index } of page.rows) {
      const tr = element('tr', 'data-row');
      tr.dataset.row = String(index);
      tr.append(cell(companySerial(row), 'company-sn'), makeClientCell(row, index));
      const plantCell = element('td');
      plantCell.append(element('span', `plant-badge ${row.plant === 'plant2' ? 'plant2' : ''}`, labelPlant(row.plant)));
      tr.append(plantCell);
      tr.append(cell(dateLabel(savedCertificateDetails(row).inward_date), 'field-value'));
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
}

function renderWorkbookIssues() {
  const notes = $('#workbook-maintenance-note');
  const plants = [...new Set((app.state.maintenance_notes || []).map(note => labelPlant(note.plant)))];
  notes.hidden = plants.length === 0;
  notes.textContent = plants.length
    ? `${plants.join(' and ')}: Excel row order was kept. Sort dates in Excel if needed; certificates can still be created.` : '';
  const messages = (app.state.workbook_issues || []).map(row => row.detail || '').filter(Boolean);
  if (app.state.setup_warning) messages.push(app.state.setup_warning);
  if (app.workbookSyncError) messages.push(app.workbookSyncError);
  if (app.state.scan_complete === false && !app.state.workbooks_disconnected && !messages.length) messages.push('Both Excel files could not be checked. Choose the correct plant files in Settings, then refresh.');
  const unique = [...new Set(messages)];
  $('#workbook-issues').hidden = unique.length === 0;
  renderWarningList($('#workbook-issue-list'), unique, true);
}

function setView(view) {
  clearRecordFilterValues();
  if (view === 'attention') $('#status-filter').value = 'needs_attention';
  restartRecordList();
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
    app.workbookSyncError = '';
    setState(response.state);
    if (response.state?.workbooks_disconnected) toast('No Excel files connected.');
    else if (response.state?.already_up_to_date) toast('Everything is up to date.');
    else if (response.state?.scan_complete === false || response.state?.setup_warning) toast('Some Excel files could not be checked. See the workbook message above the company list.', true);
    else toast('Workbooks refreshed.');
  } catch (error) { toast(error.message, true); }
  finally { setBusy(false); }
}

function applyReminderReview() {
  if (!app.reminderReviewPending || app.loadingWorkspace || app.generationBusy || app.detailsBusy
      || app.backupBusy || !$('#busy-overlay').hidden || document.querySelector('dialog[open]')) return;
  app.reminderReviewPending = false;
  setWorkspaceView('records');
  setView('attention');
  renderFolderInventory();
  window.focus?.();
}

async function pollReminderReview() {
  if (app.pollingReminderReview || app.stopping) return;
  app.pollingReminderReview = true;
  app.reviewClientId ||= `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
  try {
    // This lightweight presence check must also run during an edit. It does
    // not reread Excel or show connection-error popups in the background.
    const response = await fetch('/api/dashboard-presence', {method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({client_id:app.reviewClientId, visible:!document.hidden, acknowledged:app.lastReminderReview,
        preview_session:$('#details-modal').open ? app.editor?.previewSession : null})});
    if (!response.ok) return;
    const result = await response.json();
    if (result.review_request?.id && result.review_request.id !== app.lastReminderReview) {
      app.lastReminderReview = result.review_request.id;
      app.reminderReviewPending = true;
      if (document.querySelector('dialog[open]')) toast('Needs attention will open when you finish this panel. Your edits are kept.');
    }
    applyReminderReview();
  } catch { /* The next presence check retries; certificate actions report errors normally. */ }
  finally { app.pollingReminderReview = false; }
}

async function syncWorkbookChanges() {
  if (app.loadingWorkspace || app.syncingWorkbooks || app.stopping || !$('#offline-screen').hidden || !$('#busy-overlay').hidden
      || document.hidden || document.querySelector('dialog[open]')) return;
  app.syncingWorkbooks = true;
  const stateVersion = app.stateVersion || 0;
  try {
    const response = await api('/api/workbook-changes', undefined, {background:true});
    if (stateVersion !== (app.stateVersion || 0) || app.completionBusy
        || app.removalBusy || app.entryDeletionBusy || document.querySelector('dialog[open]')) return;
    const hadError = Boolean(app.workbookSyncError);
    app.workbookSyncError = '';
    if (response.state) setState(response.state);
    else if (hadError) renderWorkbookIssues();
  } catch (error) {
    if (app.workbookSyncError !== error.message) {
      app.workbookSyncError = error.message;
      renderWorkbookIssues();
    }
  } finally { app.syncingWorkbooks = false; }
}

function rowWarningSummary(row) {
  const detail = String(row.detail || '').trim();
  if (Array.isArray(row.warnings) && row.warnings.length) return uniqueWarnings(row.warnings);
  if (row.outcome === 'needs_review' && stateOf(row) === 'ready') return [detail || 'Review this row before issuing.'];
  if (row.system_alert && detail && stateOf(row) !== 'generated') return [detail];
  return [];
}

function showWarnings(warnings, continuation) {
  const entries = warnings.filter(row => typeof row !== 'string');
  const messages = warnings.filter(row => typeof row === 'string');
  if (messages.length) entries.unshift({supplier:continuation.payload?.details?.supplier || 'Certificate', warnings:messages});
  app.pendingConfirmation = { ...continuation, hasWarnings: entries.length > 0 };
  $('#warning-go-back').textContent = continuation.fromDetails ? 'No, go back' : 'No, cancel';
  const confirm = $('#confirm-warnings');
  confirm.textContent = entries.length ? 'Yes, continue anyway' : 'Yes, create';
  confirm.classList.toggle('button-danger', entries.length > 0);
  confirm.classList.toggle('button-primary', entries.length === 0);
  const count = continuation.payload?.selections?.length || 1;
  $('#warning-title').textContent = entries.length ? 'Check before creating' : count > 1 ? `Create ${count} certificate PDFs?` : 'Create this certificate PDF?';
  $('#warning-consequence').hidden = entries.length === 0;
  $('#warning-description').textContent = entries.length ? entries.length === 1
    ? 'Check each issue below before creating this certificate.' : `${entries.length} certificates need a check. Review the issues for each company.` : 'This will mark the selected certificates as Generated.';
  const list = $('#warning-list');
  list.replaceChildren();
  for (const row of entries) {
    const item = element('div', 'warning-item');
    item.append(element('strong', '', `${row.plant ? `${labelPlant(row.plant)} · ` : ''}${row.supplier || row.company || `Client ${row.sn ?? ''}`}`));
    const issues = rowWarningSummary(row);
    const details = issues.length ? issues : [row.detail || 'Please review this row.'];
    const bullets = element('ul');
    renderWarningList(bullets, details);
    item.append(bullets);
    list.append(item);
  }
  list.hidden = entries.length === 0;
  $('#warning-modal').showModal();
  $('#warning-go-back').focus();
}

function returnFromWarnings() {
  const next = app.pendingConfirmation;
  $('#warning-modal').close();
  app.pendingConfirmation = null;
  if (next?.fromDetails && $('#details-modal').open) {
    setEditorMode(true);
    const form = $('#details-form');
    form.elements.namedItem('supplier').focus();
  }
}

async function submitGeneration(payload, confirmed = false, endpoint = '/api/generate', fromDetails = false) {
  if (app.generationBusy || app.detailsBusy || app.backupBusy || app.completionBusy
      || app.clientBusy || app.settingsBusy || app.stopping) return;
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
    if (created || response.unissued?.length) {
      if (created && fromDetails) { clearWorkingDraft(); app.editor.baseline = detailsBody($('#details-form')); $('#details-modal').close(); }
      showGenerationSuccess(response);
    } else toast('No certificates were generated. Check the selected rows.');
  } catch (error) {
    if (fromDetails && endpoint === '/api/manual/generate' && error.workspaceRecovered
        && app.state.scan_complete === true && !app.state.refresh_pending && app.editor?.linkedHddSaveSnapshot
        && app.editor.linkedHddSaveSnapshot === linkedHddSaveSnapshot()) {
      app.editor.workspaceRevision = app.workspaceRevision;
      invalidatePreview();
      saveWorkingDraft();
      showDetailsError('Excel was refreshed. The linked inward and HDD record still match. Review your retained edits, then select the save button again.');
    } else if (fromDetails) showDetailsError(error.message);
    else toast(error.message, true);
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
      app.attachments = { ...(response.annexures || {}), ...(app.attachments || {}) };
      app.templatesLoaded = true;
      for (const select of [$('#template-filter'), $('#saved-template-filter'), $('#details-form').elements.namedItem('certificate_type')]) {
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
    values[input.dataset.certificateField] = isDateInput(input) ? normalizeDate(input.value) || input.value : input.value;
  }
}

function updateAnnexureDisplay() {
  const values = app.editor.fields[app.editor.template] || {};
  const attachment = app.editor.annexures[values._annexure] || app.attachments?.[values._annexure];
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
    (app.attachments ||= {})[attachment.id] = attachment;
    app.editor.annexures[attachment.id] = attachment;
    app.editor.annexureId = attachment.id;
    app.editor.fields[app.editor.template]._annexure = attachment.id;
    updateAnnexureDisplay();
    invalidatePreview();
    // The attachment changes the workspace. Reread the rows instead of
    // treating the picker's token as permission to save cached Excel data.
    // Keep this editor's reviewed snapshot frozen if external data changed.
    await refreshState();
  } catch (error) { toast(error.message, true); }
  finally { setDetailsBusy(false); setBusy(false); schedulePreview(); }
}

function deviceReceiptEditor(raw, key) {
  const section = element('section', 'device-receipts');
  section.append(element('h3', '', 'HDD / data devices received'), element('p', 'form-hint', 'Enter the received date and whole number of HDDs or other data devices for each receipt. Enter dates as DD/MM/YY.'));
  const storage = element('textarea'); storage.hidden = true; storage.dataset.certificateField = key;
  let rows = [];
  try {
    const parsed = JSON.parse(raw || '[]');
    if (Array.isArray(parsed)) rows = parsed.map(row => ({
      date: typeof row?.date === 'string' ? row.date : '',
      count: typeof row?.count === 'string' || typeof row?.count === 'number' ? row.count : '',
    }));
  } catch { /* Keep incomplete entries editable. */ }
  if (!rows.length) rows = [{date:'', count:''}];
  const list = element('div', 'device-receipt-list');
  const total = element('p', 'device-receipt-total');
  total.setAttribute('role', 'status'); total.setAttribute('aria-live', 'polite');
  const add = element('button', 'button button-quiet', '+ Add receipt'); add.type = 'button';
  function sync(changed = false) {
    storage.value = JSON.stringify(rows);
    const active = rows.filter(row => row.date || row.count);
    const validCount = row => /^[0-9]+$/.test(String(row.count)) && Number(row.count) > 0 && Number(row.count) <= 1000000000;
    const count = active.reduce((sum, row) => sum + (validCount(row) ? Number(row.count) : 0), 0);
    const incomplete = active.filter(row => !normalizeDate(row.date) || !validCount(row)).length;
    total.textContent = active.length ? `Total devices: ${count.toLocaleString()} across ${active.length} receipt${active.length === 1 ? '' : 's'}` : 'Add a received date and device count to begin.';
    if (incomplete) total.textContent += ` · ${incomplete} receipt${incomplete === 1 ? '' : 's'} need${incomplete === 1 ? 's' : ''} details`;
    if (rows.length > 9) total.textContent += ' · Only 9 rows fit. Remove the extra rows.';
    total.classList.toggle('incomplete', Boolean(incomplete) || rows.length > 9);
    add.disabled = rows.length >= 9;
    if (changed) storage.dispatchEvent(new Event('input', {bubbles:true}));
  }
  function renderRows() {
    list.replaceChildren();
    rows.forEach((row, index) => {
      const line = element('div', 'device-receipt-row');
      for (const [name, title, type] of [['date','Date received','date'], ['count','Number of devices','number']]) {
        const label = element('label', 'field'); label.append(element('span', '', title));
        const input = element('input'); input.type = type; input.value = row[name] || '';
        if (name === 'count') { input.min = '1'; input.max = '1000000000'; input.step = '1'; }
        input.setAttribute('aria-label', `${title}, receipt ${index + 1}`);
        label.append(input);
        if (type === 'date') configureDateInput(input, row[name]);
        input.addEventListener('input', () => { row[name] = name === 'date' ? normalizeDate(input.value) || input.value : input.value; sync(true); });
        line.append(label);
      }
      const remove = element('button', 'text-button', 'Remove'); remove.type = 'button';
      remove.setAttribute('aria-label', `Remove receipt ${index + 1}`);
      remove.addEventListener('click', () => { rows.splice(index, 1); renderRows(); sync(true); });
      line.append(remove); list.append(line);
    });
  }
  add.addEventListener('click', () => { if (rows.length < 9) { rows.push({date:'', count:''}); renderRows(); sync(true); } });
  renderRows(); sync(); section.append(storage, list, total, add);
  return section;
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
  app.editor.template = id;
  const typeSelect = form.elements.namedItem('certificate_type');
  const destructionOption = Array.from(typeSelect.options || []).find(option => option.value === 'destruction');
  if (destructionOption) { destructionOption.hidden = false; destructionOption.disabled = false; }
  const inward = editorInwardRow();
  // A linked HDD is a separate record. Its other choice opens the inward;
  // it never changes the HDD record into a recycling certificate.
  typeSelect.disabled = Boolean(app.editor.existing && id === 'destruction' && !inward);
  for (const option of typeSelect.options || []) {
    option.disabled = Boolean(id === 'destruction' && inward && !['destruction', templateId(inward)].includes(option.value));
  }
  const form6Required = id === 'standard' && (app.editor?.firstGeneration
    || (app.editor && form.elements.namedItem('source_kind').value === 'workbook' && !app.editor.existing));
  $('#form6-hint').textContent = form6Required ? 'Required for certificate' : (id === 'standard' ? 'Expected' : 'Optional');
  const values = app.editor.fields[id] ||= {};
  if (spec && ['sita', 'siemens'].includes(id) && !Object.hasOwn(values, 'company_display')) {
    values.company_display = spec?.default_company || '';
  }
  if (id === 'siemens') {
    if (app.editor.firstGeneration && app.editor.sourceLocation) values.site_location = app.editor.sourceLocation;
    else if (!values.site_location) values.site_location = form.elements.namedItem('location').value || '';
  }
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
      if (field.type === 'receipts') {
        container.append(deviceReceiptEditor(values[field.key] || '', field.key));
        continue;
      }
      if (['weight_display', 'company_display'].includes(field.key)) continue;
      if (group === 'wording' && field.key !== 'company_display') continue;
      if (group === 'details' && commonTemplateKeys.has(field.key)) {
        const input = form.elements.namedItem(field.key);
        if (input && !app.editor.existing && !input.value && !app.editor.touched.has(field.key)) {
          input.value = isDateInput(input) && field.default ? dateLabel(field.default) : String(field.default ?? '');
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
      const date = group === 'details' && (field.type === 'date' || field.key.endsWith('_date'));
      const input = element(date ? 'input' : 'textarea');
      if (date) input.type = 'date';
      else input.rows = field.multiline ? 3 : 1;
      input.dataset.certificateField = field.key;
      const value = Object.hasOwn(values, field.key) ? values[field.key] : field.key === 'company_display' ? '' : field.default;
      if (!date) input.value = String(value ?? '');
      if (field.key === 'company_display') input.placeholder = 'Leave blank to use Company / supplier and City';
      if (field.key === 'site_location' && app.editor.firstGeneration && app.editor.sourceLocation) {
        label.append(element('small', '', 'Filled from Excel. You will be asked before saving changes to Excel.'));
      }
      label.append(input);
      if (date) configureDateInput(input, value);
      if (field.help) label.append(element('small', '', field.help));
      container.append(label);
    }
  }
  updateAnnexureDisplay();
  updateDetailsButtons();
}

function editorInwardRow() {
  const editor = app.editor;
  const localId = editor?.destructionSource?.record_id || editor?.localInwardId;
  if (localId) return app.state.rows.find(row => isManual(row) && row.record_id === localId && templateId(row) !== 'destruction') || null;
  const link = editor?.destructionSource || (editor?.sourceSn != null ? {plant:editor.sourcePlant, case_id:editor.sourceSn} : null);
  if (!link) return null;
  const rows = app.state.rows.filter(row => hasInward(row) && row.plant === link.plant
    && String(row.sn).trim() === String(link.case_id).trim());
  return rows.length === 1 && !recordIdentityIssue(rows[0]) ? rows[0] : null;
}

function editorDestructionRequest() {
  const editor = app.editor;
  const link = editor?.destructionSource || (editor?.localInwardId ? {record_id:editor.localInwardId}
    : editor?.sourceSn != null ? {plant:editor.sourcePlant, case_id:editor.sourceSn} : null);
  return linkedDestructionRequest(link);
}

function linkedDestructionRequest(link) {
  if (!link) return null;
  const rows = app.state.rows.filter(row => isManual(row) && templateId(row) === 'destruction'
    && (link.record_id ? (row.destruction_source || row.raw_details?.destruction_source)?.record_id === link.record_id
      : !(row.destruction_source || row.raw_details?.destruction_source)?.record_id
        && (row.destruction_source || row.raw_details?.destruction_source)?.plant === link.plant
        && String((row.destruction_source || row.raw_details?.destruction_source)?.case_id) === String(link.case_id)));
  return rows.length === 1 ? rows[0] : null;
}

function activeDestructionRequest(link) {
  const request = linkedDestructionRequest(link);
  return request && request.destruction_requested !== false && stateOf(request) !== 'na' ? request : null;
}

function preferredCertificateRow(row) {
  if (stateOf(row) !== 'na' || templateId(row) === 'destruction' || recordIdentityIssue(row)) return row;
  const link = isManual(row) ? (row.record_id ? {record_id:row.record_id} : null)
    : hasInward(row) ? {plant:row.plant, case_id:row.sn} : null;
  return activeDestructionRequest(link) || row;
}

function linkedHddSaveSnapshot() {
  const editor = app.editor;
  if (editor?.template !== 'destruction' || !editor.destructionSource?.plant || editor.destructionSource.record_id) return null;
  const inward = editorInwardRow();
  const request = editorDestructionRequest();
  if (!inward || inward.workbook_read_only || !request?.raw_details || request.record_id !== editor.recordId) return null;
  return JSON.stringify({
    inward: workbookRowPayload(inward), file: inward.file, sheet: inward.sheet, row: inward.row,
    columns: inward.column_mapping, inwardStatus: stateOf(inward), hddNumber: inward.hdd_cod_no,
    hdd: request.raw_details, hddStatus: stateOf(request), requested: request.destruction_requested,
    pdf: request.pdf_relative_path || request.pdf_filename, pdfMissing: request.pdf_missing,
  });
}

function switchesLinkedCertificate(type) {
  return Boolean(app.editor?.existing && type !== app.editor.template
    && ((app.editor.sourceSn != null || app.editor.localInwardId) && type === 'destruction'
      || app.editor.template === 'destruction' && app.editor.destructionSource && type !== 'destruction'));
}

function switchLinkedCertificate(type) {
  // Restore the current value before checking for unsaved edits. Choosing the
  // other certificate is navigation, not an edit to this certificate's type.
  $('#details-form').elements.namedItem('certificate_type').value = app.editor.template;
  if (app.detailsBusy || app.generationBusy || app.switchingCertificate) return;
  const inward = editorInwardRow();
  if (!inward) { showDetailsError('This client’s Excel entry could not be found. Refresh the list and try again.'); return; }
  const editor = app.editor;
  const currentRow = editor.template === 'destruction' ? editorDestructionRequest() : inward;
  return requestDetailsClose(async () => {
    app.switchingCertificate = true;
    setDetailsBusy(true);
    const previousStatus = $('#preview-status').textContent;
    try {
      let target = type === 'destruction' ? editorDestructionRequest() : inward;
      if (type === 'destruction' && !target) {
        clearTimeout(app.previewTimer);
        $('#preview-status').textContent = 'Opening Data Destruction…';
        const response = await api('/api/destruction/request', {plant:inward.plant, sn:String(inward.sn),
          ...(isManual(inward) ? {record_id:inward.record_id} : {}), expected_revision:editor.workspaceRevision});
        if (response.state) setState(response.state);
        target = app.state.rows.find(row => isManual(row) && row.record_id === response.record_id && templateId(row) === 'destruction');
        if (!target) throw new Error('The HDD request could not be opened. Refresh the list and try again.');
      }
      clearWorkingDraft();
      $('#details-modal').close();
      setDetailsBusy(false);
      await openDetails(target, true, true);
    } catch (error) {
      setDetailsBusy(false);
      if (!$('#details-modal').open) await openDetails(currentRow || inward, true, true);
      else {
        $('#preview-status').textContent = previousStatus;
        if (!app.editor?.previewFacts) schedulePreview();
      }
      showDetailsError(error.message);
    } finally { app.switchingCertificate = false; setDetailsBusy(false); }
  }, {keepOpen: true});
}

function bothCertificatesUnchecked() {
  return !$('#primary-certificate-needed-card').hidden && !$('#primary-certificate-needed').checked
    && !$('#destruction-needed').disabled && !$('#destruction-needed').checked;
}

function updateDestructionChoice(payload = null) {
  const editor = app.editor;
  const linkedRequestEditor = editor?.template === 'destruction' && Boolean(editor.destructionSource);
  const available = ((editor?.sourceSn != null || editor?.localInwardId) && editor.template !== 'destruction') || linkedRequestEditor;
  $('#destruction-needed-card').hidden = !available;
  $('#destruction-needed-label').textContent = linkedRequestEditor
    ? 'Keep this Data Destruction request active' : 'Data Destruction needed';
  const checkbox = $('#destruction-needed');
  const request = available ? editorDestructionRequest() : null;
  const issued = Boolean(request && certificateWasGenerated(request));
  $('#destruction-request-toggle').hidden = issued;
  $('#destruction-issued-label').hidden = !issued;
  checkbox.disabled = !available || issued || Boolean(editor?.identityIssue) || Boolean(linkedRequestEditor && !request);
  if (request && (issued || editor.destructionChoiceRecordId !== request.record_id)) {
    checkbox.checked = request.destruction_requested !== false;
    editor.destructionChoiceRecordId = request.record_id;
  }
  const inward = available ? editorInwardRow() : null;
  const primaryCard = $('#primary-certificate-needed-card');
  if (primaryCard) {
    primaryCard.hidden = !inward || templateId(inward) === 'destruction' || !canSetPrimaryCertificateStatus(inward);
    if (!primaryCard.hidden) {
      const label = templateLabel(inward);
      const needed = stateOf(inward) !== 'na' && !editor.pendingPrimaryNa;
      if (!editor.savingPrimaryChoice) $('#primary-certificate-needed').checked = needed;
      $('#primary-certificate-needed-label').textContent = `${label} certificate needed`;
      $('#primary-certificate-needed-status').textContent = editor.pendingPrimaryNa
        ? `Save to mark ${label} N/A and keep the inward.` : needed
        ? `Uncheck to skip ${label}. The inward and any HDD request are kept.`
        : `Check to return ${label} to certificate review.`;
    }
  }
  const companion = $('#open-companion-certificate');
  companion.hidden = !inward || Boolean(editor?.identityIssue);
  companion.textContent = linkedRequestEditor && inward
    ? `Open ${templateLabel(inward)}` : 'Open Data Destruction (HDD)';
  const requested = Boolean(request && request.destruction_requested !== false);
  $('#destruction-choice-status').textContent = issued
    ? unavailableCertificatePdf(request) ? 'PDF missing. Recreate it with the same certificate number.'
      : stateOf(request) === 'completed' ? 'Marked Completed.' : 'PDF ready to send.'
    : linkedRequestEditor
    ? !request ? 'This request could not be matched to its inward. Refresh before changing it.'
      : !checkbox.checked && requested ? 'Save to remove this request. Any existing HDD files and details will be kept.'
      : checkbox.checked && requested ? 'This request stays in the follow-up list until it is Completed.'
      : checkbox.checked ? 'Save to return this request to the follow-up list.'
      : 'This request is removed from follow-up. Its saved details and any files are kept.'
    : !checkbox.checked && requested ? 'Save to remove this request. Existing HDD files and details will be kept.'
    : request && checkbox.checked ? 'Open the HDD certificate whenever you’re ready.'
    : checkbox.checked ? 'Save this request below to add it to follow-up.'
    : 'You can create the HDD certificate first. Check this box to save an HDD request for later.';
  const pendingRequest = available && !issued && checkbox.checked !== requested;
  const excelChanges = editor?.baseline && payload && fieldDifferences(editor.excelBaseline || editor.baseline, payload).length > 0;
  const changes = editor?.baseline && payload ? fieldDifferences(editor.excelBaseline || editor.baseline, payload) : [];
  const newEdits = editor?.baseline && payload ? fieldDifferences(editor.baseline, payload) : [];
  const onlyInwardChanges = inwardChanges(payload).length > 0 && newEdits.every(change => Object.hasOwn(followupFields, change.key));
  const canSaveLinkedRequest = linkedRequestEditor && Boolean(request) && pendingRequest;
  const canSaveInward = available && !linkedRequestEditor
    && (editor.localInwardId ? pendingRequest : !editor.issuedCertificate || pendingRequest || onlyInwardChanges);
  const bothUnchecked = bothCertificatesUnchecked();
  const pendingCertificateChoices = (bothUnchecked && pendingRequest) || editor?.pendingPrimaryNa;
  $('#save-details').hidden = Boolean(editor?.uploadRow || editor?.noCertificate || bothUnchecked || editor?.pendingPrimaryNa);
  $('#save-inward-details').hidden = !(canSaveLinkedRequest || canSaveInward || pendingCertificateChoices) || Boolean(editor.identityIssue);
  $('#save-inward-details').textContent = pendingCertificateChoices
    ? bothUnchecked ? 'Save as N/A' : 'Save certificate choices'
    : linkedRequestEditor || editor?.localInwardId || (pendingRequest && !onlyInwardChanges && (editor.issuedCertificate || !excelChanges))
    ? 'Save request' : 'Save to Excel';
  if (bothUnchecked) $('#change-scope').textContent = pendingCertificateChoices
    ? 'Saves both certificate choices as N/A and keeps the inward.'
    : 'Both certificates are marked N/A. The inward is kept.';
}

function changeDestructionChoice() {
  const checkbox = $('#destruction-needed');
  const request = editorDestructionRequest();
  if (checkbox.disabled || (request && certificateWasGenerated(request))) {
    updateDestructionChoice();
    return;
  }
  renderChangeSummary(); saveWorkingDraft();
}

function destructionRequestNeeded() {
  const request = editorDestructionRequest();
  return request && certificateWasGenerated(request)
    ? request.destruction_requested !== false : Boolean($('#destruction-needed').checked);
}

async function changePrimaryCertificateNeeded() {
  const row = editorInwardRow();
  if (!row || !canSetPrimaryCertificateStatus(row) || app.detailsBusy) return;
  const editor = app.editor;
  const checkbox = $('#primary-certificate-needed');
  const needed = checkbox.checked;
  if ((editor.pendingPrimaryNa && needed) || (!needed && bothCertificatesUnchecked())) {
    editor.pendingPrimaryNa = !needed;
    renderChangeSummary();
    if (editorDirty()) saveWorkingDraft();
    else if (clearWorkingDraft()) editor.workingDraftSaved = false;
    return;
  }
  editor.savingPrimaryChoice = true;
  setDetailsBusy(true);
  try {
    const response = await persistPrimaryCertificateChoice(row, needed);
    if (editor.template !== 'destruction') {
      editor.noCertificate = !needed;
      editor.certificateStatus = needed ? '' : 'N/A';
      invalidatePreview();
      setEditorMode(editor.editingDetails);
      editor.baseline.starting_cod = undefined;
      schedulePreview();
    }
    updateDestructionChoice(detailsBody($('#details-form')));
    if (editorDirty()) saveWorkingDraft();
    toast(needed ? `Certificate status is ${response.status}.` : 'Certificate marked N/A. The inward was kept.');
  } catch (error) {
    checkbox.checked = !needed;
    showDetailsError(error.message);
  } finally {
    editor.savingPrimaryChoice = false;
    setDetailsBusy(false);
  }
}

async function persistPrimaryCertificateChoice(row, needed) {
  const response = await api('/api/clients/status', {plant:row.plant, sn:String(row.sn),
    expected_revision:app.editor.workspaceRevision, action:needed ? 'undo_na' : 'mark_na'});
  if (response.state) setState(response.state);
  app.editor.workspaceRevision = response.workspace_revision ?? app.workspaceRevision;
  return response;
}

function updateSourceControls() {
  const form = $('#details-form');
  const kind = form.elements.namedItem('certificate_type').value;
  if (kind === 'destruction') form.elements.namedItem('source_kind').value = 'manual';
  const manual = form.elements.namedItem('source_kind').value === 'manual';
  const certificateWeight = form.elements.namedItem('certificate_weight');
  const useCertificateWeight = form.elements.namedItem('use_certificate_weight');
  const existingWorkbook = app.editor?.existing && app.editor?.sourceSn != null && !manual;
  const manualRecord = manual && app.editor?.recordId
    ? app.state.rows.find(row => isManual(row) && row.record_id === app.editor.recordId) : null;
  const existingIssuedManual = manual && app.editor?.existing && app.editor?.recordId
    && (app.editor.issuedCertificate || manualRecord?.has_issued_certificate
      || ['generated', 'generated_manual', 'completed'].includes(stateOf(manualRecord || {})));
  $('#edit-scope-field').hidden = true;
  if (!manual && !existingWorkbook) form.elements.namedItem('edit_scope').value = 'update_excel';
  form.elements.namedItem('case_id').readOnly = Boolean(existingWorkbook);
  form.elements.namedItem('plant').disabled = Boolean(existingWorkbook || existingIssuedManual);
  form.elements.namedItem('plant').title = existingIssuedManual
    ? 'Plant is fixed for an issued certificate.' : '';
  if (existingIssuedManual && (manualRecord?.plant || app.editor.sourcePlant)) {
    form.elements.namedItem('plant').value = manualRecord?.plant || app.editor.sourcePlant;
    app.editor.issuedCertificate = true;
  }
  if (existingWorkbook) {
    form.elements.namedItem('case_id').value = app.editor.sourceSn;
    form.elements.namedItem('plant').value = app.editor.sourcePlant;
  }
  form.elements.namedItem('source_kind').disabled = Boolean(existingWorkbook || app.editor?.localInwardId || kind === 'destruction');
  const sourceChoice = form.elements.namedItem('source_kind').closest?.('.field');
  if (sourceChoice) sourceChoice.hidden = kind === 'destruction';
  const plantChoice = form.elements.namedItem('plant').closest?.('.field');
  if (plantChoice) plantChoice.hidden = manual;
  $('#record-options').hidden = Boolean(existingWorkbook);
  $('#edit-scope-hint').textContent = existingWorkbook
    ? 'Excel still records the certificate number, date and Generated status. Cancel does not save changes.'
    : 'A new Excel row is saved only when the PDF is created successfully.';
  form.elements.namedItem('case_id').required = !manual;
  form.elements.namedItem('form6').required = !manual && kind === 'standard' && (!app.editor?.existing || Boolean(app.editor?.firstGeneration));
  form.elements.namedItem('reference').required = kind !== 'destruction';
  form.elements.namedItem('reference').disabled = kind === 'destruction';
  $('#certificate-number-label').textContent = kind === 'destruction' ? 'Certificate number (HDD)' : 'Certificate number (COD)';
  $('#case-required').hidden = manual;
  $('#case-optional').hidden = !manual;
  $('#pick-workbook').hidden = manual || Boolean(existingWorkbook);
  $('#details-plant-label').textContent = manual ? 'Plant' : 'Plant *';
  $('#material-options').hidden = manual;
  $('#save-draft').hidden = !manual;
  $('#draft-options').hidden = !manual;
  $('#certificate-weight-choice').hidden = true;
  certificateWeight.disabled = false;
  certificateWeight.required = !manual && Boolean(useCertificateWeight?.checked);
  certificateWeight.closest('.field').hidden = false;
  form.elements.namedItem('case_id').closest('.field').hidden = manual;
  form.elements.namedItem('inward_date').closest('.field').hidden = manual && kind === 'siemens';
  form.elements.namedItem('form6').closest('.field').hidden = kind !== 'standard';
  form.elements.namedItem('reference').closest('.field').hidden = kind === 'destruction';
  certificateWeight.closest('.field').querySelector('span').textContent = certificateWeight.required
    ? 'Certificate Weight (kg) · Required' : 'Certificate Weight (kg)';
  certificateWeight.placeholder = manual ? 'Enter the weight in kg' : 'Enter a different weight in kg';
  $('#edit-items-hint').hidden = manual || app.editor.firstGeneration || !app.editor.itemHint;
  $('#details-description').textContent = manual
    ? 'Enter the company details below. You do not need an Excel row.'
    : app.editor?.firstGeneration && !form.elements.namedItem('form6').value.trim()
      ? 'Enter the Form 6 number now. It will be saved with the certificate.'
    : existingWorkbook ? 'These details came from Excel. Change only what you need, then check the preview.'
    : 'Enter the company details. The app will save a new Excel row with the PDF.';
  $('#save-details').textContent = 'Create PDF';
  $('#weight-options').open = manual || Boolean(useCertificateWeight?.checked);
  $('#weight-options').hidden = kind === 'destruction';
  if (kind === 'destruction') {
    form.elements.namedItem('inward_date').closest('.field').hidden = true;
    form.elements.namedItem('form6').required = false;
    form.elements.namedItem('form6').closest('.field').hidden = true;
  }
  $('#weight-options-title').textContent = 'Certificate Weight';
  $('#certificate-weight-hint').textContent = manual ? 'Enter the total weight to print on this certificate.'
    : 'Enter the weight to print. Clear it to use the Excel weight or material total.';
  $('#case-id-hint').textContent = existingWorkbook ? 'This identifies the original Excel record, so it cannot be changed here.'
    : 'Filled in automatically when your Excel SNs are numeric. You can change it.';
  suggestWorkbookId();
  updatePrintedWeight();
  setEditorMode(app.editor?.editingDetails ?? true);
  updateDestructionChoice();
}

function nextWorkbookId(plant) {
  const ids = app.state.rows.filter(row => !isManual(row) && row.plant === plant)
    .map(row => String(row.sn ?? row.case_id ?? '').trim()).filter(Boolean);
  if (!ids.length || ids.some(id => !/^\d+$/.test(id) || !Number.isSafeInteger(Number(id)))) return '';
  const next = ids.reduce((maximum, id) => Math.max(maximum, Number(id)), 0) + 1;
  return Number.isSafeInteger(next) ? String(next) : '';
}

function suggestWorkbookId() {
  if (!app.editor || app.editor.sourceSn != null || app.editor.touched.has('case_id')) return;
  const form = $('#details-form');
  if (form.elements.namedItem('source_kind').value !== 'workbook') return;
  const input = form.elements.namedItem('case_id');
  const suggestion = nextWorkbookId(form.elements.namedItem('plant').value);
  if (!input.value || input.value === app.editor.suggestedCaseId) input.value = suggestion;
  app.editor.suggestedCaseId = suggestion;
  $('#case-id-hint').textContent = suggestion ? 'Next SN filled in. You can change it. Excel is checked again before saving.'
    : 'Enter a unique SN for this new Excel record.';
}

function printedWeight() {
  const form = $('#details-form');
  const value = name => form.elements.namedItem(name)?.value || '';
  const override = value('source_kind') === 'manual'
    ? Boolean(value('certificate_weight').trim()) : Boolean(form.elements.namedItem('use_certificate_weight')?.checked);
  const total = materialWeightTotal();
  const excelWeight = app.editor?.excelCertificateWeight ?? '';
  const hasExcelWeight = String(excelWeight).trim() !== '';
  return {weight: override ? value('certificate_weight') : hasExcelWeight ? excelWeight
    : total.state === 'complete' ? readableNumber(total.total) : '',
    source: override ? 'entered here' : hasExcelWeight ? 'Certificate Weight in Excel' : 'material total'};
}

function applyCertificateWeight(value) {
  const form = $('#details-form');
  const input = form.elements.namedItem('certificate_weight');
  const entered = String(value).trim() !== '';
  form.elements.namedItem('use_certificate_weight').checked = entered;
  input.disabled = false;
  input.required = form.elements.namedItem('source_kind').value === 'workbook' && entered;
  app.editor.touched.add('certificate_weight');
  const fields = app.editor.fields[app.editor.template] ||= {};
  if (Object.hasOwn(fields, 'weight_display')) fields.weight_display = '';
}

function usesSummaryEditor() {
  return Boolean(app.editor?.uploadRow || (app.editor?.existing && !app.editor.firstGeneration && app.editor.template !== 'destruction'));
}

function summaryEditorVisible() {
  return Boolean((app.editor?.firstGeneration || usesSummaryEditor()) && !app.editor?.editingDetails);
}

const summaryFields = {Company:'supplier', City:'location', Received:'inward_date', 'Reference number':'reference',
  'Form 6':'form6', 'COD number (preview)':'cod_no', 'COD date':'cod_date', 'Certificate Weight':'certificate_weight',
  'Company on PDF':'certificate_fields.company_display'};

function inlineSource(key) {
  if (key.startsWith('certificate_fields.')) return [...$('#template-editor').querySelectorAll('[data-certificate-field]')]
    .find(input => input.dataset.certificateField === key.slice('certificate_fields.'.length));
  return $('#details-form').elements.namedItem(key);
}

function finishInlineEdit() {
  if (!app.editor?.inlineEdit) return;
  app.editor.inlineEdit = null;
  renderEditorSummary();
}

function applyInlineValue(key, value) {
  if (!usesSummaryEditor() || !app.editor.inlineEdit) return;
  const form = $('#details-form');
  const source = inlineSource(key);
  if (!source) return;
  source.value = value;
  app.editor.touched.add(key);
  if (key === 'certificate_weight') applyCertificateWeight(value);
  captureTemplateFields(); invalidatePreview(); schedulePreview();
}

function inlineInput(key, label, initialValue) {
  const source = inlineSource(key);
  if (!source) return null;
  const multiline = source.type === 'textarea' || source.tagName?.toLowerCase() === 'textarea';
  const input = element(multiline ? 'textarea' : 'input', 'inline-fact-input');
  if (multiline) input.rows = source.rows;
  else input.type = source.type || 'text';
  input.inputMode = source.inputMode;
  if (source.maxLength >= 0) input.maxLength = source.maxLength;
  input.required = Boolean(source.required);
  input.placeholder = source.placeholder || '';
  input.value = initialValue ?? source.value;
  input.setAttribute('aria-label', label);
  if (isDateInput(source)) configureDateInput(input);
  // The original form controls remain the only named fields. Opening a pencil
  // does not change the payload, numeric precision, or source data.
  const update = event => { event.stopPropagation?.(); applyInlineValue(key, input.value); };
  input.addEventListener('input', update); input.addEventListener('change', update);
  input.addEventListener('keydown', event => {
    if (event.key === 'Enter' && (!multiline || source.rows === 1)) { event.preventDefault(); finishInlineEdit(); }
  });
  return input;
}

function renderInlineMaterials(container) {
  container.replaceChildren();
  const items = materialValues();
  items.forEach((item, index) => {
    const line = element('div', 'inline-material');
    const heading = element('div', 'inline-material-heading'); heading.append(element('strong', '', `Material ${index + 1}`));
    if (items.length > 1) {
      const remove = element('button', 'text-button', 'Remove'); remove.type = 'button';
      remove.setAttribute('aria-label', `Remove material ${index + 1}`);
      remove.addEventListener('click', () => {
        const next = materialValues(); next.splice(index, 1); renderMaterialRows(next);
        app.editor.touched.add('items');
        if (index === 0) for (const key of ['item', 'quantity', 'per_piece_weight', 'weight']) app.editor.touched.add(key);
        renderInlineMaterials(container); invalidatePreview(); schedulePreview();
      }); heading.append(remove);
    }
    line.append(heading);
    const {grid, inputs} = materialWeightEditor(item, index, (key, calculated) => {
      if (!key) return;
      for (const [field, input] of Object.entries(inputs)) app.materialRows[index].inputs[field].value = input.value;
      app.editor.touched.add(materialInputName(index, key)); app.editor.touched.add('items');
      if (calculated || ['quantity', 'per_piece_weight'].includes(key)) app.editor.touched.add(materialInputName(index, 'weight'));
      updateMaterialWeightTotal();
      const total = materialWeightTotal();
      if (app.editor.inlineEdit?.total) app.editor.inlineEdit.total.textContent = total.state === 'complete' ? `Material total: ${readableNumber(total.total)} kg` : 'Check each material weight.';
      invalidatePreview(); schedulePreview();
    });
    for (const [key, input] of Object.entries(inputs)) input.parentElement?.classList.add(`inline-material-${key}`);
    line.append(grid); container.append(line);
  });
  const total = materialWeightTotal();
  const note = element('p', 'form-hint', total.state === 'complete' ? `Material total: ${readableNumber(total.total)} kg` : 'Check each material weight.');
  app.editor.inlineEdit.total = note; container.append(note);
  const add = element('button', 'button button-quiet', '+ Add another material'); add.type = 'button'; add.disabled = items.length >= 100;
  add.addEventListener('click', () => {
    renderMaterialRows([...materialValues(), {}]); app.editor.touched.add('items');
    renderInlineMaterials(container); invalidatePreview(); schedulePreview();
  }); container.append(add);
}

function startInlineEdit(key, label) {
  if (!usesSummaryEditor() || app.detailsBusy) return;
  if (key === 'cod_no' || (key === 'cod_date' && !app.editor?.issuedCertificate)) return;
  if (!['materials', ...Object.values(summaryFields)].includes(key) && !key.startsWith('certificate_fields.')) return;
  if (key !== 'materials' && !inlineSource(key)) return;
  finishInlineEdit();
  const host = key === 'materials' ? $('#editor-materials') : app.editor.summaryHosts?.[key];
  if (!host) return;
  const editor = element('div', 'inline-field-editor');
  app.editor.inlineEdit = {key, label, container:editor};
  host.replaceChildren(editor);
  let input;
  if (key === 'materials') renderInlineMaterials(editor);
  else {
    input = inlineInput(key, label, key === 'certificate_weight' ? printedWeight().weight : undefined);
    if (input) {
      editor.append(input);
      if (isDateInput(input)) configureDateInput(input);
    }
    if (key === 'supplier' && $('#details-form').elements.namedItem('certificate_type').value === 'standard') {
      const companyInput = inlineInput('certificate_fields.company_display', 'Company line on PDF (optional)');
      if (companyInput) {
        const field = element('label', 'field'); field.append(element('span', '', 'Company line on PDF (optional)'), companyInput); editor.append(field);
      }
    }
    if (key === 'certificate_weight') {
      if ($('#details-form').elements.namedItem('source_kind').value === 'workbook') {
        const reset = element('button', 'text-button', String(app.editor.excelCertificateWeight ?? '').trim() ? 'Use Excel weight' : 'Use material total'); reset.type = 'button';
        reset.addEventListener('click', () => {
          $('#details-form').elements.namedItem('use_certificate_weight').checked = false;
          const weight = inlineSource('certificate_weight'); weight.value = ''; weight.required = false; weight.disabled = false;
          const fields = app.editor.fields[app.editor.template] ||= {};
          if (Object.hasOwn(fields, 'weight_display')) fields.weight_display = '';
          app.editor.touched.add('use_certificate_weight'); finishInlineEdit(); invalidatePreview(); schedulePreview();
        }); editor.append(reset);
      }
    }
  }
  const done = element('button', 'button button-quiet inline-done', 'Done'); done.type = 'button';
  done.setAttribute('aria-label', `Done editing ${label.toLowerCase()}`); done.addEventListener('click', finishInlineEdit); editor.append(done);
  input?.focus();
}

function updatePrintedWeight() {
  if (!app.editor) return;
  const {weight, source} = printedWeight();
  const note = $('#current-weight');
  const form = $('#details-form');
  const manual = form.elements.namedItem('source_kind').value === 'manual';
  const total = materialWeightTotal();
  const assetNote = total.state === 'complete' ? `Assets total: ${readableNumber(total.total)} kg.`
    : total.state === 'invalid' ? 'Asset weights need checking.'
      : total.state === 'incomplete' ? 'Asset total is incomplete.' : 'No asset weights entered yet.';
  note.hidden = true;
  if (source === 'entered here') {
    note.textContent = `${assetNote} ${String(weight).trim() !== ''
      ? `Certificate will use your entered Certificate Weight: ${readableNumber(weight)} kg.`
      : 'Enter your Certificate Weight below; this is the weight that will print.'}`;
  } else if (source === 'Certificate Weight in Excel') {
    note.textContent = `${assetNote} Certificate will use Excel Certificate Weight: ${readableNumber(weight)} kg.`;
  } else if (total.state === 'complete') {
    note.textContent = `${assetNote} Certificate will use this total.`;
  } else {
    note.textContent = `${assetNote} Certificate will use the asset total once all weights are entered and valid.`;
  }
}

function updateSingleCodControls() {
  if (!app.editor) return;
  const form = $('#details-form');
  const input = form.elements.namedItem('starting_cod');
  const numbering = app.editor.numbering || app.state.numbering;
  const knownCod = app.state.rows.some(row => String(row.certificate_no || '').trim())
    || (app.state.certificate_inventory?.entries || []).some(entry => entry.cods?.length);
  const fixed = form.elements.namedItem('cod_no').value.trim() || app.editor.issuedCertificate;
  const needsStart = !app.editor.noCertificate && !fixed && (numbering ? numbering.needs_start : !knownCod);
  app.editor.needsStartingCod = Boolean(needsStart);
  $('#single-start-cod-section').hidden = !needsStart;
  input.required = Boolean(needsStart);
  input.disabled = !needsStart;
  input.setCustomValidity('');
  const value = input.value.trim();
  const valid = /^[0-9]{1,12}$/.test(value) && Number(value) >= 1;
  if (needsStart && value && !valid) input.setCustomValidity('Enter a starting COD number (1 or higher).');
  const format = number => (numbering?.prefix || '') + String(number).padStart(numbering?.width || 4, '0');
  $('#single-start-cod-hint').textContent = valid
    ? `First COD: ${format(Number(value))}. Next: ${format(Number(value) + 1)}.`
    : 'No previous COD found. Enter the number to start from.';
  if (needsStart && !value && !app.previewLoading) $('#preview-status').textContent = 'Enter the starting COD number to preview.';
}

function renderEditorSummary() {
  if (!app.editor?.baseline) return;
  renderEditorFollowup();
  // Keep the active input mounted while automatic previews and draft summaries
  // refresh, so typing never loses focus or the caret position.
  if (app.editor.inlineEdit) return;
  const form = $('#details-form');
  const value = name => form.elements.namedItem(name)?.value || '';
  const {weight, source} = printedWeight();
  const preview = app.editor.previewFacts;
  const codDate = preview?.certificate_date || value('cod_date');
  const facts = $('#editor-facts'); facts.replaceChildren();
  const weightText = preview?.printed_details?.certificate_weight || (String(weight).trim() !== '' ? `${readableNumber(weight)} kg` : '');
  const rows = [['Company', value('supplier') || 'Company name missing'],
    ['COD number (preview)', preview?.certificate_no || value('cod_no') || (app.editor.needsStartingCod ? 'Enter the starting COD above' : 'Assigned automatically')],
    ['COD date', codDate ? dateLabel(codDate) : 'Set automatically']];
  if (value('certificate_type') !== 'destruction' || (app.editor.uploadRow && value('reference'))) {
    rows.push(['Reference number', value('reference') || 'Not entered']);
  }
  if (value('certificate_type') === 'standard' || (app.editor.uploadRow && value('form6'))) {
    rows.push(['Form 6', value('form6') || 'Not entered']);
  }
  rows.push(['City', value('location') || 'Not entered'], ['Received', dateLabel(value('inward_date'))],
    ['Plant', labelPlant(value('plant'))], ['SN', value('case_id') || 'Not entered'],
    ['Certificate Weight', weightText ? (app.editor.uploadRow ? weightText : `${weightText} · ${source}`) : 'Not entered']);
  const companyOnPdf = preview?.printed_details?.company_display;
  const expectedCompany = [value('supplier'), value('location').trim()].filter(Boolean).join(', ').toUpperCase();
  if (companyOnPdf && companyOnPdf !== expectedCompany && value('certificate_type') === 'standard' && !usesSummaryEditor()) rows.splice(1, 0, ['Company on PDF', companyOnPdf]);
  if (usesSummaryEditor()) {
    const spec = app.templates.find(template => template.id === value('certificate_type'));
    for (const field of spec?.fields || []) {
      if (commonTemplateKeys.has(field.key) || ['weight_display', 'company_display'].includes(field.key)) continue;
      if (field.key === 'company_display' && value('certificate_type') === 'standard'
          && !app.editor.fields[app.editor.template]?.company_display && (!companyOnPdf || companyOnPdf === expectedCompany)) continue;
      if (!inlineSource(`certificate_fields.${field.key}`)) continue;
      rows.push([field.key === 'company_display' ? 'Company on PDF' : field.label,
        app.editor.fields[app.editor.template]?.[field.key] || (field.key === 'company_display' ? companyOnPdf || expectedCompany : 'Not entered'), `certificate_fields.${field.key}`]);
    }
  }
  const attachment = app.editor.fields?.[app.editor.template]?._annexure;
  if (attachment) rows.push(['Attached list', app.editor.annexures[attachment]?.name || app.attachments?.[attachment]?.name || 'PDF list attached']);
  app.editor.summaryHosts = {};
  for (const [label, text, fieldKey] of rows) {
    const key = fieldKey || summaryFields[label];
    const fact = element('div', label === 'Certificate Weight' ? 'weight-fact' : '');
    const heading = element('dt', 'fact-heading', label), display = element('dd', '', text);
    if (usesSummaryEditor() && key) app.editor.summaryHosts[key] = display;
    fact.append(heading, display); facts.append(fact);
  }
  const materials = $('#editor-materials'); materials.replaceChildren();
  const sourceNote = $('#editor-source-note');
  const notes = app.editor.importNotes || [];
  sourceNote.hidden = true;
  sourceNote.textContent = app.editor.sourceSheet
    ? [`Excel: ${app.editor.sourceSheet} · headings on row ${app.editor.headerRow || 1}.`, ...notes].join(' ') : '';
  sourceNote.title = Object.entries(app.editor.columnMapping || {}).map(([key, column]) => `${fieldLabels[key] || key}: ${column.column} (${column.heading})`).join('\n');
  const items = materialValues().filter(item => Object.values(item).some(entry => String(entry).trim()));
  const heading = element('h4', 'fact-heading', usesSummaryEditor() ? 'Materials' : 'Materials from Excel');
  materials.append(heading);
  if (!items.length) materials.append(element('p', 'form-hint', 'No material details entered.'));
  else {
    const list = element('ul', 'editor-material-list');
    for (const item of items) {
      const line = element('li');
      const total = materialWeightTotal([item]);
      const details = [];
      if (item.quantity.trim()) details.push(`Quantity: ${readableNumber(item.quantity)}`);
      if (item.per_piece_weight.trim()) details.push(`${readableNumber(item.per_piece_weight)} kg per piece`);
      details.push(total.state === 'complete' ? `${readableNumber(total.total)} kg total` : 'Weight needs checking');
      line.append(element('strong', '', item.item || 'Material name not entered'), element('span', '', details.join(' · ')));
      list.append(line);
    }
    materials.append(list);
  }
}

function setEditorMode(editing = false) {
  const first = Boolean(app.editor?.firstGeneration);
  if (app.editor) app.editor.editingDetails = Boolean(editing);
  const summaryVisible = summaryEditorVisible();
  const form = $('#details-form');
  const number = form.elements.namedItem('cod_no').value.trim();
  $('#certificate-selection-status').textContent = `${templateNames[app.editor?.template] || 'Certificate'} · ${number
    ? `${app.editor.template === 'destruction' ? 'HDD' : 'COD'} ${number}`
    : 'New number when created'}${app.editor?.certificateStatus ? ` · ${app.editor.certificateStatus}` : ''}`;
  updateSingleCodControls();
  $('#editor-fields').hidden = false;
  $('#editor-summary').hidden = !summaryVisible;
  $('#editor-summary').setAttribute('aria-label', first ? 'Details from Excel' : 'Certificate details');
  $('#excel-detail-fields').hidden = summaryVisible;
  $('#template-editor').hidden = summaryVisible;
  $('#edit-items-hint').hidden = summaryVisible || !app.editor?.itemHint;
  $('#editor-summary-note').textContent = first && editing
    ? (!form.elements.namedItem('form6').value.trim()
      ? 'The Form 6 number will be saved to Excel with the certificate.'
      : 'Choose whether to save these changes to Excel when you save.')
    : '';
  $('#editor-summary-note').hidden = !$('#editor-summary-note').textContent;
  $('#change-details').hidden = true;
  $('#finish-editing').hidden = true;
  $('#edit-pdf-details').hidden = !(first || usesSummaryEditor());
  if (app.editor?.uploadRow) $('#edit-pdf-details').hidden = true;
  $('#edit-pdf-details').textContent = editing ? 'Done editing' : '✎ Edit details';
  $('#edit-pdf-details').setAttribute('aria-pressed', String(Boolean(editing)));
  $('#editor-mode-title').textContent = editing ? 'Edit details' : app.editor?.existing ? 'Review & edit' : 'Certificate details';
  $('#details-description').textContent = first
    ? 'Review the inward details and create this certificate when you’re ready.'
    : app.editor?.existing
    ? 'Review the certificate, then save corrections with its current COD.'
    : 'Choose a template, enter the details, and preview your certificate.';
  for (const name of ['supplier', 'location', 'inward_date', 'form6', 'reference', 'cod_no', 'cod_date']) {
    form.elements.namedItem(name).readOnly = name === 'cod_no'
      || (name === 'cod_date' && !app.editor?.issuedCertificate)
      || (first && !editing && ['supplier', 'location', 'inward_date', 'form6', 'reference'].includes(name));
  }
  if (first) form.elements.namedItem('case_id').readOnly = true;
  if (editing && form.elements.namedItem('source_kind').value === 'workbook'
      && !form.elements.namedItem('use_certificate_weight').checked && !app.editor.touched.has('certificate_weight')) {
    form.elements.namedItem('certificate_weight').value = printedWeight().weight;
  }
  $('#save-details').textContent = first ? 'Generate certificate' : app.editor?.existing ? 'Save corrected certificate' : 'Create PDF';
  const noCertificate = Boolean(app.editor?.noCertificate);
  $('#save-details').hidden = noCertificate || bothCertificatesUnchecked() || Boolean(app.editor?.pendingPrimaryNa || app.editor?.uploadRow);
  $('#replace-uploaded-pdf').hidden = !app.editor?.uploadRow || Boolean(app.editor.identityIssue || app.editor.uploadRow.workbook_read_only);
  $('#preview-details').hidden = noCertificate || Boolean(app.editor?.uploadRow);
  if (noCertificate) {
    $('#details-description').textContent = app.editor.sourceSn != null
      ? `The ${templateNames[app.editor.template]} certificate is marked N/A. Use Save to Excel to update the inward.`
      : 'This certificate is marked N/A. Choose Undo N/A in Companies to create its PDF.';
    $('#preview-status').textContent = `${templateNames[app.editor.template]} certificate marked N/A.`;
  }
  if (app.editor?.template === 'destruction' && !noCertificate) {
    $('#save-details').textContent = app.editor.issuedCertificate ? 'Save corrected certificate' : 'Create data destruction PDF';
    $('#details-description').textContent = app.editor.destructionSource
      ? `Data Destruction for ${labelPlant(app.editor.destructionSource.plant)} inward ${app.editor.destructionSource.case_id}. Check the actual destruction details and both preview pages.`
      : 'Enter the company and actual destruction details, then check both preview pages. An inward entry is optional.';
  }
  if (app.editor?.uploadRow) $('#details-description').textContent = 'Review the saved details and the original uploaded certificate.';
  renderEditorSummary();
}

function validateEditor() {
  const form = $('#details-form');
  const start = form.elements.namedItem('starting_cod');
  if (start.required && !start.checkValidity()) {
    start.reportValidity(); start.focus(); return false;
  }
  if (form.checkValidity()) return true;
  setEditorMode(true);
  if (usesSummaryEditor()) {
    const invalid = [...form.querySelectorAll(':invalid')][0];
    if (invalid?.name && inlineSource(invalid.name)) {
      startInlineEdit(invalid.name, fieldLabels[invalid.name] || invalid.name);
      const input = app.editor.inlineEdit?.container.querySelector('input, textarea');
      if (input) { input.reportValidity(); return false; }
    }
  }
  // A required field may be in a closed disclosure. Reveal its correction path.
  for (const input of form.querySelectorAll(':invalid')) {
    for (let parent = input.parentElement; parent && parent !== form; parent = parent.parentElement) {
      if (parent.tagName === 'DETAILS') parent.open = true;
    }
  }
  form.reportValidity();
  return false;
}

async function finishEditing() {
  if (app.detailsBusy || !validateEditor()) return;
  setEditorMode(false);
  await previewDetails();
}

function updateDetailsButtons() {
  const unavailable = app.detailsBusy || !app.templatesLoaded || !app.templates.some(template => template.id === $('#details-form').elements.namedItem('certificate_type').value);
  for (const id of ['save-details', 'save-inward-details', 'save-draft', 'preview-details', 'pick-annexure', 'remove-annexure', 'keep-working-draft']) $(`#${id}`).disabled = unavailable;
  $('#open-companion-certificate').disabled = Boolean(app.detailsBusy || app.generationBusy);
  $('#preview-details').disabled = unavailable || Boolean(app.previewLoading);
  if ($('#details-form').elements.namedItem('source_kind').value === 'workbook' && app.editor?.issuanceBlocked) $('#save-details').disabled = true;
  if (app.editor?.identityIssue) {
    $('#preview-details').disabled = true;
    showDetailsError(app.editor.identityIssue);
    $('#preview-status').textContent = 'Preview unavailable';
  }
}

function setDetailsBusy(on) {
  app.detailsBusy = on;
  $('#details-form').inert = on;
  updateDetailsButtons();
}

function invalidatePreview() {
  clearAttachmentPreview();
  $('#preview-detail-images').replaceChildren();
  clearTimeout(app.previewTimer);
  app.previewRevision++;
  if (app.editor) app.editor.previewFacts = null;
  $('#details-error').hidden = true;
  $('#preview-frame').removeAttribute('src');
  $('#preview-frame').hidden = true;
  $('#preview-image').removeAttribute('src');
  $('#preview-image').hidden = true;
  $('#preview-image-note').hidden = true;
  $('#preview-full-pdf').textContent = 'Enlarge preview';
  $('#preview-full-pdf').setAttribute('aria-expanded', 'false');
  $('#preview-link').removeAttribute('href');
  $('#preview-document').hidden = true;
  $('#preview-number').hidden = true;
  $('#preview-warnings').hidden = true;
  $('#editor-checks').hidden = true;
  $('#preview-warnings-title').hidden = true;
  $('#preview-warnings').replaceChildren();
  $('#preview-status').textContent = app.editor?.noCertificate ? `${templateNames[app.editor.template]} certificate marked N/A.` : 'Preview updates automatically.';
  $('#preview-details').textContent = 'Update preview';
  if (app.editor?.baseline && $('#details-modal').open) {
    updatePrintedWeight();
    renderChangeSummary();
    if (editorDirty()) saveWorkingDraft();
    else if (app.editor.workingDraftSaved && clearWorkingDraft()) {
      app.editor.workingDraftSaved = false;
      $('#working-draft-status').textContent = 'No unsaved field changes. Excel is unchanged.';
    }
  }
  if (app.editor?.identityIssue) updateDetailsButtons();
  if (app.editor?.uploadRow && $('#details-modal').open) showUploadedCertificate(app.editor.uploadRow);
}

function showUploadedCertificate(row) {
  const url = certificateUrl({filename:row.pdf_filename, relative_path:row.pdf_relative_path});
  if (!url) return;
  const pdf = new URL(url);
  pdf.searchParams.set('sha256', row.uploaded_pdf_sha256);
  const query = new URLSearchParams({path:row.pdf_relative_path, sha256:row.uploaded_pdf_sha256});
  const image = $('#preview-image');
  $('#preview-link').href = pdf.href;
  $('#preview-document').hidden = false;
  $('#preview-status').textContent = 'Original uploaded certificate';
  $('#preview-number').textContent = `${row.certificate_no}${row.certificate_date ? ` · ${dateLabel(row.certificate_date)}` : ''}`;
  $('#preview-number').hidden = !row.certificate_no;
  image.hidden = false;
  image.src = `/api/certificate-preview-image?${query}`;
}

function schedulePreview({initial = false} = {}) {
  clearTimeout(app.previewTimer);
  const editor = app.editor;
  if (!editor || editor.uploadRow || editor.identityIssue || editor.noCertificate || !$('#details-modal').open || !app.templatesLoaded) return;
  app.previewTimer = setTimeout(async () => {
    if (app.editor !== editor || !$('#details-modal').open || app.detailsBusy || app.generationBusy || app.pendingConfirmation) return;
    if (app.previewLoading) { schedulePreview({initial}); return; }
    const form = $('#details-form');
    if (!form.checkValidity() || !form.elements.namedItem('supplier').value.trim()) return;
    const {weight} = printedWeight();
    if (!initial && editor.template !== 'destruction' && !String(weight).trim()) return;
    await previewDetails({automatic:true}); // Read-only: never allocates COD or writes an Excel row.
  }, 1000);
}

async function openDetails(row = null, skipRecovery = false, startEditing = false) {
  if (app.loadingWorkspace || app.generationBusy || app.backupBusy || app.completionBusy
      || app.clientBusy || app.settingsBusy || app.stopping) return;
  if ($('#details-modal').open || app.detailsBusy) return;
  if (showWorkingDraftRecovery && !skipRecovery && offerDraftRecovery(row, true, startEditing)) return;
  const sourceRow = row;
  const uploadRow = row?.uploaded_pdf_sha256 && row.pdf_relative_path ? row : null;
  const issued = row && (row.has_issued_certificate || ['generated', 'generated_manual', 'completed'].includes(stateOf(row)));
  if (row?.last_certificate_details && !recordIdentityIssue(row)) {
    row = { ...row, ...row.last_certificate_details };
  }
  const form = $('#details-form');
  form.reset();
  const rowItems = row && Array.isArray(row.items) ? row.items : [];
  renderMaterialRows(rowItems.length ? rowItems : [{ item: row?.item ?? '', quantity: row?.quantity ?? '',
    per_piece_weight: row?.per_piece_weight ?? '', weight: row?.weight ?? '' }]);
  for (const disclosure of form.querySelectorAll('details')) disclosure.open = false;
  const id = templateId(row || {});
  app.editor = { existing: Boolean(row), recordId: row?.record_id || null, template: id,
    uploadRow,
    noCertificate: Boolean(sourceRow && stateOf(sourceRow) === 'na'),
    certificateStatus: sourceRow ? issued ? stateLabel(sourceRow) : 'Not created' : '',
    destructionSource: row?.destruction_source || row?.raw_details?.destruction_source || null,
    localInwardId: row && isManual(row) && id !== 'destruction' ? row.record_id : null,
    editingDetails: !uploadRow && Boolean(startEditing || !row || (row && !isManual(row) && !issued && (!row.form6 || !row.reference))),
    pdfEditUsed: Boolean(startEditing && row && !isManual(row) && !issued),
    firstGeneration: Boolean(row && !isManual(row) && !issued), sourceLocation: sourceRow?.location || '',
    reviewEdits: Boolean(row && !isManual(row)),
    issuanceBlocked: Boolean(recordIdentityIssue(sourceRow) || app.state.scan_complete === false),
    identityIssue: recordIdentityIssue(sourceRow),
    sourceSheet: sourceRow?.sheet || '', headerRow: sourceRow?.source_header_row || 1,
    importNotes: sourceRow?.import_notes || [], columnMapping: sourceRow?.column_mapping || {},
    sourceSn: row && !isManual(row) ? String(row.sn) : null, sourcePlant: row?.plant || null,
    excelCertificateWeight: !isManual(row || {}) ? sourceRow?.certificate_weight ?? '' : '',
    issuedCertificate: Boolean(issued),
    fields: { [id]: { ...(row?.certificate_fields || {}) } }, annexures: {}, annexureId: row?.certificate_fields?._annexure || null,
    touched: new Set(), sourceExplicit: false, rawNumeric: {}, rawDetails: isManual(row || {}) ? { ...(row.raw_details || {}) } : {}, workspaceRevision: app.workspaceRevision ?? app.state.workspace_revision ?? '' };
  app.editor.linkedHddSaveSnapshot = linkedHddSaveSnapshot();
  form.elements.namedItem('certificate_type').value = id;
  form.elements.namedItem('source_kind').value = row ? (isManual(row) ? 'manual' : 'workbook') : 'workbook';
  form.elements.namedItem('edit_scope').value = row && !isManual(row) ? 'certificate_only' : 'update_excel';
  invalidatePreview();
  $('#details-title').textContent = row ? `Review & edit: ${row.supplier || `Client ${row.sn}`}` : 'Create a certificate';
  $('#save-details').innerHTML = row ? 'Save changes and create PDF <span>→</span>' : 'Save and create PDF <span>→</span>';
  const items = rowItems;
  const hint = $('#edit-items-hint');
  if (items.length > 1) {
    hint.hidden = false;
    hint.textContent = `${items.length} materials from Excel are listed below. Add, remove or change any line.`;
  } else hint.hidden = true;
  if (row) {
    const values = {
      plant: row.plant, case_id: isManual(row) ? row.case_id || '' : row.sn ?? row.case_id, supplier: row.supplier, location: row.location,
      inward_date: row.inward_date, item: items[0]?.item ?? row.item ?? '', quantity: items[0]?.quantity ?? row.quantity,
      per_piece_weight: items[0]?.per_piece_weight ?? row.per_piece_weight, weight: items[0]?.weight ?? row.weight,
      certificate_weight: row.certificate_weight, form6: row.form6, reference: row.reference,
      cod_no: row.cod_no ?? row.certificate_no, cod_date: row.cod_date ?? row.certificate_date,
      ...Object.fromEntries(Object.keys(followupFields).map(key => [key, sourceRow?.[key] || ''])),
    };
    if (isManual(row)) Object.assign(values, row.raw_details || {});
    for (const [name, value] of Object.entries(values)) {
      const input = form.elements.namedItem(name);
      if (input) {
        let normalized = value == null ? '' : String(value);
        if (isDateInput(input)) normalized = normalizeDate(value) ? dateLabel(value) : normalized;
        if (['quantity', 'per_piece_weight', 'weight', 'certificate_weight'].includes(name) && normalized !== '' && Number.isFinite(Number(normalized))) {
          app.editor.rawNumeric[name] = normalized;
          normalized = readableNumber(normalized);
        }
        input.value = normalized;
      }
    }
    const savedWeight = row.last_certificate_details?.certificate_weight;
    if (savedWeight && String(savedWeight) !== String(sourceRow?.certificate_weight ?? '')) {
      form.elements.namedItem('use_certificate_weight').checked = true;
    }
  } else {
    const selectedPlant = app.workspaceView === 'files' ? $('#saved-plant-filter').value : $('#plant-filter').value;
    if (['plant1', 'plant2'].includes(selectedPlant)) form.elements.namedItem('plant').value = selectedPlant;
    // Leave the receipt date empty so the user enters the actual date; never
    // silently substitute today's date for the material receipt date.
    form.elements.namedItem('inward_date').value = '';
  }
  const inwardTotal = items.length && items.every(item => item.weight != null && Number.isFinite(Number(item.weight)))
    ? items.reduce((sum, item) => sum + Number(item.weight), 0) : null;
  app.editor.itemHint = !hint.hidden ? hint.textContent : '';
  app.editor.inwardPlaceholder = inwardTotal != null
    ? `Leave blank to use ${readableNumber(inwardTotal)} kg from Excel` : 'Enter weight, or leave blank to use Excel';
  updateSourceControls();
  renderTemplateEditor();
  formatFormDates(form);
  updateDestructionChoice();
  app.editor.baseline = detailsBody(form);
  app.editor.excelBaseline = sourceRow && !isManual(sourceRow) ? workbookRowPayload(sourceRow) : null;
  renderChangeSummary();
  $('#working-draft-status').textContent = app.editor.firstGeneration
    ? (!form.elements.namedItem('form6').value.trim()
      ? 'The Form 6 number will be saved to Excel when you create the certificate.'
      : 'Edits are saved on this computer. You will be asked before Excel is updated.')
    : 'Edits are saved on this computer. Choose whether to update Excel when saving corrections.';
  $('#details-error').hidden = true;
  $('#details-modal').showModal();
  setEditorMode(app.editor.editingDetails);
  if (uploadRow) showUploadedCertificate(uploadRow);
  $('#record-options').open = !row;
  const editor = app.editor;
  const initialPreviewRevision = app.previewRevision;
  const templatesReady = await loadTemplates();
  if (app.editor === editor && $('#details-modal').open) {
    const changed = editorDirty();
    captureTemplateFields(); renderTemplateEditor();
    if (!changed) editor.baseline = detailsBody(form);
    renderChangeSummary();
    if (editor.identityIssue) {
      updateDetailsButtons();
    } else if (!templatesReady) {
      $('#details-error').textContent = 'Certificate templates could not be loaded. Close and reopen this certificate to try again.';
      $('#details-error').hidden = false;
    } else if (row && !changed && app.previewRevision === initialPreviewRevision
        && !app.generationBusy && !app.pendingConfirmation) schedulePreview({initial:true});
    else if (!row || changed) schedulePreview();
  }
}

function comparisonFields(payload) {
  const fields = { plant: payload.plant, source_kind: payload.source_kind, ...(payload.details || {}) };
  if (payload.starting_cod != null) fields.starting_cod = payload.starting_cod;
  const numeric = value => {
    const raw = String(value ?? '').trim();
    const grouping = '(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+|[0-9]{1,2}(?:,[0-9]{2})+,[0-9]{3})';
    const valid = new RegExp(`^[+-]?(?:${grouping}(?:\\.[0-9]*)?|\\.[0-9]+)(?:[eE][+-]?[0-9]+)?$`).test(raw);
    const number = valid ? Number(raw.replaceAll(',', '')) : NaN;
    return Number.isFinite(number) ? String(number) : raw;
  };
  for (const key of ['quantity', 'per_piece_weight', 'weight', 'certificate_weight']) {
    if (Object.hasOwn(fields, key)) fields[key] = numeric(fields[key]);
  }
  if (Array.isArray(fields.items)) fields.items = fields.items.map(item => ({item:String(item.item ?? ''),
    quantity:numeric(item.quantity), per_piece_weight:numeric(item.per_piece_weight), weight:numeric(item.weight)}));
  delete fields.certificate_fields;
  for (const [key, value] of Object.entries(payload.details?.certificate_fields || {})) fields[`certificate_fields.${key}`] = value;
  return fields;
}

function fieldDifferences(before, after) {
  const original = comparisonFields(before || {}), current = comparisonFields(after || {});
  const token = value => value && typeof value === 'object' ? JSON.stringify(value) : String(value ?? '');
  return [...new Set([...Object.keys(original), ...Object.keys(current)])].filter(key => token(original[key]) !== token(current[key]))
    .map(key => ({ key, before: token(original[key]), after: token(current[key]) }));
}

function changeDisplay(key, value) {
  if (key === 'device_receipts' || key === 'certificate_fields.device_receipts') {
    try {
      const rows = Array.isArray(value) ? value : JSON.parse(value || '[]');
      if (!Array.isArray(rows)) throw new Error('Invalid list');
      return rows.filter(row => row && typeof row === 'object' && (row.date || row.count))
        .map(row => {
          const count = String(row.count ?? '').trim();
          const amount = !count ? 'Number not entered' : /^[0-9]+$/.test(count) && Number(count) > 0
            ? `${Number(count).toLocaleString()} ${Number(count) === 1 ? 'device' : 'devices'}` : 'Check the number of devices';
          const day = !row.date ? 'Date not entered' : typeof row.date === 'string' && normalizeDate(row.date)
            ? dateLabel(row.date) : 'Check the received date';
          return `${day} · ${amount}`;
        }).join('; ') || 'No devices entered';
    } catch { return 'Check the received dates and number of devices'; }
  }
  if (key.endsWith('_date') && normalizeDate(value)) return dateLabel(value);
  if (key !== 'items') return value || '(blank)';
  try {
    const items = JSON.parse(value);
    const shown = items.filter(item => Object.values(item).some(field => String(field ?? '').trim()))
      .map(item => `${item.item || 'Unnamed material'}${item.weight ? ` · ${readableNumber(item.weight)} kg` : ''}`);
    return shown.length ? shown.join('; ') : '(no materials)';
  } catch { return value || '(blank)'; }
}

function editorDirty() {
  if (!app.editor?.baseline || !$('#details-modal').open) return false;
  const current = detailsBody($('#details-form'));
  return Boolean(app.editor.pendingPrimaryNa) || fieldDifferences(app.editor.baseline, current).length > 0 || current.edit_scope !== app.editor.baseline.edit_scope
    || Boolean(current.data_destruction_needed) !== Boolean(app.editor.baseline.data_destruction_needed);
}

function scopeDescription(payload) {
  if (workbookTypeChanged(payload)) return app.editor.issuedCertificate
    ? 'The certificate type differs from Excel. Update the inward to match when saving the corrected certificate.'
    : 'The certificate type differs from Excel. When saving, update the inward to match or delete that inward and keep the PDF as a local certificate.';
  if (payload.source_kind === 'manual') return app.editor?.destructionSource?.plant && !app.editor.destructionSource.record_id
    ? 'Saves the PDF and records its HDD number and date in Excel.'
    : 'Saves the PDF and updates local tracking. No Excel row is changed.';
  const plant = labelPlant(payload.plant);
  if (app.editor?.firstGeneration) return 'Save to Excel now, or create the PDF. You will be asked before Excel details are changed.';
  if (app.editor?.existing && app.editor.sourceSn != null) return 'When saving corrections, choose whether to update this inward in Excel too.';
  if (payload.edit_scope === 'certificate_only') return `Edits appear on the PDF only. ${plant} Excel still records the certificate number, date and Generated status.`;
  return app.editor?.sourceSn != null
    ? `Creating the PDF also updates this record in ${plant} Excel.`
    : `Creates the PDF and adds a new record to ${plant} Excel.`;
}

function renderChangeSummary() {
  if (!app.editor?.baseline) return;
  const payload = detailsBody($('#details-form'));
  $('#change-scope').textContent = app.editor.uploadRow ? '' : scopeDescription(payload);
  const changes = fieldDifferences(app.editor.baseline, payload);
  const list = $('#change-list'); list.replaceChildren();
  for (const change of changes) {
    const key = change.key.replace(/^certificate_fields\./, '');
    const spec = app.templates.find(template => template.id === app.editor.template)?.fields.find(field => field.key === key);
    const label = key === '_annexure' ? 'Attached PDF list' : fieldLabels[key] || spec?.label || key.replace(/_/g, ' ');
    const display = value => key === '_annexure' && value ? app.editor.annexures[value]?.name || value : changeDisplay(key, value);
    const entry = element('div', 'field-change');
    const values = element('dd');
    values.append(element('span', 'change-before', `Before: ${display(change.before)}`), element('span', 'change-after', `After: ${display(change.after)}`));
    entry.append(element('dt', '', label), values); list.append(entry);
  }
  $('#change-empty').hidden = changes.length > 0;
  $('#changes-disclosure').hidden = Boolean(app.editor.firstGeneration) || usesSummaryEditor() || changes.length === 0;
  renderEditorSummary();
  updateDestructionChoice(payload);
}

function readWorkingDraft() {
  try {
    const saved = JSON.parse(localStorage.getItem(workingDraftKey) || 'null');
    if (app.state.history_reset && saved?.historyReset !== app.state.history_reset) {
      localStorage.removeItem(workingDraftKey);
      return null;
    }
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
  values.certificate_type = form.elements.namedItem('certificate_type').value;
  values.use_certificate_weight = Boolean(form.elements.namedItem('use_certificate_weight')?.checked);
  values.data_destruction_needed = destructionRequestNeeded();
  // Disabled override controls are omitted by FormData, but their local edit
  // must survive recovery. This is not sent to Excel unless explicitly ticked.
  values.certificate_weight = form.elements.namedItem('certificate_weight')?.value || '';
  try {
    const {inlineEdit, summaryHosts, previewFacts, ...savedEditor} = app.editor;
    localStorage.setItem(workingDraftKey, JSON.stringify({ version: 1, historyReset:app.state.history_reset || '', savedAt: new Date().toISOString(), values,
      editor: { ...savedEditor, touched: [...app.editor.touched] } }));
    app.editor.workingDraftSaved = true;
    $('#working-draft-status').textContent = app.editor.firstGeneration
      ? 'Your certificate choices are saved on this computer.'
      : 'Edits saved on this computer. Choose whether to update Excel when saving corrections.';
    $('#keep-working-draft').hidden = true;
    return true;
  } catch {
    if (!showWorkingDraftRecovery) return false;
    $('#working-draft-status').textContent = 'The browser could not save your working draft. Keep this form open to avoid losing edits.';
    $('#draft-options').hidden = false;
    $('#keep-working-draft').hidden = false;
    $('#keep-working-draft').textContent = 'Retry saving edits';
    return false;
  }
}

function offerDraftRecovery(row = null, resumeEditor = false, startEditing = false) {
  const draft = readWorkingDraft();
  if (!draft) return false;
  app.draftRecovery = { draft, row, resumeEditor, startEditing };
  const savedDate = new Date(draft.savedAt);
  $('#draft-recovery-description').textContent = `A draft for ${draft.values.supplier || 'a certificate'}${Number.isNaN(savedDate.getTime()) ? '' : ` was saved ${dateTimeLabel(draft.savedAt)}`}. ${resumeEditor ? 'Restore it or discard it to continue with the certificate you selected.' : 'Restore it to continue editing, or discard it to stay on the dashboard.'}`;
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
    renderMaterialRows(materialValuesFromFormValues(saved.values));
    app.editor = { ...saved.editor, inlineEdit:null, summaryHosts:{}, previewFacts:null, touched: new Set(saved.editor.touched) };
    const sourceRow = app.state.rows.find(row => !isManual(row)
      && row.plant === app.editor.sourcePlant && String(row.sn) === String(app.editor.sourceSn));
    app.editor.pdfEditUsed = Boolean(saved.editor.pdfEditUsed);
    app.editor.reviewEdits = Boolean(sourceRow) || Boolean(saved.editor.reviewEdits);
    if (sourceRow) {
      app.editor.issuedCertificate = Boolean(sourceRow.has_issued_certificate || ['generated', 'generated_manual', 'completed'].includes(stateOf(sourceRow)));
      app.editor.firstGeneration = !app.editor.issuedCertificate;
      app.editor.sourceLocation = sourceRow.location || '';
      app.editor.identityIssue = recordIdentityIssue(sourceRow);
    }
    const workbookDraft = saved.values.source_kind === 'workbook';
    const missingSource = workbookDraft && app.editor.sourceSn != null && !sourceRow;
    const statusRow = sourceRow || app.state.rows.find(row => isManual(row) && row.record_id === app.editor.recordId);
    app.editor.noCertificate = Boolean(statusRow && stateOf(statusRow) === 'na');
    if (missingSource) app.editor.identityIssue = 'This inward could not be found. Refresh Excel and reopen the record.';
    app.editor.issuanceBlocked = Boolean(workbookDraft && (missingSource || app.state.scan_complete === false
      || app.editor.identityIssue));
    if (app.editor.excelCertificateWeight == null) {
      app.editor.excelCertificateWeight = app.state.rows.find(row => !isManual(row)
        && row.plant === app.editor.sourcePlant && String(row.sn) === String(app.editor.sourceSn))?.certificate_weight ?? '';
    }
    app.editor.workingDraftSaved = true;
    for (const [name, value] of Object.entries(saved.values)) {
      const input = form.elements.namedItem(name); if (input) {
        if (input.type === 'checkbox') input.checked = Boolean(value);
        else input.value = value;
      }
    }
    if (sourceRow) for (const key of Object.keys(followupFields)) {
      const value = String(sourceRow[key] || '');
      if (!Object.hasOwn(saved.values, key)) {
        const input = form.elements.namedItem(key); if (input) input.value = value;
        if (app.editor.baseline?.details) app.editor.baseline.details[key] = value;
      }
      if (app.editor.excelBaseline?.details && !Object.hasOwn(app.editor.excelBaseline.details, key)) app.editor.excelBaseline.details[key] = value;
    }
    if (sourceRow && app.editor.firstGeneration && !app.editor.pdfEditUsed && !saved.editor.reviewEdits) {
      // Drafts made before the PDF-only edit control cannot carry source edits into first issuance.
      const source = workbookRowPayload(sourceRow).details;
      for (const name of ['supplier', 'location', 'inward_date', 'form6', 'reference', 'cod_no', 'cod_date']) {
        form.elements.namedItem(name).value = source[name];
      }
      app.editor.rawNumeric = {};
      renderMaterialRows(sourceRow.items || []);
      form.elements.namedItem('use_certificate_weight').checked = false;
      app.editor.excelCertificateWeight = sourceRow.certificate_weight ?? '';
      const fields = app.editor.fields[app.editor.template] ||= {};
      for (const key of ['weight_display', ...(app.editor.template === 'standard' ? ['company_display'] : [])]) {
        if (key === 'company_display' && sourceRow.certificate_type !== 'standard') delete fields[key];
        else if (Object.hasOwn(sourceRow.certificate_fields || {}, key)) fields[key] = sourceRow.certificate_fields[key];
        else delete fields[key];
      }
    }
    for (const disclosure of form.querySelectorAll('details')) disclosure.open = false;
    $('#details-title').textContent = 'Continue your working draft';
    $('#edit-items-hint').textContent = app.editor.itemHint || '';
    formatFormDates(form);
    updateSourceControls(); renderTemplateEditor(); invalidatePreview(); renderChangeSummary();
    $('#working-draft-status').textContent = 'Working draft restored. Review the changes and preview before creating.';
    $('#details-error').hidden = true;
    $('#draft-recovery-modal').close(); $('#details-modal').showModal();
    setEditorMode(true);
    updateDetailsButtons();
    updateDestructionChoice(detailsBody(form));
    schedulePreview();
    app.draftRecovery = null;
  } finally { app.restoringDraft = false; }
}

function requestDetailsClose(afterClose = null, {keepOpen = false} = {}) {
  if (app.detailsBusy || app.generationBusy) return false;
  if ($('#details-modal').open && editorDirty()) {
    app.afterDraftClose = afterClose && (() => {
      if (!keepOpen) $('#details-modal').close();
      afterClose();
    });
    if (!$('#unsaved-modal').open) $('#unsaved-modal').showModal();
    return false;
  }
  if (!keepOpen) $('#details-modal').close();
  if (afterClose) afterClose();
  return true;
}

function finishDraftClose(discard) {
  if (!discard) {
    app.afterDraftClose = null;
    $('#unsaved-modal').close();
    return;
  }
  if (!clearWorkingDraft()) return;
  const continuation = app.afterDraftClose;
  app.afterDraftClose = null;
  $('#unsaved-modal').close();
  if (continuation) continuation();
  else $('#details-modal').close();
}

function certificateUrl(certificate) {
  const filename = String(certificate.filename || '');
  const relativePath = String(certificate.relative_path || certificate.pdf_relative_path || filename).replaceAll('\\', '/');
  const parts = relativePath.split('/');
  if (!relativePath || parts.some(part => !part || part === '.' || part === '..' || part.includes(':') || /[\x00-\x1f]/.test(part))
      || !/\.pdf$/i.test(relativePath)
      || (filename && parts.at(-1) !== filename)) return null;
  // Match the server's quote(..., safe='/') encoding, including characters
  // encodeURIComponent normally leaves literal in readable company names.
  const encodePart = part => encodeURIComponent(part).replace(/[!'()*]/g,
    character => `%${character.charCodeAt(0).toString(16).toUpperCase()}`);
  const expectedPath = `/api/certificates/${parts.map(encodePart).join('/')}`;
  const url = new URL(certificate.pdf_url || expectedPath, window.location.href);
  return url.origin === window.location.origin && url.pathname === expectedPath && !url.search && !url.hash ? url.href : null;
}

async function openCertificateLocally(relativePath, openPdf = true) {
  if (!certificateUrl({relative_path:relativePath})) {
    toast('This PDF path is unavailable. Refresh the certificate list and try again.', true);
    return;
  }
  try {
    const response = await api('/api/open-certificate', {relative_path:relativePath, reveal:true, open_pdf:openPdf});
    if (response.selected) toast(openPdf
      ? 'PDF opened and selected in its folder.' : 'PDF selected in its folder.');
  } catch (error) { toast(error.message, true); }
}

function showGenerationSuccess(response) {
  const certificates = Array.isArray(response.certificates) ? response.certificates : [];
  const unissued = Array.isArray(response.unissued) ? response.unissued : [];
  const saved = response.created ?? certificates.length;
  $('#success-title').textContent = unissued.length
    ? saved ? 'Some certificates were not created' : 'No certificates were created'
    : saved === 1 ? 'Certificate generated' : 'Certificates generated';
  $('#success-description').textContent = unissued.length
    ? `${saved} of ${saved + unissued.length} certificates saved. Resolve the issues below, then click Refresh and select only the unfinished companies.`
    : certificates.length
    ? `${certificates.length} PDF${certificates.length === 1 ? '' : 's'} saved. Check the actual files below before sending.`
    : `${response.created} PDF${response.created === 1 ? '' : 's'} generated. Open the folder to check the saved files.`;
  for (const note of new Set([response.state?.recovery_note, response.state?.setup_warning].filter(Boolean))) {
    $('#success-description').textContent += ` ${note}`;
  }
  $('#unissued-certificates').hidden = unissued.length === 0;
  $('#unissued-list').replaceChildren(...unissued.map(row => {
    const item = element('li', 'issue-card');
    item.append(element('strong', 'issue-title', `${row.supplier || 'Company'} · ${labelPlant(row.plant)} · SN ${row.sn}`),
      element('p', 'issue-message', noticeText(row.detail)));
    return item;
  }));
  const list = $('#created-certificates'); list.replaceChildren();
  for (const certificate of certificates) {
    const entry = element('li', 'saved-file');
    const copy = element('div', 'saved-file-copy');
    const filename = element('a', 'saved-pdf-name', certificate.filename);
    const meta = element('div', 'saved-file-meta');
    meta.append(element('span', 'saved-file-cod', `${labelPlant(certificate.plant)} · COD ${certificate.certificate_no || 'number not provided'}`),
      element('span', 'saved-file-type', templateNames[certificate.certificate_type] || 'Type not identified'));
    copy.append(element('strong', '', certificate.supplier || 'Certificate'), meta, filename);
    entry.append(copy);
    let url = null;
    try { url = certificateUrl(certificate); } catch { /* Keep file details visible even if a URL is invalid. */ }
    filename.href = url || '#';
    filename.addEventListener('click', event => { event.preventDefault(); openCertificateLocally(certificate.relative_path || certificate.filename); });
    const actions = element('div', 'saved-file-actions');
    if (url) {
      const link = element('button', 'button button-quiet', 'Open PDF'); link.type = 'button';
      link.addEventListener('click', () => openCertificateLocally(certificate.relative_path || certificate.filename));
      link.setAttribute('aria-label', `Open PDF for ${certificate.supplier || certificate.filename}`); actions.append(link);
    }
    const excel = excelTarget(certificate);
    if (excel) actions.append(excelOpenButton(excel, certificate.supplier));
    const info = element('button', 'text-button', 'Details'); info.type = 'button';
    info.addEventListener('click', () => showSavedFile({path:`Certificates/${certificate.relative_path || certificate.filename}`,present:true,cods:[certificate.certificate_no].filter(Boolean),references:[],certificate_type:certificate.certificate_type,
      ...(excel ? {managed_owner:{source_kind:'workbook',plant:excel.plant,key:excel.sn}} : {})}));
    actions.append(info); entry.append(actions);
    list.append(entry);
  }
  app.lastGeneratedCertificatePath = certificates[0]?.relative_path || certificates[0]?.filename || '';
  $('#success-modal').showModal();
}

function detailsBody(form) {
  const data = new FormData(form);
  const source = form.elements.namedItem('source_kind').value;
  const useCertificateWeight = source === 'manual'
    || Boolean(form.elements.namedItem('use_certificate_weight')?.checked);
  const details = {};
  for (const [key, value] of data.entries()) {
    if (key === 'plant' || key === 'source_kind' || key === 'edit_scope' || key === 'use_certificate_weight' || key === 'starting_cod' || key === 'data_destruction_needed') continue;
    details[key] = ['inward_date', 'cod_date'].includes(key) ? normalizeDate(value) || String(value).trim() : String(value).trim();
  }
  for (const key of Object.keys(followupFields)) {
    if (source === 'workbook' && app.editor?.sourceSn != null) {
      const input = form.elements.namedItem(key);
      if (input) details[key] = String(input.value ?? '').trim();
    } else delete details[key];
  }
  const extraMaterials = new Map();
  let hasMaterialFields = ['item', 'quantity', 'per_piece_weight', 'weight'].some(key => Object.hasOwn(details, key));
  for (const [name, value] of Object.entries(details)) {
    const match = name.match(/^material_(\d+)_(item|quantity|per_piece_weight|weight)$/);
    if (!match) continue;
    hasMaterialFields = true;
    const index = Number(match[1]);
    extraMaterials.set(index, { ...(extraMaterials.get(index) || { item: '', quantity: '', per_piece_weight: '', weight: '' }), [match[2]]: value });
    delete details[name];
  }
  // Clean up floating-point tails on screen, but preserve the original numeric
  // value unless the user actually edits it.
  for (const [key, value] of Object.entries(app.editor.rawNumeric || {})) {
    if (!app.editor.touched.has(key)) details[key] = value;
  }
  if (source === 'workbook' && !useCertificateWeight) delete details.certificate_weight;
  if (hasMaterialFields && (source === 'workbook' || app.editor.touched.has('items')
      || (app.editor.existing && materialValues().some(item => Object.values(item).some(value => String(value).trim()))))) {
    const rows = [{ item: details.item || '', quantity: details.quantity || '',
      per_piece_weight: details.per_piece_weight || '', weight: details.weight || '' }];
    for (const [index, item] of extraMaterials) rows[index] = item;
    details.items = rows.map(item => item || { item: '', quantity: '', per_piece_weight: '', weight: '' });
  }
  // Date pickers cannot display malformed saved dates. Preserve them until
  // explicitly edited so changing an unrelated field cannot erase a warning.
  for (const [key, value] of Object.entries(app.editor.rawDetails || {})) {
    const input = form.elements.namedItem(key);
    if (input?.type === 'date' && value && !normalizeDate(value) && !app.editor.touched.has(key)) details[key] = value;
  }
  captureTemplateFields();
  details.certificate_fields = { ...app.editor.fields[app.editor.template] };
  // Disabled selectors remain authoritative and do not appear in FormData.
  details.certificate_type = form.elements.namedItem('certificate_type').value;
  if (source === 'manual' && !app.editor.existing) details.case_id = '';
  if (app.editor.template === 'destruction' && app.editor.destructionSource) details.destruction_source = app.editor.destructionSource;
  // Old certificates may contain a printed-name override. An explicit edit
  // to the company (or Standard location) must use the main fields instead.
  if (app.editor.touched.has('supplier') || (app.editor.template === 'standard' && app.editor.touched.has('location'))) {
    details.certificate_fields.company_display = '';
  }
  const spec = app.templates.find(template => template.id === app.editor.template);
  for (const field of spec?.fields || []) {
    if (field.group === 'details' && commonTemplateKeys.has(field.key)) delete details.certificate_fields[field.key];
  }
  const excelType = app.editor?.excelBaseline?.details?.certificate_type || 'standard';
  if (source === 'workbook' && app.editor?.reviewEdits && details.certificate_type !== excelType
      && ['sita','siemens'].includes(details.certificate_type)) details.form6 = '';
  return { plant: source === 'manual' ? (app.editor.destructionSource && !app.editor.destructionSource.record_id ? app.editor.destructionSource.plant : null)
      : data.get('plant') || form.elements.namedItem('plant').value, details, source_kind: source,
    ...(app.editor.needsStartingCod && form.elements.namedItem('starting_cod').value.trim()
      ? {starting_cod:form.elements.namedItem('starting_cod').value.trim()} : {}),
    ...(app.editor?.firstGeneration && app.editor.pdfEditUsed && !app.editor.reviewEdits ? {pdf_only_edits:true} : {}),
    ...(source === 'workbook' && app.editor.reviewEdits ? {review_edits:true} : {}),
    ...((source === 'workbook' && app.editor.sourceSn != null) || app.editor?.localInwardId
      || (app.editor?.template === 'destruction' && app.editor?.destructionSource) ? {data_destruction_needed:destructionRequestNeeded()} : {}),
    ...(source === 'workbook' && app.editor.sourceSn == null ? {create_new:true} : {}),
    use_certificate_weight: source === 'workbook' && Boolean(form.elements.namedItem('use_certificate_weight')?.checked),
    edit_scope: data.get('edit_scope') || 'certificate_only',
    ...(app.editor.workspaceRevision != null ? { expected_revision: app.editor.workspaceRevision } : {}),
    ...(source === 'workbook' && app.editor.sourceSn != null ? { source_sn: app.editor.sourceSn, source_plant: app.editor.sourcePlant } : {}),
    record_id: source === 'manual' ? app.editor.recordId : null };
}

async function saveDetails(event) {
  event.preventDefault();
  if (app.editor?.uploadRow) return;
  if (app.editor?.noCertificate || app.editor?.pendingPrimaryNa || bothCertificatesUnchecked()) return saveInwardDetails();
  if (app.detailsBusy || !app.templatesLoaded) return;
  if (app.editor?.identityIssue) { updateDetailsButtons(); return; }
  finishInlineEdit();
  if (!validateEditor()) return;
  const payload = detailsBody(event.currentTarget);
  if (payload.create_new && duplicateSerialNumber(payload.plant, payload.details.case_id)) {
    setEditorMode(true); $('#record-options').open = true;
    showDetailsError('This Serial Number already exists in this plant.');
    event.currentTarget.elements.namedItem('case_id').focus(); return;
  }
  if (payload.source_kind === 'workbook' && app.editor.existing
      && fieldDifferences(app.editor.excelBaseline || app.editor.baseline, payload).length) {
    showExcelSaveChoice(payload, 'create');
    return;
  }
  await submitGeneration(payload, false, payload.source_kind === 'manual' ? '/api/manual/generate' : '/api/create', true);
}

function workbookTypeChanged(payload) {
  if (payload.source_kind !== 'workbook' || app.editor?.sourceSn == null) return false;
  const source = (app.editor.excelBaseline || app.editor.baseline)?.details;
  return Boolean(source && (source.certificate_type || 'standard') !== (payload.details.certificate_type || 'standard'));
}

function showExcelSaveChoice(payload, mode) {
  app.excelSavePayload = {payload, mode};
  const saveOnly = mode === 'save_only' || mode === 'inward_only';
  $('#excel-save-title').textContent = saveOnly ? 'Save these changes to Excel?' : 'Also update these details in Excel?';
  $('#excel-save-description').textContent = saveOnly
    ? 'Yes saves the details below. No takes you back to editing.'
    : 'Choose Yes to save your changes in Excel too, or No to use them only on this PDF.';
  $('#excel-save-note').textContent = saveOnly
    ? 'A backup is kept before updating Excel. This saves the inward details without creating a PDF.'
    : 'Both choices record the certificate number, date and Generated status. A backup is kept before updating Excel.';
  $('#excel-save-no').textContent = saveOnly ? 'No, keep editing' : 'No, keep Excel details';
  const source = (app.editor.excelBaseline || app.editor.baseline)?.details || {};
  const requiredDetails = app.editor.firstGeneration && !saveOnly
    ? [['form6','Form 6 number'],['reference','reference number']].filter(([key]) => !source[key] && payload.details[key]).map(([,label]) => label) : [];
  if (requiredDetails.length) {
    $('#excel-save-description').textContent = `The ${requiredDetails.join(' and ')} must be recorded in Excel. Save the other changes too?`;
    $('#excel-save-no').textContent = 'No, only save required details';
    $('#excel-save-note').textContent = `Both choices save the ${requiredDetails.join(' and ')}, certificate number, date and Generated status. A backup is kept before updating Excel.`;
  }
  const list = $('#excel-save-change-list'); list.replaceChildren();
  const changes = fieldDifferences(app.editor.excelBaseline || app.editor.baseline, payload)
    .filter(change => mode !== 'inward_only' || Object.hasOwn(followupFields, change.key));
  for (const change of changes) {
    const key = change.key.replace(/^certificate_fields\./, '');
    const spec = app.templates.find(template => template.id === app.editor.template)?.fields.find(field => field.key === key);
    const entry = element('div', 'field-change');
    entry.append(element('dt', '', fieldLabels[key] || spec?.label || key.replace(/_/g, ' ')),
      element('dd', '', `From: ${changeDisplay(key, change.before)} → To: ${changeDisplay(key, change.after)}`));
    list.append(entry);
  }
  if (!saveOnly && workbookTypeChanged(payload)) {
    const from = source.certificate_type || 'standard';
    const to = payload.details.certificate_type || 'standard';
    $('#excel-save-title').textContent = 'What should happen to the original Excel inward?';
    $('#excel-save-description').textContent = `Excel has ${templateNames[from] || from}; this PDF will use ${templateNames[to] || to}. Update this selected inward to match the reviewed details, or delete it from Excel and keep the PDF in the app.`;
    $('#excel-save-note').textContent = 'Deletion removes this inward and all its materials from Excel after saving the PDF. Other inwards stay unchanged. A workbook backup is kept; certificate history and linked HDD records are retained. Cancel keeps editing.';
    $('#excel-save-no').textContent = 'Delete Excel inward; keep PDF';
    const inward = editorInwardRow();
    if (app.editor.issuedCertificate || (inward && !canDeleteEntry(inward))) {
      $('#excel-save-title').textContent = 'Update this inward’s certificate type in Excel?';
      $('#excel-save-description').textContent = `Excel has ${templateNames[from] || from}; the corrected PDF will use ${templateNames[to] || to}. Choose Yes to update this inward to match.`;
      $('#excel-save-note').textContent = 'An inward with a generated certificate cannot be deleted. Its COD stays the same. Cancel or No keeps editing.';
      $('#excel-save-no').textContent = 'No, keep editing';
    }
  }
  $('#excel-save-modal').showModal();
  $('#excel-save-no').focus();
}

async function persistDestructionChoice(link, needed) {
  const editor = app.editor;
  const response = await api('/api/destruction/request', {plant:link.plant, sn:String(link.case_id),
    ...(link.record_id ? {record_id:link.record_id} : {}), expected_revision:editor.workspaceRevision, needed});
  if (response.state) setState(response.state);
  editor.workspaceRevision = response.workspace_revision ?? app.workspaceRevision;
  editor.baseline.data_destruction_needed = needed;
  return response;
}

async function saveCertificateChoices(inward, requestChanged, needed) {
  const editor = app.editor;
  let saveError = null;
  setDetailsBusy(true);
  $('#details-error').hidden = true;
  try {
    if (!inward || !canSetPrimaryCertificateStatus(inward)) throw new Error('Refresh and reopen this inward before changing its certificate status.');
    if (editor.pendingPrimaryNa) {
      if (stateOf(inward) !== 'na') await persistPrimaryCertificateChoice(inward, false);
      editor.pendingPrimaryNa = false;
      if (editor.template !== 'destruction') {
        editor.noCertificate = true;
        editor.certificateStatus = 'N/A';
      }
    }
    if (requestChanged) {
      await persistDestructionChoice({plant:inward.plant, case_id:inward.sn}, needed);
      if (editor.template === 'destruction') {
        editor.noCertificate = !needed;
        editor.certificateStatus = needed ? 'Not created' : 'N/A';
      }
    }
    toast(needed ? 'Certificate choices saved.' : 'Both certificates marked N/A. The inward was kept.');
  } catch (error) {
    saveError = error;
  } finally {
    if (editor.noCertificate) invalidatePreview();
    setEditorMode(editor.editingDetails);
    renderChangeSummary();
    if (editorDirty()) saveWorkingDraft();
    else if (clearWorkingDraft()) editor.workingDraftSaved = false;
    setDetailsBusy(false);
    if (saveError) showDetailsError(saveError.message);
    else schedulePreview();
  }
}

async function saveInwardDetails() {
  const linkedRequestEditor = app.editor?.template === 'destruction' && Boolean(app.editor.destructionSource);
  if (app.detailsBusy || !app.templatesLoaded || (app.editor?.sourceSn == null && !app.editor?.localInwardId && !linkedRequestEditor) || app.editor.identityIssue) return;
  finishInlineEdit();
  const form = $('#details-form');
  const payload = detailsBody(form);
  const request = editorDestructionRequest();
  const needed = Boolean(payload.data_destruction_needed);
  const requestChanged = needed !== Boolean(request && request.destruction_requested !== false);
  if (app.editor.pendingPrimaryNa || (bothCertificatesUnchecked() && requestChanged)) {
    return saveCertificateChoices(editorInwardRow(), requestChanged, needed);
  }
  const changes = fieldDifferences(app.editor.excelBaseline || app.editor.baseline, payload);
  if (!app.editor.localInwardId && inwardChanges(payload).length && fieldDifferences(app.editor.baseline, payload).every(change => Object.hasOwn(followupFields, change.key))) {
    showExcelSaveChoice(payload, 'inward_only'); return;
  }
  const requestOnly = requestChanged && (linkedRequestEditor || app.editor.issuedCertificate || !changes.length);
  if (requestOnly) {
    let openRequest = null;
    setDetailsBusy(true);
    try {
      const link = linkedRequestEditor ? app.editor.destructionSource
        : {plant:app.editor.sourcePlant, case_id:app.editor.sourceSn, ...(app.editor.localInwardId ? {record_id:app.editor.localInwardId} : {})};
      const response = await persistDestructionChoice(link, needed);
      updateDestructionChoice(detailsBody(form));
      renderChangeSummary();
      if (editorDirty()) saveWorkingDraft();
      else if (clearWorkingDraft()) {
        app.editor.workingDraftSaved = false;
        const saved = needed && !linkedRequestEditor ? activeDestructionRequest(link) : null;
        if (saved && saved.record_id === response.record_id) openRequest = saved;
      }
      toast(needed ? 'Data Destruction request saved.' : 'Data Destruction request removed.');
    } catch (error) { showDetailsError(error.message); }
    finally { setDetailsBusy(false); }
    if (openRequest) {
      $('#details-modal').close();
      await openDetails(openRequest, true, true);
    }
    return;
  }
  if (linkedRequestEditor || app.editor.issuedCertificate || app.editor.localInwardId) return;
  if (!fieldDifferences(app.editor.excelBaseline || app.editor.baseline, payload).length) {
    toast('No changes to save.'); return;
  }
  // An inward may be saved before the Form 6/reference or first COD is known.
  const optional = ['form6','reference','starting_cod'].map(name => form.elements.namedItem(name));
  const required = optional.map(input => input.required);
  let valid;
  try { optional.forEach(input => { input.required = false; }); valid = validateEditor(); }
  finally { optional.forEach((input,index) => { input.required = required[index]; }); }
  if (valid) showExcelSaveChoice(detailsBody(form), 'save_only');
}

function openCompanionCertificate() {
  const inward = editorInwardRow();
  if (!inward || app.editor.identityIssue) return;
  switchLinkedCertificate(app.editor.template === 'destruction' ? templateId(inward) : 'destruction');
}

function workbookRowPayload(row) {
  const details = {};
  for (const key of ['supplier', 'location', 'form6', 'reference', 'certificate_type', ...Object.keys(followupFields)]) details[key] = String(row[key] ?? (key === 'certificate_type' ? 'standard' : ''));
  details.case_id = String(row.sn ?? row.case_id ?? '');
  details.inward_date = normalizeDate(row.inward_date);
  details.cod_no = String(row.cod_no ?? row.certificate_no ?? '');
  details.cod_date = normalizeDate(row.cod_date ?? row.certificate_date);
  details.items = (row.items || []).map(item => Object.fromEntries(['item', 'quantity', 'per_piece_weight', 'weight'].map(key => [key, String(item[key] ?? '')])));
  Object.assign(details, details.items[0] || {item:'', quantity:'', per_piece_weight:'', weight:''});
  details.certificate_fields = {...(row.certificate_fields || {})};
  return {plant:row.plant, source_kind:'workbook', details};
}

async function saveExcelChoice(updateExcel) {
  const next = app.excelSavePayload;
  if (!next || app.generationBusy) return;
  if (JSON.stringify(next.payload) !== JSON.stringify(detailsBody($('#details-form')))) {
    app.excelSavePayload = null; $('#excel-save-modal').close();
    toast('The details changed. Save again to review the current corrections.'); return;
  }
  app.excelSavePayload = null;
  $('#excel-save-modal').close();
  const addedDestruction = next.payload.data_destruction_needed && !app.editor.baseline?.data_destruction_needed;
  if (next.mode === 'create' && !updateExcel && workbookTypeChanged(next.payload)) {
    const inward = editorInwardRow();
    if (app.editor.issuedCertificate || (inward && !canDeleteEntry(inward))) return;
    const source = (app.editor.excelBaseline || app.editor.baseline)?.details || {};
    const approved = await confirmAction({title:'Delete the original Excel inward and keep the PDF?',
      message:`${labelPlant(next.payload.plant)} · SN ${app.editor.sourceSn} · ${source.supplier || ''}. This removes the entire inward and its materials from Excel. The PDF, certificate number, history and linked HDD records will be kept in the app. A workbook backup is saved first.`,
      confirmLabel:'Delete inward and save PDF', cancelLabel:'Keep editing'});
    if (!approved) return;
    if (JSON.stringify(next.payload) !== JSON.stringify(detailsBody($('#details-form')))) {
      showDetailsError('The details changed. Save again to review the current inward before deleting it.'); return;
    }
    await submitGeneration({...next.payload, confirmed_excel_removal:true, edit_scope:'remove_excel'}, false, '/api/create', true);
    return;
  }
  if (next.mode === 'inward_only') {
    if (!updateExcel) return;
    setDetailsBusy(true);
    try {
      const values = Object.fromEntries(Object.keys(followupFields).map(key => [key, next.payload.details[key] || '']));
      const response = await api('/api/clients/followup', {...next.payload, sn:app.editor.sourceSn, details:values, confirmed_excel:true});
      if (response.state) setState(response.state);
      app.editor.workspaceRevision = response.workspace_revision ?? app.workspaceRevision;
      for (const key of Object.keys(followupFields)) {
        app.editor.baseline.details[key] = values[key];
        if (app.editor.excelBaseline) app.editor.excelBaseline.details[key] = values[key];
      }
      if (response.state?.setup_warning) { showDetailsError(response.state.setup_warning); renderChangeSummary(); }
      else {
        const request = addedDestruction
          ? activeDestructionRequest({plant:next.payload.plant, case_id:app.editor.sourceSn}) : null;
        clearWorkingDraft(); $('#details-modal').close();
        toast('Inward details saved to Excel.');
        if (request) { setDetailsBusy(false); await openDetails(request, true, true); }
      }
    } catch (error) { showDetailsError(error.message); }
    finally { setDetailsBusy(false); }
    return;
  }
  if (next.mode === 'save_only') {
    if (!updateExcel) return;
    setDetailsBusy(true);
    try {
      const response = await api('/api/clients/update', {...next.payload, confirmed_excel:true});
      if (response.state) setState(response.state);
      const request = addedDestruction && !response.state?.setup_warning
        ? activeDestructionRequest({plant:next.payload.plant, case_id:response.sn ?? app.editor.sourceSn}) : null;
      clearWorkingDraft();
      app.editor.baseline = detailsBody($('#details-form'));
      $('#details-modal').close();
      const moved = response.sn && response.sn !== app.editor.sourceSn;
      toast(response.state?.setup_warning || (moved ? `Changes saved to Excel. This inward is now SN ${response.sn}.`
        : 'Changes saved to Excel. No PDF was created.'), Boolean(response.state?.setup_warning));
      if (request) { setDetailsBusy(false); await openDetails(request, true, true); }
    } catch (error) { showDetailsError(error.message); }
    finally { setDetailsBusy(false); }
    return;
  }
  await submitGeneration({...next.payload, confirmed_excel:updateExcel, edit_scope:updateExcel ? 'update_excel' : 'certificate_only'}, false, '/api/create', true);
}

async function saveDraft() {
  const form = $('#details-form');
  if (app.detailsBusy || !app.templatesLoaded || form.elements.namedItem('source_kind').value !== 'manual') return;
  // A draft does not issue or reserve a certificate number.
  const start = form.elements.namedItem('starting_cod');
  const wasDisabled = start.disabled;
  let valid;
  try { start.disabled = true; valid = form.reportValidity(); }
  finally { start.disabled = wasDisabled; }
  if (!valid) return;
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

function clearAttachmentPreview() {
  app.previewAttachment = null;
  $('#preview-attachment').hidden = true;
  $('#preview-attachment-image').replaceChildren();
  $('#preview-attachment-status').textContent = '';
  $('#preview-attachment-retry').hidden = true;
}

function showAttachmentPage(index) {
  const state = app.previewAttachment;
  if (!state || state.revision !== app.previewRevision || !Number.isInteger(index) || index < 0 || index >= state.pages.length) return;
  state.page = index;
  $('#preview-attachment-count').textContent = `Page ${index + 1} of ${state.pages.length}`;
  $('#preview-attachment-previous').disabled = index === 0;
  $('#preview-attachment-next').disabled = index === state.pages.length - 1;
  const status = $('#preview-attachment-status'); status.hidden = false; status.textContent = 'Loading page…';
  $('#preview-attachment-retry').hidden = true;
  const image = element('img'); image.hidden = true;
  image.alt = `${state.name}, page ${index + 1} of ${state.pages.length}`;
  state.image = image;
  const isCurrent = () => app.previewAttachment === state && state.image === image && state.revision === app.previewRevision;
  image.addEventListener('load', () => { if (isCurrent()) { image.hidden = false; status.hidden = true; } });
  image.addEventListener('error', () => {
    if (!isCurrent()) return;
    image.hidden = true; status.hidden = false;
    status.textContent = 'This page could not load. Try again or open the full preview.';
    $('#preview-attachment-retry').hidden = false;
  });
  $('#preview-attachment-image').replaceChildren(image);
  image.src = state.pages[index];
}

function renderAttachmentPreview(attachment, previewPath) {
  clearAttachmentPreview();
  if (!attachment) return;
  $('#preview-attachment').hidden = false;
  $('#preview-attachment-name').textContent = attachment.name || 'Attached PDF';
  const id = previewPath.match(/^\/previews\/([a-f0-9]{32})\.pdf$/)?.[1];
  try {
    if (!id || !Array.isArray(attachment.pages) || !attachment.pages.length || attachment.pages.length > 2000) throw new Error();
    const pages = attachment.pages.map(source => {
      const candidate = new URL(source, window.location.href);
      if (candidate.origin !== window.location.origin || !new RegExp(`^/previews/${id}/page/[1-9][0-9]{0,3}\\.png$`).test(candidate.pathname) || candidate.search || candidate.hash) throw new Error();
      return candidate.href;
    });
    app.previewAttachment = {pages, name:attachment.name || 'Attached PDF', page:0, revision:app.previewRevision};
    showAttachmentPage(0);
  } catch {
    $('#preview-attachment-status').hidden = false;
    $('#preview-attachment-status').textContent = 'Open the full preview to check the attached PDF.';
    $('#preview-attachment-count').textContent = '';
    $('#preview-attachment-previous').disabled = true;
    $('#preview-attachment-next').disabled = true;
  }
}

async function previewDetails({automatic = false} = {}) {
  if (app.editor?.uploadRow) return;
  if (app.editor?.noCertificate) return;
  const form = $('#details-form');
  if (app.editor?.needsStartingCod && !form.elements.namedItem('starting_cod').value.trim()) {
    $('#preview-status').textContent = 'Enter the starting COD number to preview.';
    return;
  }
  if (app.editor?.identityIssue) { updateDetailsButtons(); return; }
  if (app.detailsBusy || app.previewLoading || !app.templatesLoaded) return;
  if (automatic ? !form.checkValidity() : !validateEditor()) return;
  const payload = detailsBody(form);
  invalidatePreview();
  const revision = app.previewRevision;
  app.previewLoading = true;
  $('#preview-details').disabled = true;
  $('#preview-status').textContent = 'Preparing your PDF preview…';
  try {
    app.editor.previewSession ||= `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
    app.reviewClientId ||= `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
    const response = await api('/api/preview', {...payload, preview_session:app.editor.previewSession, preview_client:app.reviewClientId}, {background:automatic});
    if (revision !== app.previewRevision || !$('#details-modal').open) return;
    if (!response.preview_url) throw new Error('The PDF preview is not available. Please try again.');
    const url = new URL(response.preview_url, window.location.href);
    if (url.origin !== window.location.origin || !['http:', 'https:'].includes(url.protocol)) throw new Error('The PDF preview address is not available in this workspace.');
    app.editor.previewFacts = {certificate_no: response.certificate_no, certificate_date: response.certificate_date, printed_details: response.printed_details};
    if (response.numbering) app.editor.numbering = response.numbering;
    updateSingleCodControls();
    renderEditorSummary();
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
    const detailPages = $('#preview-detail-images'); detailPages.replaceChildren();
    for (const [index, source] of (response.preview_detail_images || []).entries()) {
      try {
        const candidate = new URL(source, window.location.href);
        if (candidate.origin !== window.location.origin || !/^\/previews\/[a-f0-9]{32}(?:\.png|\/page\/[1-9][0-9]{0,3}\.png)$/.test(candidate.pathname) || candidate.search || candidate.hash) continue;
        const page = element('img'); page.src = candidate.href;
        page.alt = `Device receipt details, page ${index + 2}`; detailPages.append(page);
      } catch { /* The full PDF remains available. */ }
    }
    renderAttachmentPreview(response.preview_attachment, url.pathname);
    $('#preview-status').textContent = imageUrl ? 'Preview ready' : 'PDF ready. Open the full preview.';
    $('#preview-number').textContent = `${app.editor.template === 'destruction' ? 'HDD' : 'COD'} ${response.certificate_no || 'Not available'}${response.certificate_date ? ` · ${dateLabel(response.certificate_date)}` : ''}`;
    $('#preview-number').hidden = false;
    const warnings = Array.isArray(response.warnings) ? response.warnings.filter(value => !isRoutineNotice(value)) : [];
    renderWarningList($('#preview-warnings'), warnings, true);
    $('#preview-warnings').hidden = warnings.length === 0;
    $('#editor-checks').hidden = warnings.length === 0;
    $('#preview-warnings-title').hidden = warnings.length === 0;
    $('#preview-warnings-title').textContent = 'Check before creating';
  } catch (error) {
    if (revision === app.previewRevision && $('#details-modal').open) {
      const showRepair = /Serial Number|Excel|workbook/i.test(error.message);
      const inline = !automatic || showRepair;
      $('#preview-status').textContent = inline ? 'Preview unavailable' : noticeText(error.message);
      if (inline) showDetailsError(error.message, {previewNotice:automatic});
    }
  } finally {
    app.previewLoading = false;
    updateDetailsButtons();
    if (revision !== app.previewRevision && $('#details-modal').open) schedulePreview();
  }
}

function toggleFullPdfPreview() {
  const frame = $('#preview-frame');
  if ($('#preview-document').hidden || !$('#preview-link').href) return;
  frame.hidden = !frame.hidden;
  if (!frame.hidden) frame.src = $('#preview-link').href;
  else frame.removeAttribute('src');
  $('#preview-full-pdf').textContent = frame.hidden ? 'Enlarge preview' : 'Close enlarged preview';
  $('#preview-full-pdf').setAttribute('aria-expanded', String(!frame.hidden));
}

async function confirmWarnings() {
  const next = app.pendingConfirmation;
  if (!next) return;
  $('#warning-modal').close();
  app.pendingConfirmation = null;
  if (!next) return;
  if (next.fromDetails && next.revision !== app.previewRevision) {
    toast('Details have changed. Create the PDF again to review the current details.');
    return;
  }
  await submitGeneration({ ...next.payload, review_token: next.reviewToken,
    override_reason: next.hasWarnings ? 'Warnings reviewed; user chose Continue anyway.' : 'No warnings were reported; confirmed creation.' }, true, next.endpoint, next.fromDetails);
}

async function openInwardRemoval(row) {
  if (app.removalBusy || app.removalPending || app.generationBusy || app.backupBusy || app.detailsBusy || app.clientBusy) return;
  app.removalBusy = true;
  setBusy(true, 'Checking this inward…');
  try {
    const payload = {plant:row.plant, sn:String(row.sn), sheet:row.sheet, row:row.row,
      request_id:Array.from({length:32}, () => Math.floor(Math.random() * 16).toString(16)).join(''),
      expected_revision:app.workspaceRevision ?? app.state.workspace_revision ?? ''};
    const preview = await api('/api/clients/remove-preview', payload);
    app.removalPending = {...payload, review_token:preview.review_token, expected_revision:preview.expected_revision};
    const facts = $('#remove-inward-facts'); facts.replaceChildren();
    for (const [label, value] of [['Company', preview.entry.supplier || 'Company name missing'], ['SN', preview.entry.sn],
      ['Plant', labelPlant(preview.entry.plant)], ['Received', dateLabel(preview.entry.inward_date)], ['Materials', preview.entry.materials]]) {
      const fact = element('div'); fact.append(element('dt', '', label), element('dd', '', String(value))); facts.append(fact);
    }
    $('#remove-inward-checked').checked = false;
    $('#confirm-remove-inward').disabled = true;
    $('#remove-inward-error').hidden = true;
    $('#remove-inward-modal').showModal();
    $('#keep-inward').focus();
  } catch (error) { app.removalPending = null; toast(error.message, true); }
  finally { app.removalBusy = false; setBusy(false); }
}

function entryDeletionIdentity(row) {
  return isManual(row) ? {source_kind:'manual',plant:row.plant,record_id:row.record_id}
    : {source_kind:'workbook',plant:row.plant,sn:String(row.sn),sheet:row.sheet,row:row.row};
}

function openEntryDeletion({showDialog = true} = {}) {
  if (app.entryDeletionBusy || app.loadingWorkspace || app.generationBusy || app.backupBusy || app.detailsBusy || app.clientBusy || app.settingsBusy) return;
  const rows = app.state.rows.filter(row => (isManual(row) || ['plant1','plant2'].includes(row.plant)) && canDeleteEntry(row)
    && (isManual(row) ? Boolean(row.record_id) : Boolean(row.sn && row.sheet && row.row)));
  app.entryDeletion = {rows,checked:new Set(),labels:[],revision:app.workspaceRevision ?? app.state.workspace_revision ?? ''};
  $('#entry-delete-search').value = ''; $('#entry-delete-hdd').checked = false;
  $('#entry-delete-choices').replaceChildren(...rows.map((row,index) => {
    const label=element('label','bulk-choice'),input=element('input'); input.type='checkbox';
    input.setAttribute('aria-label',`${row.supplier || 'Unnamed company'}, ${labelPlant(row.plant)}, ${isManual(row) ? row.certificate_no || 'local record' : 'SN '+row.sn}`);
    input.addEventListener('change',()=>{if(input.checked)app.entryDeletion.checked.add(index);else app.entryDeletion.checked.delete(index);filterEntryDeletion();});
    const copy=element('span','bulk-client-copy');
    copy.append(element('strong','',row.supplier || 'Company name missing'),element('small','',`${labelPlant(row.plant)} · ${isManual(row) ? 'Local record' : 'SN '+row.sn} · COD ${row.certificate_no || 'not assigned'} · ${row.location || ''}`));
    const meta=element('span','bulk-client-meta'); meta.append(element('span','',stateLabel(row)));
    label.append(input,copy,meta);app.entryDeletion.labels.push(label);return label;
  }));
  backToEntryDeletionSelection();filterEntryDeletion();
  if (showDialog) { $('#entry-deletion-modal').showModal();$('#entry-delete-search').focus(); }
}

function filterEntryDeletion() {
  const state=app.entryDeletion;if(!state)return;
  const words=$('#entry-delete-search').value.trim().toLowerCase().split(/\s+/).filter(Boolean);
  let shown=0,hiddenSelected=0;
  state.rows.forEach((row,index)=>{
    const searchable=[row.supplier,row.sn,row.case_id,row.certificate_no,row.reference,row.form6,row.location,
      labelPlant(row.plant),stateLabel(row),row.removed_inward?.case_id].join(' ').toLowerCase();
    const visible=words.every(word=>searchable.includes(word));
    state.labels[index].hidden=!visible;
    state.labels[index].classList.toggle('bulk-choice-selected',state.checked.has(index));
    if(visible)shown++;else if(state.checked.has(index))hiddenSelected++;
  });
  $('#entry-delete-shown').textContent=`${shown} of ${state.rows.length} shown`;
  $('#entry-delete-empty').hidden=shown>0;$('#entry-delete-clear').hidden=words.length===0;
  $('#entry-delete-selected').textContent=`${state.checked.size} selected${hiddenSelected ? ` · ${hiddenSelected} hidden by search` : ''}${state.checked.size>50 ? ' · choose up to 50 at a time' : ''}`;
  $('#entry-delete-preview').disabled=app.entryDeletionBusy || !state.checked.size || state.checked.size>50;
}

function backToEntryDeletionSelection() {
  if(app.entryDeletionBusy)return;
  if(app.entryDeletion)app.entryDeletion.pending=null;
  $('#entry-delete-selection').hidden=false;$('#entry-delete-review').hidden=true;
  $('#entry-delete-preview').hidden=false;$('#entry-delete-confirm').hidden=true;
  $('#entry-delete-back').hidden=true;$('#entry-delete-checked').checked=false;
  $('#entry-delete-confirm').disabled=true;$('#entry-delete-error').hidden=true;
}

async function reviewEntryDeletion() {
  const state=app.entryDeletion;
  if(!state || app.entryDeletionBusy || !state.checked.size || state.checked.size>50)return;
  const payload={selected:[...state.checked].map(index=>entryDeletionIdentity(state.rows[index])),
    include_linked_hdd:Boolean($('#entry-delete-hdd').checked), expected_revision:state.revision,
    request_id:Array.from({length:32},()=>Math.floor(Math.random()*16).toString(16)).join('')};
  app.entryDeletionBusy=true;$('#entry-delete-preview').disabled=true;$('#entry-delete-error').hidden=true;
  try {
    const response=await api('/api/entries/delete-preview',payload);
    if(app.entryDeletion!==state)return;
    state.pending={...payload,review_token:response.review_token};
    const list=$('#entry-delete-review-list');list.replaceChildren();
    for(const entry of response.review.entries){
      const section=element('section');
      section.append(element('strong','',entry.supplier || 'Unnamed company'),element('p','',`${labelPlant(entry.plant)} · ${entry.source_kind==='workbook' ? 'SN '+entry.sn+' · '+entry.materials+(entry.materials===1?' material row':' material rows') : 'Local certificate record'} · COD ${entry.cod || 'not assigned'} · ${templateNames[entry.certificate_type] || entry.certificate_type}`));
      if(entry.status || entry.reference || entry.form6) section.append(element('p','',[
        entry.status ? stateLabel({certificate_status:entry.status}) : '',entry.received ? 'Received '+dateLabel(entry.received) : '',entry.reference ? 'Reference '+entry.reference : '',entry.form6 ? 'Form 6 '+entry.form6 : ''
      ].filter(Boolean).join(' · ')));
      list.append(section);
    }
    const pdfList=element('ul');for(const path of response.review.pdfs)pdfList.append(element('li','',path));
    list.append(element('h3','',`${response.review.pdfs.length} PDF${response.review.pdfs.length===1?'':'s'} to remove`),pdfList);
    const kept=response.review.kept_hdd || [];
    $('#entry-delete-kept-hdd').textContent=kept.length ? `${kept.length} linked HDD ${kept.length===1?'record will':'records will'} be kept as standalone: ${kept.map(row=>row.supplier+(row.cod?' · COD '+row.cod:'')).join('; ')}` : 'No linked HDD records will be left behind.';
    $('#entry-delete-selection').hidden=true;$('#entry-delete-review').hidden=false;
    $('#entry-delete-preview').hidden=true;$('#entry-delete-confirm').hidden=false;
    $('#entry-delete-back').hidden=false;$('#entry-delete-checked').checked=false;
    $('#entry-delete-confirm').disabled=true;
    if (!$('#entry-deletion-modal').open) $('#entry-deletion-modal').showModal();
    $('#entry-delete-cancel').focus();
  }catch(error){
    if (error.workspaceRecovered) {
      $('#entry-deletion-modal').close();app.entryDeletion=null;
      toast(error.message,true);return;
    }
    $('#entry-delete-error').textContent=error.message;$('#entry-delete-error').hidden=false;
    if (!$('#entry-deletion-modal').open) $('#entry-deletion-modal').showModal();
  }
  finally{app.entryDeletionBusy=false;filterEntryDeletion();}
}

async function confirmEntryDeletion() {
  const state=app.entryDeletion;
  if(!state?.pending || app.entryDeletionBusy || !$('#entry-delete-checked').checked)return;
  app.entryDeletionBusy=true;$('#entry-delete-confirm').disabled=true;
  try{
    const response=await api('/api/entries/delete',{...state.pending,checked:true});
    if(response.state)setState(response.state);
    app.entryDeletionBusy=false;$('#entry-deletion-modal').close();app.entryDeletion=null;
    if (response.deleted) rememberDeletion({kind:'entries', request_id:state.pending.request_id});
    toast(response.state?.setup_warning || 'Selected entries deleted.', false, response.deleted === true);
  }catch(error){
    if (error.workspaceRecovered) {
      $('#entry-deletion-modal').close();app.entryDeletion=null;
      toast(error.message,true);return;
    }
    $('#entry-delete-error').textContent=error.message;$('#entry-delete-error').hidden=false;
    $('#entry-delete-checked').checked=false;
  }finally{app.entryDeletionBusy=false;}
}

async function confirmInwardRemoval() {
  if (!app.removalPending || app.removalBusy || !$('#remove-inward-checked').checked || $('#confirm-remove-inward').disabled) return;
  app.removalBusy = true;
  $('#confirm-remove-inward').disabled = true;
  setBusy(true, 'Removing this inward…');
  try {
    const pending = app.removalPending;
    const result = await api('/api/clients/remove', {...pending, checked:true});
    app.removalPending = null;
    $('#remove-inward-modal').close();
    app.expanded = null;
    setState(result.state);
    if (result.removed) rememberDeletion({kind:'inward', request_id:pending.request_id});
    toast('Inward removed from Excel.', false, result.removed === true);
  } catch (error) {
    if (error.workspaceRecovered) {
      app.removalPending = null;
      $('#remove-inward-modal').close();
      toast(error.message, true);
      return;
    }
    $('#remove-inward-error').textContent = error.message;
    $('#remove-inward-error').hidden = false;
    $('#remove-inward-checked').checked = false;
    $('#keep-inward').focus();
  } finally { app.removalBusy = false; setBusy(false); }
}

function rememberDeletion(target) {
  const history = app.deletionHistory ||= [];
  if (target) {
    if (!history.some(entry => entry.request_id === target.request_id)) history.push(target);
    if (history.length > 50) history.shift();
  } else {
    history.pop();
  }
  updateDeletionUndo();
}

function updateDeletionUndo() {
  app.deletionUndo = app.deletionHistory.at(-1);
  $('#undo-last-deletion').hidden = !app.deletionUndo;
  try {
    if (app.deletionUndo) localStorage.setItem(deletionUndoKey, JSON.stringify(app.deletionHistory));
    else localStorage.removeItem(deletionUndoKey);
  } catch { /* The action still works for this open dashboard. */ }
}

async function undoLastDeletion() {
  if (!app.deletionUndo || app.removalBusy || app.entryDeletionBusy || app.generationBusy
      || app.backupBusy || app.clientBusy || app.completionBusy || app.settingsBusy || app.detailsBusy) return;
  app.removalBusy = true;
  setBusy(true, 'Restoring deleted entries…');
  try {
    const response = await api('/api/entries/undo-delete', app.deletionUndo);
    if (!response.undone) throw new Error('The deletion could not be restored. Refresh the company list.');
    if (response.state) setState(response.state);
    rememberDeletion(null);
    toast(response.state?.setup_warning || 'Deletion undone.');
  } catch (error) {
    toast(error.message, true);
  } finally {
    app.removalBusy = false;
    setBusy(false);
  }
}

function renderEditorFollowup() {
  const available = app.editor?.sourceSn != null && app.editor.template !== 'destruction';
  $('#editor-inward-details').hidden = !available;
}

function inwardChanges(payload) {
  return fieldDifferences(app.editor?.excelBaseline || app.editor?.baseline, payload)
    .filter(change => Object.hasOwn(followupFields, change.key));
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
  if (action === 'mark-na' || action === 'undo-na') {
    if (app.inwardStatusBusy) return;
    const mark = action === 'mark-na';
    app.inwardStatusBusy = true;
    const manual = isManual(row);
    setBusy(true, mark ? 'Marking N/A…' : 'Restoring certificate review…',
      manual ? 'Updating the local certificate record.' : 'Updating the status in Excel.');
    try {
      const result = await api('/api/clients/status', {plant:row.plant, sn:String(row.sn),
        ...(manual ? {record_id:row.record_id} : {}), action:mark ? 'mark_na' : 'undo_na'});
      setState(result.state);
      toast(mark ? 'Marked N/A.' : `N/A undone. Status is ${result.status}.`);
    } catch (error) { toast(error.message, true); }
    finally { app.inwardStatusBusy = false; setBusy(false); }
    return;
  }
  if (action === 'fill-details') { openDetails(row, false, true); return; }
  if (action === 'delete-entry') {
    if (!canDeleteEntry(row)) {toast('Generated certificates cannot be deleted.',true);return;}
    openEntryDeletion();app.entryDeletion.checked.add(app.entryDeletion.rows.indexOf(row));await reviewEntryDeletion();return;
  }
  if (action === 'remove-inward') {
    if (!canDeleteEntry(row)) { toast('Generated certificates cannot be deleted.', true); return; }
    const pendingHdd = app.state.rows.some(item => isManual(item) && (item.destruction_source || item.raw_details?.destruction_source)?.plant === row.plant
      && String((item.destruction_source || item.raw_details?.destruction_source)?.case_id) === String(row.sn));
    if (pendingHdd) {
      openEntryDeletion({showDialog:false}); const selected = app.entryDeletion.rows.indexOf(row);
      app.entryDeletion.checked.add(selected); $('#entry-delete-hdd').checked = true;
      await reviewEntryDeletion();
    } else await openInwardRemoval(row);
    return;
  }
  if (action === 'details' || action === 'edit-details') {
    await openDetails(preferredCertificateRow(row), false, action === 'edit-details');
    return;
  }
  if (action === 'generate') {
    await submitGeneration(isManual(row) ? { plant: row.plant, record_id: row.record_id } : { plant: row.plant, sn: row.sn }, false,
      isManual(row) ? '/api/manual/generate' : '/api/generate');
    return;
  }
  if (action === 'complete' || action === 'undo-completed') {
    if (app.completionBusy || app.generationBusy || app.backupBusy) return;
    const undo = action === 'undo-completed';
    const expectedRevision = app.workspaceRevision ?? app.state.workspace_revision ?? '';
    const approved = await confirmAction(undo ? { title: 'Undo Completed?',
      message: `This returns the ${templateLabel(row)} certificate for ${row.supplier || `Client ${row.sn}`}${row.certificate_no ? ` (COD ${row.certificate_no})` : ''} to the send queue.`,
      confirmLabel: 'Yes, undo', cancelLabel: 'No, keep Completed' } : { title: 'Mark certificate Completed?',
      message: `Confirm that you have sent the certificate for ${row.supplier || `Client ${row.sn}`}. This marks it Completed in ${isManual(row) ? 'the local certificate register' : 'Excel'} and stops its reminder.`,
      confirmLabel: 'Mark Completed', cancelLabel: 'Go back' });
    if (!approved || app.completionBusy || app.generationBusy || app.backupBusy) return;
    app.completionBusy = true;
    setBusy(true, isManual(row) ? 'Updating your local register…' : 'Updating the workbook…');
    try { const response = await api(isManual(row) ? '/api/manual/complete' : '/api/complete', { expected_revision: expectedRevision,
      ...(undo ? {undo_completed:true} : {}),
      ...(isManual(row) ? { record_id: row.record_id } : { plant: row.plant, sn: row.sn }) }); setState(response.state); toast(undo ? 'Completion undone.' : 'Marked Completed.'); }
    catch (error) { toast(error.message, true); }
    finally { app.completionBusy = false; setBusy(false); }
    return;
  }
  if (action === 'open-workbook' && !isManual(row)) {
    await openClientInExcel(row.plant, row.sn, row.sheet, row.row);
  }
}

function openSettings() {
  if (app.generationBusy || app.backupBusy || app.completionBusy || app.clientBusy
      || app.settingsBusy || app.detailsBusy || app.stopping) return;
  updateSettingsDisplay();
  app.settingsBaseline = JSON.stringify(settingsValues());
  for (const disclosure of $('#settings-form').querySelectorAll('details')) disclosure.open = false;
  $('#settings-status').textContent = '';
  $('#settings-modal').showModal();
}

function settingsValues() {
  const form = $('#settings-form');
  const cod = form.elements.namedItem('next_cod');
  return {
    workbook_paths: [form.elements.namedItem('plant1_path').value.trim(), form.elements.namedItem('plant2_path').value.trim()],
    reminders_enabled: Boolean(form.elements.namedItem('reminders_enabled').checked),
    reminder_time: form.elements.namedItem('reminder_time').value,
    ...(cod ? {next_cod:cod.value.trim()} : {}),
  };
}

async function requestSettingsClose() {
  if (app.settingsBusy || app.backupBusy) return;
  if (app.settingsBaseline && app.settingsBaseline !== JSON.stringify(settingsValues())) {
    const approved = await confirmAction({title:'Discard these changes?', message:'These Settings changes have not been saved.', confirmLabel:'Yes, discard changes', cancelLabel:'No, keep editing'});
    if (!approved) return;
  }
  updateSettingsDisplay();
  $('#settings-modal').close();
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
    if (response.settings) {
      app.settings = { ...app.settings, ...response.settings };
      // Choosing a backup folder changes settings.local.json and its revision.
      // Read the rows before allowing another action against that new snapshot.
      await refreshState();
    }
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
    ? dateTimeLabel(timestamp)
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
      const current = await api('/api/check', {}); app.settings = current.settings || app.settings; setState(current.state);
    }
    $('#backup-restore-modal').close();
    updateSettingsDisplay();
    app.settingsBaseline = JSON.stringify(settingsValues());
    await loadBackupStatus();
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
  return ({complete:'Generated',deleted:'Deleted mistaken entry',rolled_back:'Recovered after an error',prepared:'Creation started',published:'PDF saved; tracking pending',source_committed:'Tracking saved; completion pending'})[state] || String(state || 'Unknown state').replace(/_/g, ' ');
}

function activityTime(record) {
  const value = record.completed_at || record.started_at;
  const date = new Date(value);
  return dateTimeLabel(value);
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
      const voidedBy = records.find(candidate => candidate.voided_certificate
        && candidate.source_kind === record.source_kind && candidate.plant === record.plant
        && String(candidate.key) === String(record.key)
        && String(candidate.voided_certificate.cod) === String(record.cod));
      const label = record.voided_certificate
        ? `Replacement issued · previous COD ${record.voided_certificate.cod}`
        : voidedBy ? `Replaced · new COD ${voidedBy.cod}` : activityState(record.state);
      const item = element('li', 'history-entry'), copy = element('div');
      copy.append(element('strong', '', `Certificate ${record.cod || 'number not recorded'} · ${record.after?.supplier || record.before?.supplier || record.key || 'Company'}`), element('span', '', `${label} · ${activityTime(record)}`));
      const button = element('button', 'button button-quiet', 'View changes'); button.type = 'button';
      button.addEventListener('click', () => showHistoryDetail({...record, voided_by:voidedBy || null})); item.append(copy, button); list.append(item);
    }
    $('#history-status').textContent = records.length ? `${records.length} recent record${records.length === 1 ? '' : 's'}.` : 'No certificate activity recorded yet.';
  } catch (error) { $('#history-status').textContent = error.message; }
  finally { app.historyBusy = false; }
}

function showHistoryDetail(record) {
  $('#history-detail-title').textContent = `Certificate ${record.cod || 'activity'}`;
  const state = record.voided_certificate ? `Replacement issued; previous COD ${record.voided_certificate.cod}`
    : record.voided_by ? `Replaced; new COD ${record.voided_by.cod}` : activityState(record.state);
  $('#history-detail-description').textContent = `${state} · ${activityTime(record)}`;
  const detail = $('#history-detail'); detail.replaceChildren();
  const facts = element('dl', 'history-facts');
  for (const [label, value] of [['Company', record.after?.supplier || record.before?.supplier || 'Not recorded'], ['Plant', labelPlant(record.plant)], ['Source', record.source_kind === 'manual' ? 'Local certificate register' : 'Excel'], ['Operator', record.operator || 'Not recorded'], ['PDF file', String(record.pdf_path || '').split(/[\\/]/).pop() || 'Not recorded'], ['Accepted warning reason', record.override_reason || 'No reason recorded']]) {
    const fact = element('div'); fact.append(element('dt', '', label), element('dd', '', value)); facts.append(fact);
  }
  if (record.voided_certificate) {
    for (const [label, value] of [['Previous COD', record.voided_certificate.cod || 'Not recorded'],
      ['Replacement note', record.voided_certificate.reason || 'Replaced through Review & edit'],
      ['Archived PDF', record.voided_certificate.archive_path || record.voided_certificate.pdf_path || 'Not recorded']]) {
      const fact = element('div'); fact.append(element('dt', '', label), element('dd', '', value)); facts.append(fact);
    }
  } else if (record.voided_by) {
    const fact = element('div'); fact.append(element('dt', '', 'Replacement COD'), element('dd', '', record.voided_by.cod || 'Not recorded')); facts.append(fact);
    const reason = element('div'); reason.append(element('dt', '', 'Replacement note'), element('dd', '', record.voided_by.voided_certificate.reason || 'Replaced through Review & edit')); facts.append(reason);
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
    renderWarningList(warnings, record.accepted_warnings);
    detail.append(warnings);
  }
  if (record.previous_pdfs?.length) detail.append(element('p', '', `${record.previous_pdfs.length} previous PDF${record.previous_pdfs.length === 1 ? '' : 's'} retained in the recovery archive.`));
  $('#history-detail-modal').showModal();
}

async function saveSettings(event, repairReminder = false) {
  event.preventDefault();
  if (app.settingsBusy) return;
  const form = event.currentTarget;
  if (!form.checkValidity()) {
    $('#settings-paths').open = true;
    if (!form.elements.namedItem('reminder_time').validity?.valid) form.elements.namedItem('reminder_time').closest('.field').hidden = false;
    form.reportValidity(); return;
  }
  const values = settingsValues();
  const payload = {...values, update_reminder: repairReminder
    || values.reminders_enabled !== Boolean(app.settings.reminders_enabled)
    || values.reminder_time !== (app.settings.reminder_time || '11:00')};
  app.settingsBusy = true; form.inert = true;
  $('#settings-status').textContent = 'Saving…';
  setBusy(true, 'Saving settings…', 'Checking the workbook paths and updating your Windows reminder.');
  try {
    const response = await api('/api/settings', payload);
    app.settings = response.settings;
    updateSettingsDisplay();
    app.settingsBaseline = JSON.stringify(settingsValues());
    if (response.warning) {
      $('#settings-status').textContent = response.warning;
      toast(response.warning, true);
    } else {
      $('#settings-modal').close();
      toast(response.message || 'Settings saved.');
    }
    await refreshState();
  } catch (error) { $('#settings-status').textContent = error.message; toast(error.message, true); }
  finally { app.settingsBusy = false; form.inert = false; setBusy(false); }
}

async function stopApp() {
  if (app.detailsBusy || app.generationBusy || app.backupBusy || app.completionBusy || app.stopping) return;
  if (app.settingsBusy) return;
  if ($('#details-modal').open && editorDirty()) { requestDetailsClose(stopApp); return; }
  const approved = await confirmAction({ title: 'Close dashboard?',
    message: 'This closes the certificate dashboard on this computer. Your saved files and working drafts are kept. Open Prepare Certificates to use it again.',
    confirmLabel: 'Close dashboard now', cancelLabel: 'Keep open' });
  if (!approved) return;
  if (app.detailsBusy || app.generationBusy || app.backupBusy || app.completionBusy || app.stopping) return;
  app.stopping = true;
  try {
    await api('/api/stop', {});
    $('#offline-title').textContent = 'Dashboard closed';
    $('#offline-message').textContent = 'You can close this browser tab now. To use the dashboard again, double-click Prepare Certificates.';
    $('#reconnect').hidden = true;
    $('#offline-screen').hidden = false;
  }
  catch (error) { toast(error.message, true); }
  finally { app.stopping = false; }
}

async function reconnectDashboard() {
  if (app.reconnecting) return;
  app.reconnecting = true;
  for (const id of ['reconnect', 'dialog-reconnect']) $(`#${id}`).disabled = true;
  try {
    const response = await api('/api/check', {});
    const pendingSettings = $('#settings-modal').open ? settingsValues() : null;
    app.settings = response.settings || app.settings; setState(response.state, {preserveSettings:false});
    if (pendingSettings) {
      // A lost response may already have saved Settings. Keep those committed
      // values as the baseline, without erasing the operator's pending inputs.
      app.settingsBaseline = JSON.stringify(settingsValues());
      const form = $('#settings-form');
      form.elements.namedItem('plant1_path').value = pendingSettings.workbook_paths[0];
      form.elements.namedItem('plant2_path').value = pendingSettings.workbook_paths[1];
      form.elements.namedItem('reminders_enabled').checked = pendingSettings.reminders_enabled;
      form.elements.namedItem('reminder_time').value = pendingSettings.reminder_time;
      if (form.elements.namedItem('next_cod')) form.elements.namedItem('next_cod').value = pendingSettings.next_cod;
      updateSettingsFiles(); updateReminderControls();
      $('#settings-status').textContent = 'Reconnected. Saved settings were checked; your pending changes are still here.';
    }
    $('#offline-screen').hidden = true;
    if ($('#connection-modal').open) $('#connection-modal').close();
    toast('Connected. Your edits are kept. Check Saved certificates before repeating a creation.');
  } catch { /* Never automatically retry a write after an uncertain response. */ }
  finally {
    app.reconnecting = false;
    for (const id of ['reconnect', 'dialog-reconnect']) $(`#${id}`).disabled = false;
  }
}

function wire() {
  app.toastDialogs = [...document.querySelectorAll('dialog')].filter(dialog => dialog.open);
  if (typeof MutationObserver !== 'undefined') {
    app.toastDialogObserver = new MutationObserver(records => {
      for (const {target} of records) {
        app.toastDialogs = app.toastDialogs.filter(dialog => dialog !== target && dialog.open);
        if (target.open) app.toastDialogs.push(target);
      }
      syncToastLayer();
    });
    for (const dialog of document.querySelectorAll('dialog')) {
      app.toastDialogObserver.observe(dialog, {attributes:true, attributeFilter:['open']});
    }
  }
  $('#view-records').addEventListener('click', () => setWorkspaceView('records'));
  $('#view-analytics').addEventListener('click', () => setWorkspaceView('analytics'));
  $('#analytics-workspace').addEventListener('analytics:review', event => {
    const detail = event.detail || {};
    clearRecordFilterValues();
    app.analyticsAsOf = detail.asOf;
    setWorkspaceView('records');
    $('#plant-filter').value = $('#analytics-plant').value;
    $('#search').value = detail.search || '';
    $('#status-filter').value = detail.status || 'all';
    $('#template-filter').value = detail.type || 'all';
    restartRecordList();
  });
  $('#analytics-workspace').addEventListener('analytics:files', () => {
    Object.keys(savedFilterDefaults).forEach(clearSavedFilter);
    $('#saved-plant-filter').value = $('#analytics-plant').value;
    app.savedFilesPage = 1;
    setWorkspaceView('files');
  });
  $('#view-files').addEventListener('click', () => {
    setWorkspaceView('files');
    syncWorkbookChanges();
  });
  $('#reset-record-filters').addEventListener('click', resetRecordFilters);
  $('#dismiss-toast').addEventListener('click', dismissToast);
  $('#undo-deletion-toast').addEventListener('click', undoLastDeletion);
  $('#undo-last-deletion').addEventListener('click', undoLastDeletion);
  try {
    const saved = JSON.parse(localStorage.getItem(deletionUndoKey) || 'null');
    app.deletionHistory = (Array.isArray(saved) ? saved : [saved])
      .filter(entry => entry && ['inward', 'entries'].includes(entry.kind) && /^[a-f0-9]{32}$/.test(entry.request_id))
      .slice(-50);
    updateDeletionUndo();
  } catch { /* Ignore an unavailable or invalid stored action. */ }
  $('#saved-file-excel').addEventListener('click', () => {
    const button = $('#saved-file-excel');
    if (!button.hidden && button.dataset.plant && button.dataset.sn) openClientInExcel(button.dataset.plant, button.dataset.sn);
  });
  $('#saved-type-select').addEventListener('change', () => { $('#saved-type-status').textContent = ''; });
  $('#save-saved-type').addEventListener('click', async () => {
    const entry = app.savedFileEntry;
    const button = $('#save-saved-type');
    const status = $('#saved-type-status');
    const certificateType = $('#saved-type-select').value;
    const relativePath = String(entry?.relative_path || entry?.path || '').replaceAll('\\', '/').replace(/^Certificates\//i, '');
    if (!entry || !entry.sha256 || !relativePath || !certificateType) {
      status.textContent = 'Choose a type and refresh the saved files before saving.';
      return;
    }
    button.disabled = true;
    try {
      await api('/api/certificate-type', {relative_path:relativePath, sha256:entry.sha256, certificate_type:certificateType});
      entry.certificate_type = certificateType;
      renderFolderInventory();
      showSavedFile(entry);
      status.textContent = 'Type saved.';
    } catch (error) {
      status.textContent = error.message;
    } finally { button.disabled = false; }
  });
  $('#use-saved-upload').addEventListener('click', () => {
    if (!$('#use-saved-upload').hidden && app.savedFileEntry) openUploadCertificate({entry:app.savedFileEntry});
  });
  $('#add-client-button').addEventListener('click', openAddClient);
  $('#add-client-form').addEventListener('submit', event => { event.preventDefault(); saveClient(); });
  $('#add-client-form').addEventListener('input', event => {
    $('#add-client-error').hidden = true;
    if (event.target.name === 'sn' || app.clientPendingAttempt) validateClientSerialNumber();
  });
  $('#add-client-form').addEventListener('change', event => {
    $('#add-client-error').hidden = true;
    if (event.target.name === 'plant' || app.clientPendingAttempt) validateClientSerialNumber();
  });
  $('#add-client-material').addEventListener('click', () => {
    if (!app.clientBusy && (app.clientMaterialRows || []).length < 100) renderClientMaterials([...clientMaterialValues(), {}]);
  });
  for (const id of ['reconnect', 'dialog-reconnect']) $(`#${id}`).addEventListener('click', reconnectDashboard);
  for (const id of [...Object.keys(savedFilterDefaults), 'saved-file-sort']) {
    $(`#${id}`).addEventListener(id === 'saved-file-search' ? 'input' : 'change', restartSavedList);
  }
  $('#clear-file-search').addEventListener('click', () => {
    Object.keys(savedFilterDefaults).forEach(clearSavedFilter);
    restartSavedList();
    $('#saved-file-search').focus();
  });
  document.querySelectorAll('[data-quick-status]').forEach(button => button.addEventListener('click', () => {
    $('#status-filter').value = button.dataset.quickStatus;
    restartRecordList();
  }));
  $('#confirm-action').addEventListener('click', () => settleActionConfirmation(true));
  $('#open-entry-deletion').addEventListener('click', openEntryDeletion);
  $('#entry-delete-search').addEventListener('input', filterEntryDeletion);
  $('#entry-delete-clear').addEventListener('click',()=>{$('#entry-delete-search').value='';filterEntryDeletion();$('#entry-delete-search').focus();});
  $('#entry-delete-preview').addEventListener('click', reviewEntryDeletion);
  $('#entry-delete-back').addEventListener('click',()=>{backToEntryDeletionSelection();filterEntryDeletion();});
  $('#entry-delete-checked').addEventListener('change',()=>{$('#entry-delete-confirm').disabled=!$('#entry-delete-checked').checked || app.entryDeletionBusy;});
  $('#entry-delete-confirm').addEventListener('click', confirmEntryDeletion);
  $('#entry-deletion-modal').addEventListener('close',()=>{
    if(!app.entryDeletionBusy)app.entryDeletion=null;
    $('#entry-delete-checked').checked=false;$('#entry-delete-confirm').disabled=true;
  });
  $('#entry-deletion-modal').addEventListener('keydown',event=>{
    if(event.key==='Enter' && !['entry-delete-cancel','entry-delete-preview','entry-delete-confirm','entry-delete-back'].includes(event.target.id))event.preventDefault();
  });
  $('#remove-inward-checked').addEventListener('change', () => {
    $('#confirm-remove-inward').disabled = !$('#remove-inward-checked').checked || app.removalBusy;
  });
  $('#confirm-remove-inward').addEventListener('click', confirmInwardRemoval);
  $('#remove-inward-modal').addEventListener('keydown', event => {
    if (event.key === 'Enter' && !['keep-inward', 'confirm-remove-inward'].includes(event.target.id)) event.preventDefault();
  });
  $('#remove-inward-modal').addEventListener('close', () => {
    app.removalPending = null;
    $('#remove-inward-checked').checked = false;
    $('#confirm-remove-inward').disabled = true;
  });
  $('#action-confirmation-modal').addEventListener('close', () => {
    if (!$('#action-confirmation-modal').open && app.actionConfirmation) {
      const resolve = app.actionConfirmation; app.actionConfirmation = null; resolve(false);
    }
  });
  $('#keep-working-draft').addEventListener('click', () => { if (saveWorkingDraft()) toast('Working draft saved on this computer. Excel is unchanged.'); });
  $('#keep-editing').addEventListener('click', () => finishDraftClose(false));
  $('#discard-edits').addEventListener('click', () => finishDraftClose(true));
  $('#unsaved-modal').addEventListener('close', () => { if (!$('#unsaved-modal').open) app.afterDraftClose = null; });
  $('#restore-working-draft').addEventListener('click', restoreWorkingDraft);
  $('#discard-working-draft').addEventListener('click', () => {
    const next = app.draftRecovery;
    if (!clearWorkingDraft()) return;
    app.draftRecovery = null; $('#draft-recovery-modal').close();
    if (next?.resumeEditor) openDetails(next.row || null, true, next.startEditing || false);
  });
  $('#success-open-folder').addEventListener('click', async () => {
    if (app.lastGeneratedCertificatePath) await openCertificateLocally(app.lastGeneratedCertificatePath, false);
    else try { await api('/api/open-folder', {}); } catch (error) { toast(error.message, true); }
  });
  $('#saved-file-open').addEventListener('click', () => openCertificateLocally($('#saved-file-open').dataset.relativePath));
  $('#create-backup').addEventListener('click', () => backupAction('create'));
  $('#pick-backup-destination').addEventListener('click', () => backupAction('pick-destination'));
  $('#working-draft-status').hidden = !showWorkingDraftRecovery;
  $('#restore-backup').addEventListener('click', openBackupRestore);
  $('#confirm-backup-restore').addEventListener('click', restoreBackup);
  $('#reload-history').addEventListener('click', loadHistory);
  $('#history-section').addEventListener('toggle', () => { if ($('#history-section').open) loadHistory(); });
  window.addEventListener('beforeunload', event => {
    if (!editorDirty() && !app.generationBusy && !app.backupBusy && !app.completionBusy && !app.clientBusy && !app.settingsBusy && !app.detailsBusy) return;
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
  $('#upload-certificate').addEventListener('click', openUploadCertificate);
  $('#external-upload-file').addEventListener('change', inspectUploadedCertificate);
  $('#upload-cod').addEventListener('input', event => {
    if (app.uploadedFile && !event.target.readOnly) app.uploadedFile.manualCod = event.target.value;
  });
  $('#upload-certificate-form').addEventListener('submit', saveUploadedCertificate);
  $('#external-upload-target').addEventListener('change', () => updateUploadTarget(true));
  $('#upload-certificate-form').addEventListener('change', event => {
    if (event.target.name === 'certificate_type') {
      app.uploadTypeExplicit = Boolean(event.target.value);
      updateUploadTarget();
    }
  });
  $('#new-button').addEventListener('click', () => openDetails());
  $('#destruction-pending').addEventListener('click', () => {
    clearRecordFilterValues();
    $('#record-followup-filter').value = 'destruction_pending';
    restartRecordList();
  });
  $('#open-bulk').addEventListener('click', openBulkGeneration);
  $('#bulk-start-cod').addEventListener('input', updateBulkCodHint);
  $('#generate-all').addEventListener('click', startBulkGeneration);
  $('#bulk-select-all').addEventListener('change', event => {
    for (const index of app.bulkVisible || []) {
      if (event.target.checked) app.bulkChecked.add(index); else app.bulkChecked.delete(index);
      app.bulkCheckboxes[index].checked = event.target.checked;
    }
    updateBulkSelection();
  });
  $('#bulk-search').addEventListener('input', filterBulkClients);
  $('#bulk-clear-search').addEventListener('click', () => { $('#bulk-search').value = ''; filterBulkClients(); $('#bulk-search').focus(); });
  $('#bulk-generation-modal').addEventListener('close', () => { if (!$('#bulk-generation-modal').open) app.bulkSelection = null; });
  $('#more-filters').addEventListener('beforetoggle', event => { if (event.newState === 'open') positionRecordFilters(); });
  $('#more-filters').addEventListener('toggle', event => {
    $('#open-record-filters').setAttribute('aria-expanded', String(event.newState === 'open'));
    if (event.newState === 'open') positionRecordFilters();
  });
  window.addEventListener('resize', positionRecordFilters);
  $('#clear-filters').addEventListener('click', () => {
    $('#more-filters').hidePopover();
    resetRecordFilters();
  });
  document.querySelectorAll('[data-filter-status]').forEach(button => button.addEventListener('click', () => {
    $('#status-filter').value = button.dataset.filterStatus;
    restartRecordList();
  }));
  for (const id of [...Object.keys(recordFilterDefaults), 'attention-only', 'company-sort']) {
    $(`#${id}`).addEventListener(['search', 'serial-filter'].includes(id) ? 'input' : 'change', restartRecordList);
  }
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
      $(`#queue-body [data-expand="${index}"]`)?.focus();
    }
  });
  document.querySelectorAll('[data-view]').forEach(button => button.addEventListener('click', () => setView(button.dataset.view)));
  $('#details-form').addEventListener('submit', saveDetails);
  $('#replace-uploaded-pdf').addEventListener('click', () => {
    const replacement = app.editor?.uploadRow;
    if (replacement) requestDetailsClose(() => openUploadCertificate({replacement}));
  });
  $('#edit-pdf-details').addEventListener('click', async () => {
    if (!app.editor || app.detailsBusy || app.generationBusy) return;
    if (summaryEditorVisible()) {
      if (app.editor.firstGeneration) app.editor.pdfEditUsed = true;
      setEditorMode(true);
      $('#details-form').elements.namedItem('supplier').focus();
      return;
    }
    finishInlineEdit();
    setEditorMode(false);
    await previewDetails();
  });
  $('#excel-save-yes').addEventListener('click', () => saveExcelChoice(true));
  $('#excel-save-no').addEventListener('click', () => saveExcelChoice(false));
  $('#save-inward-details').addEventListener('click', saveInwardDetails);
  $('#open-companion-certificate').addEventListener('click', openCompanionCertificate);
  $('#primary-certificate-needed').addEventListener('change', changePrimaryCertificateNeeded);
  $('#excel-save-modal').addEventListener('close', () => { app.excelSavePayload = null; });
  $('#change-details').addEventListener('click', () => {
    setEditorMode(true); $('#details-form').elements.namedItem('supplier').focus();
  });
  $('#finish-editing').addEventListener('click', finishEditing);
  $('#add-material').addEventListener('click', addMaterialRow);
  $('#save-draft').addEventListener('click', saveDraft);
  $('#preview-details').addEventListener('click', previewDetails);
  $('#preview-full-pdf').addEventListener('click', toggleFullPdfPreview);
  $('#preview-attachment-previous').addEventListener('click', () => showAttachmentPage((app.previewAttachment?.page ?? 0) - 1));
  $('#preview-attachment-next').addEventListener('click', () => showAttachmentPage((app.previewAttachment?.page ?? 0) + 1));
  $('#preview-attachment-retry').addEventListener('click', () => showAttachmentPage(app.previewAttachment?.page ?? 0));
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
    schedulePreview();
  });
  $('#details-form').addEventListener('input', event => {
    if (event.target.id === 'primary-certificate-needed') return;
    if (event.target.name === 'certificate_type' && switchesLinkedCertificate(event.target.value)) return;
    if (event.target.name) app.editor?.touched.add(event.target.name);
    if (event.target.name === 'starting_cod') updateSingleCodControls();
    if (event.target.name === 'certificate_weight') applyCertificateWeight(event.target.value);
    updateMaterialWeightTotal();
    invalidatePreview();
    schedulePreview();
  });
  $('#details-form').addEventListener('change', event => {
    if (event.target.id === 'primary-certificate-needed') return;
    const form = event.currentTarget;
    if (event.target.name === 'data_destruction_needed') { changeDestructionChoice(); return; }
    if (event.target.name === 'certificate_type' && switchesLinkedCertificate(event.target.value)) {
      switchLinkedCertificate(event.target.value); return;
    }
    if (event.target.name) app.editor?.touched.add(event.target.name);
    if (event.target.name === 'certificate_weight') applyCertificateWeight(event.target.value);
    if (event.target.name === 'edit_scope') updateSourceControls();
    if (event.target.name === 'use_certificate_weight') updateSourceControls();
    updateMaterialWeightTotal();
    if (event.target.name === 'certificate_type') {
      finishInlineEdit();
      captureTemplateFields();
      const previousSource = form.elements.namedItem('source_kind').value;
      if (!app.editor.existing && !app.editor.sourceExplicit) {
        form.elements.namedItem('source_kind').value = event.target.value === 'standard' ? 'workbook' : 'manual';
        if (previousSource === 'manual' && form.elements.namedItem('source_kind').value === 'workbook'
            && form.elements.namedItem('certificate_weight').value.trim()) {
          form.elements.namedItem('use_certificate_weight').checked = true;
          app.editor.touched.add('use_certificate_weight');
        }
      }
      renderTemplateEditor();
      // A deliberate type change exposes that template's editable fields,
      // including the Siemens identifiers needed before creation.
      if (app.editor.firstGeneration) app.editor.pdfEditUsed = true;
      setEditorMode(true);
      updateSourceControls();
    } else if (event.target.name === 'source_kind') {
      app.editor.sourceExplicit = true;
      if (app.editor.recordId && app.editor.sourceSn == null) {
        if (event.target.value === 'workbook' && !app.editor.manualSourceIdentifiers) {
          app.editor.manualSourceIdentifiers = {
            plant:form.elements.namedItem('plant').value,
            cod_no:form.elements.namedItem('cod_no').value,
            cod_date:form.elements.namedItem('cod_date').value,
            issuedCertificate:app.editor.issuedCertificate,
            rawDate:app.editor.rawDetails?.cod_date,
          };
          form.elements.namedItem('cod_no').value = '';
          form.elements.namedItem('cod_date').value = '';
          if (app.editor.rawDetails) delete app.editor.rawDetails.cod_date;
          app.editor.issuedCertificate = false;
        } else if (event.target.value === 'manual' && app.editor.manualSourceIdentifiers) {
          const saved = app.editor.manualSourceIdentifiers;
          for (const key of ['plant', 'cod_no', 'cod_date']) form.elements.namedItem(key).value = saved[key];
          app.editor.issuedCertificate = saved.issuedCertificate;
          if (saved.rawDate !== undefined) (app.editor.rawDetails ||= {}).cod_date = saved.rawDate;
          delete app.editor.manualSourceIdentifiers;
        }
      }
      if (event.target.value === 'workbook' && form.elements.namedItem('certificate_weight').value.trim()) {
        form.elements.namedItem('use_certificate_weight').checked = true;
        app.editor.touched.add('use_certificate_weight');
      }
      updateSourceControls();
    } else if (event.target.name === 'plant') {
      suggestWorkbookId();
    }
    invalidatePreview();
    schedulePreview();
  });
  $('#details-modal').addEventListener('close', () => { if (!$('#details-modal').open) { finishInlineEdit(); invalidatePreview(); } });
  $('#warning-modal').addEventListener('close', () => { if (!$('#warning-modal').open) app.pendingConfirmation = null; });
  $('#settings-form').addEventListener('submit', saveSettings);
  $('#settings-form').addEventListener('input', updateSettingsFiles);
  $('#settings-form').addEventListener('change', updateReminderControls);
  $('#backup-options').addEventListener('toggle', () => { if ($('#backup-options').open) loadBackupStatus(); });
  $('#repair-reminder').addEventListener('click', () => saveSettings({preventDefault(){}, currentTarget:$('#settings-form')}, true));
  $('#warning-go-back').addEventListener('click', returnFromWarnings);
  $('#confirm-warnings').addEventListener('click', confirmWarnings);
  $('#open-settings').addEventListener('click', openSettings);
  $('#workbook-issue-settings').addEventListener('click', openSettings);
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
      suggestWorkbookId();
      invalidatePreview();
      schedulePreview();
      toast(`${labelPlant(result.plant)} workbook selected.`);
    } catch (error) { toast(error.message, true); }
    finally { setBusy(false); }
  });
  document.querySelectorAll('[data-pick-setting]').forEach(button => button.addEventListener('click', async () => {
    setBusy(true, 'Choose an Excel workbook…', 'The Windows file picker will open.');
    try {
      const result = await api('/api/pick-workbook', { for_settings: true });
      if (!result.cancelled) $('#settings-form').elements.namedItem(button.dataset.pickSetting).value = result.path;
      updateSettingsFiles();
    } catch (error) { toast(error.message, true); }
    finally { setBusy(false); }
  }));
  document.querySelectorAll('[data-disconnect-setting]').forEach(button => button.addEventListener('click', () => {
    if (app.settingsBusy || app.backupBusy) return;
    $('#settings-form').elements.namedItem(button.dataset.disconnectSetting).value = '';
    updateSettingsFiles();
    $('#settings-status').textContent = 'Click Save to apply the disconnection.';
  }));
  document.querySelectorAll('[data-create-setting]').forEach(button => button.addEventListener('click', async () => {
    const plant = button.dataset.createSetting;
    const label = labelPlant(plant);
    setBusy(true, `Create a ${label} workbook…`, 'Choose where to save the new Excel template.');
    try {
      const result = await api('/api/create-plant-workbook', { plant });
      if (result.cancelled) return;
      const input = $('#settings-form').elements.namedItem(`${plant}_path`);
      input.value = result.path;
      updateSettingsFiles();
      $('#settings-status').textContent = `${label} template created. Click Save to connect it.`;
      toast(`${label} template created. Click Save in Settings to connect it.`);
    } catch (error) {
      $('#settings-status').textContent = error.message;
      toast(error.message, true);
    } finally { setBusy(false); }
  }));
  const closeDialog = dialog => {
    if (dialog.id === 'entry-deletion-modal' && app.entryDeletionBusy) return;
    if (dialog.id === 'remove-inward-modal' && app.removalBusy) return;
    if (dialog.id === 'add-client-modal' && app.clientBusy) return;
    if (dialog.id === 'details-modal') requestDetailsClose();
    else if (dialog.id === 'settings-modal') requestSettingsClose();
    else if (!app.generationBusy && !app.backupBusy && !app.restoringDraft) dialog.close();
  };
  document.querySelectorAll('[data-close]').forEach(button => button.addEventListener('click', () => closeDialog($(`#${button.dataset.close}`))));
  for (const dialog of document.querySelectorAll('dialog')) {
    dialog.addEventListener('click', event => {
      if (event.target !== dialog || ['success-modal', 'add-client-modal'].includes(dialog.id)) return;
      const bounds = dialog.getBoundingClientRect();
      if (event.clientX < bounds.left || event.clientX > bounds.right || event.clientY < bounds.top || event.clientY > bounds.bottom) closeDialog(dialog);
    });
    dialog.addEventListener('cancel', event => { event.preventDefault(); closeDialog(dialog); });
  }
  window.addEventListener('keydown', event => {
    if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'k') {
      if (document.querySelector('dialog[open]')) return;
      if (app.workspaceView === 'analytics') setWorkspaceView('records');
      event.preventDefault(); $(app.workspaceView === 'files' ? '#saved-file-search' : '#search').focus();
    }
  });
}

async function init() {
  if (new URL(window.location.href).searchParams.get('view') === 'attention') $('#status-filter').value = 'needs_attention';
  app.loadingWorkspace = true;
  $('#new-button').disabled = true;
  $('#add-client-button').disabled = true;
  const now = new Date();
  $('#today-label').textContent = dateLabel(validDate(now.getFullYear(), now.getMonth() + 1, now.getDate()));
  configureDateInput($('#received-date-filter'));
  configureDateInput($('#received-to-filter'));
  const externalDate = $('#external-issued-date');
  if (externalDate) configureDateInput(externalDate);
  formatFormDates($('#add-client-form'));
  formatFormDates($('#details-form'));
  wire();
  if (new URL(window.location.href).searchParams.get('view') !== 'attention') setWorkspaceView('analytics');
  const templatesReady = loadTemplates();
  try {
    const response = await api('/api/state');
    app.settings = response.settings || {};
    setState(response.state);
    const startupStateVersion = app.stateVersion || 0;
    await templatesReady;
    const checked = await api('/api/workbook-changes');
    if (checked.state) setState(checked.state);
    else if (checked.changed === false && (app.stateVersion || 0) === startupStateVersion
        && typeof response.workspace_revision === 'string') {
      // Keep the revision paired with the initial rows. An unchanged-file
      // response carries no new revision and must not erase this snapshot.
      app.workspaceRevision = response.workspace_revision;
    }
    if (showWorkingDraftRecovery) offerDraftRecovery();
  } catch (error) {
    toast(error.message, true);
    $('#queue-body').replaceChildren();
    const tr = element('tr'); const td = element('td'); td.colSpan = 6;
    const empty = element('div', 'empty-state');
    empty.append(element('span', 'empty-icon', '!'));
    empty.append(element('strong', '', 'Could not load certificate data'));
    empty.append(element('span', '', error.message));
    td.append(empty); tr.append(td); $('#queue-body').append(tr);
  } finally {
    app.loadingWorkspace = false;
    $('#new-button').disabled = false; $('#add-client-button').disabled = false; summarize();
    if (app.workspaceView === 'analytics') window.CertificateAnalytics?.load();
  }
  window.setInterval(syncWorkbookChanges, 15000);
  pollReminderReview();
  window.setInterval(pollReminderReview, 2000);
  window.addEventListener('focus', pollReminderReview);
  document.addEventListener('visibilitychange', pollReminderReview);
  window.addEventListener('focus', syncWorkbookChanges);
  document.addEventListener('visibilitychange', syncWorkbookChanges);
}

init();
