const $ = (selector, parent = document) => parent.querySelector(selector);
const app = { state: { rows: [] }, settings: {}, filtered: [], templates: [], templatesLoaded: false,
  editor: null, materialRows: [], page: 1, savedFilesPage: 1, workspaceView: 'records', previewRevision: 0, detailsBusy: false, pendingConfirmation: null, toastTimer: null };
const templateNames = { standard: 'Standard', sita: 'Sita International', siemens: 'Siemens Medical', destruction: 'Secured Data Destruction' };
const commonTemplateKeys = new Set(['supplier', 'location', 'inward_date', 'certificate_weight', 'form6', 'reference', 'cod_no', 'cod_date']);
const revisionEndpoints = new Set(['/api/create', '/api/generate', '/api/manual/save', '/api/manual/generate', '/api/manual/complete', '/api/complete', '/api/backup/restore', '/api/clients/add']);
const workingDraftKey = 'ecoreco.certificate.working-draft.v1';
const fieldLabels = { plant: 'Plant', source_kind: 'Use details from', certificate_type: 'Certificate type', case_id: 'Serial Number', supplier: 'Company name', location: 'City', inward_date: 'Received date', certificate_weight: 'Certificate Weight (kg)', item: 'Item description', items: 'Materials and weights', quantity: 'Quantity', per_piece_weight: 'Per piece weight (kg)', weight: 'Inward weight (kg)', form6: 'Form 6 number', reference: 'Reference number', cod_no: 'Certificate number', cod_date: 'Certificate date', starting_cod:'Starting COD number' };
function isManual(row) { return row.source_kind === 'manual'; }

function excelTarget(entry) {
  const source = entry.managed_owner || entry;
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
    if (response.warning) toast('Excel opened. Check the Serial Number, then save and refresh.');
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
    const grid = element('div', 'client-material-grid');
    const inputs = {};
    const addField = ([key, label, type, placeholder], className = '') => {
      const field = element('label', `field ${className}`.trim()); field.append(element('span', '', label));
      const input = element('input'); input.type = type; input.value = item[key] ?? '';
      input.maxLength = key === 'item' ? 500 : 100;
      input.inputMode = key === 'item' || key === 'quantity' ? 'text' : 'decimal'; input.placeholder = placeholder;
      input.setAttribute('aria-label', `Material ${index + 1}: ${label}`);
      field.append(input); inputs[key] = input; return field;
    };
    grid.append(addField(clientMaterialFields[0], 'client-material-name'), addField(clientMaterialFields[3]));
    const calculation = element('div', 'material-calculation client-material-calculation');
    calculation.append(element('strong', 'client-material-calc-title', 'Calculate weight'));
    const equation = element('div', 'client-material-equation');
    equation.append(addField(clientMaterialFields[1]), element('span', 'client-material-operator', '×'),
      addField(clientMaterialFields[2]));
    calculation.append(equation);
    grid.append(calculation);
    const hint = element('small', 'client-material-hint');
    hint.hidden = true;
    grid.append(hint); Object.defineProperty(inputs, 'hint', {value:hint});
    const updateCalculatedWeight = (recalculate = true) => {
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
      updateClientMaterialWeightTotal();
    };
    for (const key of ['quantity', 'per_piece_weight']) inputs[key].addEventListener('input', updateCalculatedWeight);
    inputs.weight.addEventListener('input', () => {
      delete inputs.weight.dataset.autoCalculated;
      updateClientMaterialWeightTotal();
    });
    inputs.item.addEventListener('input', updateClientMaterialWeightTotal);
    updateCalculatedWeight(false);
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
  details.items = clientMaterialValues().map(item => Object.fromEntries(
    ['item','quantity','per_piece_weight','weight'].map(key => [key, String(item[key] ?? '').trim()])));
  return {plant:String(values.plant ?? '').trim(), details};
}

function validateClientSerialNumber() {
  const form = $('#add-client-form'), input = form.elements.namedItem('sn');
  const attempt = app.clientPendingAttempt;
  // The row may be our committed add whose response was lost. Only the exact
  // same attempted request may reach the server's idempotency check again.
  const retry = attempt && attempt.requestId === app.clientRequestId
    && attempt.signature === JSON.stringify(clientEntryPayload(form));
  const duplicate = !retry && duplicateSerialNumber(form.elements.namedItem('plant').value, input.value);
  const message = duplicate ? `Already used${duplicate.supplier ? ` by ${duplicate.supplier}` : ''}. Use another number or leave blank.` : '';
  input.setCustomValidity(message);
  input.setAttribute('aria-invalid', String(Boolean(duplicate)));
  const hint = $('#client-serial-number-hint');
  hint.textContent = message;
  hint.hidden = !message;
  hint.classList.toggle('field-error', Boolean(duplicate));
  return !duplicate;
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
  $('#add-client-modal').showModal();
  form.elements.namedItem('supplier').focus();
}

async function saveClient() {
  const form = $('#add-client-form');
  if (app.clientBusy) return;
  const serialValid = validateClientSerialNumber();
  if (!serialValid) { form.elements.namedItem('sn').focus(); return; }
  if (!form.reportValidity()) return;
  const payload = clientEntryPayload(form);
  app.clientBusy = true;
  $('#add-client-error').hidden = true;
  const controls = [...form.querySelectorAll('button, input, select')];
  controls.forEach(control => { control.disabled = true; });
  $('#save-client').disabled = true; $('#save-client').textContent = 'Adding to Excel…';
  try {
    app.clientPendingAttempt = {requestId:app.clientRequestId, signature:JSON.stringify(payload)};
    const response = await api('/api/clients/add', {...payload, request_id:app.clientRequestId, expected_revision:app.clientRevision});
    if (!response.client_added || !response.client) throw new Error('The client could not be confirmed. Check Excel before adding it again.');
    app.clientEntryInitialized = false; app.clientRevision = null; app.clientRequestId = null; app.clientPendingAttempt = null;
    $('#add-client-modal').close();
    $('#status-filter').value = 'all'; $('#template-filter').value = 'all';
    $('#plant-filter').value = response.client.plant; $('#search').value = String(response.client.sn);
    setWorkspaceView('records'); setState(response.state);
    toast(`${response.client.supplier} added to ${labelPlant(response.client.plant)}. ${response.client.status === 'ready' ? 'Ready to create a certificate.' : 'Fill the remaining details in Excel when ready.'}`);
  } catch (error) {
    $('#add-client-error').textContent = noticeText(error.message);
    $('#add-client-error').hidden = false;
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
    const fields = element('div', 'material-row-fields');
    const calculation = element('details', 'material-calculation');
    calculation.append(element('summary', '', 'Calculate from quantity (optional)'));
    calculation.append(element('p', '', 'Leave Weight blank to use quantity × weight per piece.'));
    const calculationFields = element('div', 'material-row-fields');
    calculation.append(calculationFields);
    calculation.open = !String(item?.weight ?? '').trim()
      && Boolean(String(item?.quantity ?? '').trim() || String(item?.per_piece_weight ?? '').trim());
    const inputs = {};
    for (const [key, title, type, placeholder] of materialFields) {
      const label = element('label', `field material-field material-${key}`);
      label.append(element('span', '', title));
      const input = element('input');
      input.name = materialInputName(index, key);
      input.type = type;
      input.maxLength = key === 'item' ? 500 : 100;
      input.inputMode = key === 'item' ? 'text' : 'decimal';
      input.placeholder = placeholder;
      input.value = item?.[key] == null ? '' : String(item[key]);
      label.append(input);
      (['quantity', 'per_piece_weight'].includes(key) ? calculationFields : fields).append(label);
      inputs[key] = input;
    }
    row.append(heading, fields, calculation);
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

function labelPlant(value) { return value === 'plant2' ? 'Plant 2' : value === 'plant1' ? 'Plant 1' : 'Setup'; }
function stateOf(row) {
  const raw = String(row.certificate_status || 'blank').trim().toLowerCase().replace(/[\s-]+/g, '_');
  return ({ created: 'generated', drafted: 'generated', sent: 'completed',
    check_details: 'alert', working_draft: 'alert', 'generated_manual': 'generated_manual' })[raw] || raw;
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
  return ({ ready: 'Ready', alert: 'Needs attention', draft: 'Draft', generated: 'Generated', generated_manual: 'Generated Manual', completed: 'Completed', na: 'N/A', unknown: 'Needs attention' })[state] || 'Needs attention';
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
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2])).toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' });
}
function escapeText(value) { return String(value ?? '').trim(); }

function readableNumber(value) {
  if (value === '' || value == null) return '';
  const number = Number(value);
  return Number.isFinite(number) ? String(Number(number.toPrecision(12))) : String(value);
}

function warningInfo(value) {
  const text = String(value || '').trim();
  const info = (title, message, check = '') => ({title, message, check, original: text});
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
  if (/Windows error (?:32|33)|workbook.*(?:in use|busy)|Windows still reports .* in use/i.test(text)) return {...info(
    'The Excel file is in use.', '', 'Close Excel or wait a moment, then retry.'), compact:true};
  let sourceMatch = text.match(/^(.+?\.xls[mx]) is unavailable: (.+)$/i);
  if (sourceMatch && /(?:close.*Excel|Windows error (?:32|33)|already locked)/i.test(sourceMatch[2])) return info(
    `Close ${sourceMatch[1]} in Excel`, 'Excel is using this file, so the app cannot check it yet.', 'Save and close the file in Excel, then click Refresh.');
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
  match = text.match(/^Reference (.+?) already appears in .+\. This may be the same inward or an already-created certificate\. Review it; continuing keeps this reference unchanged\.$/i);
  if (match) return info('Reference already used', `Reference ${match[1]} is on an existing certificate. Continuing keeps this reference unchanged.`,
    'Check whether this is the same inward. If it is, cancel and open the existing certificate.');
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
  return /^Excel is open or in use\. Showing its saved contents/i.test(String(value || '').trim());
}

function noticeText(value) {
  const warning = warningInfo(value);
  if (warning.compact) return [warning.title, warning.message, warning.check].filter(Boolean).join(' ');
  return simpleWarning(value);
}

function showDetailsError(message) {
  const node = $('#details-error');
  node.textContent = noticeText(message);
  node.hidden = false;
  const form = $('#details-form');
  const existingWorkbook = form.elements.namedItem('source_kind').value === 'workbook' && app.editor?.sourceSn != null;
  const plant = existingWorkbook ? app.editor.sourcePlant : form.elements.namedItem('plant').value;
  const sn = existingWorkbook ? app.editor.sourceSn : form.elements.namedItem('case_id').value;
  if (['plant1', 'plant2'].includes(plant) && sn && /Serial Number|Excel|workbook/i.test(String(message))) {
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
  if (warning.original && !warning.compact) {
    const details = element('details', 'issue-details');
    details.append(element('summary', '', 'Details'), element('p', '', warning.original)); card.append(details);
  }
  return card;
}

function renderWarningList(list, warnings) {
  list.classList.add('issue-list');
  const unique = new Map();
  for (const value of warnings) {
    if (isRoutineNotice(value)) continue;
    const warning = warningInfo(value);
    const key = [warning.title, warning.message, warning.check].join('|');
    if (!unique.has(key)) unique.set(key, value);
  }
  list.replaceChildren(...[...unique.values()].map(warningCard));
}

function nextStep(row) {
  if (recordIdentityIssue(row)) return recordIdentityIssue(row);
  if (row.destruction_pending) return row.pdf_filename
    ? 'Send the data destruction PDF, then mark this request Completed. Payment may depend on it.'
    : 'Data destruction requested. Complete the device details and create the PDF; payment may depend on it.';
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

function toast(message, isError = false) {
  const node = $('#toast');
  $('#toast-message').textContent = isError ? noticeText(message) : message;
  $('#dismiss-toast').hidden = !isError;
  node.classList.toggle('error', isError);
  node.classList.add('visible');
  syncToastLayer();
  clearTimeout(app.toastTimer);
  if (!isError) app.toastTimer = setTimeout(dismissToast, 4500);
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

async function api(path, payload) {
  if (payload !== undefined && revisionEndpoints.has(path)) payload = { ...payload, expected_revision: payload.expected_revision ?? app.workspaceRevision ?? app.state.workspace_revision ?? '' };
  const options = payload === undefined ? { method: 'GET' } : {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload),
  };
  let response;
  try { response = await fetch(path, options); }
  catch (error) {
    $('#offline-title').textContent = 'Connection lost';
    $('#offline-message').textContent = 'The app is not responding. Your saved files are safe. Check the connection, or reopen Prepare Certificates. Do not repeat creation until you have checked Saved certificates.';
    $('#reconnect').hidden = false;
    $('#offline-screen').hidden = false;
    if (document.querySelector('dialog[open]') && !$('#connection-modal').open) $('#connection-modal').showModal();
    throw new Error('The local dashboard server is not responding. Reopen Prepare Certificates.bat.');
  }
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error((data.error || `Request failed (${response.status}).`) + (response.status === 409 ? ' Refresh Excel and reopen the details to review the latest data. Your working draft is kept.' : ''));
  // A revision belongs to the rows that were actually read. Preview, history,
  // picker and cached startup responses must not bless stale Excel rows.
  const freshRows = data.state && (path === '/api/check' || path === '/api/workbook-changes' || revisionEndpoints.has(path));
  const revision = data.state?.workspace_revision ?? data.workspace_revision;
  if (freshRows && typeof revision === 'string') {
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
    updateSettingsFiles(); updateReminderControls();
  }
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

function savedFilePlant(entry) {
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
  return row?.supplier || certificateFileTitle(entry);
}

function savedFileMatches(entry) {
  const plant = $('#saved-plant-filter').value || 'all';
  if (plant !== 'all' && savedFilePlant(entry) !== plant) return false;
  const words = $('#saved-file-search').value.toLowerCase().trim().split(/\s+/).filter(Boolean);
  const haystack = [savedFileLabel(entry), entry.path, ...currentFileNumbers(entry), ...(entry.references || [])]
    .join(' ').replaceAll('_', ' ').toLowerCase();
  return words.every(word => haystack.includes(word));
}

function setWorkspaceView(view) {
  app.workspaceView = view === 'files' ? 'files' : 'records';
  const files = app.workspaceView === 'files';
  $('#records-workspace').hidden = files;
  $('#folder-inventory').hidden = !files;
  $('#view-records').setAttribute('aria-pressed', String(!files));
  $('#view-files').setAttribute('aria-pressed', String(files));
  if (files) {
    $('#saved-plant-filter').value = $('#plant-filter').value;
    renderFolderInventory();
  }
  $('#page-title').textContent = files ? 'Saved certificates' : 'Your certificates';
  $('#page-description').textContent = files ? 'Find a PDF and open it. Files added manually appear here too.'
    : 'Find a company, check its certificate, and create the PDF.';
}

function resetRecordFilters() {
  $('#search').value = '';
  $('#status-filter').value = 'all';
  $('#template-filter').value = 'all';
  app.page = 1; app.expanded = null;
  render(); // A reset must never unexpectedly switch the selected plant.
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
  if (['Auto Generated', 'Created in App', 'Manually Added'].includes(folder)) return folder;
  if (entry.managed_owner?.source_kind === 'manual') return 'Created in App';
  if (entry.managed_owner?.source_kind === 'workbook') return 'Auto Generated';
  return 'Existing PDF';
}

function showSavedFile(entry) {
  $('#saved-file-name').textContent = certificateFileTitle(entry);
  const facts = $('#saved-file-facts');
  facts.replaceChildren();
  const pathParts = String(entry.path || '').replaceAll('\\', '/').split('/').filter(Boolean);
  const folder = pathParts.length > 2 ? pathParts.slice(1, -1).join(' / ') : '';
  const details = [['Source', savedFileSource(entry)], ['Certificate number', currentFileNumbers(entry).join(', ') || 'Could not be read'],
    ['Reference number', entry.references?.join(', ') || 'Not entered (optional)']];
  if (folder) details.push(['Folder', folder]);
  details.push(['PDF file', certificateFileName(entry)]);
  for (const [label, value] of details) {
    const fact = element('div'); fact.append(element('dt', '', label), element('dd', '', value)); facts.append(fact);
  }
  renderWarningList($('#saved-file-warnings'), entry.warnings || []);
  $('#saved-file-warnings').hidden = !entry.warnings?.length;
  const url = savedFileUrl(entry);
  $('#saved-file-open').hidden = !url;
  $('#saved-file-open').dataset.relativePath = entry.relative_path || String(entry.path || '').replace(/^Certificates[\\/]/, '').replaceAll('\\', '/');
  const target = excelTarget(entry);
  $('#saved-file-excel').hidden = !target;
  $('#saved-file-excel').dataset.plant = target?.plant || '';
  $('#saved-file-excel').dataset.sn = target?.sn || '';
  if (!url) facts.append(element('p', '', 'Open the PDF folder to view this file in its subfolder.'));
  $('#saved-file-modal').showModal();
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
    ? 'All saved PDFs are shown here. Created in App and Manually Added stay in separate folders.' : 'No saved PDFs yet. Generated certificates will appear here.';
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
    copy.append(title, element('span', '', `COD ${currentFileNumbers(entry).join(', ') || 'not identified'} · ${plant === 'unknown' ? 'Plant not identified' : labelPlant(plant)}`));
    copy.append(element('span', 'saved-file-source', savedFileSource(entry)));
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
  $('#clear-file-search').hidden = !$('#saved-file-search').value.trim() && ['all', ''].includes($('#saved-plant-filter').value);
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

function summarize() {
  const destructionPending = app.state.rows.filter(row => row.destruction_pending).length;
  $('#destruction-pending').hidden = !destructionPending;
  $('#destruction-pending').textContent = `${destructionPending} data destruction request${destructionPending === 1 ? '' : 's'} awaiting completion. Payment may depend on these. Show requests`;
  const rows = app.state.rows.filter(row => matches(row, true));
  const ready = rows.filter(row => stateOf(row) === 'ready' && !hasWarnings(row)).length;
  const selectedPlant = $('#plant-filter').value;
  const workbookReady = app.state.rows.filter(row => !isManual(row) && stateOf(row) === 'ready'
    && !row.issuance_blocked && !recordIdentityIssue(row)
    && (selectedPlant === 'all' || row.plant === selectedPlant)).length;
  const attention = rows.filter(needsAttention).length;
  const generated = rows.filter(row => stateOf(row) === 'generated').length;
  const generatedManual = rows.filter(row => stateOf(row) === 'generated_manual').length;
  const completed = rows.filter(row => stateOf(row) === 'completed').length;
  const overdue = rows.filter(row => Boolean(row.overdue)).length;
  $('#count-all').textContent = rows.length;
  $('#count-blank').textContent = rows.filter(row => stateOf(row) === 'blank').length;
  $('#count-na').textContent = rows.filter(row => stateOf(row) === 'na').length;
  $('#count-ready').textContent = ready;
  $('#count-attention').textContent = attention;
  $('#count-generated').textContent = generated;
  $('#count-generated-manual').textContent = generatedManual;
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
  const count = app.state.rows.filter(row => !isManual(row) && stateOf(row) === 'ready'
    && !row.issuance_blocked && !recordIdentityIssue(row)
    && payload.plants.includes(row.plant)).length;
  if (!count) { toast('No Ready Excel rows in this plant.'); return; }
  app.bulkSelection = { ...payload, expected_revision: app.workspaceRevision ?? app.state.workspace_revision ?? '' };
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
    ? `Next COD: ${numbering.next_cod}. Numbers continue automatically.` : 'COD numbers continue automatically.';
  const plants = payload.plants.map(labelPlant).join(' and ');
  $('#bulk-description').textContent = `${count} certificate${count === 1 ? '' : 's'} from ${plants}.`;
  const warned = app.state.rows.filter(row => !isManual(row) && stateOf(row) === 'ready'
    && !row.issuance_blocked && !recordIdentityIssue(row)
    && payload.plants.includes(row.plant) && hasWarnings(row)).length;
  $('#bulk-description').textContent += warned ? ` ${warned} need${warned === 1 ? 's' : ''} warning review.` : '';
  $('#generate-count').textContent = `(${count})`;
  $('#bulk-generation-modal').showModal();
  if (needsStart) startInput.focus();
}

async function startBulkGeneration() {
  const payload = app.bulkSelection;
  if (!payload || app.generationBusy) return;
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
  const received = row => normalizeDate(row.inward_date) || '9999-99-99';
  return received(a).localeCompare(received(b))
    || String(a.plant || '').localeCompare(String(b.plant || ''))
    || String(a.sn ?? a.case_id ?? '').localeCompare(String(b.sn ?? b.case_id ?? ''), undefined, {numeric:true})
    || String(a.supplier || '').localeCompare(String(b.supplier || ''));
}

function matches(row, ignoreStatus = false) {
  const plant = $('#plant-filter').value;
  const status = ignoreStatus === true ? 'all' : $('#status-filter').value;
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
  const meta = [row.certificate_no ? `COD ${row.certificate_no}` : '', row.location || '', templateId(row) !== 'standard' ? templateLabel(row) : ''].filter(Boolean);
  if (meta.length) copy.append(element('small', '', meta.join(' · ')));
  const issues = rowWarningSummary(row);
  if (issues.length && $('#status-filter').value !== 'needs_attention' && !['generated', 'generated_manual', 'completed'].includes(stateOf(row))) {
    const warning = warningInfo(issues[0]);
    const message = warning.original ? warning.title : warning.message;
    const note = element('small', 'row-alert-note', message.length > 130 ? message.slice(0, 127) + '…' : message);
    note.title = issues.map(simpleWarning).join(' · '); copy.append(note);
  }
  wrap.append(copy);
  td.append(wrap);
  return td;
}

function attentionReason(row) {
  const state = stateOf(row);
  if (recordIdentityIssue(row)) return recordIdentityIssue(row);
  if (Number.isFinite(row.age_days) && row.age_days < 0) return 'Received date is in the future. Check the date.';
  if (row.issuance_blocked) return 'Details need fixing before this certificate can be created.';
  if (hasWarnings(row) && !['generated', 'generated_manual', 'completed'].includes(state)) {
    const issue = warningInfo(rowWarningSummary(row)[0] || 'Check the certificate details.');
    return issue.original ? `${issue.title}. Review before creating.` : issue.message;
  }
  if (['generated', 'generated_manual'].includes(state)) return row.pdf_filename
    ? 'PDF ready. Send it, then mark Completed.' : 'Saved PDF is missing. Review the certificate.';
  if (state === 'blank') return 'Fill the missing inward details in Excel.';
  if (row.reminder_due || row.overdue) return Number.isFinite(row.age_days) && row.age_days >= 0
    ? `Certificate pending · received ${row.age_days} days ago.` : 'Certificate pending.';
  return 'Check the certificate details and Excel status.';
}

function makeStatus(row) {
  const td = element('td');
  const status = element('span', `status-pill ${statusClass(row)}`, stateLabel(row));
  if (stateOf(row) === 'blank') {
    status.setAttribute('aria-label', 'No status selected');
    status.title = 'No status selected';
  }
  td.append(status);
  if (row.destruction_pending) td.append(element('small', 'row-alert-note', 'Data destruction pending'));
  if (!isManual(row)) {
    const linked = app.state.rows.filter(other => other.destruction_source?.plant === row.plant
      && String(other.destruction_source.case_id) === String(row.sn));
    if (linked.length) td.append(element('small', linked.some(other => other.destruction_pending) ? 'row-alert-note' : 'age-note',
      linked.some(other => other.destruction_pending) ? 'Data destruction requested · Pending' : 'Data destruction completed'));
  }
  if ($('#status-filter').value === 'needs_attention') td.append(element('small', 'age-note attention-reason', attentionReason(row)));
  else if (Number.isFinite(row.age_days) && row.age_days < 0 && needsAttention(row)) td.append(element('small', 'age-note', 'Received date is in the future'));
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

function makeActions(row, index) {
  const td = element('td');
  const wrap = element('div', 'row-actions');
  if (recordIdentityIssue(row) || (row.issuance_blocked && !['generated', 'generated_manual', 'completed'].includes(stateOf(row)))) {
    wrap.append(actionButton('View details', 'details', index));
  }
  else if (stateOf(row) === 'ready' || (isManual(row) && ['blank', 'draft'].includes(stateOf(row)))) {
    const correction = $('#status-filter').value === 'needs_attention' && hasWarnings(row);
    wrap.append(actionButton(correction ? 'Review details' : 'Review & create', 'details', index, true));
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
    } else wrap.append(actionButton('Edit details', 'edit-details', index));
  }
  else if (['generated', 'generated_manual'].includes(stateOf(row))) wrap.append(actionButton('Edit certificate', 'edit-details', index, true));
  else {
    wrap.append(actionButton('Review details', 'edit-details', index, true));
  }
  if (!recordIdentityIssue(row) && (row.has_issued_certificate || ['generated', 'generated_manual', 'completed'].includes(stateOf(row)))) {
    if (!wrap.children.length || ![...wrap.children].some(button => button.dataset.action === 'edit-details')) {
      wrap.append(actionButton('Edit certificate', 'edit-details', index));
    }
  }
  if (isManual(row)) wrap.append(actionButton('Copy as new', 'copy', index));
  else if (['plant1','plant2'].includes(row.plant) && String(row.sn ?? '').trim()) {
    if (!recordIdentityIssue(row)) wrap.append(actionButton('Data destruction', 'destruction', index));
    wrap.append(excelOpenButton({plant:row.plant,sn:String(row.sn),sheet:row.sheet,row:row.row}, row.supplier, 'row-action'));
  }
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
  for (const [label, value] of [['Certificate Weight', row.display_weight != null ? `${readableNumber(row.display_weight)} kg` : 'Not entered'], ['Form 6', row.form6 || 'Not entered'], ['Reference', row.reference || 'Not entered (optional)'], ['Serial Number', row.sn || 'Manual certificate']]) {
    const fact = element('div'); fact.append(element('dt', '', label), element('dd', '', value)); facts.append(fact);
  }
  content.append(facts);
  const issues = rowWarningSummary(row);
  if (issues.length) {
    const warning = element('div', 'detail-warnings');
    warning.append(element('strong', '', ['generated', 'generated_manual', 'completed'].includes(stateOf(row)) ? 'Warnings accepted when generated' : 'Check before creating'));
    const list = element('ul');
    renderWarningList(list, issues);
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
  actions.append(actionButton(recordIdentityIssue(row) ? 'View details' : 'Edit details', recordIdentityIssue(row) ? 'details' : 'edit-details', index));
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
  renderWorkbookIssues();
  const view = $('#status-filter').value === 'needs_attention' ? 'attention' : 'queue';
  $('#page-title').textContent = app.workspaceView === 'files' ? 'Saved certificates' : view === 'attention' ? 'Needs attention' : 'Your certificates';
  $('#page-description').textContent = view === 'attention'
    ? 'Check missing details and pending certificates.'
    : 'Find a company, check its certificate, and create the PDF.';
  if (app.workspaceView === 'files') $('#page-description').textContent = 'Find a PDF and open it. Files added manually appear here too.';
  $('#queue-title').textContent = ({all:'All companies', ready:'Ready', needs_attention:'Needs attention',
    generated:'Generated', generated_manual:'Generated Manual', completed:'Completed',
    blank:'No status selected', na:'N/A'})[$('#status-filter').value] || 'All companies';
  $('#queue-subtitle').textContent = ({
    generated: 'Open a PDF to check it. After sending, mark it Completed.',
    generated_manual: 'These PDFs were made manually. Send them, then mark them Completed.',
    completed: 'These certificates are marked Completed.',
    na: 'No certificate is needed for these rows. Automatic creation skips them.',
    needs_attention: 'Read the issue beside each company, then choose its next action. Completed and N/A stop date reminders.',
  })[$('#status-filter').value] || 'Choose a company to check its certificate.';
  document.querySelectorAll('[data-quick-status]').forEach(button => button.setAttribute('aria-pressed', String(button.dataset.quickStatus === $('#status-filter').value)));
  document.querySelectorAll('[data-filter-status]').forEach(button => button.setAttribute('aria-pressed', String(button.dataset.filterStatus === $('#status-filter').value)));
  $('#reset-record-filters').hidden = !$('#search').value.trim() && $('#status-filter').value === 'all' && $('#template-filter').value === 'all';

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
}

function renderWorkbookIssues() {
  const notes = $('#workbook-maintenance-note');
  const plants = [...new Set((app.state.maintenance_notes || []).map(note => labelPlant(note.plant)))];
  notes.hidden = plants.length === 0;
  notes.textContent = plants.length
    ? `${plants.join(' and ')}: Excel row order was kept. Sort dates in Excel if needed; certificates can still be created.` : '';
  const messages = (app.state.workbook_issues || []).map(row => row.detail || '').filter(Boolean);
  if (app.state.setup_warning) messages.push(app.state.setup_warning);
  if (app.state.scan_complete === false && !app.state.workbooks_disconnected && !messages.length) messages.push('Both Excel files could not be checked. Choose the correct plant files in Settings, then refresh.');
  const unique = [...new Set(messages)];
  $('#workbook-issues').hidden = unique.length === 0;
  renderWarningList($('#workbook-issue-list'), unique);
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
    if (response.state?.workbooks_disconnected) toast('No Excel files connected.');
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
      body:JSON.stringify({client_id:app.reviewClientId, visible:!document.hidden, acknowledged:app.lastReminderReview})});
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
  try {
    const response = await api('/api/workbook-changes');
    if (response.state) setState(response.state);
    app.workbookSyncError = '';
  } catch (error) {
    if (app.workbookSyncError !== error.message) toast(error.message, true);
    app.workbookSyncError = error.message;
  } finally { app.syncingWorkbooks = false; }
}

function rowWarningSummary(row) {
  const detail = String(row.detail || '').trim();
  if (Array.isArray(row.warnings) && row.warnings.length) return [...new Set(row.warnings.filter(value => !isRoutineNotice(value)))];
  if (row.outcome === 'needs_review' && stateOf(row) === 'ready') return [detail || 'Review this row before issuing.'];
  if (row.system_alert && detail && stateOf(row) !== 'generated') return [detail];
  return [];
}

function showWarnings(warnings, continuation) {
  const entries = warnings.map(row => typeof row === 'string' ? { supplier: continuation.payload?.details?.supplier || 'Certificate', warnings: [row] } : row);
  app.pendingConfirmation = { ...continuation, hasWarnings: entries.length > 0 };
  $('#warning-go-back').textContent = continuation.fromDetails ? 'Go back and edit' : 'Cancel batch';
  $('#warning-description').textContent = entries.length ? entries.length === 1
    ? 'Check each issue below before creating this certificate.' : `${entries.length} certificates need a check. Review the issues for each company.` : 'No issues were found. Confirm to create the PDF.';
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
  if (!warnings.length) list.append(element('div', 'warning-item', 'No warnings were found.'));
  $('#warning-modal').showModal();
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
    if (created) {
      if (fromDetails) { clearWorkingDraft(); app.editor.baseline = detailsBody($('#details-form')); $('#details-modal').close(); }
      showGenerationSuccess(response);
    } else toast('No certificates were generated. Check the selected rows.');
  } catch (error) {
    if (fromDetails) showDetailsError(error.message);
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
    // The attachment changes the workspace. Reread the rows instead of
    // treating the picker's token as permission to save cached Excel data.
    // Keep this editor's reviewed snapshot frozen if external data changed.
    await refreshState();
  } catch (error) { toast(error.message, true); }
  finally { setDetailsBusy(false); setBusy(false); schedulePreview(); }
}

function deviceReceiptEditor(raw, key) {
  const section = element('section', 'device-receipts');
  section.append(element('h3', '', 'Devices received'), element('p', 'form-hint', 'Enter only the devices covered by this destruction certificate.'));
  const storage = element('textarea'); storage.hidden = true; storage.dataset.certificateField = key;
  let rows = [];
  try { const parsed = JSON.parse(raw || '[]'); if (Array.isArray(parsed)) rows = parsed; } catch { /* Keep incomplete entries editable. */ }
  if (!rows.length) rows = [{date:'', count:''}];
  const list = element('div', 'device-receipt-list');
  const total = element('p', 'form-hint');
  const add = element('button', 'button button-quiet', '+ Add receipt'); add.type = 'button';
  function sync(changed = false) {
    storage.value = JSON.stringify(rows);
    const count = rows.reduce((sum, row) => sum + (/^[0-9]+$/.test(String(row.count)) ? Number(row.count) : 0), 0);
    total.textContent = `Total devices: ${count}`;
    add.disabled = rows.length >= 100;
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
        input.addEventListener('input', () => { row[name] = input.value; sync(true); });
        label.append(input); line.append(label);
      }
      const remove = element('button', 'text-button', 'Remove'); remove.type = 'button';
      remove.setAttribute('aria-label', `Remove receipt ${index + 1}`);
      remove.addEventListener('click', () => { rows.splice(index, 1); renderRows(); sync(true); });
      line.append(remove); list.append(line);
    });
  }
  add.addEventListener('click', () => { if (rows.length < 100) { rows.push({date:'', count:''}); renderRows(); sync(true); } });
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
  if (destructionOption) destructionOption.disabled = Boolean(app.editor.sourceSn != null && form.elements.namedItem('source_kind').value === 'workbook');
  typeSelect.disabled = Boolean(app.editor.existing && (app.editor.destructionSource || id === 'destruction'));
  const form6Required = app.editor?.firstGeneration
    || (app.editor && form.elements.namedItem('source_kind').value === 'workbook' && !app.editor.existing);
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
      if (field.key === 'site_location' && app.editor.firstGeneration && app.editor.sourceLocation) {
        label.append(element('small', '', 'Filled from Excel. Changes apply to this PDF only.'));
      }
      label.append(input);
      container.append(label);
    }
  }
  updateAnnexureDisplay();
  updateDetailsButtons();
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
    ? 'Plant is fixed for an issued certificate. Copy into a new certificate to use another plant.' : '';
  if (existingIssuedManual && (manualRecord?.plant || app.editor.sourcePlant)) {
    form.elements.namedItem('plant').value = manualRecord?.plant || app.editor.sourcePlant;
    app.editor.issuedCertificate = true;
  }
  if (existingWorkbook) {
    form.elements.namedItem('case_id').value = app.editor.sourceSn;
    form.elements.namedItem('plant').value = app.editor.sourcePlant;
  }
  form.elements.namedItem('source_kind').disabled = Boolean(existingWorkbook || kind === 'destruction');
  $('#record-options').hidden = Boolean(existingWorkbook);
  $('#edit-scope-hint').textContent = existingWorkbook
    ? 'Excel still records the certificate number, date and Generated status. Cancel does not save changes.'
    : 'A new Excel row is saved only when the PDF is created successfully.';
  form.elements.namedItem('case_id').required = !manual;
  form.elements.namedItem('form6').required = !manual && (!app.editor?.existing || Boolean(app.editor?.firstGeneration));
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
  form.elements.namedItem('form6').closest('.field').hidden = manual && kind !== 'standard';
  form.elements.namedItem('reference').closest('.field').hidden = manual && kind === 'siemens';
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
    : 'Filled in automatically when your Excel Serial Numbers are numeric. You can change it.';
  suggestWorkbookId();
  updatePrintedWeight();
  setEditorMode(app.editor?.editingDetails ?? true);
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
  $('#case-id-hint').textContent = suggestion ? 'Next Serial Number filled in. You can change it. Excel is checked again before saving.'
    : 'Enter a unique Serial Number for this new Excel record.';
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
  return Boolean(app.editor?.existing && !app.editor.firstGeneration && app.editor.template !== 'destruction');
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
  const input = element(source.tagName?.toLowerCase() === 'textarea' ? 'textarea' : 'input', 'inline-fact-input');
  input.type = source.type || 'text'; input.inputMode = source.inputMode;
  if (source.maxLength >= 0) input.maxLength = source.maxLength;
  input.required = Boolean(source.required);
  input.placeholder = source.placeholder || '';
  input.value = initialValue ?? source.value;
  input.setAttribute('aria-label', label);
  // The original form controls remain the only named fields. Opening a pencil
  // does not change the payload, numeric precision, or source data.
  const update = event => { event.stopPropagation?.(); applyInlineValue(key, input.value); };
  input.addEventListener('input', update); input.addEventListener('change', update);
  input.addEventListener('keydown', event => {
    if (event.key === 'Enter' && input.type !== 'textarea') { event.preventDefault(); finishInlineEdit(); }
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
    const grid = element('div', 'inline-material-grid');
    for (const [key, label] of materialFields) {
      const field = element('label', `field inline-material-${key}`); field.append(element('span', '', label));
      const input = element('input'); input.type = 'text'; input.inputMode = key === 'item' ? 'text' : 'decimal';
      input.maxLength = key === 'item' ? 500 : 100; input.value = item[key];
      input.setAttribute('aria-label', `${label} for material ${index + 1}`);
      const update = event => {
        event.stopPropagation?.(); app.materialRows[index].inputs[key].value = input.value;
        app.editor.touched.add(index === 0 ? key : materialInputName(index, key)); app.editor.touched.add('items');
        updateMaterialWeightTotal();
        const total = materialWeightTotal();
        app.editor.inlineEdit.total.textContent = total.state === 'complete' ? `Material total: ${readableNumber(total.total)} kg` : 'Check each material weight.';
        invalidatePreview(); schedulePreview();
      };
      input.addEventListener('input', update); input.addEventListener('change', update);
      field.append(input); grid.append(field);
    }
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
    if (input) editor.append(input);
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
  const needsStart = !fixed && (numbering ? numbering.needs_start : !knownCod);
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
    ['COD date', codDate ? dateLabel(codDate) : 'Set automatically'],
    ['Reference number', value('reference') || 'Not entered']];
  if (value('certificate_type') === 'standard') rows.push(['Form 6', value('form6') || 'Not entered']);
  rows.push(['City', value('location') || 'Not entered'], ['Received', dateLabel(value('inward_date'))],
    ['Plant', labelPlant(value('plant'))], ['Serial Number', value('case_id') || 'Not entered'],
    ['Certificate Weight', weightText ? `${weightText} · ${source}` : 'Weight missing in Excel']);
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
  if (attachment) rows.push(['Attached list', app.editor.annexures[attachment]?.name || 'PDF list attached']);
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
      : 'Changes apply to this PDF.')
    : '';
  $('#editor-summary-note').hidden = !$('#editor-summary-note').textContent;
  $('#change-details').hidden = true;
  $('#finish-editing').hidden = true;
  $('#edit-pdf-details').hidden = !(first || usesSummaryEditor());
  $('#edit-pdf-details').textContent = editing ? 'Done editing' : '✎ Edit PDF details';
  $('#edit-pdf-details').setAttribute('aria-pressed', String(Boolean(editing)));
  $('#editor-mode-title').textContent = first ? (editing ? 'Edit PDF details' : 'Check details from Excel') : app.editor?.existing ? 'Edit certificate' : 'Certificate details';
  $('#details-description').textContent = 'Check the preview, then save.';
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
  if (app.editor?.template === 'destruction') {
    $('#save-details').textContent = app.editor.issuedCertificate ? 'Save corrected certificate' : 'Create data destruction PDF';
    $('#details-description').textContent = app.editor.destructionSource
      ? `Additional certificate for ${labelPlant(app.editor.destructionSource.plant)} inward ${app.editor.destructionSource.case_id}. Check the actual destruction details and both preview pages.`
      : 'Enter the company and actual destruction details, then check both preview pages. An inward entry is optional.';
  }
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
  for (const id of ['save-details', 'save-draft', 'preview-details', 'pick-annexure', 'remove-annexure', 'keep-working-draft']) $(`#${id}`).disabled = unavailable;
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
  $('#preview-warnings-title').hidden = true;
  $('#preview-warnings').replaceChildren();
  $('#preview-status').textContent = 'Preview updates automatically.';
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
}

function schedulePreview() {
  clearTimeout(app.previewTimer);
  const editor = app.editor;
  if (!editor || editor.identityIssue || !$('#details-modal').open || !app.templatesLoaded) return;
  app.previewTimer = setTimeout(async () => {
    if (app.editor !== editor || !$('#details-modal').open || app.detailsBusy || app.generationBusy || app.pendingConfirmation) return;
    if (app.previewLoading) { schedulePreview(); return; }
    const form = $('#details-form');
    if (!form.checkValidity() || !form.elements.namedItem('supplier').value.trim()) return;
    const {weight} = printedWeight();
    if (editor.template !== 'destruction' && !String(weight).trim()) return;
    await previewDetails(); // Read-only: never allocates COD or writes an Excel row.
  }, 1000);
}

async function openDetails(row = null, copyAsNew = false, skipRecovery = false, startEditing = false) {
  if (app.loadingWorkspace || app.generationBusy || app.backupBusy || app.completionBusy
      || app.clientBusy || app.settingsBusy || app.stopping) return;
  if ($('#details-modal').open || app.detailsBusy) return;
  if (!skipRecovery && offerDraftRecovery(row, copyAsNew, true, startEditing)) return;
  const sourceRow = row;
  const issued = row && (row.has_issued_certificate || ['generated', 'generated_manual', 'completed'].includes(stateOf(row)));
  if (row?.last_certificate_details && !copyAsNew && !recordIdentityIssue(row)) {
    row = { ...row, ...row.last_certificate_details };
  }
  const form = $('#details-form');
  form.reset();
  const rowItems = row && Array.isArray(row.items) ? row.items : [];
  renderMaterialRows(rowItems.length ? rowItems : [{ item: row?.item ?? '', quantity: row?.quantity ?? '',
    per_piece_weight: row?.per_piece_weight ?? '', weight: row?.weight ?? '' }]);
  for (const disclosure of form.querySelectorAll('details')) disclosure.open = false;
  const id = templateId(row || {});
  app.editor = { existing: Boolean(row) && !copyAsNew, recordId: copyAsNew ? null : row?.record_id || null, template: id,
    destructionSource: copyAsNew ? null : row?.destruction_source || row?.raw_details?.destruction_source || null,
    editingDetails: Boolean(startEditing || !row || copyAsNew || (row && !isManual(row) && !issued && !row.form6)),
    pdfEditUsed: Boolean(startEditing && row && !isManual(row) && !copyAsNew && !issued),
    firstGeneration: Boolean(row && !isManual(row) && !copyAsNew && !issued), sourceLocation: sourceRow?.location || '',
    issuanceBlocked: Boolean(recordIdentityIssue(sourceRow) || (!issued && sourceRow?.issuance_blocked) || app.state.scan_complete === false),
    identityIssue: copyAsNew ? '' : recordIdentityIssue(sourceRow),
    sourceSheet: sourceRow?.sheet || '', headerRow: sourceRow?.source_header_row || 1,
    importNotes: sourceRow?.import_notes || [], columnMapping: sourceRow?.column_mapping || {},
    sourceSn: row && !isManual(row) && !copyAsNew ? String(row.sn) : null, sourcePlant: row?.plant || null,
    excelCertificateWeight: !copyAsNew && !isManual(row || {}) ? sourceRow?.certificate_weight ?? '' : '',
    issuedCertificate: !copyAsNew && Boolean(issued),
    fields: { [id]: { ...(row?.certificate_fields || {}) } }, annexures: {}, annexureId: row?.certificate_fields?._annexure || null,
    touched: new Set(), sourceExplicit: false, rawNumeric: {}, rawDetails: isManual(row || {}) ? { ...(row.raw_details || {}) } : {}, workspaceRevision: app.workspaceRevision ?? app.state.workspace_revision ?? '' };
  form.elements.namedItem('certificate_type').value = id;
  form.elements.namedItem('source_kind').value = row ? (isManual(row) ? 'manual' : 'workbook') : 'workbook';
  form.elements.namedItem('edit_scope').value = row && !copyAsNew && !isManual(row) ? 'certificate_only' : 'update_excel';
  invalidatePreview();
  $('#details-title').textContent = copyAsNew ? 'Copy into a new certificate'
    : row ? `${issued ? 'Edit certificate for' : 'Create certificate for'} ${row.supplier || `Client ${row.sn}`}` : 'Create a certificate';
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
    const savedWeight = row.last_certificate_details?.certificate_weight;
    if (savedWeight && String(savedWeight) !== String(sourceRow?.certificate_weight ?? '')) {
      form.elements.namedItem('use_certificate_weight').checked = true;
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
  app.editor.baseline = detailsBody(form);
  app.editor.excelBaseline = sourceRow && !isManual(sourceRow) ? workbookRowPayload(sourceRow) : null;
  renderChangeSummary();
  $('#working-draft-status').textContent = app.editor.firstGeneration
    ? (!form.elements.namedItem('form6').value.trim()
      ? 'The Form 6 number will be saved to Excel when you create the certificate.'
      : 'Certificate choices are saved on this computer. Excel details stay unchanged.')
    : 'Edits are saved on this computer. Choose whether to update Excel when saving corrections.';
  $('#details-error').hidden = true;
  $('#details-modal').showModal();
  setEditorMode(app.editor.editingDetails);
  $('#record-options').open = !row || copyAsNew;
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
    } else if (row && !copyAsNew && !changed && app.previewRevision === initialPreviewRevision
        && !app.generationBusy && !app.pendingConfirmation) await previewDetails();
    else if (!row || copyAsNew || changed) schedulePreview();
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
  if (key === 'certificate_fields.device_receipts') {
    try {
      return JSON.parse(value || '[]').filter(row => row.date || row.count)
        .map(row => `${dateLabel(row.date)}: ${row.count || 'count missing'} devices`).join('; ') || '(no device receipts)';
    } catch { return 'Device receipts need checking'; }
  }
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
  return fieldDifferences(app.editor.baseline, current).length > 0 || current.edit_scope !== app.editor.baseline.edit_scope;
}

function scopeDescription(payload) {
  if (payload.source_kind === 'manual') return 'Saves the PDF and updates local tracking. No Excel row is changed.';
  const plant = labelPlant(payload.plant);
  if (app.editor?.firstGeneration) return `Excel details stay unchanged. ${plant} records the COD, certificate date and Generated status automatically.`;
  if (app.editor?.existing && app.editor.sourceSn != null) return 'When saving corrections, choose whether to update this inward in Excel too.';
  if (payload.edit_scope === 'certificate_only') return `Edits appear on the PDF only. ${plant} Excel still records the certificate number, date and Generated status.`;
  return app.editor?.sourceSn != null
    ? `Creating the PDF also updates this record in ${plant} Excel.`
    : `Creates the PDF and adds a new record to ${plant} Excel.`;
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
    const display = value => key === '_annexure' && value ? app.editor.annexures[value]?.name || value : changeDisplay(key, value);
    const entry = element('div', 'field-change');
    const values = element('dd');
    values.append(element('span', 'change-before', `Before: ${display(change.before)}`), element('span', 'change-after', `After: ${display(change.after)}`));
    entry.append(element('dt', '', label), values); list.append(entry);
  }
  $('#change-empty').hidden = changes.length > 0;
  $('#changes-disclosure').hidden = Boolean(app.editor.firstGeneration) || usesSummaryEditor() || changes.length === 0;
  renderEditorSummary();
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
    $('#working-draft-status').textContent = 'The browser could not save your working draft. Keep this form open to avoid losing edits.';
    $('#draft-options').hidden = false;
    $('#keep-working-draft').hidden = false;
    $('#keep-working-draft').textContent = 'Retry saving edits';
    return false;
  }
}

function offerDraftRecovery(row = null, copyAsNew = false, resumeEditor = false, startEditing = false) {
  const draft = readWorkingDraft();
  if (!draft) return false;
  app.draftRecovery = { draft, row, copyAsNew, resumeEditor, startEditing };
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
    renderMaterialRows(materialValuesFromFormValues(saved.values));
    app.editor = { ...saved.editor, inlineEdit:null, summaryHosts:{}, previewFacts:null, touched: new Set(saved.editor.touched) };
    const sourceRow = app.state.rows.find(row => !isManual(row)
      && row.plant === app.editor.sourcePlant && String(row.sn) === String(app.editor.sourceSn));
    app.editor.pdfEditUsed = Boolean(saved.editor.pdfEditUsed);
    if (sourceRow) {
      app.editor.issuedCertificate = Boolean(sourceRow.has_issued_certificate || ['generated', 'generated_manual', 'completed'].includes(stateOf(sourceRow)));
      app.editor.firstGeneration = !app.editor.issuedCertificate;
      app.editor.sourceLocation = sourceRow.location || '';
      app.editor.identityIssue = recordIdentityIssue(sourceRow);
    }
    const workbookDraft = saved.values.source_kind === 'workbook';
    const missingSource = workbookDraft && app.editor.sourceSn != null && !sourceRow;
    if (missingSource) app.editor.identityIssue = 'This inward could not be found. Refresh Excel and reopen the record.';
    app.editor.issuanceBlocked = Boolean(workbookDraft && (missingSource || app.state.scan_complete === false
      || app.editor.identityIssue || (app.editor.firstGeneration && sourceRow?.issuance_blocked)));
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
    if (sourceRow && app.editor.firstGeneration && !app.editor.pdfEditUsed) {
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
    updateSourceControls(); renderTemplateEditor(); invalidatePreview(); renderChangeSummary();
    $('#working-draft-status').textContent = 'Working draft restored. Review the changes and preview before creating.';
    $('#details-error').hidden = true;
    $('#draft-recovery-modal').close(); $('#details-modal').showModal();
    setEditorMode(true);
    updateDetailsButtons();
    schedulePreview();
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

function finishDraftClose(discard) {
  if (!discard) {
    app.afterDraftClose = null;
    $('#unsaved-modal').close();
    return;
  }
  if (!clearWorkingDraft()) return;
  const continuation = app.afterDraftClose;
  app.afterDraftClose = null;
  $('#unsaved-modal').close(); $('#details-modal').close();
  if (continuation) continuation();
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
  $('#success-title').textContent = certificates.length === 1 ? 'Certificate generated' : 'Certificates generated';
  $('#success-description').textContent = certificates.length
    ? `${certificates.length} PDF${certificates.length === 1 ? '' : 's'} saved. Check the actual files below before sending.`
    : `${response.created} PDF${response.created === 1 ? '' : 's'} generated. Open the folder to check the saved files.`;
  const list = $('#created-certificates'); list.replaceChildren();
  for (const certificate of certificates) {
    const entry = element('li', 'saved-file');
    const copy = element('div', 'saved-file-copy');
    const filename = element('a', 'saved-pdf-name', certificate.filename);
    copy.append(element('strong', '', certificate.supplier || 'Certificate'),
      element('span', '', `${labelPlant(certificate.plant)} · COD ${certificate.certificate_no || 'number not provided'}`), filename);
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
    info.addEventListener('click', () => showSavedFile({path:`Certificates/${certificate.relative_path || certificate.filename}`,present:true,cods:[certificate.certificate_no].filter(Boolean),references:[],
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
    if (key === 'plant' || key === 'source_kind' || key === 'edit_scope' || key === 'use_certificate_weight' || key === 'starting_cod') continue;
    details[key] = String(value).trim();
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
  return { plant: data.get('plant') || form.elements.namedItem('plant').value, details, source_kind: source,
    ...(app.editor.needsStartingCod && form.elements.namedItem('starting_cod').value.trim()
      ? {starting_cod:form.elements.namedItem('starting_cod').value.trim()} : {}),
    ...(app.editor?.firstGeneration && app.editor.pdfEditUsed ? {pdf_only_edits:true} : {}),
    ...(source === 'workbook' && app.editor.sourceSn == null ? {create_new:true} : {}),
    use_certificate_weight: source === 'workbook' && Boolean(form.elements.namedItem('use_certificate_weight')?.checked),
    edit_scope: app.editor.firstGeneration ? 'certificate_only' : data.get('edit_scope') || 'certificate_only',
    ...(app.editor.workspaceRevision != null ? { expected_revision: app.editor.workspaceRevision } : {}),
    ...(source === 'workbook' && app.editor.sourceSn != null ? { source_sn: app.editor.sourceSn, source_plant: app.editor.sourcePlant } : {}),
    record_id: source === 'manual' ? app.editor.recordId : null };
}

async function saveDetails(event) {
  event.preventDefault();
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
  if (payload.source_kind === 'workbook' && app.editor.existing && !app.editor.firstGeneration
      && fieldDifferences(app.editor.baseline, payload).length) {
    app.excelSavePayload = {payload};
    const list = $('#excel-save-change-list'); list.replaceChildren();
    for (const change of fieldDifferences(app.editor.excelBaseline || app.editor.baseline, payload)) {
      const key = change.key.replace(/^certificate_fields\./, '');
      const spec = app.templates.find(template => template.id === app.editor.template)?.fields.find(field => field.key === key);
      const entry = element('div', 'field-change');
      entry.append(element('dt', '', fieldLabels[key] || spec?.label || key.replace(/_/g, ' ')),
        element('dd', '', `Excel: ${changeDisplay(key, change.before)} → Certificate: ${changeDisplay(key, change.after)}`));
      list.append(entry);
    }
    $('#excel-save-modal').showModal();
    return;
  }
  await submitGeneration(payload, false, payload.source_kind === 'manual' ? '/api/manual/generate' : '/api/create', true);
}

function workbookRowPayload(row) {
  const details = {};
  for (const key of ['supplier', 'location', 'form6', 'reference', 'certificate_type']) details[key] = String(row[key] ?? (key === 'certificate_type' ? 'standard' : ''));
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
  await submitGeneration({...next.payload, edit_scope:updateExcel ? 'update_excel' : 'certificate_only'}, false, '/api/create', true);
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

async function previewDetails() {
  const form = $('#details-form');
  if (app.editor?.needsStartingCod && !form.elements.namedItem('starting_cod').value.trim()) {
    $('#preview-status').textContent = 'Enter the starting COD number to preview.';
    return;
  }
  if (app.editor?.identityIssue) { updateDetailsButtons(); return; }
  if (app.detailsBusy || app.previewLoading || !app.templatesLoaded || !validateEditor()) return;
  const payload = detailsBody(form);
  invalidatePreview();
  const revision = app.previewRevision;
  app.previewLoading = true;
  $('#preview-details').disabled = true;
  $('#preview-status').textContent = 'Preparing your PDF preview…';
  try {
    const response = await api('/api/preview', payload);
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
        if (candidate.origin !== window.location.origin || !/^\/previews\/[a-f0-9]{32}\.png$/.test(candidate.pathname) || candidate.search || candidate.hash) continue;
        const page = element('img'); page.src = candidate.href;
        page.alt = `Device receipt details, page ${index + 2}`; detailPages.append(page);
      } catch { /* The full PDF remains available. */ }
    }
    $('#preview-status').textContent = imageUrl ? 'Preview ready' : 'PDF ready. Open the full preview.';
    $('#preview-number').textContent = `COD ${response.certificate_no || 'Not available'}${response.certificate_date ? ` · ${dateLabel(response.certificate_date)}` : ''}`;
    $('#preview-number').hidden = false;
    const warnings = Array.isArray(response.warnings) ? response.warnings.filter(value => !isRoutineNotice(value)) : [];
    renderWarningList($('#preview-warnings'), warnings);
    $('#preview-warnings').hidden = warnings.length === 0;
    $('#preview-warnings-title').hidden = warnings.length === 0;
    $('#preview-warnings-title').textContent = 'Check before creating';
  } catch (error) {
    if (revision === app.previewRevision && $('#details-modal').open) {
      $('#preview-status').textContent = 'Preview unavailable';
      showDetailsError(error.message);
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
  if (action === 'destruction') {
    if (app.detailsBusy) return;
    setDetailsBusy(true);
    try {
      const response = await api('/api/destruction/request', {plant: row.plant, sn: row.sn, expected_revision: app.workspaceRevision});
      if (response.state) setState(response.state);
      const request = app.state.rows.find(entry => entry.record_id === response.record_id);
      setDetailsBusy(false);
      if (request) await openDetails(request, false, false, true);
      else toast('Request saved. Refresh to open it.', true);
    } catch (error) { toast(error.message, true); }
    finally { setDetailsBusy(false); }
    return;
  }
  if (action === 'copy') { openDetails(row, true); return; }
  if (action === 'details') { openDetails(row); return; }
  if (action === 'edit-details') { openDetails(row, false, false, true); return; }
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
  return {
    workbook_paths: [form.elements.namedItem('plant1_path').value.trim(), form.elements.namedItem('plant2_path').value.trim()],
    reminders_enabled: Boolean(form.elements.namedItem('reminders_enabled').checked),
    reminder_time: form.elements.namedItem('reminder_time').value,
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

async function clearOldHistory() {
  if (app.backupBusy || app.generationBusy || app.detailsBusy || app.settingsBusy
      || app.clientBusy || app.completionBusy || app.stopping || $('#details-modal').open) return;
  app.backupBusy = true;
  $('#clear-old-history').disabled = true;
  const status = $('#history-cleanup-status');
  try {
    setBusy(true, 'Checking old history…');
    const endpoint = '/api/history/clear';
    const preview = await api(endpoint, {expected_revision:app.workspaceRevision ?? app.state.workspace_revision ?? ''});
    setBusy(false);
    const counts = preview.cleanup;
    if (!preview.requires_confirmation || !counts) throw new Error('Refresh the app before clearing old history.');
    const approved = await confirmAction({title:'Clear old history?',
      message:`Remove ${counts.history_removed} old certificate records, ${counts.pdf_links_removed} deleted-PDF links and ${counts.local_records_removed} old local records? Current Excel files, PDFs and their active records stay. A recovery copy is kept.`,
      confirmLabel:'Clear old history'});
    if (!approved) return;
    setBusy(true, 'Clearing old history…');
    const result = await api(endpoint, {confirmed:true, review_token:preview.review_token,
      expected_revision:preview.workspace_revision});
    if (!result.cleared) throw new Error('The history changed. Review it again before clearing.');
    clearWorkingDraft();
    if (result.state) setState(result.state);
    $('#history-list').replaceChildren();
    $('#history-detail').replaceChildren();
    status.textContent = result.message;
    status.title = result.recovery_folder || '';
    toast(result.message);
  } catch (error) { status.textContent = error.message; toast(error.message, true); }
  finally { app.backupBusy = false; $('#clear-old-history').disabled = false; setBusy(false); }
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
  $('#view-files').addEventListener('click', () => setWorkspaceView('files'));
  $('#open-help').addEventListener('click', () => $('#help-modal').showModal());
  $('#reset-record-filters').addEventListener('click', resetRecordFilters);
  $('#dismiss-toast').addEventListener('click', dismissToast);
  $('#saved-file-excel').addEventListener('click', () => {
    const button = $('#saved-file-excel');
    if (!button.hidden && button.dataset.plant && button.dataset.sn) openClientInExcel(button.dataset.plant, button.dataset.sn);
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
  for (const id of ['saved-file-search', 'saved-plant-filter', 'saved-file-sort']) {
    $(`#${id}`).addEventListener(id === 'saved-file-search' ? 'input' : 'change', () => {
      app.savedFilesPage = 1;
      if (id === 'saved-plant-filter' && ['all', 'plant1', 'plant2'].includes($('#saved-plant-filter').value)) {
        $('#plant-filter').value = $('#saved-plant-filter').value; app.page = 1; render();
      }
      renderFolderInventory();
    });
  }
  $('#clear-file-search').addEventListener('click', () => {
    $('#saved-file-search').value = ''; $('#saved-plant-filter').value = 'all';
    $('#plant-filter').value = 'all'; app.savedFilesPage = 1; app.page = 1; render(); renderFolderInventory();
  });
  document.querySelectorAll('[data-quick-status]').forEach(button => button.addEventListener('click', () => {
    $('#status-filter').value = button.dataset.quickStatus;
    app.page = 1; app.expanded = null; render();
  }));
  $('#confirm-action').addEventListener('click', () => settleActionConfirmation(true));
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
    if (next?.resumeEditor) openDetails(next.row || null, next.copyAsNew || false, true, next.startEditing || false);
  });
  $('#success-open-folder').addEventListener('click', async () => {
    if (app.lastGeneratedCertificatePath) await openCertificateLocally(app.lastGeneratedCertificatePath, false);
    else try { await api('/api/open-folder', {}); } catch (error) { toast(error.message, true); }
  });
  $('#saved-file-open').addEventListener('click', () => openCertificateLocally($('#saved-file-open').dataset.relativePath));
  $('#create-backup').addEventListener('click', () => backupAction('create'));
  $('#clear-old-history').addEventListener('click', clearOldHistory);
  $('#pick-backup-destination').addEventListener('click', () => backupAction('pick-destination'));
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
  $('#new-button').addEventListener('click', () => openDetails());
  $('#destruction-pending').addEventListener('click', () => {
    $('#plant-filter').value = 'all'; $('#search').value = '';
    $('#template-filter').value = 'destruction'; $('#status-filter').value = 'needs_attention';
    app.page = 1; render();
  });
  $('#open-bulk').addEventListener('click', openBulkGeneration);
  $('#bulk-start-cod').addEventListener('input', updateBulkCodHint);
  $('#generate-all').addEventListener('click', startBulkGeneration);
  $('#bulk-generation-modal').addEventListener('close', () => { if (!$('#bulk-generation-modal').open) app.bulkSelection = null; });
  $('#clear-filters').addEventListener('click', () => {
    $('#status-filter').value = 'all'; $('#template-filter').value = 'all';
    app.page = 1; app.expanded = null; render();
  });
  const restartList = () => { app.page = 1; app.expanded = null; render(); };
  document.querySelectorAll('[data-filter-status]').forEach(button => button.addEventListener('click', () => {
    $('#status-filter').value = button.dataset.filterStatus;
    restartList();
  }));
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
      $(`#queue-body [data-expand="${index}"]`)?.focus();
    }
  });
  document.querySelectorAll('[data-view]').forEach(button => button.addEventListener('click', () => setView(button.dataset.view)));
  $('#details-form').addEventListener('submit', saveDetails);
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
  $('#excel-save-modal').addEventListener('close', () => { app.excelSavePayload = null; });
  $('#change-details').addEventListener('click', () => {
    setEditorMode(true); $('#details-form').elements.namedItem('supplier').focus();
  });
  $('#finish-editing').addEventListener('click', finishEditing);
  $('#add-material').addEventListener('click', addMaterialRow);
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
    schedulePreview();
  });
  $('#details-form').addEventListener('input', event => {
    if (event.target.name) app.editor?.touched.add(event.target.name);
    if (event.target.name === 'starting_cod') updateSingleCodControls();
    if (event.target.name === 'certificate_weight') applyCertificateWeight(event.target.value);
    updateMaterialWeightTotal();
    invalidatePreview();
    schedulePreview();
  });
  $('#details-form').addEventListener('change', event => {
    const form = event.currentTarget;
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
    if (dialog.id === 'add-client-modal' && app.clientBusy) return;
    if (dialog.id === 'details-modal') requestDetailsClose();
    else if (dialog.id === 'settings-modal') requestSettingsClose();
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
      if (document.querySelector('dialog[open]')) return;
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
  } finally { app.loadingWorkspace = false; $('#new-button').disabled = false; $('#add-client-button').disabled = false; summarize(); }
  window.setInterval(syncWorkbookChanges, 15000);
  pollReminderReview();
  window.setInterval(pollReminderReview, 2000);
  window.addEventListener('focus', pollReminderReview);
  document.addEventListener('visibilitychange', pollReminderReview);
  window.addEventListener('focus', syncWorkbookChanges);
  document.addEventListener('visibilitychange', syncWorkbookChanges);
}

init();
