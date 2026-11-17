// Pure UI regression tests; no browser, real workbook, or PDF is modified.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

class Node {
  constructor(tag) {
    this.tag = tag; this.children = []; this.dataset = {}; this.style = {}; this.attributes = {}; this.textContent = ''; this.value = '';
    this.classes = new Set(); this.classList = {
      toggle: (name, force) => { const add = force === undefined ? !this.classes.has(name) : force; if (add) this.classes.add(name); else this.classes.delete(name); return add; },
      add: (...names) => names.forEach(name => this.classes.add(name)),
      remove: (...names) => names.forEach(name => this.classes.delete(name)),
      contains: name => this.classes.has(name),
    };
    this.listeners = {};
  }
  append(...nodes) { for (const node of nodes) { if (node.parentElement) node.parentElement.children = node.parentElement.children.filter(child => child !== node); node.parentElement = this; this.children.push(node); } }
  insertBefore(node, before) { node.parentElement = this; this.children.splice(this.children.indexOf(before), 0, node); }
  replaceChildren(...nodes) { this.children = nodes; }
  contains(node) { return this === node || this.children.some(child => child.contains(node)); }
  setAttribute(key, value) { this.attributes[key] = value; }
  getAttribute(key) { return this.attributes[key] ?? null; }
  querySelectorAll() { return []; }
  addEventListener(event, fn) { (this.listeners[event] ||= []).push(fn); }
  dispatch(event, detail = {}) { for (const fn of this.listeners[event] || []) fn({target: this, currentTarget: this, ...detail}); }
  removeAttribute(key) { delete this.attributes[key]; if (key === 'src' || key === 'href') delete this[key]; }
  toggleAttribute(key, value) { if (value) this.attributes[key] = ''; else delete this.attributes[key]; }
  closest() { return null; }
  getBoundingClientRect() { return {left:10,right:100,top:10,bottom:100}; }
  focus() { this.focused = true; }
  setCustomValidity(message) { this.validationMessage = message; }
  checkValidity() { return !this.validationMessage; }
  reportValidity() { return this.checkValidity(); }
  showModal() { this.open = true; }
  close() { this.open = false; this.dispatch('close'); }
  showPopover() { this.popoverOpen = true; }
  hidePopover() { this.popoverOpen = false; }
}
function all(node, predicate) {
  return [ ...(predicate(node) ? [node] : []), ...node.children.flatMap(child => all(child, predicate)) ];
}
function text(node) { return [node.textContent, ...node.children.map(text)].join(' '); }
function harness(options = {}) {
  const controls = new Map();
  for (const id of ['#plant-filter', '#status-filter', '#template-filter']) { const select = new Node('select'); select.value = 'all'; select.selectedOptions = [{textContent:'All companies'}]; controls.set(id, select); }
  controls.set('#search', new Node('input'));
  for (const id of ['#folder-inventory', '#folder-inventory-summary', '#folder-inventory-note', '#folder-inventory-warnings', '#folder-inventory-files', '#saved-file-pages', '#saved-file-page-count', '#saved-files-previous', '#saved-files-next', '#show-past-numbers', '#past-numbers-modal', '#past-number-list', '#past-number-count', '#saved-file-name', '#saved-file-facts', '#saved-file-modal']) controls.set(id, new Node('div'));
  controls.set('#past-number-search', new Node('input'));
  controls.set('#template-editor', new Node('div'));
  const html = fs.readFileSync(path.join(__dirname, '../web/index.html'), 'utf8');
  for (const match of html.matchAll(/<([\w-]+)[^>]*\bid="([^"]+)"/g)) {
    if (!controls.has(`#${match[2]}`)) controls.set(`#${match[2]}`, new Node(match[1]));
    controls.get(`#${match[2]}`).id = match[2];
  }
  const settingsForm = controls.get('#settings-form'), settingsInputs = {};
  for (const name of ['plant1_path', 'plant2_path', 'reminders_enabled', 'reminder_time']) {
    const input = new Node('input'); input.name = name;
    input.value = name === 'reminder_time' ? '11:00' : name.endsWith('_path') ? `${name}.xlsx` : '';
    const field = new Node('label'); input.closest = () => field;
    settingsInputs[name] = input; controls.set(`#settings-form [name="${name}"]`, input);
  }
  settingsForm.elements = {namedItem:name => settingsInputs[name]};
  settingsForm.checkValidity = () => true; settingsForm.reportValidity = () => true;
  const closeButtons = [...html.matchAll(/<button[^>]*data-close="([^"]+)"[^>]*>/g)].map(match => { const button = new Node('button'); button.dataset.close = match[1]; return button; });
  const document = { body:new Node('body'), createElement: tag => new Node(tag), querySelector: key => controls.get(key), querySelectorAll: selector => selector === 'dialog' ? [...controls.values()].filter(node => /-modal$/.test(node.id)) : selector === '[data-close]' ? closeButtons : [] };
  document.body.append(controls.get('#toast'));
  let source = fs.readFileSync(path.join(__dirname, '../web/app.js'), 'utf8').replace(/\ninit\(\);\s*$/, '');
  if (options.showRestoreOptions) source = source.replace('const showWorkingDraftRecovery = false;', 'const showWorkingDraftRecovery = true;');
  class FormData {
    constructor(form) { this.values = Object.fromEntries(Object.entries(form.values || {}).filter(([name]) => !form.elements?.namedItem(name)?.disabled)); }
    entries() { return Object.entries(this.values); }
    [Symbol.iterator]() { return this.entries()[Symbol.iterator](); }
    get(key) { return this.values[key]; }
  }
  const storage = new Map(), requests = [], window = new Node('window'), timers = new Map();
  let timerId = 0;
  window.location = {href:'http://127.0.0.1:8000/',origin:'http://127.0.0.1:8000'};
  window.confirm = () => { throw new Error('Native confirmations are unavailable in the in-app browser.'); };
  const context = vm.createContext({document, console, FormData, URL, window, setTimeout: (fn, delay) => {const id = ++timerId; timers.set(id, {fn,delay}); return id;}, clearTimeout:id => timers.delete(id),
    MutationObserver: class { constructor(callback) { this.callback = callback; } observe() {} },
    localStorage: {getItem:key => storage.get(key) || null, setItem:(key,value) => storage.set(key,value), removeItem:key => storage.delete(key)},
    fetch: async (endpoint, init) => {
      requests.push({endpoint, ...init, payload:init.body ? JSON.parse(init.body) : undefined});
      return options.fetch ? options.fetch(endpoint, init) : {ok:true,json:async () => ({workspace_revision:'revision-2'})};
    }});
  const ui = vm.runInContext(`${source}\n({app, stateLabel, statusClass, nextStep, matches, makeActions, makeClientCell, makeDetailRow, renderFolderInventory, simpleWarning, readableNumber, detailsBody, materialWeightTotal, updateMaterialWeightTotal, renderMaterialRows, visiblePage, pastCertificateNumbers, certificateFileTitle, showSavedFile, renderPastNumbers, api, fieldDifferences, renderChangeSummary, editorDirty, saveWorkingDraft, saveDraft, readWorkingDraft, clearWorkingDraft, offerDraftRecovery, restoreWorkingDraft, requestDetailsClose, finishDraftClose, certificateUrl, showGenerationSuccess, submitGeneration, bulkGenerationPayload, openBulkGeneration, startBulkGeneration, renderEditorSummary, setEditorMode, validateEditor, finishEditing, openDetails, beginIssuedCertificateCorrection, renderIssuedChangeReview, openIssuedChangeReview, resolveIssuedExcelChanges, saveDetails, showWarnings, confirmWarnings, pickAnnexure, previewDetails, toggleFullPdfPreview, backupAction, openBackupRestore, restoreBackup, historyDifferences, loadHistory, showHistoryDetail, invalidatePreview, confirmAction, settleActionConfirmation, handleAction, wire, stopApp, summarize})`, context);
  Object.assign(ui, vm.runInContext('({savedFilePlant, savedFileUrl, savedFileMatches, setWorkspaceView, resetRecordFilters, nextWorkbookId, suggestWorkbookId, printedWeight, updatePrintedWeight, schedulePreview, saveSettings, settingsValues, openSettings, requestSettingsClose, returnFromWarnings, toast, makeStatus, reconnectDashboard, saveExcelChoice, renderTemplateEditor})',context));
  Object.assign(ui, vm.runInContext('({startInlineEdit, finishInlineEdit, syncWorkbookChanges, setState, renderWorkbookIssues, refreshState, pollReminderReview})',context));
  Object.assign(ui, vm.runInContext('({openAddClient, renderClientMaterials, clientMaterialValues, saveClient, validateClientSerialNumber, duplicateSerialNumber, compareRows})',context));
  Object.assign(ui, vm.runInContext('({excelTarget, openClientInExcel})',context));
  Object.assign(ui, vm.runInContext('({openInwardRemoval, confirmInwardRemoval, rememberDeletion, undoLastDeletion})',context));
  Object.assign(ui, vm.runInContext('({openEntryDeletion,filterEntryDeletion,reviewEntryDeletion,confirmEntryDeletion})',context));
  Object.assign(ui, vm.runInContext('({changeDisplay})',context));
  Object.assign(ui, vm.runInContext('({saveInwardDetails, openCompanionCertificate, updateDestructionChoice, showExcelSaveChoice})',context));
  Object.assign(ui, vm.runInContext('({noticeText, showDetailsError, warningInfo, renderWarningList, hasWarnings, needsAttention})',context));
  Object.assign(ui, vm.runInContext('({missingFollowup, inwardChanges, renderEditorFollowup, filterBulkClients, updateBulkSelection})',context));
  Object.assign(ui, vm.runInContext('({editorInwardRow, editorDestructionRequest, switchesLinkedCertificate, switchLinkedCertificate})',context));
  Object.assign(ui, vm.runInContext('({updateRecordFilterButton, positionRecordFilters})',context));
  Object.assign(ui, vm.runInContext('({init})',context));
  Object.assign(ui, vm.runInContext('({changeDestructionChoice})',context));
  Object.assign(ui, vm.runInContext('({dateLabel,normalizeDate,configureDateInput,canDeleteEntry,clientEntryPayload})',context));
  return {ui, controls, storage, requests, window, context, closeButtons, timers, settingsInputs};
}

test('Undo deletion stays available in Settings and retries after a failed restore', async () => {
  let attempts = 0;
  const h = harness({fetch: async () => ({
    ok: ++attempts > 1, status: attempts > 1 ? 200 : 400,
    json: async () => attempts > 1 ? {undone:true} : {error:'Close the workbook in Excel and try again.'},
  })});
  h.ui.app.workspaceRevision = 'reviewed-revision';
  h.ui.rememberDeletion({kind:'entries', request_id:'b'.repeat(32)});
  h.ui.rememberDeletion({kind:'inward', request_id:'a'.repeat(32)});
  h.ui.toast('Inward removed.', false, true);
  assert.equal(h.controls.get('#undo-deletion-toast').hidden, false);
  assert.equal(h.controls.get('#undo-last-deletion').hidden, false);
  await h.ui.undoLastDeletion();
  assert.equal(h.controls.get('#undo-last-deletion').hidden, false);
  await h.ui.undoLastDeletion();
  assert.equal(h.requests[1].endpoint, '/api/entries/undo-delete');
  assert.equal(h.requests[1].payload.request_id, 'a'.repeat(32));
  assert.equal(h.controls.get('#undo-last-deletion').hidden, false);
  await h.ui.undoLastDeletion();
  assert.equal(h.requests[2].payload.request_id, 'b'.repeat(32));
  assert.equal(h.controls.get('#undo-last-deletion').hidden, true);
  assert.equal(h.storage.has('ecoreco-last-deletion'), false);
});

test('Saved certificates filter by the same certificate types as Companies', () => {
  const h=harness();h.controls.get('#saved-plant-filter').value='all';
  const entries=['standard','destruction','siemens','sita'].map(certificate_type=>({certificate_type,path:`Certificates/0001_${certificate_type}.pdf`,cods:['0001']}));
  h.controls.get('#saved-template-filter').value='all';assert.equal(entries.filter(h.ui.savedFileMatches).length,4);
  for(const type of ['standard','destruction','siemens','sita']){
    h.controls.get('#saved-template-filter').value=type;
    assert.deepEqual(entries.filter(h.ui.savedFileMatches).map(entry=>entry.certificate_type),[type]);
  }
  const html=fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  const options=id=>html.match(new RegExp(`<select id="${id}"[^>]*>([\\s\\S]*?)<\\/select>`))[1];
  assert.equal(options('saved-template-filter'),options('template-filter'));
});

test('Saved filters have their own removable chips, count and toolbar reset', () => {
  const h = harness(); h.ui.wire();
  const plant = h.controls.get('#saved-plant-filter');
  const template = h.controls.get('#saved-template-filter');
  const clear = h.controls.get('#clear-file-search');
  const chips = h.controls.get('#saved-filter-chips');
  h.controls.get('#plant-filter').value = 'plant1';
  h.ui.setWorkspaceView('files');
  plant.value = 'plant2'; plant.dispatch('change');
  template.value = 'destruction'; template.dispatch('change');
  h.controls.get('#saved-file-sort').value = 'company';
  assert.equal(h.controls.get('#saved-filter-count').textContent, '2');
  assert.equal(chips.children[0].attributes['aria-label'], 'Remove filter Plant 2');
  chips.children[0].dispatch('click');
  assert.equal(plant.value, 'all');
  assert.equal(template.value, 'destruction');
  assert.equal(h.controls.get('#saved-filter-count').textContent, '1');
  h.ui.setWorkspaceView('records');
  h.ui.setWorkspaceView('files');
  assert.equal(template.value, 'destruction');
  clear.dispatch('click');
  assert.equal(template.value, 'all');
  assert.equal(clear.hidden, true);
  assert.equal(chips.hidden, true);
  assert.equal(h.controls.get('#saved-file-sort').value, 'company');
  assert.equal(h.controls.get('#plant-filter').value, 'plant1');
});

test('Inward search narrows large lists and requires an exact plant and SN selection', () => {
  for (const id of ['upload-inward']) {
    const h = harness();
    uploadForm(h);
    h.controls.get('#external-upload-target').value = 'workbook';
    const picker = vm.runInContext('uploadInwardPicker', h.context);
    const input = h.controls.get(`#${id}`);
    const list = h.controls.get(`#${id}-options`);
    const rows = Array.from({length:350}, (_, index) => ({supplier:`Client ${index}`, plant:'plant1', sn:index}));
    const target = {supplier:'Northstar Electronics', plant:'plant2', sn:'349'};
    rows[349].supplier = target.supplier;
    picker.update([...rows, target]);
    input.dispatch('focus');
    assert.equal(list.children.length, 12);
    input.value = 'northstar 349 plant 2';
    input.dispatch('input');
    assert.equal(list.children.length, 1);
    assert.match(text(list.children[0]), /Northstar Electronics.*Plant 2.*SN 349/);
    assert.equal(picker.selection(), null);
    assert.equal(input.checkValidity(), false);
    input.dispatch('keydown', {key:'ArrowDown', preventDefault() {}});
    assert.equal(input.getAttribute('aria-activedescendant'), list.children[0].id);
    input.dispatch('keydown', {key:'Enter', preventDefault() {}});
    assert.deepEqual({...picker.selection()}, {plant:'plant2', sn:'349'});
    assert.equal(input.checkValidity(), true);
    assert.equal(list.hidden, true);
    input.value = 'unknown company';
    input.dispatch('input');
    assert.equal(picker.selection(), null);
    assert.equal(h.controls.get(`#${id}-results`).textContent, 'No matching inwards.');
  }
});

function uploadForm(h) {
  const form = h.controls.get('#upload-certificate-form');
  const names = ['cod_date', 'certificate_type', ...vm.runInContext('uploadDetailFields', h.context)];
  const fields = Object.fromEntries(names.map(name => [name, new Node('input')]));
  form.elements = {namedItem:name => fields[name]};
  Object.defineProperty(form, 'values', {get:() => Object.fromEntries(Object.entries(fields).map(([name, field]) => [name, field.value]))});
  return {form, fields};
}

test('Attach to a client replaces the matching issued certificate and keeps its COD', async () => {
  const h = harness({fetch:async () => ({ok:false, status:400, json:async () => ({error:'Check the PDF.'})})});
  const mode = h.controls.get('#external-upload-target');
  const input = h.controls.get('#upload-inward');
  const options = h.controls.get('#upload-inward-options');
  const {form, fields} = uploadForm(h);
  const {updateUploadTarget, saveUploadedCertificate} = vm.runInContext('({updateUploadTarget, saveUploadedCertificate})', h.context);
  h.ui.app.workspaceRevision = 'reviewed-revision';
  h.ui.app.state.rows = [
    {supplier:'Client', plant:'plant1', sn:'12', certificate_status:'ready'},
    {supplier:'Client', plant:'plant2', sn:'12', certificate_status:'completed',
      certificate_type:'siemens', certificate_no:'0456', certificate_date:'2026-11-05',
      pdf_relative_path:'Auto Generated/0456_Client.pdf', reference:'Excel reference', location:'Excel site',
      last_certificate_details:{supplier:'Name on certificate', reference:'PDF reference', certificate_weight:'8.075'}},

  ];
  mode.value = 'workbook';
  updateUploadTarget(true);
  assert.equal(h.controls.get('#upload-inward-label').textContent, 'Client');
  input.dispatch('focus');
  input.value = '0456'; input.dispatch('input');
  assert.equal(options.children.length, 1);
  assert.match(text(options.children[0]), /Plant 2.*SN 12.*COD 0456/);
  options.children[0].dispatch('click');
  assert.equal(h.controls.get('#upload-cod').value, '0456');
  assert.equal(fields.cod_date.value, '05/11/26');
  assert.equal(fields.certificate_type.value, 'siemens');
  assert.equal(fields.certificate_type.disabled, true);
  assert.equal(h.controls.get('#upload-company-fields').hidden, false);
  assert.equal(fields.supplier.value, 'Name on certificate');
  assert.equal(fields.reference.value, 'PDF reference');
  assert.equal(fields.certificate_weight.value, '8.075');
  fields.reference.value = '';
  fields.location.value = 'Revised site';
  assert.equal(h.controls.get('#save-uploaded-certificate').textContent, 'Replace PDF');
  const file = {};
  h.controls.get('#external-upload-file').files = [file];
  h.ui.app.uploadedFile = {token:'uploaded', file, cod:'0456'};
  await saveUploadedCertificate({currentTarget:form, preventDefault() {}});
  assert.deepEqual(h.requests[0].payload, {
    token:'uploaded', details:{cod_date:'05/11/26'},
    inward:null, replacement:{plant:'plant2', sn:'12'},
    certificate_edits:{supplier:'Name on certificate', location:'Revised site', inward_date:'',
      reference:'', form6:'', certificate_weight:'8.075'},
    expected_revision:'reviewed-revision',
  });
  assert.equal(h.controls.get('#upload-certificate-error').textContent, 'Check the PDF.');
});

test('Upload PDF prefills the chosen inward and preserves edited fields when data refreshes', async () => {
  const h = harness({fetch:async () => ({ok:false, status:400, json:async () => ({error:'Check the PDF.'})})});
  const {form, fields} = uploadForm(h);
  const mode = h.controls.get('#external-upload-target');
  const input = h.controls.get('#upload-inward');
  const options = h.controls.get('#upload-inward-options');
  const {updateUploadTarget, saveUploadedCertificate} = vm.runInContext('({updateUploadTarget, saveUploadedCertificate})', h.context);
  h.ui.app.workspaceRevision = 'reviewed-revision';
  const inward = {supplier:'Client', plant:'plant1', sn:'12', certificate_status:'ready', location:'Pune',
    reference:'REF-12', form6:'F6-12', inward_date:'2026-11-03', certificate_type:'sita', items:[{item:'Laptops', quantity:3, weight:4.5}]};
  h.ui.app.state.rows = [inward, {supplier:'Second', plant:'plant2', sn:'12', certificate_status:'ready'}];
  mode.value = 'workbook';
  updateUploadTarget(true);
  assert.equal(h.controls.get('#upload-company-fields').hidden, true);
  input.dispatch('focus');
  options.children[0].dispatch('click');
  assert.equal(fields.supplier.value, 'Client');
  assert.equal(fields.certificate_type.value, 'sita');
  assert.equal(fields.cod_date.value, '');
  assert.equal(fields.inward_date.value, '03/11/26');
  assert.equal(fields.reference.value, 'REF-12');
  assert.equal(fields.form6.value, 'F6-12');
  assert.match(text(h.controls.get('#upload-materials-list')), /Laptops.*Qty 3.*4.5 kg/);
  fields.reference.value = 'Changed reference';
  fields.cod_date.value = '04/11/26';
  h.ui.app.state.rows[0] = {...inward, location:'Pune - updated'};
  updateUploadTarget();
  assert.equal(fields.location.value, 'Pune - updated');
  assert.equal(fields.reference.value, 'Changed reference');
  assert.equal(fields.cod_date.value, '04/11/26');
  const file = {};
  h.controls.get('#external-upload-file').files = [file];
  h.ui.app.uploadedFile = {token:'uploaded', file, cod:'0456', reference:'PDF reference'};
  await saveUploadedCertificate({currentTarget:form, preventDefault() {}});
  assert.deepEqual(h.requests.at(-1).payload, {
    token:'uploaded', details:{certificate_type:'sita', cod_date:'04/11/26'},
    inward:{plant:'plant1', sn:'12'}, replacement:null,
    certificate_edits:{supplier:'Client', location:'Pune - updated', inward_date:'03/11/26',
      reference:'Changed reference', form6:'F6-12', certificate_weight:''}, expected_revision:'reviewed-revision',
  });
  input.dispatch('focus');
  options.children[1].dispatch('click');
  assert.equal(fields.supplier.value, 'Second');
  assert.equal(fields.reference.value, '');
  assert.equal(fields.cod_date.value, '');
  assert.equal(h.controls.get('#upload-materials').hidden, true);
});

test('Upload PDF displays automatic CODs and leaves assignment to the save request', async () => {
  const h = harness({fetch:async () => ({ok:false, status:400, json:async () => ({error:'Check the PDF.'})})});
  const html = fs.readFileSync(path.join(__dirname, '../web/index.html'), 'utf8');
  const codField = html.match(/<textarea id="upload-cod"[^>]*>/)[0];
  assert.match(codField, /\breadonly\b/);
  assert.doesNotMatch(codField, /\bname=/);
  const form = h.controls.get('#upload-certificate-form');
  form.values = {certificate_type:'standard', cod_date:'05/11/26', supplier:'Client'};
  form.elements = {namedItem:name => ({value:form.values[name]})};
  h.controls.get('#external-upload-target').value = 'local';
  h.ui.app.state.numbering = {next_cod:'0456'};
  h.ui.app.workspaceRevision = 'reviewed-revision';
  const file = {};
  h.controls.get('#external-upload-file').files = [file];
  const {updateUploadCod, saveUploadedCertificate} = vm.runInContext('({updateUploadCod, saveUploadedCertificate})', h.context);
  for (const [printed, shown] of [['0423', '0423'], ['', '']]) {
    h.ui.app.uploadedFile = {token:'uploaded', file, cod:printed};
    updateUploadCod();
    assert.equal(h.controls.get('#upload-cod').value, shown);
    assert.equal(h.controls.get('#upload-cod').readOnly, Boolean(printed));
    if (!printed) assert.equal(h.controls.get('#upload-cod').placeholder, '0456');
    await saveUploadedCertificate({currentTarget:form, preventDefault() {}});
    assert.deepEqual(h.requests.at(-1).payload, {
      token:'uploaded', details:form.values, inward:null, replacement:null,
      expected_revision:'reviewed-revision',
    });
  }
  h.controls.get('#upload-cod').value = '0424';
  await saveUploadedCertificate({currentTarget:form, preventDefault() {}});
  assert.equal(h.requests.at(-1).payload.details.cod_no, '0424');
});

test('App dates display DD/MM/YY and submit the correct ISO day', () => {
  const h=harness(),form=editing(h);
  assert.equal(h.ui.dateLabel('2026-11-02'),'02/11/26');
  assert.equal(h.ui.normalizeDate('02/11/26'),'2026-11-02');
  assert.equal(h.ui.normalizeDate('31/02/26'),'');
  form.values.inward_date='02/11/26';form.values.cod_date='03/11/26';
  assert.equal(h.ui.detailsBody(form).details.inward_date,'2026-11-02');
  assert.equal(h.ui.detailsBody(form).details.cod_date,'2026-11-03');
  const input=new Node('input');input.value='2026-11-02';h.ui.configureDateInput(input);
  assert.equal(input.type,'text');assert.equal(input.value,'02/11/26');
  input.value='31/02/26';input.dispatch('input');assert.equal(input.checkValidity(),false);
  const entry=clientEntry(h);entry.form.values.inward_date='02/11/26';
  assert.equal(h.ui.clientEntryPayload(entry.form).details.inward_date,'2026-11-02');
});

test('Add client ignores outside clicks and retains its fields until explicitly closed', () => {
  const h=harness(),{form}=clientEntry(h);h.ui.wire();h.ui.openAddClient();
  form.values.supplier='Still entering this client';
  const modal=h.controls.get('#add-client-modal');modal.dispatch('click',{clientX:-10,clientY:-10});
  assert.equal(modal.open,true);assert.equal(form.values.supplier,'Still entering this client');
  h.closeButtons.find(button=>button.dataset.close==='add-client-modal').dispatch('click');
  assert.equal(modal.open,false);h.ui.openAddClient();assert.equal(form.values.supplier,'Still entering this client');
});

test('Deletion stays available before generation and remains blocked by recycling or HDD issuance', () => {
  const h=harness(),row={source_kind:'workbook',plant:'plant1',sn:'1',sheet:'Inward Register',row:2,certificate_status:'ready',certificate_no:'0042'};
  assert.equal(h.ui.canDeleteEntry(row),true);
  for(const change of [{has_issued_certificate:true},{hdd_cod_no:'0043'},{pdf_filename:'0042.pdf'},{certificate_status:'generated'},{certificate_status:'completed'}])assert.equal(h.ui.canDeleteEntry({...row,...change}),false);
  h.ui.app.state.rows=[row,{source_kind:'manual',plant:'plant1',record_id:'hdd',certificate_status:'generated_manual',destruction_source:{plant:'plant1',case_id:'1'}}];
  assert.equal(h.ui.canDeleteEntry(row),false);
});

test('Changing an issued certificate type never offers inward deletion', async () => {
  const h=harness(),form=editing(h);
  h.ui.app.editor.existing=true;h.ui.app.editor.issuedCertificate=true;
  h.ui.app.editor.baseline=h.ui.detailsBody(form);h.ui.app.editor.excelBaseline=h.ui.detailsBody(form);
  form.values.certificate_type='sita';h.ui.app.editor.template='sita';h.ui.app.editor.fields.sita={};
  h.ui.showExcelSaveChoice(h.ui.detailsBody(form),'create');
  assert.equal(h.controls.get('#excel-save-no').textContent,'No, keep editing');
  assert.match(h.controls.get('#excel-save-note').textContent,/cannot be deleted/);
  await h.ui.saveExcelChoice(false);assert.equal(h.requests.length,0);
});

const removalRow = {source_kind:'workbook',plant:'plant1',sn:'25',supplier:'ABC Company',sheet:'Inward Register',row:3,certificate_status:'ready'};

test('Startup checks saved Excel changes before enabling actions and loads the dashboard only once', async () => {
  for (const changed of [false,true]) {
    let finishCheck;
    const h=harness({fetch:async endpoint=>{
      if(endpoint==='/api/state')return {ok:true,json:async()=>({state:{rows:[],scan_complete:true},workspace_revision:'initial',settings:{workbook_paths:['plant1.xlsx','plant2.xlsx']}})};
      if(endpoint==='/api/workbook-changes')return new Promise(resolve=>{
        finishCheck=()=>resolve({ok:true,json:async()=>changed
          ? {changed:true,state:{rows:[],scan_complete:true,workspace_revision:'checked'}} : {changed:false}});
      });
      throw new Error(`Unexpected startup request: ${endpoint}`);
    }});
    h.window.setInterval=()=>{};
    const loads=[];h.window.CertificateAnalytics={load:()=>loads.push(h.ui.app.workspaceRevision)};
    vm.runInContext('document.addEventListener=()=>{}; loadTemplates=async()=>true; pollReminderReview=()=>{};',h.context);
    const startup=h.ui.init();await new Promise(resolve=>setImmediate(resolve));
    assert.equal(h.ui.app.loadingWorkspace,true);assert.equal(h.controls.get('#new-button').disabled,true);
    assert.equal(loads.length,0);finishCheck();await startup;
    assert.equal(h.ui.app.loadingWorkspace,false);assert.equal(h.controls.get('#new-button').disabled,false);
    assert.deepEqual(loads,[changed?'checked':'initial']);
    assert.deepEqual(h.requests.map(request=>request.endpoint),['/api/state','/api/workbook-changes']);
  }
});
test('Several warning messages for one certificate never appear as several certificates', () => {
  const h=harness();
  h.ui.showWarnings(['missing supplier name','missing inward date'],{payload:{details:{supplier:'Sample company'}},fromDetails:true});
  assert.equal(h.controls.get('#warning-list').children.length,1);
  assert.match(h.controls.get('#warning-description').textContent,/this certificate/);
  assert.doesNotMatch(h.controls.get('#warning-description').textContent,/2 certificates/);
});
test('Generation offers Yes and No even when there are no warnings', () => {
  const h = harness();
  h.ui.showWarnings([], {payload:{details:{supplier:'Sample company'}}, fromDetails:true});
  assert.equal(h.controls.get('#confirm-warnings').textContent, 'Yes, create');
  assert.equal(h.controls.get('#warning-go-back').textContent, 'No, go back');
  h.ui.returnFromWarnings();
  assert.equal(h.requests.length, 0);
});
test('HDD drafts with an invalid saved receipt can still be opened and edited', () => {
  const h = harness();
  const render = vm.runInContext('deviceReceiptEditor', h.context);
  const receipt = {date:'2026-09-11', count:12};
  const editor = render(JSON.stringify([null, receipt]), 'device_receipts');
  const stored = editor.children.find(node => node.tag === 'textarea');
  assert.deepEqual(JSON.parse(stored.value), [{date:'', count:''}, receipt]);
});
test('Old shared-reference notices never display or mark a record as needing attention', () => {
  const h=harness(),list=new Node('ul');
  const first='Reference SAME-REF is also used by another inward. You can keep the same reference and continue.';
  const second='Reference SAME-REF already appears in another PDF. This may be the same inward or an already-created certificate. Review it; continuing keeps this reference unchanged.';
  h.ui.renderWarningList(list,[first,second],true);
  assert.equal(list.children.length,0);
  const row={certificate_status:'ready',warnings:[first,second]};
  assert.equal(h.ui.hasWarnings(row),false);
  assert.equal(h.ui.needsAttention(row),false);
  assert.equal(h.ui.stateLabel(row),'Ready');
});
test('Preview and confirmation keep every current check visible', () => {
  const h=harness(), list=new Node('ul');
  const warnings=['missing supplier name','missing inward date','no item rows','total weight must be a positive number'];
  h.ui.renderWarningList(list,warnings,true);
  assert.equal(list.children.length,4);
  assert.equal(all(list,n=>n.className==='issue-card').length,4);
  h.ui.renderWarningList(list,warnings);
  assert.equal(list.children.length,4,'The creation confirmation must keep every check visible');
});
test('Different wording for the same inward-date warning shows one check', () => {
  const h=harness(), list=new Node('ul');
  h.ui.renderWarningList(list,['missing inward date','Missing inward date; check the date in Excel.','no item rows']);
  assert.equal(list.children.length,2);
  assert.doesNotMatch(text(list),/other entr/);
});
const followupRow = {...removalRow, certificate_status:'completed', comments:'', transporter_details:'Eco Recycling', docket_vehicle_no:'D-42', weight_slip_scale:''};

test('Later details stay separate from certificate alerts and include completed inwards', () => {
  const h=harness();
  h.controls.get('#plant-filter').value='all'; h.controls.get('#template-filter').value='all';
  h.controls.get('#status-filter').value='details_to_finish';
  assert.equal(h.ui.matches(followupRow),true);
  assert.equal(h.ui.needsAttention(followupRow),false);
  assert.deepEqual(Array.from(h.ui.missingFollowup(followupRow)),['comments','weight_slip_scale']);
  assert.equal(h.ui.matches({...followupRow,comments:'Not applicable',weight_slip_scale:'Yes'}),false);
  assert.equal(h.ui.matches({...followupRow,source_kind:'manual'}),false);
  assert.equal(h.ui.matches({source_kind:'system',plant:'plant1'}),false);
  const actions=h.ui.makeActions(followupRow,0);
  assert.deepEqual(all(actions,n=>n.tag==='button').map(n=>n.textContent),['Review & edit']);
  assert.match(text(h.ui.makeClientCell(followupRow,0)),/Missing: Comments.*Weight slip/);
});

test('Details to finish includes N/A inwards when office details are missing', () => {
  const h=harness();
  const skipped={...followupRow, sn:'2002', certificate_status:'na'};
  const filled={...skipped, sn:'2003', comments:'Checked', weight_slip_scale:'WS-10'};
  h.ui.app.state.rows=[followupRow, skipped, filled];
  const status=h.controls.get('#status-filter');
  status.value='details_to_finish';
  h.ui.summarize();
  assert.equal(h.controls.get('#count-details-to-finish').textContent,2);
  assert.equal(h.controls.get('#count-na').textContent,2);
  assert.equal(h.ui.matches(skipped),true);
  assert.equal(h.ui.matches(filled),false);
  assert.equal(h.ui.matches(followupRow),true);
  status.value='ready';
  assert.equal(h.ui.matches(skipped),false);
  status.value='na';
  assert.equal(h.ui.matches(skipped),true);
});

test('The four inward fields remain visible in Add client and directly inside Review & edit', () => {
  const html=fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  const editor=html.slice(html.indexOf('id="details-modal"'),html.indexOf('id="excel-save-modal"'));
  const intake=html.slice(html.indexOf('id="add-client-modal"'));
  for(const key of ['comments','transporter_details','docket_vehicle_no','weight_slip_scale']) {
    assert.ok(editor.includes(`name="${key}"`)); assert.ok(intake.includes(`name="${key}"`));
  }
  assert.doesNotMatch(editor,/<details[^>]*id="editor-inward-details"|followup-modal|editor-fill-details/);
  assert.doesNotMatch(intake,/Fill later/);
});

function inlineInward(h, issued=true) {
  const form=editing(h);
  Object.assign(form.values,{comments:'Existing',transporter_details:'Eco Recycling',docket_vehicle_no:'D-42',weight_slip_scale:''});
  Object.assign(h.ui.app.editor,{existing:true,issuedCertificate:issued,sourcePlant:'plant1',sourceSn:'1001',workspaceRevision:'revision-1'});
  h.ui.app.editor.baseline=h.ui.detailsBody(form);
  h.ui.app.editor.excelBaseline=JSON.parse(JSON.stringify(h.ui.app.editor.baseline));
  h.ui.app.templatesLoaded=true;
  h.controls.get('#details-modal').open=true;
  return form;
}

test('Changing only inward details saves from the main editor after Yes and creates no PDF', async () => {
  const h=harness({fetch:async()=>({ok:true,json:async()=>({state:{rows:[]},workspace_revision:'revision-2'})})});
  vm.runInContext('setState = state => { app.state=state; };',h.context);
  const form=inlineInward(h);form.values.comments='Updated comment';
  await h.ui.saveInwardDetails();
  assert.equal(h.ui.app.excelSavePayload.mode,'inward_only');assert.equal(h.requests.length,0);
  await h.ui.saveExcelChoice(false);
  assert.equal(h.requests.length,0);assert.equal(form.values.comments,'Updated comment');
  await h.ui.saveInwardDetails();await h.ui.saveExcelChoice(true);
  assert.equal(h.requests.length,1);assert.equal(h.requests[0].endpoint,'/api/clients/followup');
  assert.deepEqual(Object.keys(h.requests[0].payload.details),['comments','transporter_details','docket_vehicle_no','weight_slip_scale']);
  assert.equal(h.requests[0].payload.details.comments,'Updated comment');
  assert.equal(h.requests[0].payload.confirmed_excel,true);
});

test('Issued PDF-only corrections do not prevent saving new inward details on their own', async () => {
  const h=harness();const form=inlineInward(h);
  h.ui.app.editor.excelBaseline.details.location='Original Excel city';
  form.values.weight_slip_scale='WS-10';
  await h.ui.saveInwardDetails();
  assert.equal(h.ui.app.excelSavePayload.mode,'inward_only');
  assert.match(text(h.controls.get('#excel-save-change-list')),/Weight slip/);
  assert.doesNotMatch(text(h.controls.get('#excel-save-change-list')),/Original Excel city/);
});

test('An inline inward save failure keeps the entered fields and error in Review & edit', async () => {
  const h=harness({fetch:async()=>({ok:false,status:400,json:async()=>({error:'Excel is busy. Try again.'})})});
  const form=inlineInward(h);form.values.comments='Keep my text';
  await h.ui.saveInwardDetails();await h.ui.saveExcelChoice(true);
  assert.equal(form.values.comments,'Keep my text');
  assert.equal(h.controls.get('#details-modal').open,true);
  assert.match(text(h.controls.get('#details-error')),/Excel file is in use.*Save and close it in Excel/);
});

function removalHarness(fail = false) {
  const h=harness({fetch:async endpoint=>({ok:!(fail && endpoint==='/api/clients/remove'),json:async()=>
    endpoint==='/api/clients/remove-preview' ? {entry:{...removalRow,inward_date:'2026-10-01',materials:2},review_token:'reviewed-entry',expected_revision:'revision-1'}
      : fail ? {error:'The files changed. Check the entry again.'} : {removed:true,state:{rows:[]}}})});
  h.ui.app.workspaceRevision='revision-1';
  vm.runInContext('setState = state => { app.state=state; };',h.context);
  return h;
}

test('Remove inward is directly visible and unavailable for issued or manual records', () => {
  const h=harness();
  const actions=h.ui.makeActions(removalRow,0);
  assert.equal(all(actions,n=>n.tag==='details').length,0);
  assert.match(text(actions),/Review & edit.*Open in Excel/);
  const trash=all(actions,n=>n.dataset.action==='remove-inward')[0];
  assert.match(trash.attributes['aria-label'],/Delete inward/); assert.match(trash.innerHTML,/<svg/);
  assert.equal(all(h.ui.makeActions({...removalRow,certificate_no:'0042'},0),n=>n.dataset.action==='remove-inward').length,1);
  assert.equal(all(actions,n=>n.dataset.action==='remove-inward').length,1);
  for(const changes of [{has_issued_certificate:true},{certificate_status:'generated'},
    {certificate_status:'completed'},{source_kind:'manual'},{duplicate_serial_number:true}]) {
    assert.equal(all(h.ui.makeActions({...removalRow,...changes},0),n=>n.dataset.action==='remove-inward').length,0);
  }
});

const deletionSearchRows = [
  {source_kind:'workbook',plant:'plant1',sn:'17',sheet:'Inward Register',row:2,supplier:'Acme Medical',location:'Mumbai',certificate_no:'0212',reference:'REF-ALPHA',form6:'F6-OLD',certificate_status:'ready'},
  {source_kind:'workbook',plant:'plant2',sn:'18',sheet:'Inward Register',row:3,supplier:'Another company',location:'Pune',certificate_no:'0213',reference:'REF-BETA',certificate_status:'ready'},
  {source_kind:'manual',plant:'plant1',record_id:'local-test',supplier:'International client',certificate_no:'0214',certificate_status:'ready'}
];

test('Ungenerated-entry deletion is inside collapsed Settings Advanced and uses a searchable checklist', () => {
  const html=fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  assert.match(html,/<details class="technical-details"><summary>Advanced<\/summary>[\s\S]*?id="open-entry-deletion"/);
  const h=harness();h.ui.app.state={rows:deletionSearchRows,workspace_revision:'search-review'};
  h.ui.openEntryDeletion();
  assert.equal(h.ui.app.entryDeletion.rows.length,3);
  assert.equal(h.ui.app.entryDeletion.checked.size,0);
  assert.equal(h.controls.get('#entry-delete-preview').disabled,true);
  assert.equal(h.controls.get('#entry-delete-hdd').checked,false);
  assert.equal(h.controls.get('#entry-delete-search').focused,true);
  for(const query of ['acme','17','0212','REF-ALPHA','F6-OLD','mumbai','ready acme']){
    h.controls.get('#entry-delete-search').value=query;h.ui.filterEntryDeletion();
    assert.equal(h.ui.app.entryDeletion.labels.filter(label=>!label.hidden).length,1,query);
    assert.equal(h.ui.app.entryDeletion.labels[0].hidden,false);
  }
  assert.equal(h.requests.length,0);
});

test('Settings no longer offers old-history cleanup or exposes its endpoint', () => {
  const html=fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  const source=fs.readFileSync(path.join(__dirname,'../web/app.js'),'utf8');
  assert.doesNotMatch(html,/clear-old-history|Clear old history/);
  assert.doesNotMatch(source,/\/api\/history\/clear|clearOldHistory/);
});

test('App-only forms submit no hidden plant and local PDFs do not infer one from their filename', () => {
  const h=harness();const form=editing(h);
  form.values.source_kind='manual';form.values.plant='plant2';
  assert.equal(h.ui.detailsBody(form).plant,null);
  assert.equal(vm.runInContext("labelPlant(null)",h.context),'App only');
  assert.equal(vm.runInContext("savedFilePlant({path:'Certificates/Manually Added/0042_Client_1.pdf',managed_owner:{source_kind:'manual',plant:null,key:'local'}})",h.context),'local');
});

test('Inward deletion keeps the Advanced checklist hidden until its review is ready', async () => {
  let finishPreview;
  const h=harness({fetch:async()=>new Promise(resolve=>{finishPreview=()=>resolve({ok:true,json:async()=>({
    review:{entries:[{plant:'plant1',sn:'17',supplier:'Acme Medical',cod:'0212',materials:1}],pdfs:[],kept_hdd:[]},review_token:'review-token'})});})});
  h.ui.app.state={rows:deletionSearchRows,workspace_revision:'search-review'};
  h.ui.openEntryDeletion({showDialog:false});
  h.ui.app.entryDeletion.checked.add(0);
  assert.equal(Boolean(h.controls.get('#entry-deletion-modal').open),false,'no searchable Settings dialog should flash');
  const pending=h.ui.reviewEntryDeletion();
  await new Promise(resolve=>setImmediate(resolve));
  assert.equal(Boolean(h.controls.get('#entry-deletion-modal').open),false,'keep it hidden during the preview request');
  finishPreview(); await pending;
  assert.equal(h.controls.get('#entry-deletion-modal').open,true);
  assert.equal(h.controls.get('#entry-delete-selection').hidden,true);
  assert.equal(h.controls.get('#entry-delete-review').hidden,false);
});

test('Deletion search keeps checked entries across searches and shows hidden selection counts', () => {
  const h=harness();h.ui.app.state={rows:deletionSearchRows,workspace_revision:'search-review'};h.ui.openEntryDeletion();
  const first=h.ui.app.entryDeletion.labels[0].children[0];first.checked=true;first.dispatch('change');
  h.controls.get('#entry-delete-search').value='another';h.ui.filterEntryDeletion();
  assert.equal(h.ui.app.entryDeletion.checked.size,1);
  assert.match(h.controls.get('#entry-delete-selected').textContent,/1 selected.*1 hidden/);
  assert.equal(h.controls.get('#entry-delete-preview').disabled,false);
  h.controls.get('#entry-delete-search').value='';h.ui.filterEntryDeletion();
  assert.equal(first.checked,true);assert.equal(h.ui.app.entryDeletion.labels[0].hidden,false);
});

test('Advanced deletion reviews exact entries and PDFs and waits for a checked confirmation', async () => {
  const review={entries:[{source_kind:'workbook',plant:'plant1',sn:'17',supplier:'Acme Medical',cod:'0212',materials:2,certificate_type:'siemens'}],pdfs:['Certificates/Auto Generated/0212_Acme_Mumbai.pdf'],kept_hdd:[{supplier:'Acme Medical',cod:'0215'}]};
  const h=harness({fetch:async(endpoint)=>({ok:true,json:async()=>endpoint.endsWith('delete-preview')?{review,review_token:'review-token'}:{deleted:true,state:{rows:[]}}})});
  h.ui.app.state={rows:deletionSearchRows,workspace_revision:'search-review'};
  vm.runInContext('setState=state=>{app.state=state;};',h.context);h.ui.wire();h.ui.openEntryDeletion();
  const input=h.ui.app.entryDeletion.labels[0].children[0];input.checked=true;input.dispatch('change');
  await h.ui.reviewEntryDeletion();
  assert.equal(h.requests.length,1);assert.equal(h.requests[0].endpoint,'/api/entries/delete-preview');
  assert.deepEqual(h.requests[0].payload.selected,[{source_kind:'workbook',plant:'plant1',sn:'17',sheet:'Inward Register',row:2}]);
  assert.match(text(h.controls.get('#entry-delete-review-list')),/0212_Acme_Mumbai.pdf/);
  assert.match(h.controls.get('#entry-delete-kept-hdd').textContent,/kept as standalone/);
  assert.equal(h.controls.get('#entry-delete-confirm').disabled,true);
  assert.equal(h.controls.get('#entry-delete-cancel').focused,true);
  await h.ui.confirmEntryDeletion();assert.equal(h.requests.length,1);
  h.controls.get('#entry-delete-checked').checked=true;h.controls.get('#entry-delete-checked').dispatch('change');
  await h.ui.confirmEntryDeletion();
  assert.equal(h.requests.length,2);assert.equal(h.requests[1].endpoint,'/api/entries/delete');
  assert.equal(h.requests[1].payload.review_token,'review-token');assert.equal(h.requests[1].payload.checked,true);
  assert.equal(h.controls.get('#entry-deletion-modal').open,false);assert.equal(h.ui.app.entryDeletion,null);
});

test('Canceling advanced deletion clears its review and never deletes an entry', async () => {
  const h=harness({fetch:async()=>({ok:true,json:async()=>({review:{entries:[],pdfs:[],kept_hdd:[]},review_token:'token'})})});
  h.ui.app.state={rows:deletionSearchRows,workspace_revision:'search-review'};h.ui.wire();h.ui.openEntryDeletion();
  h.ui.app.entryDeletion.checked.add(0);await h.ui.reviewEntryDeletion();
  h.controls.get('#entry-deletion-modal').close();
  assert.equal(h.ui.app.entryDeletion,null);assert.equal(h.requests.length,1);
  assert.equal(h.controls.get('#entry-delete-checked').checked,false);
});

test('Standalone HDD kept after deleting a mistaken inward never guesses another Excel entry', () => {
  const h=harness();h.ui.app.state.rows=[deletionSearchRows[0]];
  assert.equal(h.ui.excelTarget({source_kind:'manual',record_id:'kept-hdd',certificate_type:'destruction',supplier:'Acme Medical',location:'Mumbai',detached_inward:{plant:'plant1',case_id:'17'}}),null);
});

test('Removal review shows the exact inward, defaults to Keep entry and cannot submit unchecked', async () => {
  const h=removalHarness();await h.ui.openInwardRemoval(removalRow);
  assert.equal(h.requests.length,1);assert.equal(h.requests[0].endpoint,'/api/clients/remove-preview');
  assert.match(text(h.controls.get('#remove-inward-facts')),/ABC Company.*25.*Plant 1/);
  assert.equal(h.controls.get('#keep-inward').focused,true);
  assert.equal(h.controls.get('#remove-inward-checked').checked,false);
  assert.equal(h.controls.get('#confirm-remove-inward').disabled,true);
  await h.ui.confirmInwardRemoval();assert.equal(h.requests.length,1);
  h.controls.get('#remove-inward-checked').checked=true;
  await h.ui.confirmInwardRemoval();assert.equal(h.requests.length,1,'Disabled button cannot be bypassed by a click');
});

test('Deliberate removal sends the reviewed identity once and closes after success', async () => {
  const h=removalHarness();await h.ui.openInwardRemoval(removalRow);
  h.controls.get('#remove-inward-checked').checked=true;h.controls.get('#confirm-remove-inward').disabled=false;
  await Promise.all([h.ui.confirmInwardRemoval(),h.ui.confirmInwardRemoval()]);
  assert.equal(h.requests.filter(r=>r.endpoint==='/api/clients/remove').length,1);
  const sent=h.requests.at(-1).payload;
  assert.equal(sent.checked,true);assert.equal(sent.review_token,'reviewed-entry');
  assert.equal(sent.sn,'25');assert.equal(sent.row,3);assert.equal(sent.expected_revision,'revision-1');
  assert.equal(h.controls.get('#remove-inward-modal').open,false);
});

test('Removal errors keep the message open, clear consent and preserve the reviewed identity for retry', async () => {
  const h=removalHarness(true);await h.ui.openInwardRemoval(removalRow);
  h.controls.get('#remove-inward-checked').checked=true;h.controls.get('#confirm-remove-inward').disabled=false;
  await h.ui.confirmInwardRemoval();
  assert.equal(h.controls.get('#remove-inward-modal').open,true);
  assert.equal(h.controls.get('#remove-inward-checked').checked,false);
  assert.equal(h.controls.get('#confirm-remove-inward').disabled,true);
  assert.match(h.controls.get('#remove-inward-error').textContent,/files changed/);
  assert.equal(h.ui.app.removalPending.sn,'25');
});

test('Closing removal clears consent and Enter on the checkbox never deletes', async () => {
  const h=removalHarness();h.ui.wire();await h.ui.openInwardRemoval(removalRow);
  let prevented=false;
  h.controls.get('#remove-inward-modal').dispatch('keydown',{key:'Enter',target:h.controls.get('#remove-inward-checked'),preventDefault(){prevented=true;}});
  assert.equal(prevented,true);assert.equal(h.requests.length,1);
  h.controls.get('#remove-inward-checked').checked=true;h.controls.get('#remove-inward-checked').dispatch('change');
  assert.equal(h.controls.get('#confirm-remove-inward').disabled,false);
  h.controls.get('#remove-inward-modal').close();
  assert.equal(h.ui.app.removalPending,null);assert.equal(h.controls.get('#remove-inward-checked').checked,false);
  assert.equal(h.controls.get('#confirm-remove-inward').disabled,true);
  await h.ui.openInwardRemoval(removalRow);
  assert.equal(h.controls.get('#confirm-remove-inward').disabled,true);
});

test('HDD change review shows received dates and device numbers instead of stored JSON', () => {
  const h=harness(), form=editing(h);
  const before='[{"date":"2026-08-06","count":""}]', after='[{"date":"2026-08-06","count":"153"}]';
  form.values.certificate_type='destruction';h.ui.app.editor.template='destruction';
  h.ui.app.editor.baseline.details.certificate_type='destruction';
  h.ui.app.editor.baseline.details.certificate_fields={device_receipts:before};
  h.ui.app.editor.fields.destruction={device_receipts:after};
  h.ui.app.templates=[{id:'destruction',name:'Data destruction',fields:[{key:'device_receipts',label:'Device receipts'}]}];
  h.ui.renderChangeSummary();
  const shown=text(h.controls.get('#change-list'));
  assert.match(shown,/Device receipts.*Before:.*06\/08\/26.*Number not entered.*After:.*06\/08\/26.*153 devices/);
  assert.doesNotMatch(shown,/"date"|"count"|\[\{|device_receipts/);
  assert.equal(JSON.parse(h.ui.detailsBody(form).details.certificate_fields.device_receipts)[0].count,'153');
});

test('Both device-list key formats, multiple dates, singular counts and missing details are readable', () => {
  const h=harness(), rows=[{date:'2026-08-06',count:'1'},{date:'',count:'153'}];
  const plain=h.ui.changeDisplay('device_receipts',JSON.stringify(rows));
  assert.equal(plain,h.ui.changeDisplay('certificate_fields.device_receipts',JSON.stringify(rows)));
  assert.match(plain,/1 device; Date not entered.*153 devices/);
  assert.equal(plain,h.ui.changeDisplay('device_receipts',rows));
  for(const value of ['bad JSON','{}','null']) {
    assert.match(h.ui.changeDisplay('device_receipts',value),/Check the received dates/);
  }
  assert.equal(h.ui.changeDisplay('device_receipts','[]'),'No devices entered');
  assert.match(h.ui.changeDisplay('device_receipts','[{"date":{},"count":"0"}]'),/Check the received date.*Check the number/);
});

function editing(h) {
  const form = h.controls.get('#details-form');
  form.values = {plant:'plant1',source_kind:'workbook',edit_scope:'certificate_only',certificate_type:'standard',supplier:'Acme',weight:'100.2',case_id:'1001',use_certificate_weight:false};
  const inputs = {};
  form.elements = {namedItem: name => inputs[name] ||= {dataset:{},listeners:{},addEventListener:Node.prototype.addEventListener,type:name === 'use_certificate_weight' ? 'checkbox' : 'text', focus(){this.focused = true;}, setCustomValidity(message){this.validationMessage=message;}, checkValidity(){return !this.validationMessage;}, reportValidity(){return this.checkValidity();}, get value() {return form.values[name] || '';}, set value(value) {form.values[name] = value;}, get checked() {return Boolean(form.values[name]);}, set checked(value) {form.values[name] = Boolean(value);}}};
  form.reset = () => {};
  form.reportValidity = () => true;
  form.checkValidity = () => true;
  h.ui.app.editor = {template:'standard',fields:{standard:{}},annexures:{},touched:new Set(),rawNumeric:{weight:'100.19999999999999'},sourceSn:'1001',sourcePlant:'plant1',workspaceRevision:'revision-1'};
  h.ui.app.templatesLoaded = true;
  h.ui.app.templates = [{id:'standard',fields:[]}];
  h.controls.get('#primary-certificate-needed-card').hidden = true;
  h.controls.get('#details-modal').open = true;
  h.ui.app.editor.baseline = h.ui.detailsBody(form);
  return form;
}

test('Analytics overview navigation preserves company and saved-file workspaces', () => {
  const h = harness();
  h.ui.setWorkspaceView('analytics');
  assert.equal(h.ui.app.workspaceView, 'analytics');
  assert.equal(h.controls.get('#analytics-workspace').hidden, false);
  assert.equal(h.controls.get('#records-workspace').hidden, true);
  assert.equal(h.controls.get('#folder-inventory').hidden, true);
  assert.equal(h.controls.get('#view-records').attributes['aria-pressed'], 'false');
  h.ui.setWorkspaceView('records');
  assert.equal(h.controls.get('#analytics-workspace').hidden, true);
  assert.equal(h.controls.get('#records-workspace').hidden, false);
});

test('Analytics drilldown filters isolate pending review, dispatch and destruction follow-up', () => {
  const h = harness(), filter = h.controls.get('#status-filter');
  const generated = {plant:'plant1',certificate_status:'generated'};
  const manual = {plant:'plant1',certificate_status:'generated_manual'};
  filter.value = 'awaiting_send';
  assert.equal(h.ui.matches(generated), true);
  assert.equal(h.ui.matches(manual), true);
  assert.equal(h.ui.matches({...generated,certificate_status:'completed'}), false);
  filter.value = 'pending_review';
  assert.equal(h.ui.matches({...generated,certificate_status:'blank'}), true);
  assert.equal(h.ui.matches({...generated,certificate_status:'ready',warnings:['Check weight']}), true);
  assert.equal(h.ui.matches({...generated,certificate_status:'ready'}), false);
  assert.equal(h.ui.matches(generated), false);
  filter.value = 'destruction_pending';
  assert.equal(h.ui.matches({...manual,certificate_type:'destruction'}), true);
  assert.equal(h.ui.matches({...manual,certificate_type:'destruction',certificate_status:'completed'}), false);
  filter.value = 'overdue_unissued'; h.ui.app.analyticsAsOf = '2026-10-19';
  assert.equal(h.ui.matches({...generated,certificate_status:'ready',inward_date:'2026-08-01'}), true);
  assert.equal(h.ui.matches({...generated,certificate_status:'ready',inward_date:'2026-08-01',destruction_source:{sn:'1'}}), false);
  assert.equal(h.ui.matches({...generated,inward_date:'2026-08-01'}), false);
});

test('The dashboard inward-date drilldown shows only matching inwards for its plant', () => {
  const h = harness(); h.ui.wire();
  h.ui.setState({rows:[
    {plant:'plant1',sn:'1',supplier:'Date missing',inward_date:''},
    {plant:'plant1',sn:'2',supplier:'Invalid date',inward_date:'31/02/26'},
    {plant:'plant1',sn:'3',supplier:'Date present',inward_date:'2026-11-01'},
    {plant:'plant2',sn:'4',supplier:'Other plant',inward_date:''},
    {plant:'plant1',sn:'1',supplier:'Linked HDD',source_kind:'manual',record_id:'hdd',inward_date:''},
  ]});
  h.controls.get('#analytics-plant').value = 'plant1';
  h.controls.get('#serial-filter').value = '3';
  h.controls.get('#received-date-filter').value = '01/11/26';
  h.controls.get('#received-to-filter').value = '01/11/26';
  h.controls.get('#record-followup-filter').value = 'destruction_pending';
  h.controls.get('#analytics-workspace').dispatch('analytics:review',{detail:{status:'needs_inward_date'}});
  assert.equal(h.ui.app.workspaceView,'records');
  assert.equal(h.controls.get('#queue-title').textContent,'Inwards that need a received date');
  assert.deepEqual(Array.from(h.ui.app.filtered, item=>item.row.sn),['1','2']);
});

const duplicateIssueExamples = [
  'COD 0365 is already recorded in certificate PDF history (0365_Dummy Client 2_Pune_1.pdf). If you continue, new automatic COD 0393 will be used; existing PDFs will not be overwritten.',
  'Reference DUMMY-REF-1002 already appears in 0365_Dummy Client 2_Pune_1.pdf. This may be the same inward or an already-created certificate. Review it; continuing keeps this reference unchanged.',
];

test('Unreadable workbooks get a separate message and readable companies stay visible', () => {
  const h = harness();
  h.ui.setState({scan_complete:false, rows:[
    {source_kind:'system',plant:'plant2',sn:'',outcome:'needs_review',
      detail:'plant2.xlsx could not be read: No inward register found. Put headings within the first 20 rows, then refresh.',
      warnings:['Missing inward date; check the date in Excel.']},
    {source_kind:'workbook',plant:'plant1',sn:'A',supplier:'Acme',certificate_status:'ready',issuance_blocked:true},
  ]});
  assert.equal(h.ui.app.state.rows.length,1);
  assert.equal(h.ui.app.state.rows[0].supplier,'Acme');
  assert.equal(h.controls.get('#workbook-issues').hidden,false);
  const messages=text(h.controls.get('#workbook-issue-list'));
  assert.match(messages,/plant2.xlsx could not be read.*first 20 rows.*Refresh/);
  assert.doesNotMatch(messages,/Received date missing|Company name missing/);
  assert.equal(h.controls.get('#open-bulk').disabled,true);
  const actions=h.ui.makeActions(h.ui.app.state.rows[0],0);
  assert.match(text(actions),/Review & edit.*Open in Excel/);
  assert.doesNotMatch(text(actions),/Review & create/);
});

test('Refresh does not announce success for an incomplete workbook scan', async () => {
  const h=harness({fetch:async () => ({ok:true,json:async () => ({state:{scan_complete:false,rows:[],setup_warning:'Choose the plant files in Settings.'}})})});
  await h.ui.refreshState();
  assert.match(h.controls.get('#toast-message').textContent,/Some Excel files could not be checked/);
  assert.equal(h.controls.get('#toast').classList.contains('error'),true);
  assert.equal(h.controls.get('#workbook-issues').hidden,false);
});

test('Locked files, future dates and ambiguous totals have specific next steps', () => {
  const h=harness(); editing(h);
  const {ui}=h;
  Object.assign(ui.app.editor,{sourcePlant:'plant1',sourceSn:'1001'});
  ui.showDetailsError('plant1.xlsx is unavailable: Cannot update plant1.xlsx. Save and close it in Excel, then try again. (Windows error 32)');
  const error = h.controls.get('#details-error');
  assert.match(error.textContent,/plant1\.xlsx is in use.*Save and close it in Excel.*retry/);
  assert.equal(all(error,node=>node.tag==='button').length,0);
  assert.match(ui.simpleWarning('COD Date is in the future; check this date before creating the certificate'),
    /COD date is in the future.*later than today.*confirm this future date deliberately/);
  assert.match(ui.simpleWarning('Excel row 4: total row needs checking. It does not clearly match the materials; correct or remove the summary row before generating.'),
    /Total row needs checking.*row 4.*Keep the material rows/);
});

test('Missing source files cannot re-enable creation in the preview', () => {
  const h=harness();editing(h);
  h.ui.app.editor.issuanceBlocked=true;
  vm.runInContext('updateDetailsButtons()',h.context);
  assert.equal(h.controls.get('#save-details').disabled,true);
  assert.equal(h.controls.get('#preview-details').disabled,false);
});

test('Issued duplicate and history-conflict records remain blocked and keep Open in Excel available', async () => {
  for (const certificate_status of ['ready','generated','completed']) {
    for (const conflict of [{duplicate_serial_number:true}, {record_identity_conflict:'This Serial Number belongs to a different inward in certificate history. Check Excel before continuing.'}]) {
      const h = harness(), form = editing(h); h.controls.get('#details-modal').open = false;
      vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {};',h.context);
      const row = {source_kind:'workbook',plant:'plant1',sn:'1001',supplier:'Current company',
        certificate_status,has_issued_certificate:true,issuance_blocked:true,
        items:[{item:'Laptops',weight:10}],certificate_type:'standard',last_certificate_details:{supplier:'Wrong historical company',certificate_type:'sita'},...conflict};
      const actions = all(h.ui.makeActions(row,0), node => node.tag === 'button');
      assert.equal(actions.filter(node => node.textContent === 'Open in Excel').length,1);
      assert.equal(actions.filter(node => node.dataset.action === 'edit-details' || node.dataset.action === 'complete').length,0);
      await h.ui.openDetails(row,true,false);
      assert.equal(h.ui.app.editor.issuanceBlocked,true);
      assert.equal(h.controls.get('#save-details').disabled,true);
      assert.equal(h.controls.get('#preview-details').disabled,true);
      assert.equal(h.controls.get('#details-error').hidden,false);
      assert.match(h.controls.get('#details-error').textContent,/Serial Number/);
      assert.equal(form.values.supplier,'Current company','Conflicted receipts must not replace the live source details.');
      assert.equal(h.ui.app.editor.template,'standard','Conflicted history must not choose the certificate template.');
      h.ui.invalidatePreview();
      assert.equal(h.controls.get('#details-error').hidden,false,'Editing must not hide the identity repair message.');
      await h.ui.previewDetails();
      await h.ui.saveDetails({preventDefault(){},currentTarget:form});
      assert.equal(h.requests.length,0,'Blocked records cannot preview or submit by bypassing the disabled button.');
    }
  }
});

test('Duplicate issue cards keep the COD change and omit retired reference notices', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({preview_url:`/previews/${'a'.repeat(32)}.pdf`,
    certificate_no:'0393', warnings:duplicateIssueExamples})})}); editing(h);
  await h.ui.previewDetails();
  const list = h.controls.get('#preview-warnings');
  const cards = all(list,node => node.className === 'issue-card'); assert.equal(cards.length,1);
  const visibleCopy = card => card.children.filter(node => node.tag !== 'details').map(text).join(' ');
  assert.match(visibleCopy(cards[0]),/COD 0365 is already used.*COD 0393.*Existing PDFs stay unchanged.*inward already has a certificate/);
  assert.doesNotMatch(text(list),/Reference already used|DUMMY-REF-1002/);
  assert.doesNotMatch(visibleCopy(cards[0]),/What to do:/);
  for (let index = 0; index < cards.length; index++) {
    assert.doesNotMatch(visibleCopy(cards[index]), /\.pdf|PDF history|automatic COD/);
  }
  assert.equal(h.controls.get('#preview-warnings-title').textContent,'Check before creating');
  assert.equal(h.requests.length,1); assert.equal(h.requests[0].endpoint,'/api/preview');
  h.ui.invalidatePreview(); assert.equal(h.controls.get('#preview-warnings-title').hidden,true);
  assert.equal(list.children.length,0);
});

test('Warning review uses the same short duplicate messages and preserves the original confirmation payload', async () => {
  const h = harness(); const form = editing(h); const payload = h.ui.detailsBody(form);
  h.ui.showWarnings([{plant:'plant1',supplier:'Acme',warnings:duplicateIssueExamples}],
    {endpoint:'/api/create',payload,fromDetails:true,revision:h.ui.app.previewRevision,reviewToken:'original-review'});
  assert.equal(all(h.controls.get('#warning-list'),node => node.className === 'issue-card').length,1);
  assert.equal(h.ui.app.pendingConfirmation.payload,payload); assert.equal(h.ui.app.pendingConfirmation.reviewToken,'original-review');
  await h.ui.confirmWarnings();
  assert.equal(h.requests[0].payload.review_token,'original-review');
  assert.equal(h.requests[0].payload.override_reason,'Warnings reviewed; user chose Continue anyway.');
  assert.equal(h.requests[0].payload.confirmed,true);
});

test('Common data issues use plain language while keeping line IDs and consequences', () => {
  const {ui} = harness();
  for (const [raw, expected] of [
    ['COD 0365 is already recorded in certificate PDF history (old.pdf). Continuing will use new automatic COD 0393; existing PDFs will not be overwritten.', /0365.*0393.*Existing PDFs stay unchanged/],
    ['COD ECO-0365 belongs to another certificate. If you continue, a new automatic COD ECO-0393 will be used.', /ECO-0365.*ECO-0393/],
    ['Company name is missing; the company line will be blank.', /Company name missing.*blank on the PDF.*Add the company name/],
    ['Material receipt date is missing or invalid; it will be blank on the certificate.', /Received date missing or invalid.*blank on the PDF/],
    ['Certificate date is invalid; today\'s date will be used if you continue.', /COD date needs fixing.*today’s date will be used/],
    ['invalid weight or unsupported weight formula on row 12; enter a confirmed weight or a same-row quantity times per-piece formula', /Excel row 12.*could not read or calculate.*quantity by weight per piece from the same row/],
    ['invalid or negative per_piece_weight on row 9', /Excel row 9.*weight per piece.*below zero/],
    ['invalid or negative certificate/net weight on row 4', /Excel row 4.*material total will be used/],
    ['Material 3 has an invalid quantity; check the material list and weight before sending.', /Material 3.*quantity is invalid.*total/],
    ['Missing X-ray tube serial number; enter N/A if it does not apply', /X-ray tube serial number missing.*blank.*N\/A/],
    ['Company name is long for its space; the PDF may use small text or clip it. Shorten it and preview before sending.', /too long.*small text or cut off.*Shorten/],
    ['Plant 2 workbook is not connected; confirm that the COD does not duplicate an existing certificate before sending.', /Plant 2.*not connected.*used COD numbers.*check the COD yourself/],
  ]) assert.match(ui.simpleWarning(raw),expected,raw);
});

test('Unknown issues stay fully visible and company rows show a short useful issue label', () => {
  const h = harness(); editing(h);
  const raw = 'Unexpected <script> problem. Do not create this PDF until the operator checks it.';
  h.ui.showWarnings([raw],{endpoint:'/api/create',payload:{}});
  const card = all(h.controls.get('#warning-list'),node => node.className === 'issue-card')[0];
  assert.ok(text(card).includes(raw)); assert.equal(all(card,node => node.tag === 'details').length,0);
  assert.equal(all(card,node => node.tag === 'script').length,0);
  const client = h.ui.makeClientCell({supplier:'Acme',certificate_status:'ready',warnings:duplicateIssueExamples},0);
  const note = all(client,node => node.className === 'row-alert-note')[0];
  assert.equal(note.textContent,'COD 0365 is already used');
  assert.match(note.title,/0393/); assert.doesNotMatch(note.title,/DUMMY-REF-1002/);
  assert.doesNotMatch(note.textContent,/\.pdf|…/);
});

test('Saved PDF issues identify the certificate and keep unknown file alerts intact', () => {
  const h = harness();
  const warning = 'PDF has no readable text (possibly a scan); only its filename COD can be checked. References need manual verification.';
  h.ui.showSavedFile({path:'Certificates/0365_Dummy Client 2_Pune_1.pdf',present:true,cods:['0365'],warnings:[warning]});
  assert.match(text(h.controls.get('#saved-file-warnings')), /Scanned PDF needs a check.*filename.*COD and reference yourself/);
  assert.equal(h.controls.get('#saved-file-warnings').hidden,false);
  h.ui.showSavedFile({path:'Certificates/0393_Clear_Pune_1.pdf',present:true,cods:['0393'],warnings:[]});
  assert.equal(h.controls.get('#saved-file-warnings').hidden,true); assert.equal(h.controls.get('#saved-file-warnings').children.length,0);
  h.ui.app.state.certificate_inventory = {entries:[],warnings:[`Certificates/0365_Dummy Client 2_Pune_1.pdf: ${warning}`,'Unrecognized file error; check all certificates before sending.']};
  h.ui.renderFolderInventory();
  assert.match(text(h.controls.get('#folder-inventory-warnings')), /Certificate: Dummy Client 2 Pune/);
  assert.match(text(h.controls.get('#folder-inventory-warnings')), /Unrecognized file error; check all certificates before sending/);
});
test('Overdue does not conceal canonical certificate statuses', () => {
  const {ui} = harness();
  for (const [state, label] of [['ready','Ready'],['alert','Needs attention'],['generated','Generated'],['generated_manual','Generated Externally'],['completed','Completed'],['blank',''],['na','N/A']]) {
    assert.equal(ui.stateLabel({certificate_status:state,overdue:true}), label);
    const expectedClass = state === 'alert' ? 'status-needs_attention' : `status-${state}`;
    assert.equal(ui.statusClass({certificate_status:state,overdue:true}), expectedClass);
  }
  for (const [legacy, canonical] of [['Created','Generated'],['Sent','Completed'],['Check Details','Needs attention'],['Working Draft','Needs attention'],['Generated Manual','Generated Externally'],['Generated Externally','Generated Externally']]) {
    assert.equal(ui.stateLabel({certificate_status:legacy}), canonical);
  }
});
test('Ready card filter excludes warnings and keeps blank/N/A separate', () => {
  const {ui,controls} = harness(); controls.get('#status-filter').value = 'ready';
  assert.equal(ui.matches({certificate_status:'ready'}), true);
  for (const row of [{certificate_status:'ready',warnings:['Missing weight']},{certificate_status:'blank'},{certificate_status:'na'}]) assert.equal(ui.matches(row), false);
  assert.match(ui.nextStep({certificate_status:'blank'}), /Ready appears automatically/);
  assert.match(ui.nextStep({certificate_status:'na'}), /skipped/);
});
test('Unselected statuses show an empty accessible box without a Blank label', () => {
  const h = harness();
  const cell = h.ui.makeStatus({certificate_status:'blank'});
  const box = cell.children[0];
  assert.equal(box.textContent, '');
  assert.equal(box.className, 'status-pill status-blank');
  assert.equal(box.attributes['aria-label'], 'No status selected');
  const html = fs.readFileSync(path.join(__dirname, '../web/index.html'), 'utf8');
  assert.doesNotMatch(html, />Blank</);
});
test('Background Excel checks refresh Ready and its revision without interrupting dialogs or operations', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({changed:true, workspace_revision:'saved-excel',state:{rows:[{sn:'A',plant:'plant1',supplier:'Filled in Excel',certificate_status:'ready'}]}})})});
  h.controls.get('#busy-overlay').hidden = true;
  h.controls.get('#offline-screen').hidden = true;
  h.ui.app.loadingWorkspace = true;
  await h.ui.syncWorkbookChanges();
  assert.equal(h.requests.length, 0);
  h.ui.app.loadingWorkspace = false;
  h.controls.get('#busy-overlay').hidden = false;
  await h.ui.syncWorkbookChanges();
  assert.equal(h.requests.length, 0);
  h.controls.get('#busy-overlay').hidden = true;
  const query = h.context.document.querySelector;
  h.context.document.querySelector = selector => selector === 'dialog[open]' ? {} : query(selector);
  await h.ui.syncWorkbookChanges();
  assert.equal(h.requests.length, 0);
  h.context.document.querySelector = query;
  await h.ui.syncWorkbookChanges();
  assert.equal(h.requests.length, 1);
  assert.equal(h.requests[0].endpoint, '/api/workbook-changes');
  assert.equal(h.ui.app.state.rows[0].certificate_status, 'ready');
  assert.equal(h.ui.app.workspaceRevision, 'saved-excel');
});
test('Background workbook errors stay inline and disappear after recovery', async () => {
  let failed=true;
  const h=harness({fetch:async()=>{
    if (failed) throw new Error('Temporary connection failure');
    return {ok:true,json:async()=>({changed:false})};
  }});
  h.controls.get('#offline-screen').hidden=true;
  h.controls.get('#busy-overlay').hidden=true;
  await h.ui.syncWorkbookChanges();
  await h.ui.syncWorkbookChanges();
  assert.equal(h.controls.get('#offline-screen').hidden,true);
  assert.equal(h.controls.get('#toast').classList.contains('visible'),false);
  assert.equal(h.controls.get('#workbook-issue-list').children.length,1);
  assert.match(text(h.controls.get('#workbook-issue-list')),/could not be reached/);
  failed=false;
  await h.ui.syncWorkbookChanges();
  assert.equal(h.controls.get('#workbook-issues').hidden,true);
});

test('Automatic preview problems stay by the preview, explicit requests show the error', async () => {
  const h=harness({fetch:async()=>({ok:false,status:400,json:async()=>({error:'Enter the actual date the devices were destroyed.'})})});
  const form=editing(h);
  form.reportValidity=()=>{throw new Error('Automatic preview must not pop up field validation');};
  form.checkValidity=()=>false;
  await h.ui.previewDetails({automatic:true});
  assert.equal(h.requests.length,0);
  form.checkValidity=()=>true;
  await h.ui.previewDetails({automatic:true});
  assert.equal(h.controls.get('#details-error').hidden,true);
  assert.match(h.controls.get('#preview-status').textContent,/Destruction date missing/);
  await h.ui.previewDetails();
  assert.equal(h.controls.get('#details-error').hidden,false);
});

test('Search still finds identifiers hidden from the main table', () => {
  const {ui,controls} = harness();
  for (const [key,value] of [['reference','REF-19'],['form6','F6-20'],['sn','4018'],['certificate_no','0324']]) {
    controls.get('#search').value = value.toLowerCase();
    assert.equal(ui.matches({[key]:value}), true);
  }
});
test('Ready action opens review, never immediately generates', () => {
  const {ui} = harness(); const node = ui.makeActions({certificate_status:'ready',supplier:'Acme'},4);
  const buttons = all(node, node => node.tag === 'button');
  assert.equal(buttons[0].textContent, 'Review & edit');
  assert.equal(buttons[0].dataset.action, 'edit-details');
  assert.equal(buttons.length, 1, 'Review opens the preview directly without a Details expansion.');
});
test('Review & edit opens the editable form before or after issuance', async () => {
  const h = harness();
  const row = {certificate_status:'alert',supplier:'Acme',plant:'plant1',sn:'1001'};
  h.ui.app.state.rows = [row];
  const rowAction = all(h.ui.makeActions(row,0), node => node.tag === 'button')[0];
  const detailAction = all(h.ui.makeDetailRow(row,0), node => node.tag === 'button').find(node => node.textContent === 'Edit details');
  assert.equal(rowAction.textContent,'Review & edit'); assert.equal(rowAction.dataset.action,'edit-details');
  assert.equal(detailAction.dataset.action,'edit-details');
  vm.runInContext('globalThis.editorRequests = []; openDetails = (...args) => editorRequests.push(args);',h.context);
  await h.ui.handleAction(rowAction);
  await h.ui.handleAction({dataset:{index:'0',action:'details'}});
  assert.deepEqual(JSON.parse(JSON.stringify(h.context.editorRequests.map(args => args[2] || false))),[true,false]);
});
test('Generated Externally has its own filter and can be marked Completed', () => {
  const {ui,controls} = harness(); controls.get('#status-filter').value = 'generated_manual';
  const row = {certificate_status:'generated_manual',supplier:'Acme',pdf_filename:'0001_Acme.pdf'};
  assert.equal(ui.matches(row), true);
  const action = ui.makeActions(row,4);
  assert.ok(all(action, node => node.tag === 'button').some(node => node.textContent === 'Mark Completed'));
});
test('Full logs are retained only inside a closed technical disclosure', () => {
  const {ui} = harness(); const raw = 'Created very_long_filename.pdf; Workbook backup: old.xlsx';
  const detail = ui.makeDetailRow({certificate_status:'generated',detail:raw,display_weight:1.7999999999999998},0);
  assert.equal(detail.children[0].colSpan, 6);
  const technical = all(detail,node => node.tag === 'details').find(node => node.children[0].textContent === 'Technical details');
  assert.ok(technical); assert.notEqual(technical.open,true); assert.match(text(technical), /very_long_filename/);
  assert.match(text(detail), /1.8 kg/);
  assert.ok(all(detail,node => node.className === 'next-step').every(node => !text(node).includes('backup')));
});
test('Unknown warnings retain the full consequence instead of being hidden', () => {
  const {ui} = harness(); const unknown = 'COD 0324 is used. A new number will be assigned; existing PDFs will not be overwritten.';
  assert.equal(ui.simpleWarning(unknown), unknown);
  assert.match(ui.simpleWarning('missing Form 6 number / manifest number'), /Enter it when creating the certificate/);
});
test('Saved list contains only current PDFs, not deleted files or nested disclosures', () => {
  const {ui,controls} = harness();
  ui.app.state.certificate_inventory = {warnings:['Unreadable PDF; verify its number'],entries:[
    {path:'Certificates/0324_Acme_Mumbai_2.pdf',present:true,cods:['0324'],references:['REF-24']},
    {path:'Certificates/0340_Long_Name.pdf',present:false,cods:['0340']},
    {path:'Certificates/0350_Second_Company_Pune_1.pdf',present:true,cods:['0350']} ]};
  ui.renderFolderInventory();
  assert.notEqual(controls.get('#folder-inventory').open,true);
  assert.match(controls.get('#folder-inventory-summary').textContent,/Saved certificates \(2\).*1 file warning/);
  const list = controls.get('#folder-inventory-files');
  assert.match(text(list),/Acme Mumbai/);
  assert.match(text(list),/Second Company Pune/);
  assert.doesNotMatch(text(list),/0340|removed|reserved|\.pdf/);
  assert.equal(all(list,node => node.tag === 'li').length,2,'Each saved PDF should have its own spaced list card.');
  assert.equal(all(list,node => node.tag === 'details').length,0);
  assert.equal(controls.get('#show-past-numbers').textContent,'Past numbers (1)');
  ui.showSavedFile(ui.app.state.certificate_inventory.entries[0]);
  assert.match(text(controls.get('#saved-file-facts')),/0324_Acme_Mumbai_2.pdf/);
  assert.match(text(controls.get('#saved-file-facts')),/REF-24/);
  assert.equal(controls.get('#saved-file-modal').open,true);
});
test('Saved certificates label PDF sources in the shared folder and recognize legacy paths', () => {
  const {ui,controls} = harness();
  const entries = [
    {path:'Certificates/Created in App/0101_Client_1.pdf'},
    {relative_path:'Manually Added\\0102_Client_2.pdf'},
    {path:'Certificates/Auto Generated/0103_Client_1.pdf'},
    {path:'Certificates/0104_Client_1.pdf'},
    {path:'Certificates/0105_Client_1.pdf',managed_owner:{source_kind:'manual'}},
    {path:'Certificates/0106_Client_1.pdf',app_generated:true},
    {path:'Certificates/0107_Client_2.pdf',external:true,managed_owner:{source_kind:'workbook'}},
  ].map(entry => ({...entry,present:true}));
  ui.app.state.certificate_inventory = {entries};
  ui.renderFolderInventory();
  const labels = all(controls.get('#folder-inventory-files'), node => node.className === 'saved-file-source').map(node => node.textContent);
  assert.deepEqual(labels.sort(), ['Auto Generated','Auto Generated','Auto Generated','Auto Generated',
    'Existing PDF','Manually Added','Manually Added'].sort());
  ui.showSavedFile(entries[0]);
  assert.doesNotMatch(text(controls.get('#saved-file-facts')), /Created in App/);
  ui.showSavedFile(entries[1]);
  assert.match(text(controls.get('#saved-file-facts')), /Source Manually Added/);
});

test('Floating-point display artifacts are readable without changing source data', () => {
  const {ui} = harness();
  assert.equal(ui.readableNumber(100.19999999999999),'100.2');
  assert.equal(ui.readableNumber(0),'0');
  assert.equal(ui.readableNumber(null),'');
});
test('Material totals add entered line weights and quantity calculations, and flag incomplete or invalid lines', () => {
  const {ui,controls} = harness();
  assert.deepEqual(JSON.parse(JSON.stringify(ui.materialWeightTotal([
    {item:'Laptops',weight:'12.5'}, {item:'Monitors',quantity:'3',per_piece_weight:'2'},
  ]))), {state:'complete',total:18.5});
  assert.equal(ui.materialWeightTotal([{item:'Laptops'}]).state,'incomplete');
  assert.equal(ui.materialWeightTotal([{item:'Laptops',weight:'-1'}]).state,'invalid');
  ui.updateMaterialWeightTotal([{item:'Laptops',weight:'12'},{item:'Monitors',weight:'3'}]);
  assert.equal(controls.get('#material-weight-total').textContent,'Total weight: 15 kg');
});
test('Typing Certificate Weight applies it automatically and clearing it restores the source weight', () => {
  // Use a real harness editor state, then exercise the exact payload sent to the server.
  const h = harness(); const editable = editing(h); h.ui.wire();
  h.ui.app.templates = [{id:'standard',fields:[]}];
  editable.values.certificate_weight = '20';
  assert.equal(h.ui.detailsBody(editable).details.certificate_weight,undefined);
  assert.equal(h.ui.detailsBody(editable).use_certificate_weight,false);
  editable.values.certificate_weight = '8.5';
  h.ui.app.editor.fields.standard.weight_display = '99 KG';
  editable.dispatch('input',{target:{name:'certificate_weight',value:'8.5'}});
  const overridden = h.ui.detailsBody(editable);
  assert.equal(overridden.details.certificate_weight,'8.5');
  assert.equal(overridden.use_certificate_weight,true);
  assert.equal(overridden.details.certificate_fields.weight_display,'');
  editable.values.certificate_weight = '';
  editable.dispatch('input',{target:{name:'certificate_weight',value:''}});
  assert.equal(h.ui.detailsBody(editable).use_certificate_weight,false);
  assert.equal(h.ui.detailsBody(editable).details.certificate_weight,undefined);
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  assert.doesNotMatch(html,/Tick this|When the box is unticked|<span>Use Certificate Weight<\/span>/);
  assert.match(html,/name="use_certificate_weight" type="checkbox" hidden/);
});
test('Analytics cards count the standardized certificate statuses', () => {
  const {ui,controls} = harness();
  ui.app.state.rows = [
    {certificate_status:'ready'}, {certificate_status:'alert'},
    {certificate_status:'generated',outcome:'generated',system_alert:true,warnings:['Missing inward date; check the date in Excel.']},
    {certificate_status:'generated_manual'}, {certificate_status:'completed'}, {certificate_status:'blank'},
    {certificate_status:'na'},
  ];
  ui.summarize();
  assert.equal(controls.get('#count-ready').textContent,1);
  assert.equal(controls.get('#count-attention').textContent,3);
  controls.get('#status-filter').value = 'needs_attention';
  assert.equal(ui.matches(ui.app.state.rows[2]),true);
  assert.equal(controls.get('#count-generated').textContent,1);
  assert.equal(controls.get('#count-generated-externally').textContent,1);
  assert.equal(controls.get('#count-completed').textContent,1);
  assert.equal(controls.get('#count-overdue').textContent,0);
});
test('Existing workbook edits keep Excel unchanged by default and ask at save time', () => {
  const source = fs.readFileSync(path.join(__dirname,'../web/app.js'),'utf8');
  const editor = source.slice(source.indexOf('async function openDetails('),source.indexOf('function comparisonFields('));
  assert.match(editor,/namedItem\('edit_scope'\)\.value = .*'certificate_only'/);
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  assert.match(html,/type="hidden" name="edit_scope" value="certificate_only"/);
  assert.match(html,/No, keep Excel details/);
  assert.match(html,/Yes, update Excel/);
});
test('Saving a manual certificate draft calls the local register and confirms success', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({record_id:'manual-41'})})});
  const form = editing(h); form.values.source_kind = 'manual';
  vm.runInContext('setBusy = () => {}; setDetailsBusy = () => {};',h.context);
  await h.ui.saveDraft();
  assert.equal(h.requests.length,1); assert.equal(h.requests[0].endpoint,'/api/manual/save');
  assert.equal(h.requests[0].payload.source_kind,'manual');
  assert.equal(h.ui.app.editor.recordId,'manual-41');
  assert.equal(h.controls.get('#toast-message').textContent,'Draft saved in your local register.');
  assert.equal(h.controls.get('#toast').classList.contains('visible'),true);
});
test('Open Excel action passes the selected alert plant and row ID', async () => {
  const h = harness(); h.ui.app.state.rows = [{plant:'plant2',sn:'DUP-7',supplier:'Plant Two'}];
  await h.ui.handleAction({dataset:{index:'0',action:'open-workbook'}});
  assert.equal(h.requests.length,1); assert.equal(h.requests[0].endpoint,'/api/open-workbook');
  assert.deepEqual(JSON.parse(JSON.stringify(h.requests[0].payload)),{plant:'plant2',sn:'DUP-7'});
});
test('Optional form sections and bulk action are closed by default', () => {
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  for (const match of html.matchAll(/<details\b[^>]*>/g)) assert.doesNotMatch(match[0], /\bopen\b/);
  const editor = html.match(/<form id="details-form"[^>]*>[\s\S]*?<\/form>/)?.[0] || '';
  assert.equal((editor.match(/name="certificate_weight"/g)||[]).length,1);
  assert.equal((editor.match(/name="cod_no"/g)||[]).length,1);
  const queueHeader = html.match(/<thead><tr><th>SN<\/th><th>Company<\/th>.*?<\/thead>/s)?.[0] || '';
  assert.equal((queueHeader.match(/<th>/g)||[]).length,6);
});
test('Bulk creation payload follows the selected plant filter exactly', () => {
  const {ui,controls} = harness();
  const plantFilter = controls.get('#plant-filter');
  plantFilter.value = 'plant1';
  assert.deepEqual(JSON.parse(JSON.stringify(ui.bulkGenerationPayload())), {plants:['plant1']});
  plantFilter.value = 'plant2';
  assert.deepEqual(JSON.parse(JSON.stringify(ui.bulkGenerationPayload())), {plants:['plant2']});
  plantFilter.value = 'all';
  assert.deepEqual(JSON.parse(JSON.stringify(ui.bulkGenerationPayload())), {plants:['plant1','plant2']});
});

test('Company SN, inclusive received-date range, and sort choices work together', () => {
  const h=harness();
  const rows=[{plant:'plant1',sn:'10',supplier:'Beta',inward_date:'2026-10-02'},
    {plant:'plant1',sn:'2',supplier:'Alpha',inward_date:'2026-10-01'}];
  h.controls.get('#serial-filter').value='2';
  assert.equal(h.ui.matches(rows[0]),false);assert.equal(h.ui.matches(rows[1]),true);
  const from = h.controls.get('#received-date-filter'), to = h.controls.get('#received-to-filter');
  h.controls.get('#serial-filter').value='';from.value='2026-10-02';
  assert.equal(h.ui.matches(rows[0]),true);assert.equal(h.ui.matches(rows[1]),false);
  to.value='2026-10-02';
  assert.equal(h.ui.matches({...rows[0], inward_date:'2026-10-03'}),false);
  assert.equal(h.ui.matches(rows[0]),true);
  assert.equal(h.ui.matches({...rows[0], inward_date:''}),false);
  from.value='';to.value='01/10/26';
  assert.equal(h.ui.matches(rows[0]),false);assert.equal(h.ui.matches(rows[1]),true);
  from.value='02/10/26';h.ui.updateRecordFilterButton();
  assert.match(h.controls.get('#received-range-error').textContent,/on or before/);
  from.value='31/02/26';h.ui.updateRecordFilterButton();
  assert.match(h.controls.get('#received-range-error').textContent,/DD\/MM\/YY/);
  from.value='';to.value='';h.controls.get('#company-sort').value='serial';
  assert.equal([...rows].sort(h.ui.compareRows)[0].sn,'2');
  h.controls.get('#company-sort').value='received-newest';assert.equal([...rows].sort(h.ui.compareRows)[0].sn,'10');
  h.controls.get('#company-sort').value='company';assert.equal([...rows].sort(h.ui.compareRows)[0].supplier,'Alpha');
});

test('Linked HDD and its saved PDF open the same inward in Excel without guessing ambiguous clients', () => {
  const h=harness();
  const inward={source_kind:'workbook',plant:'plant1',sn:'12',supplier:'ABC',location:'Pune',sheet:'Inward',row:5};
  const hdd={source_kind:'manual',record_id:'hdd-1',plant:'plant1',certificate_type:'destruction',supplier:'ABC',location:'Pune',destruction_source:{plant:'plant1',case_id:'12'}};
  h.ui.app.state.rows=[inward,hdd];
  assert.deepEqual(JSON.parse(JSON.stringify(h.ui.excelTarget(hdd))),{plant:'plant1',sn:'12',sheet:'Inward',row:5});
  assert.equal(h.ui.excelTarget({managed_owner:{source_kind:'manual',plant:'plant1',key:'hdd-1'}}).sn,'12');
  assert.match(text(h.ui.makeActions(hdd,1)),/Open in Excel/);
  delete hdd.destruction_source;assert.equal(h.ui.excelTarget(hdd).sn,'12');
  h.ui.app.state.rows.push({...inward,sn:'13',row:6});assert.equal(h.ui.excelTarget(hdd),null);
});
test('Clean numeric displays preserve untouched precision in the generation payload', () => {
  const {ui} = harness();
  ui.app.editor = {template:'standard', fields:{standard:{}}, touched:new Set(), rawNumeric:{weight:'100.19999999999999'}, sourceSn:'1001', sourcePlant:'plant1'};
  const form = {values:{weight:'100.2',case_id:'1001',edit_scope:'certificate_only'}, elements:{namedItem:name => ({value:{plant:'plant1',source_kind:'workbook'}[name]})}};
  const original = ui.detailsBody(form);
  assert.equal(original.details.weight,'100.19999999999999');
  assert.equal(original.plant,'plant1');
  assert.equal(original.edit_scope,'certificate_only');
  assert.equal(original.source_sn,'1001');
  ui.app.editor.touched.add('weight'); form.values.weight = '99.5';
  assert.equal(ui.detailsBody(form).details.weight,'99.5');
});
test('Material editor sends every asset with its matching entered weight', () => {
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  assert.match(html, /id="add-material"[^>]*>\+ Add another material/);
  const {ui} = harness();
  ui.app.editor = {template:'standard',fields:{standard:{}},touched:new Set(),rawNumeric:{},rawDetails:{}};
  ui.app.templates = [{id:'standard',fields:[]}];
  const values = {plant:'plant1',source_kind:'workbook',edit_scope:'update_excel',item:'Laptops',quantity:'3',per_piece_weight:'7',weight:'21',
    material_1_item:'Monitors',material_1_quantity:'2',material_1_per_piece_weight:'6',material_1_weight:'12'};
  const form = {values,elements:{namedItem:name => ({value:values[name] || '',type:'text'})}};
  const result = ui.detailsBody(form);
  assert.deepEqual(JSON.parse(JSON.stringify(result.details.items)), [
    {item:'Laptops',quantity:'3',per_piece_weight:'7',weight:'21'},
    {item:'Monitors',quantity:'2',per_piece_weight:'6',weight:'12'},
  ]);
  assert.equal(result.details.weight,'21');
  const changes = ui.fieldDifferences({details:{items:[{item:'Laptops',weight:'21'}]}}, result);
  assert.equal(changes.some(change => change.key === 'items'),true);
});
test('Material + button adds a line and remove returns the form to one line', () => {
  const {ui,controls} = harness();
  ui.wire();
  ui.renderMaterialRows([{item:'Laptops',weight:'21'}]);
  controls.get('#add-material').dispatch('click');
  assert.equal(ui.app.materialRows.length,2);
  assert.equal(ui.app.materialRows[1].inputs.item.name,'material_1_item');
  ui.app.materialRows[0].inputs.item.value = 'Laptops';
  ui.app.materialRows[1].inputs.item.value = 'Monitors';
  ui.app.materialRows[1].element.children[0].children[1].dispatch('click');
  assert.equal(ui.app.materialRows.length,1);
  assert.equal(ui.app.materialRows[0].inputs.item.value,'Laptops');
});
test('Pages keep all rows accessible without a wall of 42 entries', () => {
  const {ui} = harness(); const items = Array.from({length:42}, (_,i) => i);
  const first = ui.visiblePage(items,1); const last = ui.visiblePage(items,99);
  assert.equal(first.rows.length,10); assert.equal(first.pages,5);
  assert.equal(last.page,5); assert.equal(last.rows.length,2); assert.equal(last.rows[0],40);
  assert.equal(ui.visiblePage(items,-1).page,1);
  assert.equal(ui.visiblePage([],1).rows.length,0);
  assert.equal(items.length,42);
});
test('Past numbers deduplicate padded IDs and exclude numbers on current PDFs', () => {
  const {ui} = harness();
  const numbers = ui.pastCertificateNumbers([
    {present:true,cods:['0009','0006'],current_cods:['0009']},
    {present:false,cods:['9','0006','6','0340']},
  ]);
  assert.equal(numbers.join(','),'0006,0340');
});
test('Saved PDF paging never hides current files behind deleted history', () => {
  const {ui,controls} = harness();
  ui.app.state.certificate_inventory = {entries:[
    ...Array.from({length:105},(_,i) => ({path:`Certificates/${i}_Removed.pdf`,present:false,cods:[String(i)]})),
    ...Array.from({length:12},(_,i) => ({path:`Certificates/${200+i}_Client.pdf`,present:true,cods:[String(200+i)]})),
  ]};
  ui.renderFolderInventory();
  assert.equal(controls.get('#folder-inventory-files').children.length,10);
  ui.app.savedFilesPage = 2; ui.renderFolderInventory();
  assert.equal(controls.get('#folder-inventory-files').children.length,2);
  assert.match(text(controls.get('#folder-inventory-files')),/210/);
});
test('Unnumbered manual file names keep the company name', () => {
  const {ui} = harness();
  assert.equal(ui.certificateFileTitle({path:'Certificates/Acme_Mumbai.pdf',cods:[]}), 'Acme Mumbai');
  assert.equal(ui.certificateFileTitle({path:'Certificates/CUSTOM-19_Acme_1.pdf',cods:['CUSTOM-19']}), 'Acme');
});
test('Past numbers are searchable without changing the inventory', () => {
  const {ui,controls} = harness();
  ui.app.state.certificate_inventory = {entries:[{path:'Certificates/old.pdf',present:false,cods:['0340','0341']}]};
  const before = JSON.stringify(ui.app.state);
  controls.get('#past-number-search').value = '0340'; ui.renderPastNumbers();
  assert.equal(controls.get('#past-number-list').children.length,1);
  assert.match(text(controls.get('#past-number-list')), /0340/);
  assert.equal(JSON.stringify(ui.app.state),before);
});

test('Saved HDD details show its own number while attached certificate numbers stay reserved', () => {
  const h = harness();
  const entry = {path:'Certificates/Auto Generated/0205_Company_Mumbai_hdd.pdf',present:true,
    cods:['0189','0205'],current_cods:['0205'],attachment_cods:['0189'],references:['REF-1'],warnings:[]};
  h.ui.showSavedFile(entry);
  const facts = text(h.controls.get('#saved-file-facts'));
  assert.match(facts,/Certificate number\s+0205/);
  assert.doesNotMatch(facts,/0189/);
  assert.doesNotMatch(facts,/Reference number|REF-1/);
  assert.equal(h.controls.get('#saved-file-warnings').hidden,true);
  assert.deepEqual(entry.cods,['0189','0205']);
});

test('Every tracked mutation carries a revision; explicit editor revision stays frozen', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({workspace_revision:'revision-2',state:{rows:[]}})})}); h.ui.app.workspaceRevision = 'revision-1';
  for (const endpoint of ['/api/create','/api/generate','/api/manual/save','/api/manual/generate','/api/manual/complete','/api/complete','/api/backup/restore','/api/destruction/request']) {
    await h.ui.api(endpoint, {});
    assert.equal(h.requests.at(-1).payload.expected_revision, endpoint === '/api/create' ? 'revision-1' : 'revision-2');
  }
  await h.ui.api('/api/create', {expected_revision:'editor-snapshot'});
  assert.equal(h.requests.at(-1).payload.expected_revision,'editor-snapshot');
  await h.ui.api('/api/state'); assert.equal(h.requests.at(-1).method,'GET');
  await h.ui.api('/api/backup/create', {}); assert.equal(h.requests.at(-1).payload.expected_revision,undefined);
});

test('Saving HDD opens the linked request and reopening it does not save again', async () => {
  const inwards = [
    {plant:'plant1',sn:'1',supplier:'First company'},
    {plant:'plant2',sn:'1',supplier:'Second company'},
  ];
  const requests = [];
  let revision = 'before-requests';
  const h = harness({fetch:async (endpoint, init) => {
    assert.equal(endpoint,'/api/destruction/request');
    const payload = JSON.parse(init.body);
    if (payload.expected_revision !== revision) return {ok:false,json:async()=>({error:'Excel has changed. Refresh and review the latest details.'})};
    const id = `request-${payload.plant}`;
    if (!requests.some(row => row.record_id === id)) {
      requests.push({record_id:id,plant:payload.plant,source_kind:'manual',certificate_type:'destruction',destruction_source:{plant:payload.plant,case_id:payload.sn}});
      revision = `with-${requests.length}-requests`;
    }
    return {ok:true,json:async()=>({record_id:id,state:{rows:[...inwards,...requests]},workspace_revision:revision})};
  }});
  vm.runInContext(`render=()=>{}; renderFolderInventory=()=>{}; updateSettingsDisplay=()=>{};
    setDetailsBusy=on=>{app.detailsBusy=on;};
    openDetails=async row=>{app.openedRequest=row.record_id; $('#details-modal').showModal();};
    editorDirty=()=>false;`,h.context);
  h.ui.setState({rows:inwards,workspace_revision:revision});
  for (const index of [0,1,0,1]) {
    const form=editing(h);
    form.values.plant=inwards[index].plant;form.values.case_id=inwards[index].sn;
    Object.assign(h.ui.app.editor,{sourceSn:inwards[index].sn,sourcePlant:inwards[index].plant,workspaceRevision:revision});
    h.controls.get('#destruction-needed').checked=false;
    h.ui.app.editor.baseline=h.ui.detailsBody(form);
    h.controls.get('#destruction-needed').checked=true;
    h.ui.app.openedRequest = null;
    if (!requests.some(row=>row.plant===inwards[index].plant)) await h.ui.saveInwardDetails();
    else await h.ui.openCompanionCertificate();
    assert.equal(h.controls.get('#details-modal').open,true);
    assert.equal(h.ui.app.openedRequest,`request-${inwards[index].plant}`);
    assert.equal(h.ui.app.state.workspace_revision,revision);
    assert.equal(h.ui.app.detailsBusy,false);
    assert.equal(h.ui.requestDetailsClose(),true);
  }
  assert.equal(requests.length,2,'Reopening must reuse each existing request.');
  assert.equal(h.requests.length,2,'An existing HDD request opens without another save or refresh.');
});

test('Review opens an active linked HDD instead of the N/A primary certificate', async () => {
  const h = harness();
  const inward = {source_kind:'workbook', plant:'plant2', sn:'1001', certificate_status:'na', certificate_type:'standard'};
  const request = {source_kind:'manual', plant:'plant2', record_id:'hdd-request', certificate_status:'ready',
    certificate_type:'destruction', destruction_source:{plant:'plant2', case_id:'1001'}};
  h.ui.app.state.rows = [inward, request];
  vm.runInContext('openDetails=async row=>{app.opened=row;};',h.context);
  const edit = {dataset:{index:'0', action:'edit-details'}};
  await h.ui.handleAction(edit);
  assert.equal(h.ui.app.opened, request);
  request.destruction_requested = false;
  await h.ui.handleAction(edit);
  assert.equal(h.ui.app.opened, inward);
  request.destruction_requested = true;
  inward.certificate_status = 'ready';
  await h.ui.handleAction(edit);
  assert.equal(h.ui.app.opened, inward);
  assert.equal(h.requests.length, 0);
});

test('A successful deletion uses the returned revision instead of an old nested report token', async () => {
  const h=harness({fetch:async()=>({ok:true,json:async()=>({workspace_revision:'after-delete',state:{rows:[],workspace_revision:'before-delete'}})})});
  h.ui.app.workspaceRevision='before-delete';
  const result=await h.ui.api('/api/clients/remove',{});
  assert.equal(h.ui.app.workspaceRevision,'after-delete');
  assert.equal(result.state.workspace_revision,'after-delete');
  await h.ui.api('/api/clients/remove-preview',{});
  assert.equal(h.requests.at(-1).payload.expected_revision,'after-delete');
});

test('A changed Excel file refreshes the queue and clears a stale deletion review without replaying it', async () => {
  let changed=false;
  const h=harness({fetch:async endpoint=>({ok:!changed || endpoint==='/api/check',json:async()=>
    endpoint==='/api/clients/remove-preview' ? {entry:{...removalRow,materials:1},review_token:'old-review',expected_revision:'old'}
    : endpoint==='/api/check' ? {workspace_revision:'latest',state:{rows:[{...removalRow,supplier:'Latest Excel company'}]}}
    : {error:'The files changed since you opened these details.'}})});
  vm.runInContext('setState=state=>{app.state=state;app.workspaceRevision=state.workspace_revision;};',h.context);
  h.ui.app.workspaceRevision='old';await h.ui.openInwardRemoval(removalRow);
  h.controls.get('#remove-inward-checked').checked=true;h.controls.get('#confirm-remove-inward').disabled=false;
  changed=true;await h.ui.confirmInwardRemoval();
  assert.equal(h.requests.filter(r=>r.endpoint==='/api/clients/remove').length,1);
  assert.equal(h.requests.at(-1).endpoint,'/api/check');
  assert.equal(h.ui.app.workspaceRevision,'latest');
  assert.equal(h.controls.get('#remove-inward-modal').open,false);
  assert.equal(h.ui.app.removalPending,null);
  assert.equal(h.ui.app.state.rows[0].supplier,'Latest Excel company');
});

test('A real external change still blocks a destruction request without opening or retrying it', async () => {
  const h = harness({fetch:async()=>({ok:false,json:async()=>({error:'Excel has changed. Refresh and review the latest details.'})})});
  h.ui.app.workspaceRevision='reviewed'; h.ui.app.state.rows=[{plant:'plant1',sn:'1'}];
  editing(h);h.ui.app.editor.workspaceRevision='reviewed';h.controls.get('#destruction-needed').checked=true;
  vm.runInContext('setDetailsBusy=on=>{app.detailsBusy=on;};',h.context);
  await h.ui.saveInwardDetails();
  assert.equal(h.requests.length,2); assert.equal(h.requests[0].payload.expected_revision,'reviewed');
  assert.equal(h.requests[1].endpoint,'/api/check','refresh the queue without retrying the rejected save');
  assert.equal(h.controls.get('#details-modal').open,true);
  assert.equal(h.ui.app.workspaceRevision,'reviewed'); assert.equal(h.ui.app.detailsBusy,false);
  assert.match(h.controls.get('#details-error').textContent,/Excel has changed/);
});

test('A late background check cannot roll back a newly opened data destruction request', async () => {
  let finish;
  const h = harness({fetch:async endpoint=>endpoint === '/api/workbook-changes'
    ? new Promise(resolve=>{finish=resolve;})
    : {ok:true,json:async()=>({record_id:'new-request',workspace_revision:'request-saved',state:{rows:[{record_id:'new-request',source_kind:'manual'}]}})}
  });
  vm.runInContext(`render=()=>{}; renderFolderInventory=()=>{}; updateSettingsDisplay=()=>{};
    setDetailsBusy=on=>{app.detailsBusy=on;}; openDetails=async()=>{};`,h.context);
  h.controls.get('#offline-screen').hidden=true; h.controls.get('#busy-overlay').hidden=true;
  h.ui.setState({rows:[{plant:'plant1',sn:'1'}],workspace_revision:'before-request'});
  const pending=h.ui.syncWorkbookChanges();
  editing(h);h.ui.app.editor.workspaceRevision='before-request';h.controls.get('#destruction-needed').checked=true;
  await h.ui.saveInwardDetails();
  finish({ok:true,json:async()=>({changed:true,workspace_revision:'before-request',state:{rows:[{plant:'plant1',sn:'1'}]}})});
  await pending;
  assert.equal(h.ui.app.workspaceRevision,'request-saved');
  assert.equal(h.ui.app.state.rows[0].record_id,'new-request');
});

test('Choosing a backup folder rereads rows before another certificate action; cancel does not', async () => {
  for (const cancelled of [false,true]) {
    const h = harness({fetch:async endpoint=>({ok:true,json:async()=>
      endpoint === '/api/backup/pick-destination' ? cancelled ? {cancelled:true} : {settings:{backup_destination:'D:\\Recovery'},workspace_revision:'settings-only'}
      : endpoint === '/api/check' ? {state:{rows:[{supplier:'Checked company'}]},workspace_revision:'checked-after-settings'}
      : {backups:[],workspace_revision:'settings-only'}
    })});
    vm.runInContext('setBusy=()=>{}; render=()=>{}; renderFolderInventory=()=>{}; updateSettingsDisplay=()=>{};',h.context);
    h.ui.app.workspaceRevision='before-settings';
    await h.ui.backupAction('pick-destination');
    assert.equal(h.ui.app.workspaceRevision,cancelled ? 'before-settings' : 'checked-after-settings');
    assert.equal(h.requests.filter(request=>request.endpoint==='/api/check').length,cancelled ? 0 : 1);
    assert.equal(h.ui.app.backupBusy,false);
  }
});

test('Change summary shows exact field differences and both tracking scopes', () => {
  const h = harness(); const form = editing(h);
  form.values.supplier = 'Acme Updated'; form.values.cod_no = 'CUSTOM-81';
  h.ui.app.editor.fields.standard.serial_no = 'SER-42';
  h.ui.renderChangeSummary();
  const summary = text(h.controls.get('#change-list'));
  assert.match(summary,/Before: Acme/); assert.match(summary,/After: Acme Updated/);
  assert.match(summary,/Certificate number.*Before: \(blank\).*After: CUSTOM-81/);
  assert.match(summary,/SER-42/); assert.doesNotMatch(summary,/Inward weight/);
  assert.match(h.controls.get('#change-scope').textContent,/PDF only.*Excel still records.*Generated/);
  form.values.edit_scope = 'update_excel'; h.ui.renderChangeSummary();
  assert.match(h.controls.get('#change-scope').textContent,/PDF.*updates this record in Plant 1 Excel/);
  form.values.source_kind = 'manual'; h.ui.renderChangeSummary();
  assert.match(h.controls.get('#change-scope').textContent,/local tracking.*No Excel row/);
});

test('Working drafts remain local and can be recovered without a backend write', () => {
  const h = harness(); const form = editing(h); form.values.supplier = 'Draft company';
  h.ui.wire();
  assert.equal(h.ui.editorDirty(),true); assert.equal(h.ui.saveWorkingDraft(),true);
  const saved = h.ui.readWorkingDraft();
  assert.equal(saved.values.supplier,'Draft company'); assert.equal(saved.editor.workspaceRevision,'revision-1');
  assert.equal(saved.editor.baseline.details.supplier,'Acme');
  assert.equal(saved.values.plant,'plant1'); assert.equal(h.requests.length,0);
  assert.equal(h.ui.requestDetailsClose(),false); assert.equal(h.controls.get('#unsaved-modal').open,true);
  h.controls.get('#keep-editing').dispatch('click');
  assert.equal(h.controls.get('#details-modal').open,true); assert.equal(h.controls.get('#unsaved-modal').open,false);
  assert.equal(h.ui.readWorkingDraft().values.supplier,'Draft company');
  h.controls.get('#details-modal').close();
  assert.equal(h.ui.offerDraftRecovery(),true); assert.equal(h.controls.get('#draft-recovery-modal').open,true);
  assert.equal(h.requests.length,0);
});

test('Draft restoration restores data and revision with zero network mutations', async () => {
  const h = harness(); const form = editing(h); form.values.supplier = 'Recovered company';
  form.values.item = 'Laptops'; form.values.weight = '21';
  form.values.material_1_item = 'Monitors'; form.values.material_1_weight = '12';
  h.ui.app.editor.touched.add('supplier'); h.ui.saveWorkingDraft();
  h.controls.get('#details-modal').close(); h.ui.offerDraftRecovery();
  vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {}; invalidatePreview = () => {};',h.context);
  form.values.supplier = 'Other company'; h.ui.app.workspaceRevision = 'revision-999';
  await h.ui.restoreWorkingDraft();
  assert.equal(form.values.supplier,'Recovered company');
  assert.equal(h.ui.app.editor.workspaceRevision,'revision-1');
  assert.equal(h.ui.app.editor.touched.has('supplier'),true);
  assert.equal(h.ui.app.materialRows.length,2);
  assert.equal(h.ui.app.materialRows[1].inputs.item.value,'Monitors');
  assert.equal(h.ui.app.materialRows[1].inputs.weight.value,'12');
  assert.equal(h.controls.get('#details-modal').open,true); assert.equal(h.ui.editorDirty(),true);
  assert.equal(h.requests.length,0);
});

test('Discarding startup recovery stays on the dashboard without opening a blank editor', () => {
  const h = harness(); const form = editing(h); form.values.supplier = 'Startup draft';
  h.ui.saveWorkingDraft(); h.controls.get('#details-modal').close(); h.ui.wire();
  h.ui.offerDraftRecovery();
  assert.equal(h.ui.app.draftRecovery.resumeEditor,false);
  assert.match(h.controls.get('#draft-recovery-description').textContent,/discard it to stay on the dashboard/);
  vm.runInContext('globalThis.editorRequests = []; openDetails = (...args) => editorRequests.push(args);',h.context);
  h.controls.get('#discard-working-draft').dispatch('click');
  assert.equal(h.ui.readWorkingDraft(),null); assert.equal(h.ui.app.draftRecovery,null);
  assert.equal(h.controls.get('#draft-recovery-modal').open,false);
  assert.equal(h.controls.get('#details-modal').open,false); assert.equal(h.context.editorRequests.length,0);
  assert.equal(h.requests.length,0);
});

test('When restore prompts are enabled, discarding the old draft opens the requested editor', async () => {
  for (const action of ['new','details','edit-details']) {
    const h = harness({showRestoreOptions:true}); const form = editing(h); form.values.supplier = 'Previous draft';
    h.ui.saveWorkingDraft(); h.controls.get('#details-modal').close(); h.ui.wire();
    const row = {supplier:'Selected company',plant:'plant2',sn:'2001'}; h.ui.app.state.rows = [row];
    if (action === 'new') h.controls.get('#new-button').dispatch('click');
    else await h.ui.handleAction({dataset:{index:'0',action}});
    assert.equal(h.ui.app.draftRecovery.resumeEditor,true);
    vm.runInContext('globalThis.editorRequests = []; openDetails = (...args) => editorRequests.push(args);',h.context);
    h.controls.get('#discard-working-draft').dispatch('click');
    assert.equal(h.context.editorRequests.length,1);
    const [selected, skipRecovery, startEditing] = h.context.editorRequests[0];
    assert.equal(selected,action === 'new' ? null : row); assert.equal(skipRecovery,true);
    assert.equal(startEditing || false,action === 'edit-details');
    assert.equal(h.ui.readWorkingDraft(),null); assert.equal(h.requests.length,0);
  }
});

test('Yes discards edits; failure to clear local draft keeps the editor open', () => {
  const h = harness(); const form = editing(h); form.values.supplier = 'Discard me';
  h.ui.wire(); h.ui.saveWorkingDraft(); h.ui.requestDetailsClose();
  h.controls.get('#discard-edits').dispatch('click');
  assert.equal(h.ui.readWorkingDraft(),null); assert.equal(h.requests.length,0);
  assert.equal(h.controls.get('#details-modal').open,false);
  const unavailable = harness(); const unavailableForm = editing(unavailable); unavailableForm.values.supplier = 'Discard me';
  unavailable.ui.wire(); unavailable.ui.saveWorkingDraft(); unavailable.ui.requestDetailsClose();
  vm.runInContext('localStorage.removeItem = () => { throw new Error("quota"); };',unavailable.context);
  unavailable.controls.get('#discard-edits').dispatch('click');
  assert.equal(unavailable.controls.get('#details-modal').open,true);
  assert.match(unavailable.controls.get('#toast-message').textContent,/could not remove/);
  assert.equal(unavailable.controls.get('#toast').classList.contains('visible'),true);
});

test('Escape, backdrop and tab closing protect dirty details; inside dialog whitespace does not close', () => {
  const h = harness(); const form = editing(h); h.ui.wire(); form.values.supplier = 'Unsaved';
  const dialog = h.controls.get('#details-modal'); let prevented = false;
  dialog.dispatch('cancel',{preventDefault() {prevented = true;}});
  assert.equal(prevented,true); assert.equal(h.controls.get('#unsaved-modal').open,true); assert.equal(dialog.open,true);
  h.controls.get('#unsaved-modal').close();
  for (const button of h.closeButtons.filter(button => button.dataset.close === 'details-modal')) {
    button.dispatch('click'); assert.equal(h.controls.get('#unsaved-modal').open,true); assert.equal(dialog.open,true);
    h.controls.get('#unsaved-modal').close();
  }
  dialog.dispatch('click',{clientX:50,clientY:50}); assert.equal(h.controls.get('#unsaved-modal').open,false);
  dialog.dispatch('click',{clientX:0,clientY:0}); assert.equal(h.controls.get('#unsaved-modal').open,true);
  prevented = false; const event = {preventDefault() {prevented = true;}};
  h.window.dispatch('beforeunload',event);
  assert.equal(prevented,true); assert.equal(h.ui.readWorkingDraft().values.supplier,'Unsaved');
  assert.equal(h.requests.length,0);
});

test('Dashboard closing waits for the unsaved-draft choice before stopping', async () => {
  const h = harness(); const form = editing(h); form.values.supplier = 'Unsaved';
  h.ui.wire();
  await h.ui.stopApp();
  assert.equal(h.requests.length,0); assert.equal(h.controls.get('#unsaved-modal').open,true);
  assert.notEqual(h.controls.get('#action-confirmation-modal').open,true);
  h.controls.get('#keep-editing').dispatch('click');
  assert.notEqual(h.controls.get('#action-confirmation-modal').open,true);
  await h.ui.stopApp();
  h.controls.get('#discard-edits').dispatch('click');
  assert.equal(h.requests.length,0); assert.equal(h.controls.get('#action-confirmation-modal').open,true);
  h.controls.get('#confirm-action').dispatch('click');
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(h.requests[0].endpoint,'/api/stop');
  assert.equal(h.ui.readWorkingDraft(),null);
});

test('Dashboard close requires in-page approval; Keep open, Escape, backdrop and X never stop it', async () => {
  const h = harness(); h.ui.wire();
  const dialog = h.controls.get('#action-confirmation-modal');
  const closeButtons = h.closeButtons.filter(button => button.dataset.close === dialog.id);
  for (const cancel of [() => closeButtons[0].dispatch('click'), () => closeButtons[1].dispatch('click'),
    () => dialog.dispatch('cancel',{preventDefault() {}}), () => dialog.dispatch('click',{clientX:0,clientY:0})]) {
    const pending = h.ui.stopApp();
    assert.equal(dialog.open,true); assert.equal(h.controls.get('#action-confirmation-title').textContent,'Close dashboard?');
    assert.equal(h.controls.get('#cancel-action').textContent,'Keep open');
    assert.equal(h.controls.get('#confirm-action').textContent,'Close dashboard now');
    assert.equal(h.controls.get('#cancel-action').focused,true);
    cancel(); await pending;
    assert.equal(h.requests.length,0); assert.equal(h.ui.app.actionConfirmation,null);
  }
  const pending = h.ui.stopApp(); h.controls.get('#confirm-action').dispatch('click'); await pending;
  assert.equal(h.requests.length,1); assert.equal(h.requests[0].endpoint,'/api/stop');
  assert.equal(h.controls.get('#offline-screen').hidden,false);
});

test('Mark Completed confirms company and consequences and preserves the reviewed revision', async () => {
  for (const manual of [false,true]) {
    const h = harness(); h.ui.wire();
    vm.runInContext('setBusy = () => {}; setState = state => { app.state = state; };',h.context);
    h.ui.app.workspaceRevision = 'reviewed-row';
    h.ui.app.state.rows = [{supplier:'Acme',sn:'1001',plant:'plant2',source_kind:manual ? 'manual' : 'workbook',record_id:'manual-81'}];
    const pending = h.ui.handleAction({dataset:{index:'0',action:'complete'}});
    assert.equal(h.requests.length,0);
    assert.match(h.controls.get('#action-confirmation-message').textContent,/Acme.*Completed.*stops its reminder/);
    assert.match(h.controls.get('#action-confirmation-message').textContent,manual ? /local certificate register/ : /Excel/);
    h.ui.app.workspaceRevision = 'changed-while-confirming';
    h.controls.get('#confirm-action').dispatch('click'); await pending;
    assert.equal(h.requests.length,1); assert.equal(h.requests[0].endpoint,manual ? '/api/manual/complete' : '/api/complete');
    assert.equal(h.requests[0].payload.expected_revision,'reviewed-row');
    assert.equal(h.requests[0].payload[manual ? 'record_id' : 'sn'],manual ? 'manual-81' : '1001');
  }
});

test('Canceled or overlapping action confirmations never mark a certificate Completed', async () => {
  const h = harness(); h.ui.wire();
  h.ui.app.state.rows = [{supplier:'Acme',sn:'1001',plant:'plant1'}];
  const rowButton = {dataset:{index:'0',action:'complete'}};
  const pending = h.ui.handleAction(rowButton);
  await h.ui.handleAction(rowButton);
  assert.equal(h.requests.length,0);
  h.controls.get('#action-confirmation-modal').close(); await pending;
  assert.equal(h.requests.length,0); assert.equal(h.ui.app.actionConfirmation,null);
  const source = fs.readFileSync(path.join(__dirname,'../web/app.js'),'utf8');
  assert.doesNotMatch(source,/\b(?:window\.)?(?:confirm|alert|prompt)\s*\(/);
  assert.match(source,/beforeunload/);
});

test('Undo Completed waits for Yes for both workbook and local certificates', async () => {
  for (const manual of [false,true]) {
    const h = harness(); h.ui.wire();
    vm.runInContext('setBusy = () => {}; setState = state => { app.state = state; };',h.context);
    h.ui.app.workspaceRevision = 'reviewed-completion';
    const row = {supplier:'Acme',sn:'1001',plant:'plant2',source_kind:manual ? 'manual' : 'workbook',
      record_id:'manual-81',certificate_status:'completed',certificate_no:'0801'};
    h.ui.app.state.rows = [row];
    const button = {dataset:{index:'0',action:'undo-completed'}};
    let pending = h.ui.handleAction(button);
    assert.equal(h.requests.length,0);
    assert.match(h.controls.get('#action-confirmation-message').textContent,/Acme.*0801.*send queue/);
    assert.equal(h.controls.get('#confirm-action').textContent,'Yes, undo');
    assert.equal(h.controls.get('#cancel-action').textContent,'No, keep Completed');
    h.closeButtons.find(control => control.dataset.close==='action-confirmation-modal').dispatch('click');
    await pending;
    assert.equal(h.requests.length,0);
    assert.equal(row.certificate_status,'completed');

    pending = h.ui.handleAction(button);
    h.ui.app.workspaceRevision = 'changed-during-confirmation';
    h.controls.get('#confirm-action').dispatch('click'); await pending;
    assert.equal(h.requests.length,1);
    assert.equal(h.requests[0].endpoint,manual ? '/api/manual/complete' : '/api/complete');
    assert.equal(h.requests[0].payload.undo_completed,true);
    assert.equal(h.requests[0].payload.expected_revision,'reviewed-completion');
    assert.equal(h.requests[0].payload[manual ? 'record_id' : 'sn'],manual ? 'manual-81' : '1001');
  }
});

test('A late background response cannot replace Completed with an older Generated status or revision', async () => {
  let finish;
  const h = harness({fetch:async () => new Promise(resolve => {finish=resolve;})});
  vm.runInContext('render=()=>{}; renderFolderInventory=()=>{}; updateSettingsDisplay=()=>{};',h.context);
  h.controls.get('#offline-screen').hidden=true;h.controls.get('#busy-overlay').hidden=true;
  h.ui.setState({rows:[{plant:'plant1',sn:'1001',certificate_status:'generated'}],workspace_revision:'old'});
  const pending=h.ui.syncWorkbookChanges();
  h.ui.setState({rows:[{plant:'plant1',sn:'1001',certificate_status:'completed'}],workspace_revision:'completed'});
  finish({ok:true,json:async()=>({changed:true,state:{rows:[{plant:'plant1',sn:'1001',certificate_status:'generated'}],workspace_revision:'old'},workspace_revision:'old'})});
  await pending;
  assert.equal(h.ui.app.state.rows[0].certificate_status,'completed');
  assert.equal(h.ui.app.workspaceRevision,'completed');
});

test('Warning continuation needs one click, uses the response revision and token, and keeps names/consequences', async () => {
  let attempt = 0;
  const h = harness({fetch:async () => ({ok:true,json:async () => ++attempt === 1
    ? {requires_confirmation:true,workspace_revision:'review-revision',review_token:'review-81',warnings:[{plant:'plant2',supplier:'Acme',warnings:['Replacement weight will be used. Verify before sending.']}]}
    : {workspace_revision:'generated-revision',created:1,certificates:[{filename:'actual.pdf',supplier:'Acme',certificate_no:'0082',plant:'plant2'}]}})});
  vm.runInContext('setBusy = () => {};',h.context);
  h.ui.app.workspaceRevision = 'old-revision';
  await h.ui.submitGeneration({sn:'1001'});
  assert.equal(h.requests[0].payload.expected_revision,'old-revision');
  assert.match(text(h.controls.get('#warning-list')),/Plant 2.*Acme.*Replacement weight will be used.*Verify before sending/);
  await h.ui.confirmWarnings();
  assert.equal(h.requests.length,2);
  const payload = h.requests[1].payload;
  assert.equal(payload.expected_revision,'review-revision'); assert.equal(payload.review_token,'review-81');
  assert.equal(payload.override_reason,'Warnings reviewed; user chose Continue anyway.'); assert.equal(payload.confirmed,true);
  assert.equal(h.controls.get('#success-modal').open,true);
});

test('Empty warning confirmations record that no warnings were reported', async () => {
  const h = harness(); vm.runInContext('setBusy = () => {};',h.context);
  h.ui.showWarnings([],{endpoint:'/api/generate',payload:{expected_revision:'revision-1'},reviewToken:'no-warnings',fromDetails:false});
  await h.ui.confirmWarnings();
  assert.equal(h.requests.length,1); assert.equal(h.requests[0].payload.override_reason,'No warnings were reported; confirmed creation.');
});

test('Stale workspace errors keep the edited form and recoverable draft open', async () => {
  const h = harness({fetch:async () => ({ok:false,status:409,json:async () => ({error:'Workspace changed. Refresh and review.'})})});
  const form = editing(h); form.values.supplier = 'Keep me'; h.ui.saveWorkingDraft();
  vm.runInContext('setBusy = () => {}; setDetailsBusy = on => { app.detailsBusy = on; };',h.context);
  await h.ui.submitGeneration(h.ui.detailsBody(form),false,'/api/create',true);
  assert.equal(h.controls.get('#details-modal').open,true); assert.equal(h.ui.readWorkingDraft().values.supplier,'Keep me');
  assert.equal(h.controls.get('#toast').classList.contains('visible'),false);
  assert.match(text(h.controls.get('#details-error')),/Refresh.*review/);
  assert.equal(h.controls.get('#details-error').hidden,false);
});

test('Success shows actual server-issued certificates and safely encoded local PDF links', () => {
  const h = harness();
  h.ui.showGenerationSuccess({created:2,state:{recovery_note:'Close the previous PDF, then refresh to finish.',
    setup_warning:'Data Destruction request could not be updated. Open Review & edit and choose Save request.'},
    certificates:[{filename:'0082_Acme_Long_Recovery_Services_Mumbai_India_1.pdf',certificate_no:'0082',supplier:'Acme',plant:'plant1'},
    {filename:'CUSTOM-84_Second_Company_Pune_2.pdf',certificate_no:'CUSTOM-84',supplier:'Second',plant:'plant2',pdf_url:'/api/certificates/CUSTOM-84_Second_Company_Pune_2.pdf'}]});
  const list = h.controls.get('#created-certificates');
  assert.match(text(list),/0082.*CUSTOM-84/); assert.doesNotMatch(text(list),/Suggested COD/);
  assert.equal(all(list,node => node.tag === 'li').length,2);
  assert.equal(all(list,node => node.tag === 'button' && node.textContent === 'Open PDF').length,2);
  assert.deepEqual(all(list,node => node.className === 'saved-pdf-name').map(node => node.textContent),
    ['0082_Acme_Long_Recovery_Services_Mumbai_India_1.pdf','CUSTOM-84_Second_Company_Pune_2.pdf']);
  const css = fs.readFileSync(path.join(__dirname,'../web/styles.css'),'utf8');
  assert.match(css,/\.saved-files\s*\{[^}]*gap:\s*18px/);
  assert.match(css,/\.saved-file-copy span\s*\{[^}]*overflow-wrap:\s*anywhere/);
  assert.match(css,/\.saved-file-copy \.saved-pdf-name\s*\{[^}]*border-top:\s*1px solid/);
  const links = all(list,node => node.tag === 'a'); assert.equal(links.length,2);
  assert.equal(links[0].href,'http://127.0.0.1:8000/api/certificates/0082_Acme_Long_Recovery_Services_Mumbai_India_1.pdf');
  assert.equal(h.controls.get('#success-modal').open,true);
  assert.match(h.controls.get('#success-description').textContent,/Close the previous PDF, then refresh to finish/);
  assert.match(h.controls.get('#success-description').textContent,/Data Destruction request.*Save request/);
  for (const certificate of [{filename:'../secret.pdf'},{filename:'a.pdf',pdf_url:'https://example.com/api/certificates/a.pdf'},
    {filename:'a.pdf',pdf_url:'/api/preview/a.pdf'},{filename:'a.pdf',pdf_url:'/api/certificates/b.pdf'}]) assert.equal(h.ui.certificateUrl(certificate),null);
});

test('Batch results show every unfinished company and explain how to retry it', async () => {
  const unissued = [{plant:'plant2',sn:'21',supplier:'Northstar',detail:'The certificate output folder became read-only.'}];
  const h = harness({fetch:async () => ({ok:true,json:async () => ({created:0,unissued})})});
  h.ui.showGenerationSuccess({created:1,certificates:[{filename:'0082_Acme.pdf',supplier:'Acme'}],unissued});
  assert.equal(h.controls.get('#success-title').textContent,'Some certificates were not created');
  assert.match(h.controls.get('#success-description').textContent,/1 of 2 certificates saved.*Refresh.*only the unfinished companies/);
  assert.match(text(h.controls.get('#created-certificates')),/Acme/);
  assert.match(text(h.controls.get('#unissued-list')),/Northstar.*Plant 2.*SN 21.*read-only/);
  h.controls.get('#success-modal').close();
  await h.ui.submitGeneration({selections:[{plant:'plant2',sn:'21'}]});
  assert.equal(h.controls.get('#success-modal').open,true);
  assert.equal(h.controls.get('#success-title').textContent,'No certificates were created');
  assert.match(h.controls.get('#success-description').textContent,/0 of 1/);
  assert.match(text(h.controls.get('#unissued-list')),/Northstar.*read-only/);
});

test('Open PDF and filename clicks open the local document and select it; Open folder only selects it', async () => {
  for (const folder of ['Auto Generated','Created in App','Manually Added','']) {
    const h = harness(); h.ui.wire();
    const filename = '0082_Acme Mumbai_2.pdf';
    const relativePath = folder ? `${folder}/${filename}` : filename;
    const row = {certificate_status:'generated',supplier:'Acme',plant:'plant2',pdf_filename:filename,pdf_relative_path:relativePath};
    all(h.ui.makeActions(row,0),node => node.textContent === 'Open PDF')[0].dispatch('click');
    const entry = {present:true,path:`Certificates/${relativePath}`,cods:['0082']};
    h.ui.app.workspaceView = 'files'; h.ui.app.state.certificate_inventory = {entries:[entry]};
    h.ui.renderFolderInventory();
    all(h.controls.get('#folder-inventory-files'),node => node.textContent === 'Open PDF')[0].dispatch('click');
    h.ui.showSavedFile(entry); h.controls.get('#saved-file-open').dispatch('click');
    h.ui.showGenerationSuccess({created:1,certificates:[{filename,relative_path:relativePath,certificate_no:'0082',plant:'plant2'}]});
    const list = h.controls.get('#created-certificates');
    all(list,node => node.textContent === 'Open PDF')[0].dispatch('click');
    let prevented = false;
    all(list,node => node.className === 'saved-pdf-name')[0].dispatch('click',{preventDefault(){prevented = true;}});
    h.controls.get('#success-open-folder').dispatch('click');
    await Promise.resolve();
    assert.equal(prevented,true,'The filename must not navigate to a browser PDF.');
    const opens = h.requests.filter(request => request.endpoint === '/api/open-certificate');
    assert.equal(opens.length,6);
    for (const [index, request] of opens.entries()) assert.deepEqual(request.payload,
      {relative_path:relativePath,reveal:true,open_pdf:index < 5});
    assert.equal(h.requests.some(request => request.endpoint === '/api/open-folder'),false);
  }
});

test('Backup destination is chosen server-side and restoration sends explicit confirmation', async () => {
  const h = harness({fetch:async endpoint => ({ok:true,json:async () => ({workspace_revision:'backup-revision',...(endpoint === '/api/backup/list'
    ? {backup_destination:'D:\\Recovery',backups:[{backup_id:'snapshot-81',created_at:'2026-10-01 10:30',valid:true}]}
    : endpoint === '/api/backup/restore' ? {settings:{},state:{rows:[]},message:'Restored with recovery snapshot',restored:{warnings:['Verify certificate continuity before issuing.']}}
    : endpoint === '/api/backup/create' ? {backup:{backup_id:'snapshot-81',path:'D:\\Recovery\\snapshot.zip'},backup_destination:'D:\\Recovery',health:'Backup verified.'}
    : {backup_destination:'D:\\Recovery',settings:{backup_destination:'D:\\Recovery'}})})})});
  h.ui.app.workspaceRevision = 'backup-revision';
  vm.runInContext('setBusy = () => {}; updateSettingsDisplay = () => {}; setState = state => { app.state = state; };',h.context);
  await h.ui.backupAction('pick-destination'); await h.ui.backupAction('create');
  assert.deepEqual(h.requests.find(request => request.endpoint === '/api/backup/pick-destination').payload,{});
  assert.deepEqual(h.requests.find(request => request.endpoint === '/api/backup/create').payload,{});
  h.controls.get('#details-modal').open = false;
  await h.ui.openBackupRestore(); assert.equal(h.controls.get('#backup-restore-modal').open,true);
  h.controls.get('#backup-selection').value = 'snapshot-81'; await h.ui.restoreBackup();
  assert.deepEqual(h.requests.find(request => request.endpoint === '/api/backup/restore').payload,{backup_id:'snapshot-81',confirmed:true,expected_revision:'backup-revision'});
  assert.match(h.controls.get('#backup-status').textContent,/recovery snapshot/);
  assert.match(h.controls.get('#backup-status').textContent,/Verify certificate continuity/);
  assert.match(h.controls.get('#backup-destination').textContent,/D:\\Recovery/);
  assert.equal(h.controls.get('#backup-health').textContent,'Backup verified.');
});

test('Missing and damaged backups cannot be selected for restoration', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({backups:[{backup_id:'broken',valid:false,error:'Missing archive'}]})})});
  vm.runInContext('setBusy = () => {};',h.context);
  await h.ui.openBackupRestore();
  assert.equal(h.controls.get('#backup-selection').children[0].disabled,true);
  assert.notEqual(h.controls.get('#backup-restore-modal').open,true);
  assert.match(h.controls.get('#backup-status').textContent,/missing or damaged/);
});

test('Backup options show localized date/time, file counts and pre-restore purpose; damaged copies stay disabled', async () => {
  const timestamp = '2026-10-01T08:23:16.515748+00:00';
  const backups = [
    {backup_id:'recovery',created_at:timestamp,file_count:24,purpose:'pre-restore',valid:true},
    {backup_id:'manual',created_at:timestamp,file_count:1,purpose:'manual',valid:true},
    {backup_id:'damaged',created_at:timestamp,file_count:0,purpose:'pre-restore',valid:false},
    {backup_id:'undated',created_at:'invalid date',valid:true},
  ];
  const h = harness({fetch:async () => ({ok:true,json:async () => ({backups})})});
  vm.runInContext('setBusy = () => {};',h.context);
  await h.ui.openBackupRestore();
  const options = h.controls.get('#backup-selection').children;
  const d=new Date(timestamp); const localized=`${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getFullYear()).slice(-2)} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
  assert.ok(options[0].textContent.startsWith(localized));
  assert.match(options[0].textContent,/24 files.*Pre-restore recovery copy/);
  assert.match(options[1].textContent,/1 file · Manual backup/);
  assert.match(options[2].textContent,/0 files.*Pre-restore recovery copy.*Unavailable/);
  assert.equal(options[2].disabled,true); assert.equal(options[0].disabled,false);
  assert.equal(options[3].textContent,'Date not recorded');
  for (const option of options) assert.doesNotMatch(option.textContent,/2026-10-01T|\+00:00|515748/);
  assert.equal(h.controls.get('#backup-selection').value,'recovery');
});

test('Activity history shows human field changes, accepted warnings and recovery facts', async () => {
  const record = {cod:'0082',state:'complete',plant:'plant2',source_kind:'workbook',operator:'Local operator',completed_at:'2026-10-01T04:30:00Z',
    before:{supplier:'Acme',items:[{weight:100}],certificate_fields:{serial_no:'OLD'}},after:{supplier:'Acme revised',items:[{weight:95}],certificate_fields:{serial_no:'NEW'}},
    accepted_warnings:['Replacement weight used. Check before sending.'],override_reason:'Confirmed by customer',pdf_path:'Certificates/0082_Acme.pdf',previous_pdfs:[{path:'Certificates/old.pdf'}]};
  const h = harness({fetch:async () => ({ok:true,json:async () => ({history:[record]})})});
  await h.ui.loadHistory();
  assert.equal(h.requests[0].endpoint,'/api/history'); assert.equal(h.requests[0].method,'GET');
  assert.match(text(h.controls.get('#history-list')),/0082.*Acme revised.*Generated/);
  h.ui.showHistoryDetail(record);
  const detail = text(h.controls.get('#history-detail'));
  assert.match(detail,/Local operator.*0082_Acme.pdf.*Confirmed by customer/);
  assert.match(detail,/Before: Acme.*After: Acme revised/);
  assert.match(detail,/Item 1.*Before: 100.*After: 95/);
  assert.match(detail,/Before: OLD.*After: NEW/);
  assert.match(detail,/Replacement weight used.*Check before sending/);
  assert.match(detail,/previous PDF retained/); assert.doesNotMatch(detail,/\{"/);
  assert.equal(h.controls.get('#history-detail-modal').open,true); assert.equal(h.requests.length,1);
});

test('New manual history skips absent, null, empty and whitespace-only blank changes', () => {
  const h = harness();
  const record = {source_kind:'manual',before:{quantity:null,reference:'',certificate_fields:{weight_display:null}},
    after:{supplier:'New company',form6:'',quantity:'',case_id:null,item:' \t ',reference:'\n',certificate_fields:{weight_display:'',company_display:'',optional_line:null}}};
  const original = JSON.stringify(record);
  const changes = h.ui.historyDifferences(record);
  assert.equal(changes.length,1); assert.equal(changes[0].key,'supplier');
  assert.equal(changes[0].before,'(blank)'); assert.equal(changes[0].after,'New company');
  h.ui.showHistoryDetail(record);
  const rendered = text(h.controls.get('#history-detail'));
  assert.match(rendered,/Before: \(blank\).*After: New company/);
  assert.doesNotMatch(rendered,/Before: \(blank\)\s+After: \(blank\)/);
  assert.equal(JSON.stringify(record),original);
});

test('History uses friendly printed-field and template labels without losing edits, removals or zero', () => {
  const h = harness();
  h.ui.app.templates = [{id:'sita',fields:[{key:'customer_account',label:'Customer account'}]}];
  const record = {before:{certificate_type:'sita',supplier:'Old company',quantity:5,form6:'F6-81',certificate_fields:{company_display:'Old printed name',weight_display:'100 kg',customer_account:'Account A'}},
    after:{certificate_type:'sita',supplier:'Revised company',quantity:0,form6:'',certificate_fields:{company_display:'New printed name',weight_display:'95 kg',customer_account:'Account B'}}};
  h.ui.showHistoryDetail(record);
  const rendered = text(h.controls.get('#history-detail'));
  assert.match(rendered,/Company line.*Before: Old printed name.*After: New printed name/);
  assert.match(rendered,/Weight label.*Before: 100 kg.*After: 95 kg/);
  assert.match(rendered,/Customer account.*Before: Account A.*After: Account B/);
  assert.match(rendered,/Before: Old company.*After: Revised company/);
  assert.match(rendered,/Quantity.*Before: 5.*After: 0/);
  assert.match(rendered,/Form 6 number.*Before: F6-81.*After: \(blank\)/);
  assert.doesNotMatch(rendered,/certificate[_ ]fields|company_display|weight_display|customer_account/i);
  assert.equal(h.ui.historyDifferences(record).length,6);
});

test('Entirely blank history reports no field changes instead of blank-to-blank rows', () => {
  const h = harness();
  h.ui.showHistoryDetail({before:{supplier:null},after:{supplier:'',certificate_fields:{company_display:' ',weight_display:null}}});
  const rendered = text(h.controls.get('#history-detail'));
  assert.match(rendered,/No field changes recorded/); assert.doesNotMatch(rendered,/Before:|After:/);
});

test('Reverting edits clears an obsolete autosaved draft instead of restoring stale values', () => {
  const h = harness(); const form = editing(h);
  form.values.supplier = 'Temporary edit'; h.ui.invalidatePreview();
  assert.equal(h.ui.readWorkingDraft().values.supplier,'Temporary edit');
  form.values.supplier = 'Acme'; h.ui.invalidatePreview();
  assert.equal(h.ui.editorDirty(),false); assert.equal(h.ui.readWorkingDraft(),null);
  assert.equal(h.requests.length,0);
});

test('Warning review has no reason input and records acceptance automatically', async () => {
  const h = harness(); vm.runInContext('setBusy = () => {};',h.context);
  h.ui.showWarnings([{supplier:'Acme',warnings:['Check weight.']}],{endpoint:'/api/generate',payload:{},reviewToken:'review-token'});
  assert.equal(h.controls.has('#override-reason'),false);
  assert.equal(h.controls.has('#warning-reason-error'),false);
  await h.ui.confirmWarnings();
  assert.equal(h.requests.length,1);
  assert.equal(h.requests[0].payload.override_reason,'Warnings reviewed; user chose Continue anyway.');
});

test('Guarded PDF attachment selection sends the editor snapshot and adopts its new revision', async () => {
  const h = harness({fetch:async endpoint => ({ok:true,json:async () => endpoint === '/api/pick-annexure'
    ? {workspace_revision:'with-attachment',annexure:{id:'attachment-1',name:'List.pdf',pages:2}}
    : {workspace_revision:'with-attachment',state:{rows:[]}}})});
  const form = editing(h);
  vm.runInContext('setBusy = () => {}; setDetailsBusy = on => { app.detailsBusy = on; };',h.context);
  await h.ui.pickAnnexure();
  assert.equal(h.requests[0].payload.expected_revision,'revision-1');
  assert.equal(h.ui.detailsBody(form).expected_revision,'with-attachment');
  assert.match(h.controls.get('#annexure-name').textContent,/List.pdf.*2 pages/);
  assert.equal(h.ui.readWorkingDraft().editor.workspaceRevision,'with-attachment');
  assert.equal(h.requests[1].endpoint,'/api/check');
  assert.equal(h.ui.app.workspaceRevision,'with-attachment');
});

test('Raster first-page preview is default; full embedded PDF loads only on request', async () => {
  const id = 'a'.repeat(32);
  const h = harness({fetch:async () => ({ok:true,json:async () => ({preview_url:`/previews/${id}.pdf`,preview_image_url:`/previews/${id}.png`,certificate_no:'0091',warnings:[]})})});
  editing(h); await h.ui.previewDetails();
  assert.equal(h.controls.get('#preview-image').src,`http://127.0.0.1:8000/previews/${id}.png`);
  assert.equal(h.controls.get('#preview-image').hidden,false); assert.equal(h.controls.get('#preview-frame').hidden,true);
  assert.equal(h.controls.get('#preview-frame').src,undefined);
  assert.equal(h.controls.get('#preview-status').textContent,'Preview ready');
  assert.match(h.controls.get('#preview-number').textContent,/COD 0091/);
  h.ui.toggleFullPdfPreview();
  assert.equal(h.controls.get('#preview-frame').hidden,false); assert.equal(h.controls.get('#preview-frame').src,h.controls.get('#preview-link').href);
  assert.equal(h.controls.get('#preview-full-pdf').attributes['aria-expanded'],'true');
  h.ui.invalidatePreview();
  assert.equal(h.controls.get('#preview-image').src,undefined); assert.equal(h.controls.get('#preview-image').hidden,true);
  assert.equal(h.controls.get('#preview-frame').src,undefined); assert.equal(h.controls.get('#preview-document').hidden,true);
});

test('Missing, unsafe or failed raster images retain a working full-PDF option without a black iframe', async () => {
  for (const previewImage of [undefined,'https://example.com/previews/a.png','/api/certificates/not-an-image.pdf','http://[']) {
    const h = harness({fetch:async () => ({ok:true,json:async () => ({preview_url:`/previews/${'b'.repeat(32)}.pdf`,preview_image_url:previewImage,warnings:[]})})});
    editing(h); await h.ui.previewDetails();
    assert.equal(h.controls.get('#preview-image').hidden,true); assert.equal(h.controls.get('#preview-image-note').hidden,false);
    assert.equal(h.controls.get('#preview-frame').hidden,true); assert.match(h.controls.get('#preview-link').href,/\.pdf$/);
  }
  const h = harness(); editing(h); h.ui.wire();
  h.controls.get('#preview-image').hidden = false;
  h.controls.get('#preview-image').dispatch('error');
  assert.equal(h.controls.get('#preview-image').hidden,true); assert.equal(h.controls.get('#preview-image-note').hidden,false);
  assert.match(h.controls.get('#preview-status').textContent,/could not be displayed.*full PDF/);
});

test('Attached PDF pages appear in the preview with Previous and Next and clear after editing', async () => {
  const id='e'.repeat(32),pages=[2,3].map(page=>`/previews/${id}/page/${page}.png`);
  const h=harness({fetch:async()=>({ok:true,json:async()=>({preview_url:`/previews/${id}.pdf`,preview_image_url:`/previews/${id}.png`,preview_attachment:{name:'Device list.pdf',pages},warnings:[]})})});
  editing(h);h.ui.wire();await h.ui.previewDetails();
  assert.equal(h.controls.get('#preview-attachment').hidden,false);
  assert.equal(h.controls.get('#preview-attachment-name').textContent,'Device list.pdf');
  assert.equal(h.controls.get('#preview-attachment-count').textContent,'Page 1 of 2');
  assert.equal(h.controls.get('#preview-attachment-previous').disabled,true);
  const first=h.controls.get('#preview-attachment-image').children[0];first.dispatch('load');
  assert.equal(first.hidden,false);assert.match(first.src,/page\/2\.png$/);
  h.controls.get('#preview-attachment-next').dispatch('click');
  const second=h.controls.get('#preview-attachment-image').children[0];second.dispatch('load');
  assert.match(second.src,/page\/3\.png$/);assert.equal(h.controls.get('#preview-attachment-count').textContent,'Page 2 of 2');
  assert.equal(h.controls.get('#preview-attachment-next').disabled,true);
  first.dispatch('error');assert.equal(h.controls.get('#preview-attachment-status').hidden,true,'old page failures cannot replace the current page');
  h.controls.get('#preview-attachment-previous').dispatch('click');
  assert.equal(h.controls.get('#preview-attachment-count').textContent,'Page 1 of 2');
  h.ui.invalidatePreview();
  assert.equal(h.controls.get('#preview-attachment').hidden,true);assert.equal(h.controls.get('#preview-attachment-image').children.length,0);
  second.dispatch('load');assert.equal(h.controls.get('#preview-attachment').hidden,true);
});

test('Attachment preview failures offer retry and unsafe or mismatched page URLs never load', async () => {
  const id='f'.repeat(32);
  for(const pages of [[`https://example.com/previews/${id}/page/2.png`],[`/previews/${'a'.repeat(32)}/page/2.png`],[`/previews/${id}/page/2.png?other=1`]]) {
    const h=harness({fetch:async()=>({ok:true,json:async()=>({preview_url:`/previews/${id}.pdf`,preview_attachment:{name:'List.pdf',pages},warnings:[]})})});
    editing(h);await h.ui.previewDetails();
    assert.equal(h.controls.get('#preview-attachment-image').children.length,0);
    assert.match(h.controls.get('#preview-attachment-status').textContent,/full preview/);
  }
  const h=harness({fetch:async()=>({ok:true,json:async()=>({preview_url:`/previews/${id}.pdf`,preview_attachment:{name:'List.pdf',pages:[`/previews/${id}/page/2.png`]},warnings:[]})})});
  editing(h);h.ui.wire();await h.ui.previewDetails();
  h.controls.get('#preview-attachment-image').children[0].dispatch('error');
  assert.equal(h.controls.get('#preview-attachment-retry').hidden,false);
  h.controls.get('#preview-attachment-retry').dispatch('click');
  h.controls.get('#preview-attachment-image').children[0].dispatch('load');
  assert.equal(h.controls.get('#preview-attachment-status').hidden,true);assert.equal(h.controls.get('#preview-attachment-retry').hidden,true);
});

test('UI keeps a responsive paired editor without avatars, stat icons or editable backup paths', () => {
  const {ui} = harness();
  assert.equal(all(ui.makeClientCell({supplier:'Acme'},0),node => node.className === 'client-avatar').length,0);
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  const css = fs.readFileSync(path.join(__dirname,'../web/styles.css'),'utf8');
  assert.doesNotMatch(html,/class="stat-icon"/);
  assert.match(html,/class="editor-workspace"/); assert.match(html,/id="preview-frame"/);
  assert.match(css,/@media \(max-width: 760px\)[\s\S]*\.editor-workspace \{ grid-template-columns: minmax\(0, 1fr\)/);
  const restoreDescription = html.match(/<p id="backup-restore-description">([^<]+)<\/p>/)?.[1] || '';
  assert.match(restoreDescription,/workbooks.*certificate records.*PDFs.*settings/);
  assert.match(restoreDescription,/recovery copy.*saved first/);
  assert.match(restoreDescription,/recovery archive.*certificate numbers stay reserved/);
  assert.doesNotMatch(html,/<input[^>]*(?:backup_destination|backup_path)/);
});

test('left navigation is removed and Settings stays visible in the main top bar', () => {
  const h = harness();
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  const css = fs.readFileSync(path.join(__dirname,'../web/styles.css'),'utf8');
  const app = fs.readFileSync(path.join(__dirname,'../web/app.js'),'utf8');
  assert.doesNotMatch(html,/class="sidebar"|class="side-link"/);
  assert.match(html,/<header class="topbar">[\s\S]*id="open-settings"[^>]*>[\s\S]*Settings/);
  assert.match(html,/id="count-attention"/);
  assert.doesNotMatch(css,/\.sidebar\s*\{/);
  assert.doesNotMatch(app,/#side-(?:ready|attention)|#side-link-/);
  h.ui.wire();
  h.controls.get('#open-settings').dispatch('click');
  assert.equal(h.controls.get('#settings-modal').open,true);
  assert.equal(h.requests.length,0,'Extra backup tools should not load until requested.');
  h.controls.get('#backup-options').open = true; h.controls.get('#backup-options').dispatch('toggle');
  assert.ok(h.requests.some(request => request.endpoint === '/api/backup/list'));
});

test('Settings can create a new Plant workbook template without silently connecting it', () => {
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  const app = fs.readFileSync(path.join(__dirname,'../web/app.js'),'utf8');
  assert.equal((html.match(/data-create-setting="plant[12]"/g)||[]).length,2);
  assert.match(html,/data-create-setting="plant1"[^>]*>Create (?:template|blank Excel)/);
  assert.match(app,/api\/create-plant-workbook/);
  assert.match(app,/Click Save to connect it/);
});

test('Home uses one status filter bar with counts beside the company list', () => {
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  assert.match(html,/<div class="quick-views" id="status-filters" aria-label="Status filters">/);
  assert.equal((html.match(/class="quick-views"/g)||[]).length,1);
  assert.doesNotMatch(html,/summary-panel|stat-card|attention-strip|review-alerts/);
  assert.match(html,/<div[^>]*id="more-filters"/);
  assert.match(html,/id="more-filters" popover="auto"/);
  assert.doesNotMatch(html,/>More filters</);
  assert.doesNotMatch(html,/<details[^>]*id="more-filters"/);
  assert.doesNotMatch(html,/id="create-another"|>Create another<\/button>/);
  const toolbar = html.slice(html.indexOf('<div class="toolbar">'),html.indexOf('<div class="quick-views"'));
  assert.match(toolbar,/id="search"/); assert.match(toolbar,/id="plant-filter"/);
  assert.match(toolbar,/id="open-record-filters" popovertarget="more-filters"/);
  assert.doesNotMatch(toolbar,/id="status-filter"|id="template-filter"|id="generate-all"/);
  const bulk = html.slice(html.indexOf('id="bulk-generation-modal"'),html.indexOf('id="saved-file-modal"'));
  assert.match(bulk,/Only checked clients will get a PDF/);
  assert.match(bulk,/id="bulk-select-all"/);
});

test('Plant alone lights the filter badge and shows Clear all in the toolbar', () => {
  const h = harness(); h.ui.wire();
  const plant = h.controls.get('#plant-filter');
  const chips = h.controls.get('#record-filter-chips');
  const clear = h.controls.get('#reset-record-filters');
  const badge = h.controls.get('#record-filter-count');
  h.ui.updateRecordFilterButton();
  assert.equal(clear.hidden,true);
  plant.value='plant1'; plant.dispatch('change');
  assert.equal(clear.hidden,false);
  assert.equal(badge.textContent,'1');
  assert.equal(chips.children[0].attributes['aria-label'],'Remove filter Plant 1');
  clear.dispatch('click');
  assert.equal(plant.value,'all');
  assert.equal(clear.hidden,true);
  assert.equal(badge.hidden,true);
});

test('The filter popover stays inside a narrow viewport when its trigger is off-center', () => {
  const h = harness();
  h.window.innerWidth = 430;
  h.controls.get('#open-record-filters').getBoundingClientRect = () => ({right:320, bottom:190});
  h.controls.get('#more-filters').getBoundingClientRect = () => ({width:398});
  h.ui.positionRecordFilters();
  assert.equal(h.controls.get('#more-filters').style.right, '16px');
});

test('Single filter removal keeps other choices and either clear button resets every filter', () => {
  for (const clear of ['#clear-filters', '#reset-record-filters']) {
    const h = harness(); h.ui.wire();
    const active = {'#status-filter':'completed', '#plant-filter':'plant2', '#search':'Acme', '#template-filter':'destruction',
      '#serial-filter':'4', '#received-date-filter':'2026-10-22', '#received-to-filter':'2026-10-31', '#record-followup-filter':'pdf_missing'};
    for (const [key,value] of Object.entries(active)) h.controls.get(key).value=value;
    h.controls.get('#company-sort').value='company';
    h.ui.updateRecordFilterButton();
    assert.equal(h.controls.get('#open-record-filters').attributes['aria-label'],'Filters, 7 active');
    h.controls.get('#record-filter-chips').children[0].dispatch('click');
    assert.equal(h.controls.get('#plant-filter').value,'all');
    assert.equal(h.controls.get('#search').value,'Acme');
    assert.equal(h.controls.get('#record-filter-count').textContent,'6');
    if (clear === '#clear-filters') h.controls.get('#more-filters').showPopover();
    h.controls.get(clear).dispatch('click');
    if (clear === '#clear-filters') assert.equal(h.controls.get('#more-filters').popoverOpen,false);
    assert.equal(h.controls.get('#record-filter-count').hidden,true);
    for (const key of ['#search', '#serial-filter', '#received-date-filter', '#received-to-filter']) assert.equal(h.controls.get(key).value,'');
    for (const key of ['#status-filter', '#plant-filter', '#template-filter', '#record-followup-filter']) assert.equal(h.controls.get(key).value,'all');
    assert.equal(h.controls.get('#company-sort').value,'company');
  }
});

test('Follow-up filters narrow by work remaining and combine with plant and status', () => {
  const h = harness(); h.ui.wire(); h.ui.app.analyticsAsOf='2026-11-16';
  const records = [
    {sn:'1', plant:'plant1', certificate_status:'generated'},
    {sn:'2', plant:'plant2', certificate_status:'generated_manual'},
    {sn:'3', plant:'plant1', certificate_status:'completed'},
    {sn:'4', plant:'plant1', certificate_status:'generated_manual', issue_code:'pdf_unavailable'},
    {sn:'5', plant:'plant1', certificate_status:'ready', pdf_missing:true},
    {sn:'6', plant:'plant1', certificate_status:'ready', inward_date:'2026-10-16'},
    {sn:'7', plant:'plant1', certificate_status:'alert', inward_date:'2026-10-17'},
    {sn:'8', plant:'plant1', certificate_status:'ready', certificate_type:'destruction', destruction_source:{case_id:'1'}},
    {sn:'9', plant:'plant1', certificate_status:'ready', certificate_type:'destruction', destruction_requested:false},
  ];
  for (const row of records) Object.assign(row, {inward_date:row.inward_date || '2026-10-26',
    form6:'F6', reference:'REF'});
  records.push({sn:'10', plant:'plant1', certificate_status:'alert', inward_date:'31/02/26', reference:'REF'});
  records.push({sn:'11', plant:'plant1', certificate_status:'ready', certificate_type:'sita', inward_date:'2026-10-26', reference:'REF'});
  records.push({sn:'12', plant:'plant1', certificate_status:'na', inward_date:'2026-10-26'});
  records.push({sn:'13', plant:'plant1', certificate_status:'generated_manual', inward_date:'2026-10-26',
    system_alert:true, detail:'The external PDF could not be matched. Open Saved certificates and link the correct PDF.'});
  records[7].reference = '';
  h.ui.setState({rows:records});
  const followup = h.controls.get('#record-followup-filter');
  for (const [value, ids] of Object.entries({
    awaiting_send:['1','2'], pdf_missing:['13','4','5'], overdue_unissued:['6'],
    destruction_pending:['8'], needs_inward_date:['10'], missing_paperwork:['10'],
  })) {
    followup.value=value; followup.dispatch('change');
    assert.deepEqual(Array.from(h.ui.app.filtered, ({row})=>row.sn).sort(),ids);
    assert.equal(h.controls.get('#count-all').textContent,ids.length);
  }
  h.controls.get('#plant-filter').value='plant2';
  followup.value='awaiting_send'; followup.dispatch('change');
  assert.deepEqual(Array.from(h.ui.app.filtered, ({row})=>row.sn),['2']);
  const status = h.controls.get('#status-filter');
  status.value='completed'; status.dispatch('change');
  assert.equal(h.ui.app.filtered.length,0);
  assert.equal(h.controls.get('#count-all').textContent,1);
});

test('Certificate footer offers Save to Excel before creating a PDF', () => {
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  const editor = html.slice(html.indexOf('id="details-modal"'),html.indexOf('id="excel-save-modal"'));
  const footer = editor.slice(editor.lastIndexOf('<div class="modal-footer">'));
  assert.match(footer,/id="replace-uploaded-pdf" hidden>Replace PDF/);
  assert.match(footer,/id="save-inward-details" hidden>Save to Excel/);
  assert.match(footer,/>Cancel<\/button>/); assert.match(footer,/id="save-details">Create PDF/);
  assert.match(editor,/<details[^>]*id="draft-options"[\s\S]*?id="keep-working-draft"[\s\S]*?id="save-draft"/);
  assert.match(editor,/<form[^>]*novalidate/);
});

test('All statuses and Details to finish are visible peers without a status dropdown', () => {
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  const statusChoices = [...html.matchAll(/data-filter-status="([^"]+)"/g)].map(match => match[1]);
  assert.doesNotMatch(html,/certificate-status-select|id="followup-filter"/);
  assert.match(html,/id="record-followup-filter"/);
  const bar=html.slice(html.indexOf('id="status-filters"'),html.indexOf('id="status-filter"'));
  assert.match(bar,/data-quick-status="details_to_finish"/);
  assert.deepEqual(statusChoices,['generated','generated_manual','completed','blank','na']);
  assert.match(html,/<input type="hidden" id="status-filter" value="all">/);
  assert.doesNotMatch(html,/<select[^>]*id="status-filter"/);
  assert.equal((html.match(/data-quick-status="all"/g)||[]).length,1);
  assert.equal((html.match(/data-quick-status="ready"/g)||[]).length,1);
  assert.equal((html.match(/data-quick-status="needs_attention"/g)||[]).length,1);
  assert.match(bar,/<input type="checkbox" id="attention-only">/);
});

test('Needs attention checkbox matches the displayed status inside the wider attention list', () => {
  const h = harness(); h.ui.wire();
  const box = h.controls.get('#attention-only');
  const filter = h.controls.get('#attention-status-filter');
  const status = h.controls.get('#status-filter');
  status.value = 'all';
  h.ui.setState({rows:[{plant:'plant1',sn:'1',supplier:'Needs review',certificate_status:'alert'},
    {plant:'plant1',sn:'2',supplier:'Ready and overdue',certificate_status:'ready',reminder_due:true},
    {plant:'plant1',sn:'3',supplier:'HDD to send',certificate_status:'generated_manual',destruction_pending:true},
    {plant:'plant1',sn:'4',supplier:'Ready with warning',certificate_status:'ready',warnings:['missing inward date']},
    {plant:'plant1',sn:'5',supplier:'No status',certificate_status:'blank'}]});
  assert.equal(filter.hidden,true);
  status.value = 'needs_attention'; status.dispatch('change');
  assert.equal(filter.hidden,false);
  assert.equal(h.ui.app.filtered.length,5);
  box.checked = true; box.dispatch('change');
  assert.equal(status.value,'needs_attention');
  assert.deepEqual(Array.from(h.ui.app.filtered, item=>item.row.sn),['1','4']);
  box.checked = false; box.dispatch('change');
  assert.equal(status.value,'needs_attention');
  assert.equal(h.ui.app.filtered.length,5);
  box.checked = true; box.dispatch('change');
  h.ui.resetRecordFilters();
  assert.equal(box.checked,false);
  assert.equal(filter.hidden,true);
});

test('Changing editor presentation does not change payload, dirty state or saved recovery', () => {
  const h = harness(); const form = editing(h);
  form.values.edit_scope = 'update_excel'; h.ui.app.editor.baseline = h.ui.detailsBody(form);
  const original = JSON.stringify(h.ui.detailsBody(form));
  h.ui.setEditorMode(false);
  assert.equal(h.controls.get('#editor-fields').hidden,false);
  assert.equal(h.controls.get('#editor-summary').hidden,true);
  h.ui.setEditorMode(true); h.ui.setEditorMode(false);
  assert.equal(JSON.stringify(h.ui.detailsBody(form)),original);
  assert.equal(h.ui.editorDirty(),false); assert.equal(h.ui.readWorkingDraft(),null);
  assert.equal(h.requests.length,0);
});

test('Summary displays material total, saved Excel Certificate Weight and explicit override correctly', () => {
  const h = harness(); const form = editing(h);
  h.ui.renderMaterialRows([{item:'Laptops',weight:'12.5'},{item:'Monitors',quantity:'3',per_piece_weight:'2'}]);
  h.ui.renderEditorSummary(); assert.match(text(h.controls.get('#editor-facts')),/18.5 kg · material total/);
  form.values.certificate_weight = '15'; h.ui.renderEditorSummary();
  h.ui.app.editor.excelCertificateWeight = '15'; h.ui.renderEditorSummary();
  assert.match(text(h.controls.get('#editor-facts')),/15 kg · Certificate Weight in Excel/);
  form.values.certificate_weight = '8.5'; form.values.use_certificate_weight = true;
  h.ui.renderEditorSummary(); assert.match(text(h.controls.get('#editor-facts')),/8.5 kg · entered here/);
  form.values.supplier = 'Updated company'; h.ui.renderChangeSummary();
  assert.match(text(h.controls.get('#editor-facts')),/Updated company/);
  assert.equal(h.controls.get('#changes-disclosure').hidden,false);
});

test('First preview exposes reference, COD/date, inward ID and all read-only material details', () => {
  const h = harness(); const form = editing(h);
  form.values.reference = 'REF-1010'; form.values.form6 = 'F6-1010';
  form.values.cod_no = '0362'; form.values.cod_date = '2026-09-30';
  h.ui.app.editor.firstGeneration = true;
  h.ui.renderMaterialRows([{item:'Laptops',quantity:'2',per_piece_weight:'3.5',weight:'7'},
    {item:'Monitors',quantity:'3',per_piece_weight:'2'}]);
  const before = JSON.stringify(h.ui.detailsBody(form));
  h.ui.setEditorMode();
  const facts = text(h.controls.get('#editor-facts'));
  for (const value of ['REF-1010','F6-1010','0362','30/09/26','SN','1001']) assert.ok(facts.includes(value));
  const materials = h.controls.get('#editor-materials');
  assert.match(text(materials), /Laptops.*Quantity: 2.*3.5 kg per piece.*7 kg total/);
  assert.match(text(materials), /Monitors.*Quantity: 3.*2 kg per piece.*6 kg total/);
  assert.equal(all(materials, node => ['input','button','details'].includes(node.tag)).length, 0);
  assert.equal(JSON.stringify(h.ui.detailsBody(form)), before, 'Summary rendering cannot change source or submit resolved preview values.');
  for (const name of ['reference','cod_no','cod_date']) assert.equal(form.elements.namedItem(name).readOnly, true);
  h.ui.app.editor.firstGeneration = false; h.ui.app.editor.existing = true; h.ui.app.editor.issuedCertificate = true;
  h.ui.setEditorMode(true);
  assert.equal(form.elements.namedItem('cod_no').readOnly, true, 'the COD remains fixed after issuance');
  assert.equal(form.elements.namedItem('cod_date').readOnly, false, 'the date can be corrected after issuance');
  form.values.certificate_type = 'sita'; h.ui.renderEditorSummary();
  assert.doesNotMatch(text(h.controls.get('#editor-facts')), /Form 6|F6-1010/);
});

test('Preview facts use the actual rendered COD, date and printed weight and clear when invalidated', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({preview_url:`/previews/${'a'.repeat(32)}.pdf`,
    certificate_no:'0042',certificate_date:'2026-09-30',printed_details:{certificate_weight:'12 KG',company_display:'ACME, MUMBAI'}})})});
  const form = editing(h); form.values.cod_no = '0001'; form.values.cod_date = '';
  form.values.inward_date = '2026-09-02';
  h.ui.app.editor.firstGeneration = true;
  h.ui.renderMaterialRows([{item:'Laptops',weight:'100.2'}]);
  h.ui.renderEditorSummary(); assert.match(text(h.controls.get('#editor-facts')), /Set automatically/);
  const before = JSON.stringify(h.ui.detailsBody(form));
  await h.ui.previewDetails();
  const facts = text(h.controls.get('#editor-facts'));
  assert.match(facts, /0042/); assert.match(facts, /30\/09\/26/); assert.match(facts, /12 KG · material total/);
  assert.doesNotMatch(facts, /0001|Set automatically/);
  assert.match(h.controls.get('#preview-number').textContent, /COD 0042.*30\/09\/26/);
  assert.equal(JSON.stringify(h.ui.detailsBody(form)), before);
  h.ui.invalidatePreview();
  assert.equal(h.ui.app.editor.previewFacts, null);
  assert.match(text(h.controls.get('#editor-facts')), /0001.*Set automatically/);
  assert.doesNotMatch(text(h.controls.get('#editor-facts')), /0042|12 KG/);
});

test('Weight note separates the asset total from the weight that will print', () => {
  const h = harness(); const form = editing(h);
  h.ui.renderMaterialRows([{item:'Laptops',weight:'12.5'},{item:'Monitors',quantity:'3',per_piece_weight:'2'}]);
  h.ui.app.editor.excelCertificateWeight = '15';
  h.ui.updatePrintedWeight();
  assert.equal(h.controls.get('#current-weight').textContent,'Assets total: 18.5 kg. Certificate will use Excel Certificate Weight: 15 kg.');
  form.values.use_certificate_weight = true; form.values.certificate_weight = '8.5';
  h.ui.updatePrintedWeight();
  assert.equal(h.controls.get('#current-weight').textContent,'Assets total: 18.5 kg. Certificate will use your entered Certificate Weight: 8.5 kg.');
  form.values.use_certificate_weight = false; h.ui.app.editor.excelCertificateWeight = '';
  h.ui.updatePrintedWeight();
  assert.equal(h.controls.get('#current-weight').textContent,'Assets total: 18.5 kg. Certificate will use this total.');
  h.ui.renderMaterialRows([{item:'Laptops',quantity:'2'}]); h.ui.updatePrintedWeight();
  assert.equal(h.controls.get('#current-weight').textContent,'Asset total is incomplete. Certificate will use the asset total once all weights are entered and valid.');
});

test('Invalid hidden fields reveal their disclosure and edit screen before preview or creation', async () => {
  const h = harness(); const form = editing(h);
  h.ui.app.state.numbering = {needs_start:false,next_cod:'0362'};
  h.ui.setEditorMode(false);
  const disclosure = new Node('details'); disclosure.tagName = 'DETAILS'; disclosure.parentElement = form;
  const invalid = new Node('input'); invalid.parentElement = disclosure;
  form.checkValidity = () => false;
  form.querySelectorAll = selector => selector === ':invalid' ? [invalid] : [];
  let validations = 0; form.reportValidity = () => { validations++; return false; };
  await h.ui.previewDetails();
  await h.ui.saveDetails({preventDefault(){},currentTarget:form});
  assert.equal(h.requests.length,0); assert.equal(disclosure.open,true);
  assert.equal(h.controls.get('#editor-fields').hidden,false); assert.equal(validations,2);
});

test('Batch selection shows exact plants and Ready count without creating anything', () => {
  for (const [plant,count,description] of [['plant1',2,'Plant 1'],['plant2',1,'Plant 2'],['all',3,'Plant 1 and Plant 2']]) {
    const h = harness();
    h.ui.app.state.rows = [{plant:'plant1',certificate_status:'ready',outcome:'ready'}, {plant:'plant1',certificate_status:'ready',outcome:'ready',warnings:['Check weight']},
      {plant:'plant2',certificate_status:'ready',outcome:'ready'}, {plant:'plant2',certificate_status:'blank'}, {plant:'plant1',certificate_status:'na'},
      {plant:'plant1',certificate_status:'ready',outcome:'needs_review',warnings:['Enter the Form 6 number']},
      {plant:'plant1',certificate_status:'ready',outcome:'ready',source_kind:'manual'}];
    h.controls.get('#plant-filter').value = plant; h.controls.get('#search').value = 'Unrelated search';
    h.controls.get('#status-filter').value = 'completed'; h.controls.get('#template-filter').value = 'siemens';
    h.ui.app.workspaceRevision = 'reviewed-batch'; h.ui.openBulkGeneration();
    assert.equal(h.controls.get('#bulk-description').textContent,`${count} ${count === 1 ? 'client' : 'clients'} ready · ${description}. Choose who needs a PDF.`);
    assert.equal(h.controls.get('#bulk-generation-modal').open,true);
    assert.equal(h.requests.length,0); assert.equal(h.ui.app.bulkSelection.expected_revision,'reviewed-batch');
  }
});

test('Batch cancellation is harmless and creation retains the reviewed plant and revision', async () => {
  const h = harness(); h.ui.wire();
  h.ui.app.state.rows = [{plant:'plant2',sn:'7',certificate_status:'ready',outcome:'ready',certificate_no:'0300'}];
  h.controls.get('#plant-filter').value = 'plant2'; h.ui.app.workspaceRevision = 'reviewed-batch';
  h.ui.openBulkGeneration(); h.controls.get('#bulk-generation-modal').close();
  await h.ui.startBulkGeneration(); assert.equal(h.requests.length,0);
  h.ui.openBulkGeneration(); h.controls.get('#plant-filter').value = 'plant1'; h.ui.app.workspaceRevision = 'changed-batch';
  vm.runInContext('submitGeneration = async payload => { app.submittedBatch = payload; };',h.context);
  h.ui.app.bulkCheckboxes[0].checked=true; h.ui.app.bulkCheckboxes[0].dispatch('change');
  await h.ui.startBulkGeneration();
  assert.deepEqual(JSON.parse(JSON.stringify(h.ui.app.submittedBatch)),{selections:[{plant:'plant2',sn:'7'}],expected_revision:'reviewed-batch'});
});

function searchableBulk() {
  const h=harness();h.ui.wire();
  h.controls.get('#plant-filter').value='all';h.ui.app.workspaceRevision='bulk-review';
  h.ui.app.state.numbering={needs_start:false,next_cod:'0209'};
  h.ui.app.state.rows=[
    {plant:'plant1',sn:'13',supplier:'Siemens Healthcare',location:'Mumbai',reference:'REF-A',certificate_status:'ready',outcome:'ready'},
    {plant:'plant2',sn:'21',supplier:'Northstar Technologies',location:'Pune',reference:'REF-B',certificate_status:'ready',outcome:'ready'},
    {plant:'plant1',sn:'14',supplier:'Harbour Education',location:'Mumbai',reference:'REF-C',certificate_status:'ready',outcome:'ready'},
  ];
  h.ui.openBulkGeneration();return h;
}

test('Create several searches names, SNs, references and cities without losing selections', () => {
  const h=searchableBulk();
  const selected=h.ui.app.bulkRows.findIndex(row=>row.sn==='21');
  h.ui.app.bulkCheckboxes[selected].checked=true;h.ui.app.bulkCheckboxes[selected].dispatch('change');
  for(const query of ['SIEMENS','13','ref-a','mumbai healthcare']) {
    h.controls.get('#bulk-search').value=query;h.controls.get('#bulk-search').dispatch('input');
    assert.deepEqual(Array.from(h.ui.app.bulkVisible).map(i=>h.ui.app.bulkRows[i].sn),['13']);
    assert.equal(h.ui.app.bulkChecked.has(selected),true);
    assert.equal(h.controls.get('#bulk-selected-count').textContent,'1 selected');
    assert.match(h.controls.get('#bulk-selection-note').textContent,/1 selected client is hidden/);
  }
  h.controls.get('#bulk-clear-search').dispatch('click');
  assert.equal(h.ui.app.bulkVisible.length,3);assert.equal(h.ui.app.bulkCheckboxes[selected].checked,true);
  assert.equal(h.controls.get('#bulk-search').focused,true);
});

test('Select shown affects only matching clients and leaves other checked clients intact', () => {
  const h=searchableBulk(),find=sn=>h.ui.app.bulkRows.findIndex(row=>row.sn===sn);
  h.ui.app.bulkCheckboxes[find('21')].checked=true;h.ui.app.bulkCheckboxes[find('21')].dispatch('change');
  h.controls.get('#bulk-search').value='Mumbai';h.controls.get('#bulk-search').dispatch('input');
  assert.equal(h.controls.get('#bulk-select-label').textContent,'Select shown');
  const selectAll=h.controls.get('#bulk-select-all');selectAll.checked=true;selectAll.dispatch('change');
  assert.equal(h.ui.app.bulkChecked.size,3);assert.equal(selectAll.indeterminate,false);
  selectAll.checked=false;selectAll.dispatch('change');
  assert.equal(h.ui.app.bulkChecked.size,1);assert.equal(h.ui.app.bulkChecked.has(find('21')),true);
  h.ui.app.bulkCheckboxes[find('13')].checked=true;h.ui.app.bulkCheckboxes[find('13')].dispatch('change');
  assert.equal(selectAll.indeterminate,true);
});

test('Empty bulk search offers clear search without changing existing choices', () => {
  const h=searchableBulk();
  h.ui.app.bulkCheckboxes[0].checked=true;h.ui.app.bulkCheckboxes[0].dispatch('change');
  h.controls.get('#bulk-search').value='Not a client';h.controls.get('#bulk-search').dispatch('input');
  assert.equal(h.controls.get('#bulk-empty').hidden,false);
  assert.equal(h.controls.get('#bulk-select-all').disabled,true);
  assert.equal(h.controls.get('#generate-all').disabled,false);
  assert.equal(h.controls.get('#bulk-visible-count').textContent,'0 of 3 shown');
  h.controls.get('#bulk-generation-modal').close();h.ui.openBulkGeneration();
  assert.equal(h.controls.get('#bulk-search').value,'');assert.equal(h.ui.app.bulkChecked.size,0);
});

test('Bulk creation sends exactly the selected clients including clearly counted hidden selections', async () => {
  const h=searchableBulk();
  h.ui.app.bulkCheckboxes[0].checked=true;h.ui.app.bulkCheckboxes[0].dispatch('change');
  h.controls.get('#bulk-search').value='Northstar';h.controls.get('#bulk-search').dispatch('input');
  const visible=h.ui.app.bulkVisible[0];h.ui.app.bulkCheckboxes[visible].checked=true;h.ui.app.bulkCheckboxes[visible].dispatch('change');
  const expected=Array.from(h.ui.app.bulkChecked).sort((a,b)=>a-b).map(i=>h.ui.app.bulkRows[i].sn);
  vm.runInContext('submitGeneration = async payload => { app.submittedBatch=payload; };',h.context);
  await h.ui.startBulkGeneration();
  assert.deepEqual(Array.from(h.ui.app.submittedBatch.selections).map(row=>row.sn),expected);
  assert.equal(h.ui.app.submittedBatch.expected_revision,'bulk-review');
});

test('Issued company actions form tidy pairs without a redundant More control', () => {
  const h=harness();
  for(const certificate_status of ['generated','completed']) {
    const actions=h.ui.makeActions({...removalRow,certificate_status,pdf_filename:'0209_Test.pdf',certificate_no:'0209'},0);
    const groups=all(actions,n=>n.className==='row-actions-main'||n.className==='row-actions-extra');
    assert.ok(groups.every(group=>group.children.length<=2));
    assert.equal(all(actions,n=>n.tag==='details').length,0);
    assert.ok(all(actions,n=>n.dataset.action==='edit-details').length);
    assert.match(text(actions),/Open PDF.*Review & edit.*Open in Excel/);
  }
});

test('Review materials use the Add client calculation layout and recalculate weights', () => {
  const h=harness(); const form=editing(h);
  h.ui.renderMaterialRows([{item:'Laptops',weight:'21',quantity:'3',per_piece_weight:'7'}]);
  const row=h.ui.app.materialRows[0];
  assert.match(text(row.element),/Material or asset.*Calculate weight.*Per Pcs Weight.*×.*Quantity in Pcs.*=.*Weight/);
  assert.equal(all(row.element,n=>n.tag==='details').length,0);
  row.inputs.quantity.value='4';row.inputs.quantity.dispatch('input');
  assert.equal(row.inputs.weight.value,'28');assert.equal(h.ui.materialWeightTotal().total,28);
  assert.equal(h.ui.app.editor.touched.has('weight'),true);
  row.inputs.quantity.value='1 lot';row.inputs.quantity.dispatch('input');
  assert.equal(row.inputs.weight.value,'');row.inputs.weight.value='12';row.inputs.weight.dispatch('input');
  assert.equal(h.ui.materialWeightTotal().total,12);
});

test('Generated certificate rows open the exact safe PDF locally and expose completion directly', () => {
  const h = harness();
  for (const status of ['generated','generated_manual','completed']) {
    const row = {certificate_status:status,supplier:'Acme',plant:'plant2',pdf_filename:'0082_Acme_Mumbai_2.pdf'};
    const actions = h.ui.makeActions(row,7); const links = all(actions,node => node.tag === 'button' && node.textContent === 'Open PDF');
    assert.equal(links.length,1); assert.equal(links[0].textContent,'Open PDF');
    const completions = all(h.ui.makeActions(row,7),node => node.tag === 'button' && node.dataset.action === 'complete');
    assert.equal(completions.length,status === 'completed' ? 0 : 1);
    if (completions.length) assert.equal(completions[0].dataset.index,'7');
  }
  assert.equal(all(h.ui.makeActions({certificate_status:'generated',pdf_filename:'../secret.pdf'},0),node => node.tag === 'button' && node.textContent === 'Open PDF').length,0);
  assert.equal(h.requests.length,0);
});

test('A recovered Ready inward with old certificate details is offered for creation again', () => {
  const h = harness();
  const row = {source_kind:'workbook',plant:'plant1',sn:'1001',supplier:'Recovered client',
    certificate_status:'ready',has_issued_certificate:false,
    last_certificate_details:{certificate_type:'siemens'}};
  const actions = all(h.ui.makeActions(row,0), node => node.tag === 'button').map(node => node.textContent);
  assert.ok(actions.includes('Review & edit'));
  assert.equal(actions.includes('Edit certificate'),false);
});

test('Alert shortcut count respects plant selection without making Blank or N/A eligible for creation', () => {
  const h = harness(); h.controls.get('#plant-filter').value = 'plant2';
  h.ui.app.state.rows = [{plant:'plant1',certificate_status:'alert'}, {plant:'plant2',certificate_status:'alert'},
    {plant:'plant2',certificate_status:'blank'}, {plant:'plant2',certificate_status:'na'}];
  h.ui.summarize(); assert.equal(h.controls.get('#count-attention').textContent,2);
  assert.equal(h.controls.get('#count-all').textContent,3);
  assert.equal(h.controls.get('#open-bulk').disabled,true); assert.equal(h.controls.get('#generate-all').disabled,true);
});

test('First issuance opens a preview with one edit control; PDF corrections leave Excel untouched', async () => {
  for (const kind of ['existing','new']) {
    const h = harness(); const form = editing(h); h.controls.get('#details-modal').open = false;
    vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {}; globalThis.previews = []; previewDetails = async () => previews.push(detailsBody(document.querySelector("#details-form")));',h.context);
    const row = {plant:'plant2',sn:'2001',supplier:'Northstar',form6:'F6-2001',reference:'REF-2001',weight:12,inward_date:'2026-10-01',items:[{item:'Laptops',weight:12}]};
    await h.ui.openDetails(kind === 'new' ? null : row);
    if (kind === 'new') assert.equal(form.values.inward_date,'','A new certificate must not assume today was the receipt date.');
    assert.equal(h.controls.get('#editor-fields').hidden,false);
    assert.equal(h.controls.get('#excel-detail-fields').hidden,kind === 'existing');
    assert.equal(h.context.previews.length,0,'The editor opens before starting its preview.');
    if (kind === 'existing') await h.timers.get(h.ui.app.previewTimer).fn();
    assert.equal(h.context.previews.length,kind === 'existing' ? 1 : 0);
    assert.equal(h.requests.length,0);
    if (kind === 'existing') {
      assert.equal(h.context.previews[0].source_sn,'2001'); assert.equal(h.context.previews[0].source_plant,'plant2');
      assert.equal(h.context.previews[0].edit_scope,'certificate_only');
      assert.equal(h.controls.get('#edit-pdf-details').hidden,false);
      h.ui.setEditorMode(true);
      assert.equal(h.controls.get('#editor-fields').hidden,false);
      assert.equal(h.controls.get('#excel-detail-fields').hidden,false,'the single edit button reveals certificate-only corrections');
      h.ui.app.editor.pdfEditUsed = true;
      assert.equal(h.ui.detailsBody(h.controls.get('#details-form')).review_edits,true);
      assert.equal(h.controls.get('#material-weight-total').textContent,'Total weight: 12 kg');
    }
  }
});

test('New certificate receipt date starts blank and can be entered manually', () => {
  const h = harness(), form = editing(h);
  h.ui.app.editor.sourceSn = null;
  form.values.inward_date = '';
  assert.equal(h.ui.validateEditor(),true,'A genuinely unknown date remains optional and can be reviewed as a warning.');
  form.values.inward_date = '2026-10-01';
  assert.equal(h.ui.validateEditor(),true);
  h.ui.app.editor.sourceSn = '1001'; h.ui.app.editor.firstGeneration = true;
  form.values.inward_date = '';
  h.ui.setEditorMode(true);
  assert.equal(form.elements.namedItem('inward_date').readOnly,false,'The date can be entered manually when Excel has no date.');
  form.values.certificate_type = 'siemens'; form.values.inward_date = '';
  assert.equal(h.ui.validateEditor(),true,'Siemens hides the receipt date and does not require it.');
});

test('An inward with deferred Form 6 opens the field for certificate creation', async () => {
  const h = harness(); const form = editing(h); h.controls.get('#details-modal').open = false;
  vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {}; previewDetails = async () => {};', h.context);
  await h.ui.openDetails({source_kind:'workbook',plant:'plant1',sn:'1001',supplier:'Acme',
    form6:'',inward_date:'2026-10-01',items:[{item:'Laptop',weight:7}]});
  assert.equal(h.controls.get('#excel-detail-fields').hidden,false);
  assert.equal(form.elements.namedItem('form6').readOnly,false);
  form.values.form6 = 'F6-AT-CREATION';
  const body = h.ui.detailsBody(form);
  assert.equal(body.details.form6,'F6-AT-CREATION');
  assert.equal(body.edit_scope,'certificate_only');
  assert.equal(h.requests.length,0,'Entering the number must not save Excel before generating.');
});

test('Source controls preserve review mode for a filled inward and explicit edit mode', () => {
  const h = harness(), form = editing(h);
  Object.assign(h.ui.app.editor, {existing:true,firstGeneration:true,editingDetails:false});
  form.values.form6 = 'F6-FILLED';
  for (const name of ['certificate_weight','case_id','inward_date','form6','reference']) {
    const field = new Node('label'), span = new Node('span');
    field.append(span); field.querySelector = () => span;
    form.elements.namedItem(name).closest = () => field;
  }
  vm.runInContext('suggestWorkbookId = () => {}; updatePrintedWeight = () => {}; updateSourceControls();',h.context);
  assert.equal(h.controls.get('#editor-summary').hidden,false);
  assert.equal(h.controls.get('#excel-detail-fields').hidden,true);
  assert.equal(form.elements.namedItem('form6').readOnly,true);
  h.ui.setEditorMode(true);
  vm.runInContext('updateSourceControls();',h.context);
  assert.equal(h.controls.get('#excel-detail-fields').hidden,false);
  assert.equal(form.elements.namedItem('form6').readOnly,false);
});

test('HDD hides and disables reference entry; switching back restores the normal requirement', () => {
  const h = harness(), form = editing(h);
  const fields = {};
  for (const name of ['certificate_weight','case_id','inward_date','form6','reference']) {
    const field = new Node('label'), span = new Node('span');
    field.append(span); field.querySelector = () => span;
    form.elements.namedItem(name).closest = () => field; fields[name] = field;
  }
  form.values.source_kind = 'manual'; form.values.reference = 'LEGACY-REF';
  vm.runInContext('suggestWorkbookId=()=>{}; updatePrintedWeight=()=>{};',h.context);
  for (const kind of ['destruction','standard','destruction','sita']) {
    form.values.certificate_type = kind; h.ui.app.editor.template = kind;
    h.ui.app.editor.fields[kind] ||= {};
    vm.runInContext('updateSourceControls()',h.context);
    const hdd = kind === 'destruction';
    assert.equal(form.elements.namedItem('reference').required, !hdd);
    assert.equal(form.elements.namedItem('reference').disabled, hdd);
    assert.equal(fields.reference.hidden, hdd);
    h.ui.renderEditorSummary();
    assert.equal(text(h.controls.get('#editor-facts')).includes('Reference number'), !hdd);
  }
  assert.equal(form.values.reference, 'LEGACY-REF', 'Changing type must not erase a manually entered reference.');
});

test('Choosing a certificate type exposes its fields from the first-issuance review', () => {
  const h = harness(), form = editing(h);
  Object.assign(h.ui.app.editor, {existing:true,firstGeneration:true,editingDetails:false});
  form.values.form6 = 'F6-FILLED';
  h.ui.wire();
  vm.runInContext('renderTemplateEditor = () => {}; updateSourceControls = () => {};',h.context);
  h.ui.setEditorMode(false);
  assert.equal(h.controls.get('#excel-detail-fields').hidden,true);
  form.values.certificate_type = 'siemens';
  form.dispatch('change', {target:{name:'certificate_type',value:'siemens'}});
  assert.equal(h.controls.get('#excel-detail-fields').hidden,false);
  assert.equal(h.controls.get('#template-editor').hidden,false);
  assert.equal(h.ui.detailsBody(form).pdf_only_edits,true);
});

test('An explicit Edit details action starts an existing record in the form', async () => {
  const h = harness(); editing(h); h.controls.get('#details-modal').open = false;
  vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {}; globalThis.previews = []; previewDetails = async () => previews.push(detailsBody(document.querySelector("#details-form")));',h.context);
  const row = {plant:'plant2',sn:'2001',supplier:'Northstar',has_issued_certificate:true,weight:12,inward_date:'2026-10-01',items:[{item:'Laptops',weight:12}]};
  await h.ui.openDetails(row,false,true);
  assert.equal(h.controls.get('#editor-fields').hidden,false);
  assert.equal(h.controls.get('#editor-summary').hidden,true);
  assert.match(h.controls.get('#details-title').textContent,/Review & edit: Northstar/);
  await h.timers.get(h.ui.app.previewTimer).fn();
  assert.equal(h.context.previews.length,1,'The read-only preview remains available in the editor.');
  assert.equal(h.requests.length,0,'Opening edit mode never creates a PDF or saves Excel changes.');
});

test('Delayed template loading cannot start an automatic preview after edits, closure or warning review', async () => {
  for (const action of ['edit','close','warning']) {
    const h = harness(); const form = editing(h); h.controls.get('#details-modal').open = false;
    vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {}; globalThis.previews = 0; previewDetails = async () => previews++; loadTemplates = () => new Promise(resolve => { globalThis.finishTemplates = resolve; });',h.context);
    const pending = h.ui.openDetails({plant:'plant1',sn:'1001',supplier:'Acme',weight:10});
    const revision = h.ui.app.previewRevision;
    if (action === 'edit') { form.values.supplier = 'Edited'; h.ui.app.editor.touched.add('supplier'); h.ui.invalidatePreview(); }
    if (action === 'close') h.controls.get('#details-modal').close();
    if (action === 'warning') h.ui.app.pendingConfirmation = {reviewToken:'keep-this-token'};
    h.context.finishTemplates(true); await pending;
    assert.equal(h.context.previews,0);
    if (action === 'edit') {
      await h.timers.get(h.ui.app.previewTimer).fn();
      assert.equal(h.context.previews,1);
    }
    if (action === 'warning') { assert.equal(h.ui.app.pendingConfirmation.reviewToken,'keep-this-token'); assert.equal(h.ui.app.previewRevision,revision); }
  }
});

test('A delayed preview cannot display old values after editing or closing', async () => {
  for (const action of ['edit','close']) {
    let finish;
    const h = harness({fetch:() => new Promise(resolve => {finish = resolve;})}); const form = editing(h);
    const pending = h.ui.previewDetails();
    if (action === 'edit') { form.values.supplier = 'New name'; h.ui.invalidatePreview(); }
    else { h.controls.get('#details-modal').close(); h.ui.invalidatePreview(); }
    finish({ok:true,json:async () => ({preview_url:`/previews/${'c'.repeat(32)}.pdf`,preview_image_url:`/previews/${'c'.repeat(32)}.png`})});
    await pending;
    assert.equal(h.controls.get('#preview-document').hidden,true);
    assert.equal(h.controls.get('#preview-image').src,undefined);
  }
});

test('A corrected edit clears the previous creation error without hiding new preview warnings', () => {
  const h = harness(); editing(h);
  h.controls.get('#details-error').hidden = false; h.controls.get('#details-error').textContent = 'Old row-capacity error';
  h.ui.invalidatePreview(); assert.equal(h.controls.get('#details-error').hidden,true);
  assert.equal(h.controls.get('#preview-document').hidden,true);
});

test('Unticking a changed weight override restores the original Excel weight in the summary', () => {
  const h = harness(); const form = editing(h);
  h.ui.app.editor.excelCertificateWeight = 15;
  h.ui.renderMaterialRows([{item:'Laptops',weight:'20'}]);
  form.values.use_certificate_weight = true; form.values.certificate_weight = '8.5';
  h.ui.renderEditorSummary(); assert.match(text(h.controls.get('#editor-facts')),/8.5 kg · entered here/);
  form.values.use_certificate_weight = false; h.ui.renderEditorSummary();
  assert.match(text(h.controls.get('#editor-facts')),/15 kg · Certificate Weight in Excel/);
  assert.doesNotMatch(text(h.controls.get('#editor-facts')),/8.5/);
  assert.equal(h.ui.detailsBody(form).details.certificate_weight,undefined);
  h.ui.app.editor.excelCertificateWeight = ''; h.ui.renderEditorSummary();
  assert.match(text(h.controls.get('#editor-facts')),/20 kg · material total/);
});

test('Older inactive weight drafts show the effective source weight without applying stale values', async () => {
  const h = harness(); const form = editing(h);
  h.ui.app.editor.excelCertificateWeight = 15;
  form.values.certificate_weight = '8.5'; form.values.use_certificate_weight = false;
  form.elements.namedItem('certificate_weight').disabled = true;
  assert.equal(h.ui.detailsBody(form).details.certificate_weight,undefined);
  assert.equal(h.ui.saveWorkingDraft(),true);
  assert.equal(h.ui.readWorkingDraft().values.certificate_weight,'8.5');
  h.ui.offerDraftRecovery(); form.values.certificate_weight = '';
  vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {};',h.context);
  await h.ui.restoreWorkingDraft();
  assert.equal(String(form.values.certificate_weight),'15');
  assert.equal(h.ui.app.editor.excelCertificateWeight,15);
  h.ui.renderEditorSummary(); assert.match(text(h.controls.get('#editor-facts')),/15 kg · Certificate Weight in Excel/);
  assert.equal(h.ui.detailsBody(form).details.certificate_weight,undefined);
  assert.equal(h.requests.length,0);
});

test('Status selection preserves search, plant and certificate type', () => {
  const h = harness();
  const alertFilter = new Node('button'); alertFilter.dataset.quickStatus = 'needs_attention';
  const selectors = h.context.document.querySelectorAll;
  h.context.document.querySelectorAll = selector => selector === '[data-quick-status]' ? [alertFilter] : selectors(selector);
  h.ui.wire();
  h.controls.get('#plant-filter').value = 'plant2'; h.controls.get('#template-filter').value = 'sita';
  h.controls.get('#search').value = 'Acme';
  h.ui.app.state.rows = [{supplier:'Acme',plant:'plant2',certificate_type:'sita',certificate_status:'alert'}];
  vm.runInContext('render = () => {};',h.context);
  alertFilter.dispatch('click');
  assert.equal(h.controls.get('#plant-filter').value,'plant2'); assert.equal(h.controls.get('#template-filter').value,'sita');
  assert.equal(h.controls.get('#search').value,'Acme'); assert.equal(h.controls.get('#status-filter').value,'needs_attention');
  assert.equal(h.ui.matches(h.ui.app.state.rows[0]),true);
});

test('A delayed close event cannot erase a newly reopened batch selection', () => {
  const h = harness(); h.ui.wire();
  h.ui.app.state.rows = [{plant:'plant2',sn:'7',certificate_status:'ready',outcome:'ready',certificate_no:'0300'}];
  h.controls.get('#plant-filter').value = 'plant2'; h.ui.openBulkGeneration();
  const selection = h.ui.app.bulkSelection;
  h.controls.get('#bulk-generation-modal').dispatch('close');
  assert.equal(h.ui.app.bulkSelection,selection);
  h.controls.get('#bulk-generation-modal').close(); assert.equal(h.ui.app.bulkSelection,null);
});

test('Switching a new manual template to Standard keeps its entered weight as an explicit override', () => {
  const h = harness(); const form = editing(h); h.ui.wire();
  form.values.source_kind = 'manual'; form.values.certificate_type = 'standard'; form.values.certificate_weight = '8.5';
  form.values.use_certificate_weight = false;
  h.ui.app.editor.existing = false; h.ui.app.editor.sourceExplicit = false;
  vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {};',h.context);
  form.dispatch('change',{target:{name:'certificate_type',value:'standard'},currentTarget:form});
  assert.equal(form.values.source_kind,'workbook'); assert.equal(form.values.use_certificate_weight,true);
  assert.equal(h.ui.detailsBody(form).details.certificate_weight,'8.5');
  assert.equal(h.ui.detailsBody(form).use_certificate_weight,true);
  assert.equal(h.requests.length,0);
});

test('A new editor cannot freeze an incomplete revision while startup is loading workbooks', async () => {
  const h = harness(); h.ui.app.loadingWorkspace = true;
  await h.ui.openDetails(); assert.notEqual(h.controls.get('#details-modal').open,true);
  assert.equal(h.ui.app.editor,null); assert.equal(h.requests.length,0);
});

test('Companies and Saved certificates are separate tasks without duplicate dropdowns or a buried file list', () => {
  const h = harness();
  h.ui.setWorkspaceView('files');
  assert.equal(h.controls.get('#page-title').textContent,'Saved certificates');
  assert.equal(h.controls.get('#records-workspace').hidden,true);
  assert.equal(h.controls.get('#folder-inventory').hidden,false);
  assert.equal(h.controls.get('#view-files').attributes['aria-pressed'],'true');
  h.ui.setWorkspaceView('records');
  assert.equal(h.controls.get('#records-workspace').hidden,false);
  assert.equal(h.controls.get('#folder-inventory').hidden,true);
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  assert.match(html,/<button[^>]*id="view-files"[^>]*>Saved certificates/);
  assert.match(html,/<section[^>]*id="folder-inventory"[^>]*aria-label="Saved certificates"/);
  assert.match(html,/<section[^>]*id="folder-inventory"/);
  assert.doesNotMatch(html,/<details[^>]*id="folder-inventory"/);
  assert.equal(h.requests.length,0);
});

test('Saved PDF search matches names, COD and references across words without changing files', () => {
  const h = harness(); const entry = {present:true,path:'Certificates/0324_Acme_Long_Name_Mumbai_2.pdf',cods:['0324'],references:['REF-19']};
  for (const search of ['acme mumbai','0324','ref-19','long name']) {
    h.controls.get('#saved-file-search').value = search; assert.equal(h.ui.savedFileMatches(entry),true);
  }
  h.controls.get('#saved-file-search').value = 'unrelated'; assert.equal(h.ui.savedFileMatches(entry),false);
  h.controls.get('#saved-file-search').value = ''; h.controls.get('#saved-plant-filter').value = 'plant1';
  assert.equal(h.ui.savedFileMatches(entry),false);
  h.controls.get('#saved-plant-filter').value = 'plant2'; assert.equal(h.ui.savedFileMatches(entry),true);
  assert.equal(h.requests.length,0);
});

test('Saved PDF links support categorized folders without exposing escaped paths', () => {
  const h = harness();
  assert.equal(h.ui.savedFileUrl({present:true,path:'Certificates/0082_Acme_1.pdf'}),'http://127.0.0.1:8000/api/certificates/0082_Acme_1.pdf');
  assert.equal(h.ui.savedFileUrl({present:true,path:'Certificates\\0082_Acme_1.pdf'}),'http://127.0.0.1:8000/api/certificates/0082_Acme_1.pdf');
  const nested = {present:true,path:'Certificates/Auto Generated/0082_Acme_1.pdf',cods:['0082']};
  assert.equal(h.ui.savedFileUrl(nested),'http://127.0.0.1:8000/api/certificates/Auto%20Generated/0082_Acme_1.pdf');
  for (const entry of [{present:false,path:'Certificates/0082_Acme_1.pdf'},
    {present:true,path:'Certificates/../secret.pdf'}, {present:true,path:'Certificates/Auto Generated/../../secret.pdf'},
    {present:true,path:'Other/secret.pdf'}, {present:true,path:'C:/Certificates/secret.pdf'}]) assert.equal(h.ui.savedFileUrl(entry),null);
  h.ui.showSavedFile(nested); assert.equal(h.controls.get('#saved-file-open').hidden,false);
  assert.match(text(h.controls.get('#saved-file-facts')),/Auto Generated/);
});

test('Saved rows open PDFs locally, keep warnings, and sort newest without hiding manual files', () => {
  const h = harness(); h.ui.app.workspaceView = 'files';
  h.ui.app.state.certificate_inventory = {entries:[
    {present:true,path:'Certificates/0001_Older_Company_1.pdf',cods:['0001'],mtime_ns:10},
    {present:true,path:'Certificates/0300_Manual_Company_2.pdf',cods:['0300'],mtime_ns:30,warnings:['Verify this certificate number.']},
    {present:false,path:'Certificates/9999_Removed_1.pdf',cods:['9999'],mtime_ns:50},
  ]};
  h.ui.renderFolderInventory(); const list = h.controls.get('#folder-inventory-files');
  assert.match(text(list.children[0]),/Manual Company/);
  assert.equal(all(list,node => node.tag === 'button' && node.textContent === 'Open PDF').length,2);
  assert.match(text(list),/1 file warning/); assert.doesNotMatch(text(list),/Removed/);
  h.controls.get('#saved-file-search').value = 'older'; h.ui.renderFolderInventory();
  assert.equal(all(h.controls.get('#folder-inventory-files'),node => node.tag === 'button' && node.textContent === 'Open PDF').length,1);
  assert.equal(h.controls.get('#saved-visible-count').textContent,'1–1 of 1 PDFs');
});

test('Unknown plant PDFs are explicitly identifiable and never attributed to Plant 1 by default', () => {
  const h = harness();
  assert.equal(h.ui.savedFilePlant({path:'Certificates/0123_Company.pdf'}),'unknown');
  assert.equal(h.ui.savedFilePlant({path:'Certificates/0123_Company_2.pdf'}),'plant2');
  assert.equal(h.ui.savedFilePlant({path:'Certificates/0123_Company.pdf',managed_owner:{plant:'plant1'}}),'plant1');
});

test('Clearing the plant allows either plant in batch selection without selecting any clients', () => {
  const h = harness(); h.controls.get('#plant-filter').value = 'plant2';
  h.ui.app.state.rows = ['plant1','plant2'].map(plant => ({plant, sn:'1', certificate_status:'ready', outcome:'ready'}));
  h.controls.get('#search').value = 'old'; h.controls.get('#status-filter').value = 'generated';
  h.controls.get('#template-filter').value = 'sita'; h.ui.resetRecordFilters();
  h.ui.openBulkGeneration();
  assert.deepEqual(Array.from(h.ui.app.bulkRows, row=>row.plant),['plant1','plant2']);
  assert.equal(h.ui.app.bulkChecked.size,0);
  assert.equal(h.controls.get('#generate-all').disabled,true);
});

test('New Excel ID suggestions include all statuses, are plant-specific and never infer nonnumeric IDs', () => {
  const h = harness(); h.ui.app.state.rows = [
    {plant:'plant1',sn:'1001',certificate_status:'ready'}, {plant:'plant1',sn:'1009',certificate_status:'na'},
    {plant:'plant2',sn:'2003',certificate_status:'blank'}, {plant:'plant1',sn:'9999',source_kind:'manual'},
  ];
  assert.equal(h.ui.nextWorkbookId('plant1'),'1010'); assert.equal(h.ui.nextWorkbookId('plant2'),'2004');
  h.ui.app.state.rows.push({plant:'plant1',sn:'CUSTOM-ID'}); assert.equal(h.ui.nextWorkbookId('plant1'),'');
  assert.equal(h.ui.nextWorkbookId('unknown'),''); assert.equal(h.requests.length,0);
});

test('An Excel ID suggestion never changes an existing or user-edited key', () => {
  const h = harness(); const form = editing(h); h.ui.app.editor.existing = false; h.ui.app.editor.sourceSn = null;
  h.ui.app.state.rows = [{plant:'plant1',sn:'1009'},{plant:'plant2',sn:'2003'}];
  form.values.case_id = ''; h.ui.suggestWorkbookId(); assert.equal(form.values.case_id,'1010');
  form.values.plant = 'plant2'; h.ui.suggestWorkbookId(); assert.equal(form.values.case_id,'2004');
  h.ui.app.editor.touched.add('case_id'); form.values.case_id = 'custom'; h.ui.suggestWorkbookId(); assert.equal(form.values.case_id,'custom');
  h.ui.app.editor.touched.clear(); h.ui.app.editor.sourceSn = '1001'; h.ui.app.editor.existing = true; h.ui.suggestWorkbookId(); assert.equal(form.values.case_id,'custom');
});

test('Automatic previews debounce edits and issue only a read-only preview request', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({preview_url:`/previews/${'d'.repeat(32)}.pdf`})})});
  const form = editing(h); h.ui.renderMaterialRows([{item:'Laptops',weight:10}]);
  h.ui.schedulePreview(); const old = h.ui.app.previewTimer;
  form.values.supplier = 'Latest company'; h.ui.invalidatePreview(); h.ui.schedulePreview();
  assert.equal(h.timers.has(old),false);
  const timer = h.timers.get(h.ui.app.previewTimer); assert.equal(timer.delay,1000);
  await timer.fn();
  assert.equal(h.requests.length,1); assert.equal(h.requests[0].endpoint,'/api/preview');
  assert.equal(h.requests[0].payload.details.supplier,'Latest company');
  assert.equal(h.ui.app.previewLoading,false);
});

test('Automatic preview failures stay inline with one Open in Excel action and no toast', async () => {
  const error = 'Serial Number "R1-10" is repeated in sheet "Register", rows 8, 9. Open Excel and give each inward a different Serial Number, then save and refresh before creating a certificate.';
  const h = harness({fetch:async () => ({ok:false,status:400,json:async () => ({error})})});
  editing(h); h.ui.renderMaterialRows([{item:'Laptop',weight:10}]);
  h.ui.schedulePreview();
  await h.timers.get(h.ui.app.previewTimer).fn();
  assert.equal(h.requests.length,1);
  assert.equal(h.requests[0].endpoint,'/api/preview');
  assert.equal(h.controls.get('#preview-status').textContent,'Preview unavailable');
  const inline = h.controls.get('#details-error');
  assert.equal(inline.hidden,false);
  assert.match(text(inline),/Serial Number R1-10 is repeated.*Excel rows 8 and 9.*Change.*refresh/);
  assert.equal(all(inline,node => node.tag === 'button' && node.textContent === 'Open in Excel').length,1);
  assert.equal(h.controls.get('#toast').classList.contains('visible'),false);
  assert.equal(h.ui.app.previewLoading,false);
  assert.equal(h.controls.get('#details-modal').open,true);
});

test('Repeated Serial Number notices preserve the number and row locations in concise text', () => {
  const h = harness();
  for (const rows of ['8, 9','8 and 9']) {
    const raw = `Serial Number "R1-10" is repeated in sheet "Inward Register", rows ${rows}. Open Excel and give each inward a different Serial Number, then save and refresh before creating a certificate.`;
    const short = h.ui.noticeText(raw);
    assert.match(short,/Serial Number R1-10 is repeated\. Excel rows 8 and 9\./);
    assert.match(short,/Change the repeated number in Excel, then refresh\./);
    assert.doesNotMatch(short,/Inward Register|before creating a certificate|What to do:/);
  }
  const list = new Node('ul');
  h.ui.renderWarningList(list,['Serial Number "R1-10" is repeated in rows 8, 9.','Serial Number "R1-10" is repeated in rows 8, 9.']);
  assert.equal(list.children.length,1,'Repeated notices should produce one issue card.');
});

test('Routine workbook reading notices do not turn healthy companies into Needs attention', () => {
  const h = harness();
  const routine = 'Excel is open or in use. Showing its saved contents; save and close it before the app writes changes.';
  for (const certificate_status of ['ready','generated','completed','na']) {
    const row = {certificate_status,warnings:[routine],workbook_read_only:true};
    assert.equal(Boolean(h.ui.hasWarnings(row)),false);
    assert.equal(Boolean(h.ui.needsAttention(row)),false);
    assert.notEqual(h.ui.stateLabel(row),'Needs attention');
  }
  const list = new Node('ul'); h.ui.renderWarningList(list,[routine]);
  assert.equal(list.children.length,0);
  const issue = {certificate_status:'ready',warnings:[routine,'missing Form 6 number / manifest number']};
  assert.equal(h.ui.needsAttention(issue),true,'A genuine missing field must still require attention.');
  h.ui.renderWarningList(list,issue.warnings); assert.equal(list.children.length,1);
  assert.match(text(list),/Form 6 number missing/);
  assert.equal(h.ui.needsAttention({certificate_status:'ready',warnings:[routine],reminder_due:true}),true);
});

test('Automatic preview waits through incomplete entries and warning review and stops after closing', async () => {
  for (const mode of ['invalid','empty-name','empty-weight','warning','busy','closed','other-editor']) {
    const h = harness(); const form = editing(h); h.ui.renderMaterialRows([{item:'Laptops',weight:10}]);
    h.ui.schedulePreview(); const timer = h.timers.get(h.ui.app.previewTimer);
    if (mode === 'invalid') form.checkValidity = () => false;
    if (mode === 'empty-name') form.values.supplier = '';
    if (mode === 'empty-weight') h.ui.renderMaterialRows([{}]);
    if (mode === 'warning') h.ui.app.pendingConfirmation = {reviewToken:'keep'};
    if (mode === 'busy') h.ui.app.detailsBusy = true;
    if (mode === 'closed') h.controls.get('#details-modal').open = false;
    if (mode === 'other-editor') h.ui.app.editor = {...h.ui.app.editor};
    await timer.fn(); assert.equal(h.requests.length,0,mode);
  }
});

test('Automatic preview cannot overlap another preview or hide warnings when an edit is made', async () => {
  let finish; const h = harness({fetch:() => new Promise(resolve => {finish = resolve;})}); const form = editing(h);
  h.ui.renderMaterialRows([{item:'Laptops',weight:10}]);
  const first = h.ui.previewDetails(); await h.ui.previewDetails(); assert.equal(h.requests.length,1);
  form.values.supplier = 'Latest'; h.ui.invalidatePreview();
  finish({ok:true,json:async () => ({preview_url:`/previews/${'a'.repeat(32)}.pdf`})}); await first;
  assert.equal(h.controls.get('#preview-document').hidden,true);
  assert.ok(h.timers.has(h.ui.app.previewTimer),'New values should be scheduled after the old preview returns.');
});

test('Day-25 reminder appears as an alert without changing status or approving Blank/N/A for creation', () => {
  const h = harness(); h.controls.get('#status-filter').value = 'needs_attention';
  const row = {plant:'plant2',certificate_status:'ready',outcome:'ready',reminder_due:true,age_days:25};
  assert.equal(h.ui.matches(row),true); assert.equal(h.ui.stateLabel(row),'Ready');
  assert.match(text(h.ui.makeStatus(row)),/No certificate yet\. Received 25 days ago/);
  h.ui.app.state.rows = [row,{plant:'plant2',certificate_status:'blank',reminder_due:true},{plant:'plant2',certificate_status:'na'}];
  h.ui.summarize(); assert.equal(h.controls.get('#count-attention').textContent,2);
  assert.equal(h.controls.get('#generate-count').textContent,'(1)');
});

test('Attention view explains future dates, pending PDFs and missing PDFs with useful actions', () => {
  const h = harness(); h.controls.get('#status-filter').value = 'needs_attention';
  const future = {certificate_status:'alert',plant:'plant1',sn:'1004',age_days:-8};
  assert.match(text(h.ui.makeStatus(future)),/Received date is in the future/);
  assert.doesNotMatch(text(h.ui.makeStatus(future)),/-8/);
  const correction = all(h.ui.makeActions(future,0), node => node.tag === 'button');
  assert.equal(correction[0].textContent,'Review & edit');
  assert.ok(correction.some(node => node.textContent === 'Open in Excel'));
  const readyWithIssue = {...future,certificate_status:'ready',warnings:['inward date is in the future']};
  const readyActions = all(h.ui.makeActions(readyWithIssue,0), node => node.tag === 'button');
  assert.equal(readyActions[0].textContent,'Review & edit');
  assert.equal(readyActions[0].dataset.action,'edit-details','First generation supports reviewed edits.');
  assert.ok(readyActions.some(node => node.textContent === 'Open in Excel'));
  const issued = {certificate_status:'generated',reminder_due:true,age_days:50,pdf_filename:'0001_Acme.pdf'};
  assert.match(text(h.ui.makeStatus(issued)),/PDF ready.*Mark Completed/);
  const missing = {...issued,pdf_filename:''};
  assert.match(text(h.ui.makeStatus(missing)),/PDF missing\. Create it again in Review & edit/);
  assert.equal(all(h.ui.makeActions(missing,0), node => node.dataset.action === 'complete').length,0);
  assert.equal(all(h.ui.makeActions(missing,0), node => node.dataset.action === 'edit-details').length,1);
});

test('Attention messages replace recovery reports with one short instruction and retain the COD', () => {
  const h = harness(); h.controls.get('#status-filter').value = 'needs_attention';
  for (const statusNote of ['Excel was changed to Ready.', 'Ready to recreate; Excel status update is pending.', 'Still Ready to recreate.']) {
    const detail = `PDF missing from Certificates: ${statusNote} Keep COD 0188 when recreating it. Ready to generate COD 0188.`;
    for (const pdf_missing of [true, false]) {
      const row = {certificate_status:'ready',certificate_no:'0188',system_alert:true,pdf_missing,detail,warnings:[]};
      const status = h.ui.makeStatus(row);
      const notes = all(status,node => node.className === 'age-note attention-reason');
      assert.equal(notes.length,1);
      assert.equal(notes[0].textContent,'PDF missing. Create it again with COD 0188.');
      assert.doesNotMatch(text(status),/Excel was changed|Ready to recreate|Ready to generate/);
      assert.equal(row.detail,detail,'The original report remains available in technical details.');
    }
  }
  const flagged = h.ui.makeStatus({certificate_status:'ready',pdf_missing:true,certificate_no:'ECO-0190',warnings:[]});
  assert.match(text(flagged),/PDF missing\. Create it again with COD ECO-0190\./);
  const local = h.ui.makeStatus({source_kind:'manual',certificate_status:'ready',system_alert:true,pdf_missing:true,certificate_no:'0190',detail:'PDF missing. Ready to recreate with COD 0190.',warnings:[]});
  assert.equal(all(local,node => node.className === 'age-note attention-reason')[0].textContent,'PDF missing. Create it again with COD 0190.');
});

test('Attention messages use plain actions for common missing details and data destruction', () => {
  const h = harness(); h.controls.get('#status-filter').value = 'needs_attention';
  for (const [warning, expected] of [
    ['missing Form 6 number / manifest number','Form 6 missing. Add it in Review & edit.'],
    ['Enter the reference number before creating the certificate.','Reference number missing. Add it in Review & edit.'],
    ['Material receipt date is missing or invalid; it will be blank on the certificate.','Received date missing or invalid. Fix it in Review & edit.'],
    ['Weight is missing, invalid or not positive; check the material list and certificate weight.','Weight missing or invalid. Check it in Review & edit.'],
  ]) {
    const status = h.ui.makeStatus({certificate_status:'ready',warnings:[warning]});
    assert.equal(all(status,node => node.className === 'age-note attention-reason')[0].textContent,expected);
  }
  const hdd = {certificate_status:'generated_manual',destruction_pending:true,pdf_filename:'HDD-0190.pdf'};
  assert.match(text(h.ui.makeStatus(hdd)),/Send the Data Destruction PDF, then click Mark Completed\./);
  assert.doesNotMatch(text(h.ui.makeStatus(hdd)),/HDD|Payment/);
  assert.match(text(h.ui.makeStatus({...hdd,pdf_missing:true})),/Data Destruction PDF missing\. Create it again in Review & edit\./);
});

test('Attention messages preserve unfamiliar problems and show data issues during PDF recovery', () => {
  const h = harness(); h.controls.get('#status-filter').value = 'needs_attention';
  const raw = 'Unexpected import problem. Ask the operator to check this record.';
  assert.ok(text(h.ui.makeStatus({certificate_status:'ready',system_alert:true,detail:raw})).includes(raw));
  const detail = 'PDF missing from Certificates: Excel was changed to Ready. Keep COD 0188 when recreating it. Weight is missing, invalid or not positive; check the material list and certificate weight.';
  assert.match(text(h.ui.makeStatus({certificate_status:'ready',system_alert:true,pdf_missing:true,detail})),/Weight missing or invalid\. Check it in Review & edit\./);
  assert.ok(h.ui.warningInfo(detail).message.includes('Weight is missing'),'Expanded warnings retain the underlying issue.');
});

test('Reminder review reuses the current page and clears filters without a navigation', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({review_request:{id:'review-1'}})})});
  h.controls.get('#busy-overlay').hidden = true;
  h.ui.setWorkspaceView('files');
  h.controls.get('#search').value = 'Other company'; h.controls.get('#plant-filter').value = 'plant2';
  await h.ui.pollReminderReview();
  assert.equal(h.ui.app.workspaceView,'records');
  assert.equal(h.controls.get('#records-workspace').hidden,false);
  assert.equal(h.controls.get('#status-filter').value,'needs_attention');
  assert.equal(h.controls.get('#search').value,''); assert.equal(h.controls.get('#plant-filter').value,'all');
  assert.equal(h.window.location.href,'http://127.0.0.1:8000/');
  await h.ui.pollReminderReview();
  assert.equal(h.requests[1].payload.acknowledged,'review-1');
  assert.ok(h.requests.every(request => request.endpoint === '/api/dashboard-presence'));
});

test('Reminder review acknowledges receipt while preserving an unfinished edit and opens after closing it', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({review_request:{id:'review-1'}})})});
  const form = editing(h); form.values.supplier = 'Unsaved correction';
  h.controls.get('#busy-overlay').hidden = true;
  vm.runInContext('const reminderQuery = document.querySelector; document.querySelector = selector => selector === "dialog[open]" ? (reminderQuery("#details-modal").open ? reminderQuery("#details-modal") : null) : reminderQuery(selector);',h.context);
  await h.ui.pollReminderReview();
  assert.equal(h.controls.get('#status-filter').value,'all');
  assert.equal(h.controls.get('#details-modal').open,true);
  assert.equal(form.values.supplier,'Unsaved correction');
  assert.equal(h.ui.app.reminderReviewPending,true);
  h.controls.get('#details-modal').close();
  await h.ui.pollReminderReview();
  assert.equal(h.controls.get('#status-filter').value,'needs_attention');
  assert.equal(h.ui.app.reminderReviewPending,false);
  assert.equal(form.values.supplier,'Unsaved correction');
});

function configuredSettings(h) {
  h.ui.app.settings = {workbook_paths:['Plant1.xlsx','Plant2.xlsx'],reminders_enabled:false,reminder_time:'11:00'};
  h.ui.openSettings();
  vm.runInContext('setBusy = () => {}; refreshState = async () => {};',h.context);
  return h.controls.get('#settings-form');
}

test('One Settings Save schedules only changed reminder preferences, with an explicit repair option', async () => {
  for (const mode of ['paths','enable','time','repair']) {
    const h = harness({fetch:async () => ({ok:true,json:async () => ({settings:{workbook_paths:['Plant1.xlsx','Plant2.xlsx'],reminders_enabled:false,reminder_time:'11:00'}})})});
    const form = configuredSettings(h);
    if (mode === 'paths') h.settingsInputs.plant1_path.value = 'New1.xlsx';
    if (mode === 'enable') h.settingsInputs.reminders_enabled.checked = true;
    if (mode === 'time') h.settingsInputs.reminder_time.value = '12:00';
    await h.ui.saveSettings({preventDefault(){},currentTarget:form},mode === 'repair');
    assert.equal(h.requests[0].endpoint,'/api/settings'); assert.equal(h.requests[0].payload.update_reminder,mode !== 'paths');
    assert.equal(h.controls.get('#settings-modal').open,false);
  }
});

test('Settings partial success is not presented as total success and remains open for repair', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({settings:{workbook_paths:['Plant1.xlsx','Plant2.xlsx'],reminders_enabled:true,reminder_time:'11:00'},warning:'Settings saved, but Windows could not update the reminder.'})})});
  const form = configuredSettings(h); h.settingsInputs.reminders_enabled.checked = true;
  await h.ui.saveSettings({preventDefault(){},currentTarget:form});
  assert.equal(h.controls.get('#settings-modal').open,true);
  assert.match(h.controls.get('#settings-status').textContent,/Windows/);
  assert.equal(h.controls.get('#dismiss-toast').hidden,false);
  assert.equal(h.ui.app.settingsBusy,false); assert.equal(form.inert,false);
});

test('Invalid Settings reveal required paths and do not send a request', async () => {
  const h = harness(); const form = configuredSettings(h); let reported = false;
  form.checkValidity = () => false; form.reportValidity = () => {reported = true;};
  await h.ui.saveSettings({preventDefault(){},currentTarget:form});
  assert.equal(h.controls.get('#settings-paths').open,true); assert.equal(reported,true); assert.equal(h.requests.length,0);
});

test('Cancel Settings asks Yes/No only for changed preferences', async () => {
  const h = harness(); configuredSettings(h); h.settingsInputs.plant1_path.value = 'Changed.xlsx';
  const first = h.ui.requestSettingsClose(); h.ui.settleActionConfirmation(false); await first;
  assert.equal(h.controls.get('#settings-modal').open,true);
  const second = h.ui.requestSettingsClose(); h.ui.settleActionConfirmation(true); await second;
  assert.equal(h.controls.get('#settings-modal').open,false); assert.equal(h.settingsInputs.plant1_path.value,'Plant1.xlsx');
  assert.equal(h.requests.length,0);
});

test('Warning correction returns straight to editing while batch cancellation creates nothing', () => {
  const h = harness(); editing(h);
  h.ui.showWarnings(['Missing weight'],{fromDetails:true,payload:{}}); h.ui.returnFromWarnings();
  assert.equal(h.controls.get('#warning-modal').open,false); assert.equal(h.controls.get('#editor-fields').hidden,false);
  h.ui.showWarnings(['Missing weight'],{fromDetails:false,payload:{}}); h.ui.returnFromWarnings();
  assert.equal(h.ui.app.pendingConfirmation,null); assert.equal(h.requests.length,0);
});

test('Errors remain visible until dismissed, whereas routine success notifications expire', () => {
  const h = harness(); h.ui.toast('Action failed.',true);
  assert.equal(h.controls.get('#toast').classList.contains('visible'),true);
  assert.equal(h.controls.get('#dismiss-toast').hidden,false);
  assert.equal(h.timers.size,0);
  h.ui.toast('Saved.'); assert.equal(h.controls.get('#dismiss-toast').hidden,true);
  assert.equal(h.timers.size,1);
});

test('Notifications stay above nested panels, remain dismissible, and survive panel closing', () => {
  const h = harness(); h.ui.wire();
  const success = h.controls.get('#success-modal'), details = h.controls.get('#saved-file-modal');
  const notification = h.controls.get('#toast');
  const changed = target => h.ui.app.toastDialogObserver.callback([{target}]);
  success.showModal(); changed(success);
  h.ui.toast('PDF opened and selected in its folder.');
  assert.equal(notification.parentElement,success);
  assert.equal(notification.popoverOpen,true);
  details.showModal(); changed(details);
  assert.equal(notification.parentElement,details,'A panel opened after a notification must not blur it.');
  assert.equal(notification.popoverOpen,true);
  h.ui.toast('Action failed.',true);
  h.controls.get('#dismiss-toast').dispatch('click');
  assert.equal(notification.popoverOpen,false);
  assert.equal(notification.classList.contains('visible'),false);
  assert.equal(details.open,true,'Dismissing a notification must keep the panel open.');
  h.ui.toast('Action failed.',true);
  details.close(); changed(details);
  assert.equal(notification.parentElement,success);
  assert.equal(notification.popoverOpen,true);
  success.close(); changed(success);
  assert.equal(notification.parentElement.tag,'body');
  assert.equal(notification.popoverOpen,true);
  h.ui.toast('Saved.');
  h.timers.get(h.ui.app.toastTimer).fn();
  assert.equal(notification.popoverOpen,false);
  assert.equal(notification.classList.contains('visible'),false);
});

test('Connection loss offers read-only recovery, never automatic repeat of a mutation', async () => {
  const h = harness({fetch:async () => {throw new Error('Network');}});
  await assert.rejects(h.ui.api('/api/create',{details:{}}));
  assert.equal(h.controls.get('#offline-title').textContent,'Connection lost');
  assert.equal(h.controls.get('#reconnect').hidden,false);
  assert.match(h.controls.get('#offline-message').textContent,/Do not repeat creation.*Saved certificates/);
  assert.equal(h.requests.length,1);
});

test('Incidental responses cannot authorize stale cached Excel details', async () => {
  const h = harness({fetch:async endpoint => ({ok:true,json:async () => ({workspace_revision:'changed-outside',...(endpoint === '/api/check' ? {state:{rows:[]}} : {})})})});
  h.ui.app.workspaceRevision = 'rows-read-before'; h.ui.app.state.workspace_revision = 'rows-read-before';
  for (const endpoint of ['/api/preview','/api/templates','/api/history','/api/backup/list','/api/state']) {
    await h.ui.api(endpoint,endpoint === '/api/preview' ? {} : undefined);
    assert.equal(h.ui.app.workspaceRevision,'rows-read-before');
  }
  const fresh = await h.ui.api('/api/check',{});
  assert.equal(h.ui.app.workspaceRevision,'changed-outside');
  assert.equal(fresh.state.workspace_revision,'changed-outside');
});

test('New workbook payload states create-only intent, while existing and manual edits do not', () => {
  const h = harness(); const form = editing(h);
  assert.equal(h.ui.detailsBody(form).create_new,undefined);
  h.ui.app.editor.existing = false; h.ui.app.editor.sourceSn = null;
  assert.equal(h.ui.detailsBody(form).create_new,true);
  form.values.source_kind = 'manual'; assert.equal(h.ui.detailsBody(form).create_new,undefined);
});

test('New certificate form rejects an already-known ID before sending any mutation', async () => {
  const h = harness(); const form = editing(h); h.ui.app.editor.existing = false; h.ui.app.editor.sourceSn = null;
  form.values.inward_date = '2026-10-01';
  h.ui.app.state.rows = [{plant:'plant1',sn:'1001',certificate_status:'na'}];
  await h.ui.saveDetails({preventDefault(){},currentTarget:form});
  assert.equal(h.requests.length,0); assert.equal(h.controls.get('#record-options').open,true);
  assert.match(h.controls.get('#details-error').textContent,/Serial Number.*already used/);
  assert.equal(form.elements.namedItem('case_id').focused,true);
});

test('Moving an existing custom certificate to Excel is create-only, never an unguarded upsert', async () => {
  const h = harness(); const form = editing(h);
  h.ui.app.editor.existing = true; h.ui.app.editor.recordId = 'custom-1'; h.ui.app.editor.sourceSn = null;
  form.values.inward_date = '2026-10-01';
  h.ui.app.state.rows = [{plant:'plant1',sn:'1001',certificate_status:'na'}];
  assert.equal(h.ui.detailsBody(form).create_new,true);
  await h.ui.saveDetails({preventDefault(){},currentTarget:form});
  assert.equal(h.requests.length,0); assert.match(h.controls.get('#details-error').textContent,/Serial Number.*already used/);
  h.ui.app.editor.touched.clear(); form.values.case_id = ''; h.ui.suggestWorkbookId();
  assert.equal(form.values.case_id,'1002');
});

test('Attachment refresh binds new rows without blessing old cached data and then previews automatically', async () => {
  const previewUrl = `/previews/${'a'.repeat(32)}.pdf`;
  const h = harness({fetch:async endpoint => ({ok:true,json:async () => endpoint === '/api/pick-annexure'
    ? {workspace_revision:'attachment-only',annexure:{id:'a',name:'List.pdf'}}
    : endpoint === '/api/check' ? {workspace_revision:'fresh-rows',state:{rows:[{plant:'plant1',sn:'1001',supplier:'Outside edit'}]}}
    : {preview_url:previewUrl,certificate_no:'0002',warnings:[]}})});
  const form = editing(h); h.ui.renderMaterialRows([{item:'Laptops',weight:10}]);
  form.elements.namedItem('starting_cod').setCustomValidity = () => {};
  vm.runInContext('setBusy = () => {}; setDetailsBusy = on => { app.detailsBusy = on; };',h.context);
  await h.ui.pickAnnexure();
  assert.equal(h.ui.app.workspaceRevision,'fresh-rows'); assert.equal(h.ui.app.state.rows[0].supplier,'Outside edit');
  assert.equal(h.ui.app.editor.workspaceRevision,'attachment-only');
  await h.timers.get(h.ui.app.previewTimer).fn();
  assert.equal(h.requests.at(-1).endpoint,'/api/preview');
  assert.equal(h.requests.at(-1).payload.expected_revision,'attachment-only');
  assert.equal(h.requests.at(-1).payload.details.certificate_fields._annexure,'a');
  assert.equal(h.controls.get('#preview-document').hidden,false);
});

test('Reconnect rereads committed Settings without replacing pending edits or retrying a save', async () => {
  const saved = {workbook_paths:['Committed1.xlsx','Committed2.xlsx'],reminders_enabled:true,reminder_time:'12:00'};
  const h = harness({fetch:async () => ({ok:true,json:async () => ({state:{rows:[]},workspace_revision:'fresh',settings:saved})})});
  configuredSettings(h); h.settingsInputs.plant1_path.value = 'Still-editing.xlsx';
  const pending = JSON.stringify(h.ui.settingsValues());
  await h.ui.reconnectDashboard();
  assert.equal(h.requests.length,1); assert.equal(h.requests[0].endpoint,'/api/check');
  assert.equal(h.ui.app.settings.workbook_paths[0],'Committed1.xlsx');
  assert.equal(JSON.stringify(h.ui.settingsValues()),pending);
  assert.equal(h.ui.app.settingsBaseline,JSON.stringify(saved));
  assert.match(h.controls.get('#settings-status').textContent,/pending changes are still here/);
  assert.equal(h.controls.get('#settings-modal').open,true);
});

test('Settings edits made while reconnect is pending are not replaced by the earlier inputs', async () => {
  let finish;
  const h = harness({fetch:() => new Promise(resolve => {finish = resolve;})}); configuredSettings(h);
  const pending = h.ui.reconnectDashboard();
  h.settingsInputs.plant1_path.value = 'Latest-edit.xlsx';
  finish({ok:true,json:async () => ({state:{rows:[]},workspace_revision:'fresh',settings:{workbook_paths:['Saved1.xlsx','Saved2.xlsx'],reminders_enabled:false,reminder_time:'11:00'}})});
  await pending;
  assert.equal(h.settingsInputs.plant1_path.value,'Latest-edit.xlsx');
  assert.equal(h.ui.app.settings.workbook_paths[0],'Saved1.xlsx'); assert.equal(h.requests.length,1);
});

test('An edit made during a delayed preview really triggers a fresh automatic preview', async () => {
  let finish;
  const h = harness({fetch:async (_endpoint, init) => JSON.parse(init.body).details.supplier === 'Acme'
    ? new Promise(resolve => {finish = resolve;})
    : {ok:true,json:async () => ({preview_url:`/previews/${'b'.repeat(32)}.pdf`,certificate_no:'0010',warnings:[]})}});
  const form = editing(h); h.ui.renderMaterialRows([{item:'Laptops',weight:10}]);
  const pending = h.ui.previewDetails();
  form.values.supplier = 'Latest company'; h.ui.app.editor.touched.add('supplier'); h.ui.invalidatePreview();
  finish({ok:true,json:async () => ({preview_url:`/previews/${'a'.repeat(32)}.pdf`,warnings:[]})}); await pending;
  await h.timers.get(h.ui.app.previewTimer).fn();
  assert.equal(h.requests.length,2); assert.equal(h.requests[1].payload.details.supplier,'Latest company');
  assert.equal(h.controls.get('#preview-document').hidden,false);
  assert.match(h.controls.get('#preview-link').href,/b{32}/);
});

test('Turning a reminder off restores a missing time instead of hiding an invalid required field', () => {
  const h = harness(); configuredSettings(h); h.ui.wire();
  h.settingsInputs.reminders_enabled.checked = false; h.settingsInputs.reminder_time.value = '';
  h.controls.get('#settings-form').dispatch('change');
  assert.equal(h.settingsInputs.reminder_time.value,'11:00');
  assert.equal(h.settingsInputs.reminder_time.closest('.field').hidden,true);
});

test('Connection recovery inside an editor uses an accessible top-layer dialog', async () => {
  const h = harness({fetch:async () => {throw new Error('Network');}}); editing(h);
  vm.runInContext('const previousQuery = document.querySelector; document.querySelector = selector => selector === "dialog[open]" ? previousQuery("#details-modal") : previousQuery(selector);',h.context);
  await assert.rejects(h.ui.api('/api/create',{}));
  assert.equal(h.controls.get('#connection-modal').open,true);
  assert.equal(h.controls.get('#details-modal').open,true);
  assert.equal(h.ui.editorDirty(),false);
});

test('Persistent errors expose a clickable dismiss control, and keyboard search does not escape a modal', () => {
  const css = fs.readFileSync(path.join(__dirname,'../web/styles.css'),'utf8');
  assert.match(css,/\.toast\.error\.visible, #dismiss-toast \{ pointer-events: auto;/);
  const h = harness(); editing(h); h.ui.wire();
  vm.runInContext('const previousQuery = document.querySelector; document.querySelector = selector => selector === "dialog[open]" ? previousQuery("#details-modal") : previousQuery(selector);',h.context);
  let prevented = false;
  h.window.dispatch('keydown',{ctrlKey:true,key:'k',preventDefault(){prevented = true;}});
  assert.equal(prevented,false); assert.notEqual(h.controls.get('#search').focused,true);
});

test('Issued certificates expose Review & edit directly with no Details expansion or HDD action', () => {
  const h = harness();
  const row = {plant:'plant1',sn:'1001',certificate_status:'generated',pdf_filename:'0001_Acme_Mumbai_1.pdf'};
  const buttons = all(h.ui.makeActions(row,0), node => node.tag === 'button');
  assert.ok(buttons.some(button => button.textContent === 'Review & edit' && button.dataset.action === 'edit-details'));
  assert.ok(buttons.every(button => button.dataset.action !== 'destruction'));
  assert.ok(buttons.every(button => !Object.hasOwn(button.dataset,'expand')));
});

test('First generation exposes one PDF edit action and keeps Excel writes disabled', () => {
  const h = harness(), form = editing(h);
  h.ui.app.editor.firstGeneration = true;
  h.ui.setEditorMode(true);
  assert.equal(h.controls.get('#editor-fields').hidden,false);
  assert.equal(h.controls.get('#excel-detail-fields').hidden,false);
  assert.equal(h.controls.get('#editor-summary').hidden,true);
  for (const name of ['supplier','location','inward_date','form6','reference']) assert.equal(form.elements.namedItem(name).readOnly,false);
  for (const name of ['cod_no','cod_date','case_id']) assert.equal(form.elements.namedItem(name).readOnly,true);
  assert.equal(h.controls.get('#save-details').textContent,'Generate certificate');
  assert.equal(h.ui.detailsBody(form).edit_scope,'certificate_only');
  assert.equal(h.controls.get('#edit-pdf-details').textContent,'Done editing');
  assert.equal(h.requests.length,0);
  assert.equal(all(h.controls.get('#editor-facts'),node => node.tag === 'button').length,0);
  h.ui.startInlineEdit('supplier','Company'); assert.equal(h.ui.app.editor.inlineEdit,undefined);
});

test('City editor sends and summarizes the full comma-separated location', () => {
  const h = harness(), form = editing(h);
  form.values.location = 'Pune, Maharashtra, India';
  h.ui.app.editor.previewFacts = {printed_details:{company_display:'ACME, PUNE, MAHARASHTRA, INDIA'}};
  assert.equal(h.ui.detailsBody(form).details.location,'Pune, Maharashtra, India');
  h.ui.renderEditorSummary();
  assert.match(text(h.controls.get('#editor-facts')),/Pune, Maharashtra, India/);
  assert.doesNotMatch(text(h.controls.get('#editor-facts')),/Company on PDF/);
});

test('Company override has no extra input and main-field edits clear old Standard overrides', () => {
  const h = harness(), form = editing(h); h.ui.wire();
  h.ui.app.editor.fields.standard = {company_display:'Old printed company'};
  h.ui.app.templates[0].fields = [{key:'company_display',label:'Company line on certificate',group:'wording',default:''}];
  h.ui.renderTemplateEditor();
  assert.equal(all(h.controls.get('#template-editor'),node => node.dataset.certificateField === 'company_display').length,0);
  assert.equal(h.ui.detailsBody(form).details.certificate_fields.company_display,'Old printed company');
  form.values.location = 'Pune, Maharashtra, India';
  form.dispatch('input',{target:{name:'location'}});
  assert.equal(h.ui.detailsBody(form).details.certificate_fields.company_display,'');
  assert.equal(h.ui.detailsBody(form).details.location,'Pune, Maharashtra, India');
});

test('Issued editing uses the first preview layout and opening pencils changes no payload or precision', () => {
  const h = harness(), form = editing(h); Object.assign(h.ui.app.editor,{existing:true,firstGeneration:false});
  h.ui.renderMaterialRows([{item:'Laptops',weight:'100.19999999999999'}]);
  const original = JSON.stringify(h.ui.detailsBody(form)); h.ui.setEditorMode();
  assert.equal(h.controls.get('#editor-summary').hidden,false);
  assert.equal(h.controls.get('#excel-detail-fields').hidden,true);
  assert.equal(h.controls.get('#template-editor').hidden,true);
  for (const key of ['supplier','reference','certificate_weight','materials']) {
    h.ui.startInlineEdit(key,key); assert.equal(h.ui.app.editor.inlineEdit.key,key);
    assert.equal(JSON.stringify(h.ui.detailsBody(form)),original); h.ui.finishInlineEdit();
  }
  for (const key of ['cod_no','cod_date']) {
    h.ui.startInlineEdit(key,key); assert.equal(h.ui.app.editor.inlineEdit,null);
  }
  for (const key of ['plant','case_id','source_kind']) { h.ui.startInlineEdit(key,key); assert.equal(h.ui.app.editor.inlineEdit,null); }
  assert.equal(h.ui.editorDirty(),false); assert.equal(h.requests.length,0);
});

test('Inline typing updates the saved field and preview without replacing its active input or storing DOM in drafts', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({preview_url:`/previews/${'a'.repeat(32)}.pdf`,certificate_no:'0041'})})});
  const form = editing(h); Object.assign(h.ui.app.editor,{existing:true,firstGeneration:false}); h.ui.setEditorMode();
  h.ui.startInlineEdit('supplier','Company');
  const editor = h.ui.app.editor.inlineEdit.container; const input = all(editor,node => node.tag === 'input')[0];
  assert.equal(input.name,undefined); input.value = 'Corrected company'; input.dispatch('input');
  assert.equal(form.values.supplier,'Corrected company'); assert.equal(h.ui.editorDirty(),true);
  const draft = h.ui.readWorkingDraft(); assert.equal(draft.values.supplier,'Corrected company');
  for (const key of ['inlineEdit','summaryHosts','previewFacts']) assert.equal(Object.hasOwn(draft.editor,key),false);
  await h.ui.previewDetails();
  assert.equal(h.ui.app.editor.inlineEdit.container,editor); assert.equal(all(editor,node => node.tag === 'input')[0],input);
  assert.equal(input.value,'Corrected company');
  h.ui.finishInlineEdit(); assert.match(text(h.controls.get('#editor-facts')),/Corrected company/);
  assert.equal(h.controls.get('#changes-disclosure').hidden,true);
});

test('Inline weight edit explicitly overrides Excel weight only after typing and can return to Excel weight', () => {
  const h = harness(), form = editing(h); Object.assign(h.ui.app.editor,{existing:true,firstGeneration:false,excelCertificateWeight:'15'});
  h.ui.setEditorMode(); h.ui.startInlineEdit('certificate_weight','Certificate Weight');
  const input = all(h.ui.app.editor.inlineEdit.container,node => node.tag === 'input')[0];
  assert.equal(input.value,'15'); assert.equal(form.values.use_certificate_weight,false);
  assert.equal(h.ui.detailsBody(form).details.certificate_weight,undefined);
  input.value = '8.25'; input.dispatch('input');
  assert.equal(form.values.use_certificate_weight,true); assert.equal(h.ui.detailsBody(form).details.certificate_weight,'8.25');
  all(h.ui.app.editor.inlineEdit.container,node => node.tag === 'button').find(node => node.textContent === 'Use Excel weight').dispatch('click');
  assert.equal(form.values.use_certificate_weight,false); assert.equal(h.ui.detailsBody(form).details.certificate_weight,undefined);
  assert.match(text(h.controls.get('#editor-facts')),/15 kg/);
});

test('Inline material edit preserves named source controls and submits every changed material with full precision', () => {
  const h = harness(), form = editing(h); Object.assign(h.ui.app.editor,{existing:true,firstGeneration:false});
  vm.runInContext(`FormData = class {constructor(form){this.form=form;} entries(){const values={...this.form.values};
    app.materialRows.forEach(row => Object.values(row.inputs).forEach(input => values[input.name]=input.value)); return Object.entries(values);} get(key){return this.form.values[key];}}`,h.context);
  h.ui.renderMaterialRows([{item:'Laptops',weight:'100.19999999999999'},{item:'Monitors',weight:'6'}]);
  h.ui.setEditorMode(); h.ui.startInlineEdit('materials','Materials');
  const editor = h.ui.app.editor.inlineEdit.container;
  const inputs = all(editor,node => node.tag === 'input'); assert.ok(inputs.every(input => !input.name));
  const secondWeight = inputs.find(input => input.attributes['aria-label'] === 'Material 2: Weight (KG)');
  secondWeight.value = '8'; secondWeight.dispatch('input');
  const items = h.ui.detailsBody(form).details.items;
  assert.equal(items[0].weight,'100.19999999999999'); assert.equal(items[1].weight,'8');
  assert.equal(h.ui.app.editor.inlineEdit.container,editor);
  all(editor,node => node.tag === 'button').find(node => node.textContent === '+ Add another material').dispatch('click');
  assert.equal(h.ui.app.materialRows.length,3); assert.equal(h.ui.detailsBody(form).details.items.length,3);
  form.values.source_kind = 'manual';
  assert.equal(h.ui.detailsBody(form).details.items[1].weight,'8','Custom certificate edits retain the material list too.');
});

test('Siemens exposes its three manual identifiers and Excel site directly without disclosures', () => {
  const h = harness(), form = editing(h);
  form.values.certificate_type = 'siemens'; form.values.location = 'Pune';
  Object.assign(h.ui.app.editor,{firstGeneration:true,existing:true,sourceLocation:'Pune'});
  h.ui.app.templates.push({id:'siemens',default_company:'SIEMENS HEALTHCARE PRIVATE LIMITED',fields:[
    ...['weight_display','model_name','model_serial','tube_serial','site_location'].map(key => ({key,label:key,group:'details',default:''})),
    {key:'company_display',label:'Company line on certificate',group:'wording',default:''}]});
  h.ui.renderTemplateEditor();
  const container = h.controls.get('#template-editor');
  const inputs = all(container,node => Boolean(node.dataset.certificateField));
  assert.deepEqual(inputs.map(input => input.dataset.certificateField),['model_name','model_serial','tube_serial','site_location']);
  assert.equal(inputs.find(input => input.dataset.certificateField === 'site_location').value,'Pune');
  assert.equal(Boolean(inputs.find(input => input.dataset.certificateField === 'site_location').readOnly),false);
  assert.equal(h.ui.app.editor.fields.siemens.company_display,'SIEMENS HEALTHCARE PRIVATE LIMITED');
  assert.equal(all(container,node => node.tag === 'details').length,0);
});

test('Issued Siemens identifiers are available from the single details editor', () => {
  const h = harness(), form = editing(h); form.values.certificate_type = 'siemens'; form.values.location = 'Pune';
  Object.assign(h.ui.app.editor,{existing:true,firstGeneration:false});
  h.ui.app.templates.push({id:'siemens',default_company:'SIEMENS HEALTHCARE PRIVATE LIMITED',fields:[
    ...['model_name','model_serial','tube_serial','site_location'].map(key => ({key,label:key,group:'details',default:''})),
    {key:'company_display',label:'Company line on certificate',group:'wording',default:''}]});
  const template = h.controls.get('#template-editor');
  template.querySelectorAll = () => all(template,node => Boolean(node.dataset.certificateField));
  h.ui.renderTemplateEditor(); h.ui.setEditorMode();
  assert.equal(template.hidden,true);
  h.ui.startInlineEdit('certificate_fields.model_serial','Model serial number');
  const input = all(h.ui.app.editor.inlineEdit.container,node => node.tag === 'input')[0];
  input.value = 'M-123'; input.dispatch('input');
  assert.equal(h.ui.detailsBody(form).details.certificate_fields.model_serial,'M-123');
  h.ui.finishInlineEdit(); assert.match(text(h.controls.get('#editor-facts')),/M-123/);
  form.values.source_kind = 'manual'; h.ui.setEditorMode();
  assert.doesNotMatch(h.controls.get('#editor-summary-note').textContent,/update Excel/);
});

test('Completed record review uses current Excel details and retains issued certificate fields', async () => {
  const h = harness(); editing(h); h.ui.wire(); h.controls.get('#details-modal').open = false;
  vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {}; previewDetails = async () => {};',h.context);
  const row = {source_kind:'workbook',plant:'plant1',sn:'1001',supplier:'Excel company',location:'Pune',inward_date:'2026-09-01',
    form6:'F6-NEW',reference:'REF-NEW',comments:'Current comments',has_issued_certificate:true,
    excel_details_changed_since_issue:true,
    excel_change_preview:[
      {field:'Company name',issued:'Old Excel company',current:'Excel company'},
      {field:'Materials and weights',issued:'Laptop · Qty 1 · 10 kg',current:'Laptop · Qty 2 · 20 kg'},
      {field:'Comments',issued:'Old comments',current:'Current comments'},
    ],
    items:[{item:'Asset',quantity:1,per_piece_weight:10,weight:10}],certificate_weight:8,
    last_certificate_details:{supplier:'Old Excel company',location:'Mumbai',inward_date:'2026-08-31',form6:'F6-OLD',reference:'REF-OLD',
      comments:'Old comments',certificate_type:'siemens',certificate_weight:'7.25',
      certificate_fields:{model_name:'Corrected scanner'},cod_no:'0041',cod_date:'2026-10-01',
      items:[{item:'Asset',quantity:'1',per_piece_weight:'10',weight:'10'}]}};
  await h.ui.openDetails(row);
  const form = h.controls.get('#details-form');
  assert.equal(form.values.supplier,'Excel company'); assert.equal(form.values.location,'Pune');
  assert.equal(form.values.form6,'F6-NEW'); assert.equal(form.values.reference,'REF-NEW');
  assert.equal(form.values.comments,'Current comments');
  assert.equal(form.values.cod_no,'0041'); assert.equal(form.values.use_certificate_weight,true);
  assert.equal(form.values.certificate_type,'siemens');
  assert.equal(h.ui.app.editor.fields.siemens.model_name,'Corrected scanner');
  assert.equal(h.controls.get('#edit-issued-certificate').hidden,false);
  assert.equal(h.ui.app.editor.excelCertificateWeight,8);
  assert.equal(h.ui.app.editor.firstGeneration,false);
  assert.equal(h.controls.get('#excel-detail-fields').hidden,true);
  assert.equal(h.ui.app.editor.editingDetails,false);
  h.controls.get('#edit-issued-certificate').dispatch('click');
  assert.equal(h.controls.get('#issued-change-review-modal').open,true);
  const preview = h.controls.get('#issued-change-preview');
  assert.match(text(preview),/Company name.*Old Excel\s+company.*Excel\s+company/);
  assert.match(text(preview),/Materials and weights.*Laptop.*Qty\s+2/);
  assert.match(text(preview),/Comments.*Old\s+comments.*Current\s+comments/);
  assert.match(text(preview),/Changed/);
  assert.match(fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8'),/Changed values are highlighted below/);
  assert.match(fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8'),/issued PDF and status stay the same/);
});

test('Issued-detail comparison highlights changed material values and marks added or removed lines', () => {
  const h = harness();
  h.ui.renderIssuedChangeReview([{field:'Materials and weights',
    issued:'Desktop · Qty 1 · 10 kg\nCable',current:'Desktop · Qty 2 · 20 kg\nDock\nMouse'}]);
  const preview = h.controls.get('#issued-change-preview');
  const issuedChanges = all(preview,node => node.tag === 'del').map(node => node.textContent);
  const currentChanges = all(preview,node => node.tag === 'mark').map(node => node.textContent);
  assert.deepEqual(issuedChanges,['1 · 10','Cable']);
  assert.deepEqual(currentChanges,['2 · 20','Dock','Mouse']);
  assert.match(text(preview),/Not present/);
  assert.match(text(preview),/Issued value.*Excel now/);
});

test('Issued-detail resolution sends the selected explicit action and refreshes restored form values', async () => {
  for (const action of ['keep','restore']) {
    const h = harness({fetch:async () => ({ok:true,json:async () => ({
      saved:true,action,workspace_revision:'after-resolution',state:{rows:[]},
      source_details:{supplier:'Issued company',location:'Pune',inward_date:'2026-09-01',form6:'F6-OLD',
        reference:'REF-OLD',comments:'Issued comments',items:[{item:'Laptop',quantity:'1',weight:'10'}]},
    })})});
    const form = editing(h); Object.assign(h.ui.app.editor,{sourcePlant:'plant1',sourceSn:'1001'});
    if (action === 'restore') { form.values.supplier='Unsaved user edit'; h.ui.app.editor.touched.add('supplier'); }
    h.ui.app.workspaceRevision='reviewed-revision';
    await h.ui.resolveIssuedExcelChanges(action);
    assert.equal(h.requests[0].endpoint,'/api/clients/resolve-issued-changes');
    assert.equal(h.requests[0].payload.action,action);
    assert.equal(h.requests[0].payload.expected_revision,'reviewed-revision');
    assert.equal(h.controls.get('#issued-change-review-modal').open,false);
    assert.equal(h.controls.get('#edit-issued-certificate').hidden,true);
    if (action === 'restore') {
      assert.equal(form.values.supplier,'Unsaved user edit');
      assert.equal(form.values.form6,'F6-OLD');
      assert.equal(h.ui.app.materialRows.length,1);
      assert.equal(h.ui.editorDirty(),true);
    }
  }
});

test('Failed issued-detail resolution keeps the comparison open and reports the error', async () => {
  const h = harness({fetch:async () => ({ok:false,status:400,json:async () => ({error:'Excel has changed. Refresh and try again.'})})});
  editing(h); Object.assign(h.ui.app.editor,{sourcePlant:'plant1',sourceSn:'1001'});
  h.controls.get('#issued-change-review-modal').open=true;
  await h.ui.resolveIssuedExcelChanges('restore');
  assert.equal(h.controls.get('#issued-change-review-modal').open,true);
  assert.equal(h.controls.get('#issued-change-error').hidden,false);
  assert.match(h.controls.get('#issued-change-error').textContent,/changed/);
  assert.equal(h.ui.app.issuedChangeBusy,false);
});

test('Unchanged issued details do not show a correction notice for date formats or Excel weight overrides', async () => {
  const h = harness(); editing(h); h.controls.get('#details-modal').open = false;
  vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {}; previewDetails = async () => {};',h.context);
  await h.ui.openDetails({source_kind:'workbook',plant:'plant1',sn:'1002',supplier:'Acme',location:'Pune',inward_date:'01/09/26',
    form6:'F6-1',reference:'REF-1',items:[{item:'Asset',quantity:1,per_piece_weight:10,weight:10}],certificate_weight:10,
    has_issued_certificate:true,last_certificate_details:{supplier:'Acme',location:'Pune',inward_date:'2026-09-01',
      form6:'F6-1',reference:'REF-1',items:[{item:'Asset',quantity:'1.0',per_piece_weight:'10.00',weight:'10'}],
      certificate_weight:'8.5',certificate_type:'standard',cod_no:'0042'}});
  assert.equal(h.controls.get('#edit-issued-certificate').hidden,true);
});

test('Edits before and after first creation wait for Yes or No before updating Excel', async () => {
  for (const firstGeneration of [false,true]) for (const yes of [false,true]) {
    const h = harness(), form = editing(h);
    Object.assign(h.ui.app.editor,{existing:true,firstGeneration,reviewEdits:true});
    form.values.supplier = 'Corrected company';
    vm.runInContext('submitGeneration = async payload => {app.correctionPayload = payload;};',h.context);
    await h.ui.saveDetails({preventDefault(){},currentTarget:form});
    assert.equal(h.controls.get('#excel-save-modal').open,true);
    assert.equal(h.ui.app.correctionPayload,undefined);
    h.ui.app.previewRevision++; // Same data, newer raster preview.
    await h.ui.saveExcelChoice(yes);
    assert.equal(h.ui.app.correctionPayload.edit_scope,yes ? 'update_excel' : 'certificate_only');
    assert.equal(h.ui.app.correctionPayload.confirmed_excel,yes);
    assert.equal(h.ui.app.correctionPayload.review_edits,true);
    assert.equal(h.ui.app.correctionPayload.details.supplier,'Corrected company');
    assert.equal(h.controls.get('#excel-save-modal').open,false);
  }
});

test('Save to Excel before first PDF uses a separate confirmed save and No keeps editing', async () => {
  for (const yes of [false,true]) {
    const h = harness({fetch:async()=>({ok:true,json:async()=>({saved:true,created:0,state:{rows:[]},workspace_revision:'saved-revision'})})});
    const form = editing(h);
    Object.assign(h.ui.app.editor,{existing:true,firstGeneration:true,reviewEdits:true});
    form.values.supplier='Changed before any PDF';
    vm.runInContext('setState=state=>{app.state=state;};',h.context);
    await h.ui.saveInwardDetails();
    assert.equal(h.controls.get('#excel-save-modal').open,true);
    assert.equal(h.requests.length,0);
    assert.match(h.controls.get('#excel-save-title').textContent,/Save these changes to Excel/);
    await h.ui.saveExcelChoice(yes);
    assert.equal(h.requests.length,yes ? 1 : 0);
    assert.equal(h.controls.get('#details-modal').open,!yes);
    if (yes) {
      assert.equal(h.requests[0].endpoint,'/api/clients/update');
      assert.equal(h.requests[0].payload.confirmed_excel,true);
      assert.equal(h.requests[0].payload.details.supplier,'Changed before any PDF');
    }
  }
});

test('Type changes offer updating Excel or explicitly deleting the inward and keeping PDF', async () => {
  for (const firstGeneration of [false,true]) for (const yes of [false,true]) {
    const h = harness(), form = editing(h);
    Object.assign(h.ui.app.editor,{existing:true,firstGeneration,reviewEdits:true});
    h.ui.app.editor.excelBaseline = JSON.parse(JSON.stringify(h.ui.app.editor.baseline));
    form.values.certificate_type = 'siemens';
    h.ui.app.editor.template = 'siemens'; h.ui.app.editor.fields.siemens = {};
    vm.runInContext('confirmAction = async () => true; submitGeneration = async payload => {app.correctionPayload = payload;};',h.context);
    await h.ui.saveDetails({preventDefault(){},currentTarget:form});
    assert.equal(h.controls.get('#excel-save-modal').open,true);
    assert.match(h.controls.get('#excel-save-description').textContent,/Excel has Recycling; this PDF will use Medical/);
    assert.equal(h.controls.get('#excel-save-no').textContent,'Delete Excel inward; keep PDF');
    await h.ui.saveExcelChoice(yes);
    assert.equal(h.controls.get('#excel-save-modal').open,false);
    assert.equal(h.controls.get('#details-modal').open,true);
    assert.equal(form.values.certificate_type,'siemens');
    assert.equal(h.ui.app.correctionPayload.edit_scope,yes ? 'update_excel' : 'remove_excel');
    assert.equal(yes ? h.ui.app.correctionPayload.confirmed_excel : h.ui.app.correctionPayload.confirmed_excel_removal,true);
    assert.equal(h.ui.app.correctionPayload.details.certificate_type,'siemens');
  }
});

test('Previously PDF-only type changes still prompt against the actual Excel type', async () => {
  const h = harness(), form = editing(h);
  Object.assign(h.ui.app.editor,{existing:true,firstGeneration:false,reviewEdits:true});
  h.ui.app.editor.excelBaseline = JSON.parse(JSON.stringify(h.ui.app.editor.baseline));
  form.values.certificate_type = 'siemens';
  h.ui.app.editor.template = 'siemens'; h.ui.app.editor.fields.siemens = {};
  h.ui.app.editor.baseline = h.ui.detailsBody(form);
  await h.ui.saveDetails({preventDefault(){},currentTarget:form});
  assert.match(h.controls.get('#excel-save-title').textContent,/original Excel inward/);
  assert.equal(h.controls.get('#excel-save-no').textContent,'Delete Excel inward; keep PDF');
});

test('International type changes clear Form 6 and keep PDF through confirmed Excel removal', async () => {
  const h = harness(), form = editing(h);
  Object.assign(h.ui.app.editor,{existing:true,firstGeneration:true,reviewEdits:true});
  h.ui.app.editor.excelBaseline = JSON.parse(JSON.stringify(h.ui.app.editor.baseline));
  form.values.certificate_type = 'sita'; form.values.form6 = 'F6-new'; form.values.reference = 'REF-new';
  h.ui.app.editor.template = 'sita'; h.ui.app.editor.fields.sita = {};
  vm.runInContext('confirmAction = async () => true; submitGeneration = async payload => {app.correctionPayload = payload;};',h.context);
  await h.ui.saveDetails({preventDefault(){},currentTarget:form});
  assert.match(h.controls.get('#excel-save-description').textContent,/this PDF will use International/);
  assert.match(h.controls.get('#excel-save-note').textContent,/all its materials from Excel/);
  assert.equal(h.controls.get('#excel-save-no').textContent,'Delete Excel inward; keep PDF');
  await h.ui.saveExcelChoice(false);
  assert.equal(h.ui.app.correctionPayload.edit_scope,'remove_excel');
  assert.equal(h.ui.app.correctionPayload.details.form6,'');
  assert.equal(h.ui.app.correctionPayload.details.certificate_type,'sita');
});

test('Canceling type-change deletion keeps the editor and sends no generation request', async () => {
  const h = harness(), form = editing(h);
  Object.assign(h.ui.app.editor,{existing:true,firstGeneration:true,reviewEdits:true});
  form.values.certificate_type='sita'; h.ui.app.editor.template='sita'; h.ui.app.editor.fields.sita={};
  vm.runInContext('submitGeneration = async payload => {app.correctionPayload = payload;};',h.context);
  await h.ui.saveDetails({preventDefault(){},currentTarget:form});
  const pending=h.ui.saveExcelChoice(false);
  assert.equal(h.controls.get('#action-confirmation-modal').open,true);
  assert.match(h.controls.get('#action-confirmation-message').textContent,/SN 1001/);
  assert.equal(h.controls.get('#cancel-action').textContent,'Keep editing');
  assert.equal(h.controls.get('#cancel-action').focused,true);
  h.ui.settleActionConfirmation(false); await pending;
  assert.equal(h.ui.app.correctionPayload,undefined);
  assert.equal(h.controls.get('#details-modal').open,true);
  assert.equal(form.values.certificate_type,'sita');
});

test('Type-change deletion cannot submit changed details after its confirmation', async () => {
  const h=harness(),form=editing(h);
  Object.assign(h.ui.app.editor,{existing:true,reviewEdits:true});
  form.values.certificate_type='sita';h.ui.app.editor.template='sita';h.ui.app.editor.fields.sita={};
  vm.runInContext('submitGeneration=async payload=>{app.correctionPayload=payload;}; confirmAction=async()=>{document.querySelector("#details-form").values.supplier="Changed after review";return true;};',h.context);
  await h.ui.saveDetails({preventDefault(){},currentTarget:form});
  await h.ui.saveExcelChoice(false);
  assert.equal(h.ui.app.correctionPayload,undefined);
  assert.match(h.controls.get('#details-error').textContent,/details changed/i);
});

test('First-PDF confirmation explains any required Form 6 or reference writes in the No choice', async () => {
  const h=harness(),form=editing(h);
  Object.assign(h.ui.app.editor,{existing:true,firstGeneration:true,reviewEdits:true});
  form.values.form6='F6-entered';form.values.reference='REF-entered';form.values.supplier='Changed company';
  await h.ui.saveDetails({preventDefault(){},currentTarget:form});
  assert.match(h.controls.get('#excel-save-description').textContent,/Form 6 number and reference number must be recorded in Excel/);
  assert.equal(h.controls.get('#excel-save-no').textContent,'No, only save required details');
});

test('Data Destruction checkbox is separate from the main type and warns on closing unsaved choices', () => {
  const h=harness();editing(h);
  h.ui.app.editor.existing=true;
  h.controls.get('#destruction-needed').checked=true;
  const payload=h.ui.detailsBody(h.controls.get('#details-form'));
  assert.equal(payload.details.certificate_type,'standard');
  assert.equal(payload.data_destruction_needed,true);
  assert.equal(payload.details.data_destruction_needed,undefined);
  assert.equal(h.ui.editorDirty(),true);
  assert.equal(h.ui.requestDetailsClose(),false);
  assert.equal(h.controls.get('#unsaved-modal').open,true);
});

test('Saving an HDD request clears its finished draft and preserves any other unsaved edits', async () => {
  for (const otherEdits of [false,true]) {
    const request={source_kind:'manual',record_id:'hdd-request',certificate_type:'destruction',plant:'plant1',destruction_source:{plant:'plant1',case_id:'1001'}};
    const h=harness({fetch:async()=>({ok:true,json:async()=>({record_id:'hdd-request',workspace_revision:'request-saved',state:{rows:[request]}})})});
    const form=editing(h);
    h.ui.app.editor.issuedCertificate=otherEdits;
    if(otherEdits) form.values.supplier='Still being edited';
    h.controls.get('#destruction-needed').checked=true;
    h.ui.saveWorkingDraft();
    vm.runInContext('setState=state=>{app.state=state;}; openDetails=async row=>{app.openedHdd=row;};',h.context);
    await h.ui.saveInwardDetails();
    assert.equal(h.ui.editorDirty(),otherEdits);
    if(otherEdits) {
      assert.equal(h.ui.app.openedHdd,undefined);
      assert.equal(h.ui.readWorkingDraft().values.supplier,'Still being edited');
      assert.equal(h.ui.readWorkingDraft().editor.baseline.data_destruction_needed,true);
    } else {
      assert.equal(h.ui.app.openedHdd,request);
      assert.equal(h.ui.readWorkingDraft(),null);
      assert.equal(h.ui.requestDetailsClose(),true);
    }
  }
});

test('A saved HDD request stays editable and unticking saves removal without changing Excel', async () => {
  const h=harness({fetch:async()=>({ok:true,json:async()=>({state:{rows:[]},workspace_revision:'removed'})})});
  const form=editing(h);h.ui.app.editor.existing=true;h.ui.app.editor.issuedCertificate=true;
  const request={source_kind:'manual',record_id:'request',certificate_type:'destruction',certificate_status:'ready',
    destruction_source:{plant:'plant1',case_id:'1001'}};
  h.ui.app.state.rows=[request];h.ui.updateDestructionChoice();
  const checkbox=h.controls.get('#destruction-needed');assert.equal(checkbox.disabled,false);assert.equal(checkbox.checked,true);
  h.ui.app.editor.baseline=h.ui.detailsBody(form);checkbox.checked=false;
  h.ui.updateDestructionChoice(h.ui.detailsBody(form));assert.equal(checkbox.checked,false);
  assert.equal(h.controls.get('#save-inward-details').hidden,false);
  assert.equal(h.controls.get('#save-inward-details').textContent,'Save request');
  vm.runInContext('setState=state=>{app.state=state;};',h.context);
  await h.ui.saveInwardDetails();
  assert.equal(h.requests[0].endpoint,'/api/destruction/request');assert.equal(h.requests[0].payload.needed,false);
  assert.equal(h.ui.editorDirty(),false);assert.equal(h.ui.app.editor.workspaceRevision,'removed');
});

test('The linked HDD editor shows removal for an unissued request', async () => {
  const link={plant:'plant1',case_id:'1001'};
  const request={source_kind:'manual',record_id:'request',certificate_type:'destruction',certificate_status:'draft',
    destruction_requested:true,destruction_source:link};
  const h=harness({fetch:async()=>({ok:true,json:async()=>({state:{rows:[{...request,destruction_requested:false}]},workspace_revision:'removed'})})});
  const form=editing(h);
  Object.assign(h.ui.app.editor,{existing:true,template:'destruction',sourceSn:null,sourcePlant:'plant1',destructionSource:link,
    recordId:'request',issuedCertificate:false,fields:{destruction:{}}});
  form.values.source_kind='manual';form.values.certificate_type='destruction';
  h.ui.app.state.rows=[request];h.ui.updateDestructionChoice();
  const checkbox=h.controls.get('#destruction-needed');
  assert.equal(h.controls.get('#destruction-needed-card').hidden,false);
  assert.equal(h.controls.get('#destruction-needed-label').textContent,'Keep this Data Destruction request active');
  assert.equal(checkbox.checked,true);assert.equal(checkbox.disabled,false);
  h.ui.app.editor.baseline=h.ui.detailsBody(form);
  checkbox.checked=false;await h.ui.changeDestructionChoice();
  h.ui.updateDestructionChoice(h.ui.detailsBody(form));
  assert.equal(h.controls.get('#save-inward-details').hidden,false);
  assert.equal(h.controls.get('#save-inward-details').textContent,'Save request');
  assert.match(h.controls.get('#destruction-choice-status').textContent,/Save to remove this request/);
  vm.runInContext('setState=state=>{app.state=state;};',h.context);
  await h.ui.saveInwardDetails();
  assert.equal(h.requests[0].endpoint,'/api/destruction/request');
  assert.deepEqual(h.requests[0].payload,{plant:'plant1',sn:'1001',expected_revision:'revision-1',needed:false});
  assert.equal(h.ui.editorDirty(),false);
  assert.equal(h.controls.get('#save-inward-details').hidden,true);
  assert.equal(h.ui.app.editor.workspaceRevision,'removed');
});

test('An issued HDD shows its status and cannot be removed from either editor', () => {
  const h=harness(),form=editing(h);h.ui.app.editor.existing=true;
  const request={source_kind:'manual',record_id:'request',certificate_type:'destruction',certificate_status:'generated_manual',
    has_issued_certificate:true,pdf_filename:'HDD.pdf',destruction_source:{plant:'plant1',case_id:'1001'}};
  h.ui.app.state.rows=[request];
  const checkbox=h.controls.get('#destruction-needed');
  for (const kind of ['standard','destruction']) {
    if (kind === 'destruction') {
      Object.assign(h.ui.app.editor,{template:kind,sourceSn:null,recordId:request.record_id,destructionSource:request.destruction_source});
      form.values.source_kind='manual';form.values.certificate_type=kind;
    }
    h.ui.updateDestructionChoice();
    assert.equal(h.controls.get('#destruction-request-toggle').hidden,true);
    assert.equal(h.controls.get('#destruction-issued-label').hidden,false);
    assert.equal(h.controls.get('#destruction-choice-status').textContent,'PDF ready to send.');
    assert.equal(checkbox.disabled,true);
    checkbox.checked=false;
    assert.equal(h.ui.detailsBody(form).data_destruction_needed,true);
    h.ui.changeDestructionChoice();
    assert.equal(checkbox.checked,true);
  }
  assert.equal(h.requests.length,0);
});

test('Canceling the Excel choice or changing details while it is open never submits stale edits', async () => {
  for (const action of ['cancel','change']) {
    const h = harness(), form = editing(h); h.ui.wire();
    Object.assign(h.ui.app.editor,{existing:true,firstGeneration:false});
    form.values.supplier = 'Correction';
    vm.runInContext('submitGeneration = async payload => {app.correctionPayload = payload;};',h.context);
    await h.ui.saveDetails({preventDefault(){},currentTarget:form});
    if (action === 'cancel') h.controls.get('#excel-save-modal').close();
    else form.values.supplier = 'Different correction';
    await h.ui.saveExcelChoice(true);
    assert.equal(h.ui.app.correctionPayload,undefined);
    assert.equal(h.controls.get('#details-modal').open,true);
  }
});

test('Excel save comparison does not report harmless numeric formatting as a correction', () => {
  const h = harness();
  const before = {details:{quantity:'2',items:[{item:'Asset',quantity:'2',per_piece_weight:'3.5',weight:'7'}]}};
  const after = {details:{quantity:'2.0',items:[{item:'Asset',quantity:'2.0',per_piece_weight:'3.50',weight:'7.0'}]}};
  assert.equal(h.ui.fieldDifferences(before,after).length,0);
});

test('Template filters use the issued template even when Excel still says Standard', () => {
  const h = harness(); h.controls.get('#template-filter').value = 'sita';
  assert.equal(h.ui.matches({certificate_type:'standard',last_certificate_details:{certificate_type:'sita'}}),true);
  h.controls.get('#template-filter').value = 'standard';
  assert.equal(h.ui.matches({certificate_type:'standard',last_certificate_details:{certificate_type:'sita'}}),false);
});

test('Restoring an older first-generation draft keeps Excel facts and discards unsupported printed overrides', async () => {
  const h = harness(), form = editing(h);
  h.ui.app.editor.fields.standard = {company_display:'Old edited name',weight_display:'999 KG'};
  h.ui.saveWorkingDraft(); h.ui.app.draftRecovery = {draft:h.ui.readWorkingDraft()};
  h.ui.app.state.rows = [{plant:'plant1',sn:'1001',certificate_status:'ready',supplier:'Excel company',
    location:'Pune',inward_date:'2026-10-01',comments:'Already entered',transporter_details:'Eco Recycling',
    docket_vehicle_no:'DC-44',weight_slip_scale:'WS-44',items:[{item:'Asset',quantity:1,per_piece_weight:10,weight:10}]}];
  vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {};',h.context);
  await h.ui.restoreWorkingDraft();
  assert.equal(form.values.supplier,'Excel company');
  assert.equal(h.ui.app.editor.firstGeneration,true);
  assert.equal(h.controls.get('#excel-detail-fields').hidden,false);
  assert.equal(h.ui.app.editor.fields.standard.company_display,undefined);
  assert.equal(h.ui.app.editor.fields.standard.weight_display,undefined);
  assert.equal(form.values.comments,'Already entered');
  assert.equal(form.values.weight_slip_scale,'WS-44');
  assert.equal(h.requests.length,0);
});

function clientEntry(h) {
  const form = h.controls.get('#add-client-form');
  const defaults = {plant:'plant1',sn:'',supplier:'',location:'',inward_date:'',form6:'',reference:'',comments:'',transporter_details:'',docket_vehicle_no:'',weight_slip_scale:''};
  form.values = {...defaults};
  const inputs = Object.fromEntries(Object.keys(defaults).map(name => {
    const input = new Node('input'); input.name = name;
    Object.defineProperty(input,'value',{get:() => form.values[name],set:value => { form.values[name] = value; }});
    const field = new Node('label'); input.closest = () => field;
    return [name,input];
  }));
  form.elements = {namedItem:name => inputs[name]};
  form.reset = () => { form.values = {...defaults}; };
  form.reportValidity = () => Boolean(form.values.supplier.trim());
  form.querySelectorAll = () => Object.values(inputs);
  h.ui.app.workspaceRevision = 'revision-1';
  return {form,inputs};
}

test('Add client always requests an automatic SN and exposes no manual SN field', () => {
  const h=harness(), {form,inputs}=clientEntry(h);
  form.values.sn='OLD-DRAFT-NUMBER';
  assert.equal(h.ui.validateClientSerialNumber(),true);
  assert.equal(form.values.sn,''); assert.equal(inputs.sn.validationMessage,'');
  const html=fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  assert.match(html, /<input type="hidden" name="sn" value="">/);
});

test('Company order follows received date across statuses and keeps unknown dates last', () => {
  const h = harness();
  const rows = [
    {sn:'3',plant:'plant1',inward_date:'',certificate_status:'ready'},
    {sn:'2',plant:'plant1',inward_date:'2026-08-12',certificate_status:'ready'},
    {sn:'9',plant:'plant1',inward_date:'11-08-2026',certificate_status:'completed'},
    {sn:'1',plant:'plant1',inward_date:'not a date',certificate_status:'alert'},
  ];
  assert.deepEqual(rows.sort(h.ui.compareRows).map(row => row.sn),['9','2','1','3']);
});

test('Add client leaves certificate template selection to certificate creation', () => {
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  const form = html.match(/<form id="add-client-form">([\s\S]*?)<\/form>/)[1];
  assert.doesNotMatch(form,/Standard COD|name="certificate_type"/);
  assert.doesNotMatch(form,/class="client-entry-id"|Numeric quantities can use|Required for every inward/);
  assert.match(form,/id="client-serial-number-hint"[^>]*hidden/);
  assert.match(html,/<select name="certificate_type"><option value="standard">Recycling/);
});

test('Add client keeps calculation and additional fields visible without dropdowns', () => {
  const h = harness(); clientEntry(h); h.ui.openAddClient();
  assert.equal(h.controls.get('#client-serial-number-hint').hidden,true);
  assert.equal(h.controls.get('#client-serial-number-hint').textContent,'');
  const inputs = h.ui.app.clientMaterialRows[0];
  assert.equal(inputs.hint.hidden,true);
  assert.equal(inputs.hint.textContent,'');
  assert.equal(h.controls.get('#client-material-weight-total').hidden,true);
  const card = h.controls.get('#client-material-rows').children[0];
  const grid = card.children[1];
  assert.match(text(grid.children[0]),/Material or asset/);
  assert.match(text(grid.children[1]),/Weight \(KG\)/);
  const calculation = all(grid,node => node.className?.includes('client-material-calculation'))[0];
  assert.ok(calculation);
  assert.equal(all(grid,node => node.tag === 'details').length,0);
  assert.match(text(calculation),/Calculate weight.*Per Pcs Weight.*Quantity in Pcs.*Weight \(KG\)/);
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  const optional = html.match(/<section class="client-entry-more-details"[\s\S]*?<\/section>/)[0];
  for (const name of ['reference','comments','transporter_details','docket_vehicle_no','weight_slip_scale']) {
    assert.ok(optional.includes(`name="${name}"`),name);
  }
});

test('Opening Excel for repair gives a short success notice', async () => {
  const warning = 'Serial Number 10 is repeated. The matching rows are selected in Excel; give each inward a unique Serial Number.';
  const h = harness({fetch:async () => ({ok:true,json:async () => ({selected:true,selection:'repair',warning})})});
  await h.ui.openClientInExcel('plant1','10');
  assert.equal(h.controls.get('#toast-message').textContent,'Excel opened. Check the SN, then save and refresh.');
  assert.equal(h.controls.get('#dismiss-toast').hidden,true);
  assert.equal(h.controls.get('#toast').classList.contains('error'),false);
  assert.equal(h.ui.app.openingExcel,false);
});

test('Add client follows the plant and keeps unfinished inputs without a certificate preview', () => {
  const h = harness(), {form,inputs} = clientEntry(h);
  h.controls.get('#plant-filter').value = 'plant2';
  h.ui.openAddClient();
  assert.equal(form.values.plant,'plant2');
  assert.equal(h.ui.app.clientMaterialRows.length,1);
  assert.equal(h.requests.length,0);
  const requestId = h.ui.app.clientRequestId;
  form.values.supplier = 'New company'; h.ui.app.clientMaterialRows[0].item.value = 'Laptop';
  h.controls.get('#add-client-modal').close(); h.ui.openAddClient();
  assert.equal(form.values.supplier,'New company');
  assert.equal(h.ui.clientMaterialValues()[0].item,'Laptop');
  assert.equal(h.ui.app.clientRequestId,requestId);
  assert.notEqual(inputs.form6.closest().hidden,true);
});

test('Add client saves with Form 6 blank and leaves it for certificate creation', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({client_added:true,
    client:{sn:'1002',plant:'plant1',supplier:'New company',status:'blank'},
    state:{rows:[],scan_complete:true}})})});
  const {form} = clientEntry(h);
  h.ui.openAddClient(); form.values.supplier = 'New company';
  await h.ui.saveClient();
  assert.equal(h.requests.length,1);
  assert.equal(h.requests[0].payload.details.form6,'');
});

test('Add client calculates numeric quantity but leaves text quantities for manual weight', () => {
  const h = harness();
  h.ui.renderClientMaterials([{item:'Laptop',quantity:'2',per_piece_weight:'3.5'}]);
  const material = h.ui.app.clientMaterialRows[0];
  assert.equal(material.weight.value,'7');
  assert.equal(h.ui.clientMaterialValues()[0].weight,'');
  material.quantity.value = '1 lot'; material.quantity.dispatch('input');
  assert.equal(material.weight.value,'');
  assert.equal(material.hint.textContent,'Enter weight directly.');
  material.weight.value = '12';
  assert.equal(h.ui.clientMaterialValues()[0].weight,'12');
});

test('Add client saves every material once and shows the new Excel entry in the list', async () => {
  let resolveSave;
  const h = harness({fetch:() => new Promise(resolve => { resolveSave = resolve; })});
  const {form} = clientEntry(h); h.ui.openAddClient();
  form.values.supplier = 'New company'; form.values.form6 = 'FORM-NEW';
  form.values.sn = '1000'; form.values.inward_date = '2026-06-15';
  form.values.comments = 'Return visit';
  form.values.transporter_details = 'ABC Logistics';
  form.values.docket_vehicle_no = 'D-123';
  form.values.weight_slip_scale = '18.5';
  h.ui.renderClientMaterials([{item:'Laptop',quantity:'2',per_piece_weight:'3.5'},{item:'Printer',weight:'8'}]);
  const pending = h.ui.saveClient();
  await h.ui.saveClient();
  assert.equal(h.requests.length,1);
  assert.equal(h.requests[0].endpoint,'/api/clients/add');
  assert.equal(h.requests[0].payload.expected_revision,'revision-1');
  assert.equal(h.requests[0].payload.details.items.length,2);
  assert.equal(h.requests[0].payload.details.items[1].weight,'8');
  assert.equal(h.requests[0].payload.details.sn,'');
  assert.equal(h.requests[0].payload.details.inward_date,'2026-06-15');
  assert.equal(h.requests[0].payload.details.comments,'Return visit');
  assert.equal(h.requests[0].payload.details.transporter_details,'ABC Logistics');
  assert.equal(h.requests[0].payload.details.docket_vehicle_no,'D-123');
  assert.equal(h.requests[0].payload.details.weight_slip_scale,'18.5');
  assert.equal(Object.hasOwn(h.requests[0].payload.details,'certificate_type'),false);
  resolveSave({ok:true,json:async () => ({client_added:true,client:{sn:'1002',plant:'plant1',supplier:'New company',status:'ready'},
    state:{rows:[{sn:'1002',plant:'plant1',supplier:'New company',certificate_status:'ready',outcome:'needs_review',
      warnings:['Reference number is required to create the certificate.']}],scan_complete:true},workspace_revision:'revision-2'})});
  await pending;
  assert.equal(h.controls.get('#add-client-modal').open,false);
  assert.match(h.controls.get('#toast-message').textContent,/Open Review & edit to finish the certificate details/);
  assert.equal(h.controls.get('#search').value,'1002');
  assert.equal(h.controls.get('#status-filter').value,'all');
  assert.equal(h.ui.app.workspaceRevision,'revision-2');
  assert.equal(h.ui.app.clientRequestId,null);
  assert.doesNotMatch(h.controls.get('#toast-message').textContent,/Ready to create/);
});

test('Failed Add client retains the form and stable request ID for a safe retry', async () => {
  const h = harness({fetch:async () => ({ok:false,status:400,json:async () => ({error:'Save and close Plant 1.xlsx in Excel, then try again.'})})});
  const {form} = clientEntry(h); h.ui.openAddClient(); form.values.supplier = 'New company'; form.values.form6 = 'FORM-NEW';
  const requestId = h.ui.app.clientRequestId;
  await h.ui.saveClient();
  assert.equal(h.controls.get('#add-client-modal').open,true);
  assert.equal(form.values.supplier,'New company');
  assert.equal(h.controls.get('#add-client-error').hidden,false);
  assert.equal(h.controls.get('#save-client').disabled,false);
  await h.ui.saveClient();
  assert.equal(h.requests[0].payload.request_id,requestId);
  assert.equal(h.requests[1].payload.request_id,requestId);
});

test('Add client retries a stale save with the refreshed revision and the same draft', async () => {
  let adds = 0;
  const h = harness({fetch:async path => {
    if (path === '/api/check') return {ok:true,json:async () => ({
      workspace_revision:'revision-2',state:{rows:[],scan_complete:true},
    })};
    adds += 1;
    return {ok:false,status:400,json:async () => ({
      error:adds === 1 ? 'The files changed since you opened these details.' : 'Save and close Excel.',
    })};
  }});
  const {form} = clientEntry(h);
  h.ui.openAddClient();
  form.values.supplier = 'New company';
  await h.ui.saveClient();
  assert.equal(h.controls.get('#add-client-modal').open,true);
  assert.match(h.controls.get('#add-client-error').textContent,/Excel list was refreshed/);
  await h.ui.saveClient();
  assert.equal(h.requests[0].payload.expected_revision,'revision-1');
  assert.equal(h.requests[2].payload.expected_revision,'revision-2');
  assert.equal(h.requests[2].payload.request_id,h.requests[0].payload.request_id);
  assert.equal(h.requests[2].payload.details.supplier,'New company');
});

test('Add client can start a new request after the server resolves an old request', async () => {
  const h = harness({fetch:async () => ({ok:false,status:400,json:async () => ({
    client_request_ended:true,error:'This client was already added and later removed.',
  })})});
  const {form} = clientEntry(h);
  h.ui.openAddClient();
  form.values.supplier = 'New company';
  await h.ui.saveClient();
  assert.equal(h.requests.length,1);
  assert.equal(h.controls.get('#add-client-modal').open,true);
  await h.ui.saveClient();
  assert.notEqual(h.requests[1].payload.request_id,h.requests[0].payload.request_id);
  assert.equal(h.requests[1].payload.details.supplier,'New company');
});

test('Every Excel client has one direct Open in Excel action, regardless of status', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({selected:true})})});
  for (const status of ['ready','alert','blank','generated','completed','na']) {
    const row = {source_kind:'workbook',plant:'plant2',sn:'SHARED',supplier:'Exact client',certificate_status:status,pdf_filename:'0001_Exact client_2.pdf'};
    const buttons = all(h.ui.makeActions(row,0), node => node.tag === 'button' && node.textContent === 'Open in Excel');
    assert.equal(buttons.length,1,status);
    assert.match(buttons[0].attributes['aria-label'],/Plant 2.*Exact client/);
    buttons[0].dispatch('click');
    await Promise.resolve(); await Promise.resolve(); await Promise.resolve();
    // Finish the pending API call before checking the next status.
    await new Promise(resolve => setImmediate(resolve));
    assert.deepEqual(h.requests.at(-1).payload,{plant:'plant2',sn:'SHARED'});
  }
  const manual = {source_kind:'manual',plant:'plant1',record_id:'manual-1',certificate_status:'generated_manual'};
  assert.equal(all(h.ui.makeActions(manual,0), node => node.textContent === 'Open in Excel').length,0);
});

test('Saved PDFs and generation results open Excel only when linked to an exact workbook client', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({selected:true})})});
  h.ui.app.state.certificate_inventory = {entries:[
    {path:'Certificates/Auto Generated/0001_Acme_2.pdf',present:true,cods:['0001'],managed_owner:{source_kind:'workbook',plant:'plant2',key:'1001'}},
    {path:'Certificates/Created in App/0002_Manual_1.pdf',present:true,cods:['0002'],managed_owner:{source_kind:'manual',plant:'plant1',key:'manual-1'}},
    {path:'Certificates/Manually Added/0003_External_1.pdf',present:true,cods:['0003']},
  ]};
  h.ui.renderFolderInventory();
  assert.equal(all(h.controls.get('#folder-inventory-files'),node => node.textContent === 'Open in Excel').length,1);
  h.ui.showSavedFile(h.ui.app.state.certificate_inventory.entries[0]);
  assert.equal(h.controls.get('#saved-file-excel').hidden,false);
  assert.equal(h.controls.get('#saved-file-excel').dataset.sn,'1001');
  h.ui.showSavedFile(h.ui.app.state.certificate_inventory.entries[1]);
  assert.equal(h.controls.get('#saved-file-excel').hidden,true);
  h.ui.showGenerationSuccess({created:2,certificates:[
    {filename:'0001_Acme_2.pdf',relative_path:'Auto Generated/0001_Acme_2.pdf',plant:'plant2',sn:'1001',source_kind:'workbook',supplier:'Acme'},
    {filename:'0002_Manual_1.pdf',relative_path:'Created in App/0002_Manual_1.pdf',plant:'plant1',source_kind:'manual',supplier:'Manual'},
  ]});
  const buttons = all(h.controls.get('#created-certificates'),node => node.textContent === 'Open in Excel');
  assert.equal(buttons.length,1);
  buttons[0].dispatch('click'); await new Promise(resolve => setImmediate(resolve));
  assert.deepEqual(h.requests.at(-1).payload,{plant:'plant2',sn:'1001'});
});

test('Opening Excel prevents duplicate launches and announces success only after selection is confirmed', async () => {
  let resolveOpen;
  const h = harness({fetch:() => new Promise(resolve => {resolveOpen=resolve;})});
  const pending = h.ui.openClientInExcel('plant1','1001');
  await h.ui.openClientInExcel('plant2','1002');
  assert.equal(h.requests.length,1);
  resolveOpen({ok:true,json:async () => ({selected:true})}); await pending;
  assert.match(h.controls.get('#toast-message').textContent,/client’s rows selected/);
  assert.equal(h.ui.app.openingExcel,false);
  const failure = h.ui.openClientInExcel('plant1','1001');
  resolveOpen({ok:true,json:async () => ({ok:true,selected:false})}); await failure;
  assert.match(h.controls.get('#toast-message').textContent,/could not confirm/);
  assert.equal(h.controls.get('#busy-overlay').hidden,true);
});

test('Data destruction requests remain attention items after PDF creation until Completed', () => {
  const h = harness();
  const row = {source_kind:'manual',record_id:'request-1',plant:'plant1',sn:'1001',supplier:'Acme',
    certificate_type:'destruction',certificate_status:'generated_manual',pdf_filename:'0303_Acme.pdf',destruction_pending:true};
  h.ui.app.state.rows = [row];
  assert.equal(h.ui.needsAttention(row),true);
  assert.match(h.ui.nextStep(row),/Send the Data Destruction PDF.*Completed/);
  h.ui.summarize();
  assert.equal(h.controls.get('#destruction-pending').hidden,false);
  assert.match(h.controls.get('#destruction-pending').textContent,/Data Destruction · 1 to finish/);
  assert.match(h.controls.get('#destruction-pending').title,/Payment may depend/);
  row.destruction_pending = false; row.certificate_status = 'completed'; h.ui.summarize();
  assert.equal(h.controls.get('#destruction-pending').hidden,true);
});

test('Recycling and HDD statuses stay on their own rows through issuance and completion', () => {
  const h = harness();
  const inward = {source_kind:'workbook',plant:'plant1',sn:'1001',
    certificate_type:'standard',certificate_status:'ready',reminder_due:true,age_days:39};
  const hdd = {source_kind:'manual',plant:'plant1',record_id:'request',certificate_type:'destruction',
    destruction_source:{plant:'plant1',case_id:'1001'},destruction_pending:true,
    certificate_status:'ready',reminder_due:true,age_days:39};
  h.ui.app.state.rows=[inward,hdd];
  const recycling = text(h.ui.makeStatus(inward)).trim();
  assert.equal(recycling,'Ready 39 days since received · Pending');
  assert.match(text(h.ui.makeStatus(hdd)),/PDF to create/);
  Object.assign(hdd,{certificate_status:'generated',pdf_filename:'HDD.pdf'});
  assert.equal(text(h.ui.makeStatus(inward)).trim(),recycling);
  assert.equal(text(h.ui.makeStatus(hdd)).trim(),'Generated PDF ready to send');
  assert.match(text(h.ui.makeStatus({...hdd,pdf_missing:true})),/PDF missing/);
  hdd.certificate_status='completed';
  assert.equal(text(h.ui.makeStatus(hdd)).trim(),'Completed');
  assert.equal(text(h.ui.makeStatus(inward)).trim(),recycling);
});

test('Data destruction payload retains its disabled type and inward link during editing', () => {
  const h = harness(); const form = editing(h);
  form.values.source_kind = 'manual'; form.values.certificate_type = 'destruction';
  form.elements.namedItem('certificate_type').disabled = true;
  const link = {plant:'plant1',case_id:'1001'};
  Object.assign(h.ui.app.editor, {template:'destruction',destructionSource:link,recordId:'request-1',sourceSn:null,
    fields:{destruction:{destruction_date:'2026-10-16',device_receipts:'[{"date":"2026-10-01","count":"12"}]'}}});
  const payload = h.ui.detailsBody(form);
  assert.equal(payload.source_kind,'manual');
  assert.equal(payload.details.certificate_type,'destruction');
  assert.deepEqual(payload.details.destruction_source,link);
  h.ui.saveWorkingDraft();
  assert.equal(h.ui.readWorkingDraft().values.certificate_type,'destruction');
});

const linkedBase = {source_kind:'workbook',plant:'plant1',sn:'1001',supplier:'Acme',location:'Mumbai',certificate_type:'standard',cod_no:'0301',certificate_status:'completed',items:[{item:'Computers',weight:20}]};
const linkedHdd = {source_kind:'manual',plant:'plant1',record_id:'hdd-1',supplier:'Acme',location:'Mumbai',certificate_type:'destruction',cod_no:'0302',certificate_status:'generated',destruction_source:{plant:'plant1',case_id:'1001'},certificate_fields:{device_receipts:'[{"date":"2026-10-01","count":"4"}]'}};
const settleSwitch = () => new Promise(resolve => setImmediate(resolve));
function linkedEditor(h, rows=[linkedBase,linkedHdd]) {
  const form=editing(h);h.ui.app.editor.existing=true;h.ui.app.state.rows=rows;
  vm.runInContext('globalThis.opened=[]; openDetails=async row=>{opened.push(row);};',h.context);
  return form;
}

test('A stale linked HDD save can be retried after refresh only when its records still match', async () => {
  const hdd={...linkedHdd,raw_details:{supplier:'Acme',certificate_type:'destruction',cod_no:'0302'}};
  for (const changed of [false,true]) {
    const inward=changed ? {...linkedBase,supplier:'Changed in Excel'} : linkedBase;
    const h=harness({fetch:async endpoint => endpoint === '/api/check'
      ? {ok:true,json:async()=>({workspace_revision:'revision-2',state:{scan_complete:true,rows:[inward,hdd]}})}
      : {ok:false,status:400,json:async()=>({error:'The files changed since you opened these details.'})}});
    const form=linkedEditor(h,[linkedBase,hdd]);
    Object.assign(h.ui.app.editor,{template:'destruction',sourceSn:null,recordId:hdd.record_id,
      destructionSource:hdd.destruction_source,fields:{destruction:{}}});
    form.values.certificate_type='destruction';form.values.source_kind='manual';
    form.values.supplier='Edited for this certificate';
    vm.runInContext('app.editor.linkedHddSaveSnapshot=linkedHddSaveSnapshot();',h.context);
    await h.ui.submitGeneration(h.ui.detailsBody(form),false,'/api/manual/generate',true);
    assert.deepEqual(h.requests.map(request=>request.endpoint),['/api/manual/generate','/api/check']);
    assert.equal(form.values.supplier,'Edited for this certificate');
    assert.equal(h.ui.detailsBody(form).expected_revision,changed?'revision-1':'revision-2');
    assert.match(h.controls.get('#details-error').textContent,changed?/Excel has changed/:/still match/);
  }
});

test('Switching from the inward opens its own HDD record without editing either certificate', async () => {
  const h=harness(),form=linkedEditor(h);h.ui.wire();
  h.ui.updateDestructionChoice();
  h.ui.app.editor.baseline=h.ui.detailsBody(form);
  form.values.certificate_type='destruction';
  form.dispatch('input',{target:{name:'certificate_type',value:'destruction'}});
  form.dispatch('change',{target:{name:'certificate_type',value:'destruction'}});
  await settleSwitch();
  assert.equal(form.values.certificate_type,'standard');
  assert.equal(h.context.opened[0],linkedHdd);assert.equal(h.context.opened[0].cod_no,'0302');
  assert.equal(h.requests.length,0);assert.equal(h.controls.get('#unsaved-modal').open,undefined);
  assert.equal(linkedBase.cod_no,'0301');
});

test('Switching back from linked HDD opens the exact inward in its plant', async () => {
  const h=harness(),form=linkedEditor(h,[{...linkedBase,plant:'plant2',cod_no:'9999'},linkedBase,linkedHdd]);
  Object.assign(h.ui.app.editor,{template:'destruction',sourceSn:null,recordId:'hdd-1',destructionSource:linkedHdd.destruction_source});
  form.values.certificate_type='destruction';form.values.source_kind='manual';
  h.ui.updateDestructionChoice();
  h.ui.app.editor.baseline=h.ui.detailsBody(form);
  const companion=h.controls.get('#open-companion-certificate');
  assert.equal(companion.hidden,false);assert.equal(companion.textContent,'Open Recycling');
  h.ui.openCompanionCertificate();await settleSwitch();
  assert.equal(h.context.opened[0],linkedBase);assert.equal(h.requests.length,0);
});

test('After Excel removal the local certificate and HDD switch by local ID without opening a reused SN', async () => {
  const base={...linkedBase,source_kind:'manual',record_id:'local-base',removed_inward:{plant:'plant1',case_id:'1001'},certificate_type:'siemens'};
  const hdd={...linkedHdd,destruction_source:{plant:'plant1',case_id:'1001',record_id:base.record_id}};
  const h=harness(),form=linkedEditor(h,[{...linkedBase,supplier:'Different replacement inward'},base,hdd]);
  Object.assign(h.ui.app.editor,{sourceSn:null,localInwardId:base.record_id,recordId:base.record_id,template:'siemens',fields:{siemens:{}}});
  form.values.source_kind='manual';form.values.certificate_type='siemens';h.ui.app.editor.baseline=h.ui.detailsBody(form);
  assert.equal(h.ui.editorInwardRow(),base);assert.equal(h.ui.editorDestructionRequest(),hdd);
  assert.equal(h.ui.excelTarget(hdd),null);
  h.ui.switchLinkedCertificate('destruction');await settleSwitch();
  assert.equal(h.context.opened[0],hdd);
  Object.assign(h.ui.app.editor,{localInwardId:null,recordId:hdd.record_id,template:'destruction',destructionSource:hdd.destruction_source,fields:{destruction:{}}});
  form.values.certificate_type='destruction';h.ui.app.editor.baseline=h.ui.detailsBody(form);
  h.ui.switchLinkedCertificate('siemens');await settleSwitch();
  assert.equal(h.context.opened[1],base);assert.equal(h.requests.length,0);
});

test('A local certificate retained after Excel deletion can add its own HDD request', async () => {
  const base={...linkedBase,source_kind:'manual',record_id:'local-base',removed_inward:{plant:'plant1',case_id:'1001'},certificate_type:'siemens'};
  const hdd={...linkedHdd,destruction_source:{plant:'plant1',case_id:'1001',record_id:base.record_id}};
  const h=harness({fetch:async()=>({ok:true,json:async()=>({record_id:hdd.record_id,state:{rows:[base,hdd]}})})});
  const form=linkedEditor(h,[base]);
  vm.runInContext('setState=state=>{app.state=state;};',h.context);
  Object.assign(h.ui.app.editor,{sourceSn:null,localInwardId:base.record_id,recordId:base.record_id,template:'siemens',fields:{siemens:{}}});
  form.values.source_kind='manual';form.values.certificate_type='siemens';h.ui.app.editor.baseline=h.ui.detailsBody(form);
  h.ui.switchLinkedCertificate('destruction');await settleSwitch();
  assert.equal(h.requests.length,1);assert.equal(h.requests[0].endpoint,'/api/destruction/request');
  assert.equal(h.requests[0].payload.record_id,base.record_id);
  assert.equal(h.context.opened[0],hdd);
});

test('The Data Destruction checkbox works for an app-only certificate with no plant or SN', async () => {
  const base={source_kind:'manual',record_id:'local-only',plant:null,case_id:'',supplier:'Local client',
    certificate_type:'standard',certificate_status:'generated_manual'};
  const hdd={source_kind:'manual',record_id:'local-hdd',plant:null,supplier:'Local client',
    certificate_type:'destruction',certificate_status:'ready',
    destruction_source:{plant:null,case_id:base.record_id,record_id:base.record_id}};
  const h=harness({fetch:async()=>({ok:true,json:async()=>({record_id:hdd.record_id,workspace_revision:'after-request',state:{rows:[base,hdd]}})})});
  const form=editing(h);
  Object.assign(h.ui.app.editor,{existing:true,sourceSn:null,sourcePlant:null,localInwardId:base.record_id,
    recordId:base.record_id,template:'standard',issuedCertificate:true});
  form.values.source_kind='manual';form.values.certificate_type='standard';
  h.ui.app.state.rows=[base];
  vm.runInContext('setState=state=>{app.state=state;}; setDetailsBusy=on=>{app.detailsBusy=on;};',h.context);
  h.ui.wire();h.ui.updateDestructionChoice();
  const checkbox=h.controls.get('#destruction-needed');
  assert.equal(checkbox.disabled,false);assert.equal(h.controls.get('#destruction-needed-card').hidden,false);
  checkbox.checked=true;
  checkbox.dispatch('click');
  await new Promise(resolve=>setImmediate(resolve));
  assert.equal(h.controls.get('#action-confirmation-modal').open,true,'adding the request must show its yes/no confirmation');
  assert.match(h.controls.get('#action-confirmation-title').textContent,/Add a Data Destruction request/);
  assert.equal(h.requests.length,0,'do not save until the user confirms');
  h.controls.get('#confirm-action').dispatch('click');
  await new Promise(resolve=>setImmediate(resolve));
  assert.equal(h.requests.length,1,'checking the box must save the request immediately after confirmation');
  assert.equal(h.requests[0].endpoint,'/api/destruction/request');
  assert.deepEqual(h.requests[0].payload,{plant:null,sn:'local-only',record_id:'local-only',expected_revision:'revision-1',needed:true});
  assert.equal(checkbox.checked,true);assert.equal(checkbox.disabled,true);
  assert.equal(h.ui.app.editor.workspaceRevision,'after-request');
  assert.equal(h.controls.get('#details-error').hidden,true);
});

test('Canceling a switch preserves unsaved edits and the current dropdown; discarding opens only the selected certificate', async () => {
  const h=harness(),form=linkedEditor(h);form.values.supplier='Unsaved company';
  form.values.certificate_type='destruction';h.ui.switchLinkedCertificate('destruction');
  assert.equal(h.controls.get('#unsaved-modal').open,true);assert.equal(form.values.certificate_type,'standard');
  h.ui.finishDraftClose(false);await settleSwitch();
  assert.equal(h.context.opened.length,0);assert.equal(form.values.supplier,'Unsaved company');
  h.ui.switchLinkedCertificate('destruction');h.ui.finishDraftClose(true);await settleSwitch();
  assert.equal(h.context.opened[0],linkedHdd);assert.equal(h.requests.length,0);
});

test('HDD can open before Recycling is issued or the HDD checkbox is selected', async () => {
  let finish;
  const inward={...linkedBase,cod_no:'',certificate_status:'alert'};
  const h=harness({fetch:()=>new Promise(resolve=>{finish=resolve;})});linkedEditor(h,[inward]);
  vm.runInContext('setState=state=>{app.state=state;};',h.context);
  h.ui.updateDestructionChoice();
  assert.equal(h.controls.get('#open-companion-certificate').hidden,false);
  assert.equal(h.controls.get('#open-companion-certificate').textContent,'Open Data Destruction (HDD)');
  h.ui.openCompanionCertificate();h.ui.openCompanionCertificate();
  assert.equal(h.requests.length,1);assert.equal(h.requests[0].endpoint,'/api/destruction/request');
  assert.deepEqual(h.requests[0].payload,{plant:'plant1',sn:'1001',expected_revision:'revision-1'});
  const pending={...linkedHdd,cod_no:'',certificate_status:'draft'};
  finish({ok:true,json:async()=>({record_id:'hdd-1',state:{rows:[inward,pending]},workspace_revision:'revision-2'})});
  await settleSwitch();assert.equal(h.context.opened[0],pending);assert.equal(h.ui.app.workspaceRevision,'revision-2');
  assert.equal(h.ui.app.detailsBusy,false);assert.equal(h.ui.app.switchingCertificate,false);
});

test('A failed HDD request keeps the original editor open and shows the error', async () => {
  const h=harness({fetch:async()=>({ok:false,status:409,json:async()=>({error:'Excel has changed.'})})});linkedEditor(h,[linkedBase]);
  h.ui.switchLinkedCertificate('destruction');await settleSwitch();
  assert.equal(h.context.opened.length,0);assert.equal(h.controls.get('#details-modal').open,true);
  assert.match(h.controls.get('#details-error').textContent,/Excel has changed/);
  assert.equal(h.controls.get('#details-error').hidden,false);assert.equal(h.ui.app.detailsBusy,false);
});

test('Missing or ambiguous inward links never navigate by company name or another plant’s SN', () => {
  for(const rows of [[{...linkedBase,plant:'plant2'},linkedHdd],[linkedBase,{...linkedBase},linkedHdd]]) {
    const h=harness();linkedEditor(h,rows);
    Object.assign(h.ui.app.editor,{template:'destruction',sourceSn:null,recordId:'hdd-1',destructionSource:linkedHdd.destruction_source});
    h.ui.switchLinkedCertificate('standard');assert.equal(h.context.opened.length,0);assert.equal(h.requests.length,0);
    assert.equal(h.controls.get('#details-error').hidden,false);
  }
});

test('The linked HDD selector allows its base type; standalone HDD remains fixed', () => {
  for(const linked of [false,true]) {
    const h=harness(),form=editing(h);h.ui.app.state.rows=[linkedBase,linkedHdd];
    Object.assign(h.ui.app.editor,{existing:true,template:'destruction',sourceSn:null,destructionSource:linked?linkedHdd.destruction_source:null});
    form.values.certificate_type='destruction';form.values.source_kind='manual';
    const select=form.elements.namedItem('certificate_type');
    select.options=['standard','destruction','siemens','sita'].map(value=>({value}));
    h.ui.renderTemplateEditor();assert.equal(select.disabled,!linked);
    assert.equal(select.options[1].disabled,false);assert.equal(select.options[1].hidden,false);
    if(linked) {assert.equal(select.options[0].disabled,false);assert.equal(select.options[2].disabled,true);}
  }
});

test('Opening linked certificates resets warnings and loads each record’s own COD, receipt fields and identity', async () => {
  const h=harness(),form=editing(h);h.controls.get('#details-modal').open=false;
  form.reset=()=>{form.values={};};h.ui.app.state.rows=[linkedBase,linkedHdd];
  vm.runInContext('updateSourceControls=()=>{}; renderTemplateEditor=()=>{}; previewDetails=async()=>{};',h.context);
  for(const row of [linkedBase,linkedHdd,linkedBase]) {
    h.controls.get('#details-modal').open=false;h.controls.get('#preview-warnings').hidden=false;
    h.controls.get('#details-error').hidden=false;
    await h.ui.openDetails(row,true,true);
    assert.equal(form.values.cod_no,row.cod_no);assert.equal(form.values.certificate_type,row.certificate_type);
    assert.equal(h.ui.app.editor.recordId,row.record_id||null);
    assert.equal(h.ui.app.editor.sourceSn,row.source_kind==='manual'?null:'1001');
    assert.equal(h.controls.get('#details-error').hidden,true);assert.equal(h.controls.get('#preview-warnings').hidden,true);
    assert.match(h.controls.get('#certificate-selection-status').textContent,new RegExp(row.cod_no));
    if(row===linkedHdd)assert.equal(h.ui.app.editor.fields.destruction.device_receipts,linkedHdd.certificate_fields.device_receipts);
  }
});

test('Data destruction automatic preview works with device counts and no material weight', async () => {
  const h = harness(); const form = editing(h);
  form.values.source_kind = 'manual'; form.values.certificate_weight = ''; form.values.weight = '';
  h.ui.app.editor.template = 'destruction';
  vm.runInContext('globalThis.destructionPreviewCount=0; previewDetails=async()=>{destructionPreviewCount++}',h.context);
  h.ui.schedulePreview();
  const timer = [...h.timers.values()].find(timer => timer.delay === 1000);
  assert(timer); await timer.fn();
  assert.equal(vm.runInContext('destructionPreviewCount',h.context),1);
});

test('Reopened attachments show the retained filename and page count from the catalog', () => {
  const h = harness(); editing(h);
  const token = 'a'.repeat(32);
  h.ui.app.editor.fields.standard._annexure = token;
  h.ui.app.attachments = {[token]:{id:token,name:'Manually selected list.pdf',pages:2}};
  vm.runInContext('updateAnnexureDisplay()',h.context);
  assert.equal(h.controls.get('#annexure-name').hidden,false);
  assert.match(h.controls.get('#annexure-name').textContent,/Manually selected list.pdf.*2 pages/);
});

test('Reference is preserved exactly and SN is the visible record label', () => {
  const h = harness(); const form = editing(h);
  form.values.reference = 'MY MANUAL REF / 0007-A';
  assert.equal(h.ui.detailsBody(form).details.reference,'MY MANUAL REF / 0007-A');
  assert.equal(h.ui.detailsBody(form).details.reference,'MY MANUAL REF / 0007-A');
  assert.equal(vm.runInContext('fieldLabels.case_id',h.context),'SN');
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  assert.match(html,/<span>SN /);
  assert.match(html,/never generated automatically/);
});
