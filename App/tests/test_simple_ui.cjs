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
  replaceChildren(...nodes) { this.children = nodes; }
  setAttribute(key, value) { this.attributes[key] = value; }
  querySelectorAll() { return []; }
  addEventListener(event, fn) { (this.listeners[event] ||= []).push(fn); }
  dispatch(event, detail = {}) { for (const fn of this.listeners[event] || []) fn({target: this, currentTarget: this, ...detail}); }
  removeAttribute(key) { delete this.attributes[key]; if (key === 'src' || key === 'href') delete this[key]; }
  toggleAttribute(key, value) { if (value) this.attributes[key] = ''; else delete this.attributes[key]; }
  closest() { return null; }
  getBoundingClientRect() { return {left:10,right:100,top:10,bottom:100}; }
  focus() { this.focused = true; }
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
  const source = fs.readFileSync(path.join(__dirname, '../web/app.js'), 'utf8').replace(/\ninit\(\);\s*$/, '');
  class FormData {
    constructor(form) { this.values = Object.fromEntries(Object.entries(form.values || {}).filter(([name]) => !form.elements?.namedItem(name)?.disabled)); }
    entries() { return Object.entries(this.values); }
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
  const ui = vm.runInContext(`${source}\n({app, stateLabel, statusClass, nextStep, matches, makeActions, makeClientCell, makeDetailRow, renderFolderInventory, simpleWarning, readableNumber, detailsBody, materialWeightTotal, updateMaterialWeightTotal, renderMaterialRows, visiblePage, pastCertificateNumbers, certificateFileTitle, showSavedFile, renderPastNumbers, api, fieldDifferences, renderChangeSummary, editorDirty, saveWorkingDraft, saveDraft, readWorkingDraft, clearWorkingDraft, offerDraftRecovery, restoreWorkingDraft, requestDetailsClose, finishDraftClose, certificateUrl, showGenerationSuccess, submitGeneration, bulkGenerationPayload, openBulkGeneration, startBulkGeneration, renderEditorSummary, setEditorMode, validateEditor, finishEditing, openDetails, saveDetails, showWarnings, confirmWarnings, pickAnnexure, previewDetails, toggleFullPdfPreview, backupAction, openBackupRestore, restoreBackup, historyDifferences, loadHistory, showHistoryDetail, invalidatePreview, confirmAction, settleActionConfirmation, handleAction, wire, stopApp, summarize})`, context);
  Object.assign(ui, vm.runInContext('({savedFilePlant, savedFileUrl, savedFileMatches, setWorkspaceView, resetRecordFilters, nextWorkbookId, suggestWorkbookId, printedWeight, updatePrintedWeight, schedulePreview, saveSettings, settingsValues, openSettings, requestSettingsClose, returnFromWarnings, toast, makeStatus, reconnectDashboard, saveExcelChoice, renderTemplateEditor})',context));
  Object.assign(ui, vm.runInContext('({startInlineEdit, finishInlineEdit, syncWorkbookChanges, setState, renderWorkbookIssues, refreshState, pollReminderReview})',context));
  Object.assign(ui, vm.runInContext('({openAddClient, renderClientMaterials, clientMaterialValues, saveClient, updateClientType})',context));
  Object.assign(ui, vm.runInContext('({excelTarget, openClientInExcel})',context));
  return {ui, controls, storage, requests, window, context, closeButtons, timers, settingsInputs};
}

function editing(h) {
  const form = h.controls.get('#details-form');
  form.values = {plant:'plant1',source_kind:'workbook',edit_scope:'certificate_only',certificate_type:'standard',supplier:'Acme',weight:'100.2',case_id:'1001',use_certificate_weight:false};
  const inputs = {};
  form.elements = {namedItem: name => inputs[name] ||= {type:name === 'use_certificate_weight' ? 'checkbox' : 'text', focus(){this.focused = true;}, get value() {return form.values[name] || '';}, set value(value) {form.values[name] = value;}, get checked() {return Boolean(form.values[name]);}, set checked(value) {form.values[name] = Boolean(value);}}};
  form.reset = () => {};
  form.reportValidity = () => true;
  form.checkValidity = () => true;
  h.ui.app.editor = {template:'standard',fields:{standard:{}},annexures:{},touched:new Set(),rawNumeric:{weight:'100.19999999999999'},sourceSn:'1001',sourcePlant:'plant1',workspaceRevision:'revision-1'};
  h.ui.app.templatesLoaded = true;
  h.ui.app.templates = [{id:'standard',fields:[]}];
  h.controls.get('#details-modal').open = true;
  h.ui.app.editor.baseline = h.ui.detailsBody(form);
  return form;
}

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
  assert.match(text(actions),/View details.*Open in Excel/);
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
  const {ui}=harness();
  assert.match(ui.simpleWarning('plant1.xlsx is unavailable: Cannot update plant1.xlsx. Save and close it in Excel, then try again. (Windows error 32)'),
    /Close plant1.xlsx in Excel.*Save and close.*Refresh/);
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

test('Duplicate issue cards keep the COD change, reference and next steps visible without technical filenames', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({preview_url:`/previews/${'a'.repeat(32)}.pdf`,
    certificate_no:'0393', warnings:duplicateIssueExamples})})}); editing(h);
  await h.ui.previewDetails();
  const list = h.controls.get('#preview-warnings');
  const cards = all(list,node => node.className === 'issue-card'); assert.equal(cards.length,2);
  const visibleCopy = card => card.children.filter(node => node.tag !== 'details').map(text).join(' ');
  assert.match(visibleCopy(cards[0]),/COD 0365 is already used.*COD 0393.*Existing PDFs stay unchanged.*What to do:.*inward already has a certificate/);
  assert.match(visibleCopy(cards[1]),/Reference already used.*DUMMY-REF-1002.*reference unchanged.*What to do:.*same inward.*cancel/);
  for (let index = 0; index < cards.length; index++) {
    assert.doesNotMatch(visibleCopy(cards[index]), /\.pdf|PDF history|automatic COD/);
    const details = cards[index].children.find(node => node.tag === 'details');
    assert.notEqual(details.open,true); assert.ok(text(details).includes(duplicateIssueExamples[index]));
  }
  assert.equal(h.controls.get('#preview-warnings-title').textContent,'2 things to check before creating');
  assert.equal(h.requests.length,1); assert.equal(h.requests[0].endpoint,'/api/preview');
  h.ui.invalidatePreview(); assert.equal(h.controls.get('#preview-warnings-title').hidden,true);
  assert.equal(list.children.length,0);
});

test('Warning review uses the same short duplicate messages and preserves the original confirmation payload', async () => {
  const h = harness(); const form = editing(h); const payload = h.ui.detailsBody(form);
  h.ui.showWarnings([{plant:'plant1',supplier:'Acme',warnings:duplicateIssueExamples}],
    {endpoint:'/api/create',payload,fromDetails:true,revision:h.ui.app.previewRevision,reviewToken:'original-review'});
  assert.equal(all(h.controls.get('#warning-list'),node => node.className === 'issue-card').length,2);
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
  assert.equal(note.textContent,'COD 0365 is already used · 1 more issue');
  assert.match(note.title,/0393.*DUMMY-REF-1002/); assert.doesNotMatch(note.textContent,/\.pdf|…/);
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
  for (const [state, label] of [['ready','Ready'],['alert','Needs attention'],['generated','Generated'],['generated_manual','Generated Manual'],['completed','Completed'],['blank',''],['na','N/A']]) {
    assert.equal(ui.stateLabel({certificate_status:state,overdue:true}), label);
    const expectedClass = state === 'alert' ? 'status-needs_attention' : `status-${state}`;
    assert.equal(ui.statusClass({certificate_status:state,overdue:true}), expectedClass);
  }
  for (const [legacy, canonical] of [['Created','Generated'],['Sent','Completed'],['Check Details','Needs attention'],['Working Draft','Needs attention']]) {
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
  assert.equal(buttons[0].textContent, 'Review & create');
  assert.equal(buttons[0].dataset.action, 'details');
  assert.equal(buttons.length, 1, 'Review opens the preview directly without a Details expansion.');
});
test('Explicit Edit details actions open editing, while Review & create stays preview-first', async () => {
  const h = harness();
  const row = {certificate_status:'alert',supplier:'Acme',plant:'plant1',sn:'1001'};
  h.ui.app.state.rows = [row];
  const rowAction = all(h.ui.makeActions(row,0), node => node.tag === 'button')[0];
  const detailAction = all(h.ui.makeDetailRow(row,0), node => node.tag === 'button').find(node => node.textContent === 'Edit details');
  assert.equal(rowAction.textContent,'Review details'); assert.equal(rowAction.dataset.action,'edit-details');
  assert.equal(detailAction.dataset.action,'edit-details');
  vm.runInContext('globalThis.editorRequests = []; openDetails = (...args) => editorRequests.push(args);',h.context);
  await h.ui.handleAction(rowAction);
  await h.ui.handleAction({dataset:{index:'0',action:'details'}});
  assert.deepEqual(JSON.parse(JSON.stringify(h.context.editorRequests.map(args => args[3] || false))),[true,false]);
});
test('Generated Manual has its own filter and can be marked Completed', () => {
  const {ui,controls} = harness(); controls.get('#status-filter').value = 'generated_manual';
  const row = {certificate_status:'generated_manual',supplier:'Acme',pdf_filename:'0001_Acme.pdf'};
  assert.equal(ui.matches(row), true);
  const action = ui.makeActions(row,4);
  assert.ok(all(action, node => node.tag === 'button').some(node => node.textContent === 'Mark Completed'));
});
test('Full logs are retained only inside a closed technical disclosure', () => {
  const {ui} = harness(); const raw = 'Created very_long_filename.pdf; Workbook backup: old.xlsx';
  const detail = ui.makeDetailRow({certificate_status:'generated',detail:raw,display_weight:1.7999999999999998},0);
  assert.equal(detail.children[0].colSpan, 5);
  const technical = all(detail,node => node.tag === 'details').find(node => node.children[0].textContent === 'Technical details');
  assert.ok(technical); assert.notEqual(technical.open,true); assert.match(text(technical), /very_long_filename/);
  assert.match(text(detail), /1.8 kg/);
  assert.ok(all(detail,node => node.className === 'next-step').every(node => !text(node).includes('backup')));
});
test('Unknown warnings retain the full consequence instead of being hidden', () => {
  const {ui} = harness(); const unknown = 'COD 0324 is used. A new number will be assigned; existing PDFs will not be overwritten.';
  assert.equal(ui.simpleWarning(unknown), unknown);
  assert.match(ui.simpleWarning('missing Form 6 number / manifest number'), /blank on the PDF/);
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
test('Saved certificates share one list with clear folder sources and preserve legacy PDFs', () => {
  const {ui,controls} = harness();
  const entries = [
    {path:'Certificates/Created in App/0101_Client_1.pdf'},
    {relative_path:'Manually Added\\0102_Client_2.pdf'},
    {path:'Certificates/Auto Generated/0103_Client_1.pdf'},
    {path:'Certificates/0104_Client_1.pdf'},
    {path:'Certificates/0105_Client_1.pdf',managed_owner:{source_kind:'manual'}},
  ].map(entry => ({...entry,present:true}));
  ui.app.state.certificate_inventory = {entries};
  ui.renderFolderInventory();
  const labels = all(controls.get('#folder-inventory-files'), node => node.className === 'saved-file-source').map(node => node.textContent);
  assert.deepEqual(labels.sort(), ['Auto Generated','Created in App','Created in App','Existing PDF','Manually Added'].sort());
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
  assert.equal(controls.get('#material-weight-total').textContent,'Total inward weight: 15 kg');
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
  assert.equal(controls.get('#count-generated-manual').textContent,1);
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
  assert.equal((html.match(/name="certificate_weight"/g)||[]).length,1);
  assert.equal((html.match(/name="cod_no"/g)||[]).length,1);
  assert.equal((html.match(/<th>/g)||[]).length,5);
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

test('Every tracked mutation carries a revision; explicit editor revision stays frozen', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({workspace_revision:'revision-2',state:{rows:[]}})})}); h.ui.app.workspaceRevision = 'revision-1';
  for (const endpoint of ['/api/create','/api/generate','/api/manual/save','/api/manual/generate','/api/manual/complete','/api/complete','/api/backup/restore']) {
    await h.ui.api(endpoint, {});
    assert.equal(h.requests.at(-1).payload.expected_revision, endpoint === '/api/create' ? 'revision-1' : 'revision-2');
  }
  await h.ui.api('/api/create', {expected_revision:'editor-snapshot'});
  assert.equal(h.requests.at(-1).payload.expected_revision,'editor-snapshot');
  await h.ui.api('/api/state'); assert.equal(h.requests.at(-1).method,'GET');
  await h.ui.api('/api/backup/create', {}); assert.equal(h.requests.at(-1).payload.expected_revision,undefined);
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

test('Discarding recovery prompted by New, Review, Edit or Copy resumes exactly the requested editor', async () => {
  for (const action of ['new','details','edit-details','copy']) {
    const h = harness(); const form = editing(h); form.values.supplier = 'Previous draft';
    h.ui.saveWorkingDraft(); h.controls.get('#details-modal').close(); h.ui.wire();
    const row = {supplier:'Selected company',plant:'plant2',sn:'2001'}; h.ui.app.state.rows = [row];
    if (action === 'new') h.controls.get('#new-button').dispatch('click');
    else await h.ui.handleAction({dataset:{index:'0',action}});
    assert.equal(h.ui.app.draftRecovery.resumeEditor,true);
    vm.runInContext('globalThis.editorRequests = []; openDetails = (...args) => editorRequests.push(args);',h.context);
    h.controls.get('#discard-working-draft').dispatch('click');
    assert.equal(h.context.editorRequests.length,1);
    const [selected, copyAsNew, skipRecovery] = h.context.editorRequests[0];
    assert.equal(selected,action === 'new' ? null : row); assert.equal(copyAsNew,action === 'copy'); assert.equal(skipRecovery,true);
    assert.equal(h.context.editorRequests[0][3] || false,action === 'edit-details');
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
  assert.match(h.controls.get('#toast-message').textContent,/Refresh.*reopen/);
  assert.equal(h.controls.get('#details-error').hidden,false);
});

test('Success shows actual server-issued certificates and safely encoded local PDF links', () => {
  const h = harness();
  h.ui.showGenerationSuccess({created:2,certificates:[{filename:'0082_Acme_Long_Recovery_Services_Mumbai_India_1.pdf',certificate_no:'0082',supplier:'Acme',plant:'plant1'},
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
  for (const certificate of [{filename:'../secret.pdf'},{filename:'a.pdf',pdf_url:'https://example.com/api/certificates/a.pdf'},
    {filename:'a.pdf',pdf_url:'/api/preview/a.pdf'},{filename:'a.pdf',pdf_url:'/api/certificates/b.pdf'}]) assert.equal(h.ui.certificateUrl(certificate),null);
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
  const localized = new Date(timestamp).toLocaleString(undefined,{day:'2-digit',month:'short',year:'numeric',hour:'numeric',minute:'2-digit'});
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
  assert.match(h.controls.get('#preview-status').textContent,/Preview is up to date.*company and weight/);
  assert.match(h.controls.get('#preview-number').textContent,/Certificate number: 0091.*not reserved/);
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

test('Home uses one status filter bar with counts beside the company list', () => {
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  assert.match(html,/<div class="quick-views" id="status-filters" aria-label="Status filters">/);
  assert.equal((html.match(/class="quick-views"/g)||[]).length,1);
  assert.doesNotMatch(html,/summary-panel|stat-card|attention-strip|review-alerts/);
  assert.match(html,/<div[^>]*id="more-filters"/);
  assert.match(html,/<h3 class="filter-section-title">More filters<\/h3>/);
  assert.doesNotMatch(html,/<details[^>]*id="more-filters"|id="filters-summary"/);
  assert.doesNotMatch(html,/id="create-another"|>Create another<\/button>/);
  const toolbar = html.slice(html.indexOf('<div class="toolbar">'),html.indexOf('<div class="quick-views"'));
  assert.match(toolbar,/id="search"/); assert.match(toolbar,/id="plant-filter"/);
  assert.doesNotMatch(toolbar,/id="status-filter"|id="template-filter"|id="generate-all"/);
  const bulk = html.slice(html.indexOf('id="bulk-generation-modal"'),html.indexOf('id="saved-file-modal"'));
  assert.match(bulk,/Search, status and certificate-type filters do not narrow/);
  assert.match(bulk,/Empty statuses and N\/A rows are skipped/);
});

test('Certificate footer has only Cancel and Create PDF; secondary saving stays accessible', () => {
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  const editor = html.slice(html.indexOf('id="details-modal"'),html.indexOf('id="excel-save-modal"'));
  const footer = editor.slice(editor.lastIndexOf('<div class="modal-footer">'));
  assert.equal((footer.match(/<button\b/g) || []).length,2);
  assert.match(footer,/>Cancel<\/button>/); assert.match(footer,/id="save-details">Create PDF/);
  assert.match(editor,/<details[^>]*id="draft-options"[\s\S]*?id="keep-working-draft"[\s\S]*?id="save-draft"/);
  assert.match(editor,/<form[^>]*novalidate/);
});

test('Other status filters are visible buttons without repeating All, Ready or Alert', () => {
  const html = fs.readFileSync(path.join(__dirname,'../web/index.html'),'utf8');
  const statusChoices = [...html.matchAll(/data-filter-status="([^"]+)"/g)].map(match => match[1]);
  assert.deepEqual(statusChoices,['generated','generated_manual','completed','blank','na']);
  assert.match(html,/<input type="hidden" id="status-filter" value="all">/);
  assert.doesNotMatch(html,/<select[^>]*id="status-filter"/);
  assert.equal((html.match(/data-quick-status="all"/g)||[]).length,1);
  assert.equal((html.match(/data-quick-status="ready"/g)||[]).length,1);
  assert.equal((html.match(/data-quick-status="needs_attention"/g)||[]).length,1);
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
  for (const value of ['REF-1010','F6-1010','0362','2026','Inward ID / SN','1001']) assert.ok(facts.includes(value));
  const materials = h.controls.get('#editor-materials');
  assert.match(text(materials), /Laptops.*Quantity: 2.*3.5 kg per piece.*7 kg total/);
  assert.match(text(materials), /Monitors.*Quantity: 3.*2 kg per piece.*6 kg total/);
  assert.equal(all(materials, node => ['input','button','details'].includes(node.tag)).length, 0);
  assert.equal(JSON.stringify(h.ui.detailsBody(form)), before, 'Summary rendering cannot change source or submit resolved preview values.');
  for (const name of ['reference','cod_no','cod_date']) assert.equal(form.elements.namedItem(name).readOnly, true);
  form.values.certificate_type = 'sita'; h.ui.renderEditorSummary();
  assert.doesNotMatch(text(h.controls.get('#editor-facts')), /Form 6|F6-1010/);
});

test('Preview facts use the actual rendered COD, date and printed weight and clear when invalidated', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({preview_url:`/previews/${'a'.repeat(32)}.pdf`,
    certificate_no:'0042',certificate_date:'2026-09-30',printed_details:{certificate_weight:'12 KG',company_display:'ACME, MUMBAI'}})})});
  const form = editing(h); form.values.cod_no = '0001'; form.values.cod_date = '';
  h.ui.app.editor.firstGeneration = true;
  h.ui.renderMaterialRows([{item:'Laptops',weight:'100.2'}]);
  h.ui.renderEditorSummary(); assert.match(text(h.controls.get('#editor-facts')), /Set automatically/);
  const before = JSON.stringify(h.ui.detailsBody(form));
  await h.ui.previewDetails();
  const facts = text(h.controls.get('#editor-facts'));
  assert.match(facts, /0042/); assert.match(facts, /2026/); assert.match(facts, /12 KG · material total/);
  assert.doesNotMatch(facts, /0001|Set automatically/);
  assert.match(h.controls.get('#preview-number').textContent, /COD date:.*2026.*not reserved/);
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
  const h = harness(); const form = editing(h); h.ui.setEditorMode(false);
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
    h.ui.app.state.rows = [{plant:'plant1',certificate_status:'ready'}, {plant:'plant1',certificate_status:'ready',warnings:['Missing name']},
      {plant:'plant2',certificate_status:'ready'}, {plant:'plant2',certificate_status:'blank'}, {plant:'plant1',certificate_status:'na'},
      {plant:'plant1',certificate_status:'ready',source_kind:'manual'}];
    h.controls.get('#plant-filter').value = plant; h.controls.get('#search').value = 'Unrelated search';
    h.controls.get('#status-filter').value = 'completed'; h.controls.get('#template-filter').value = 'siemens';
    h.ui.app.workspaceRevision = 'reviewed-batch'; h.ui.openBulkGeneration();
    assert.equal(h.controls.get('#bulk-description').textContent,`${count} certificate${count === 1 ? '' : 's'} from ${description}.${plant === 'plant2' ? '' : ' 1 needs warning review.'}`);
    assert.equal(h.controls.get('#bulk-generation-modal').open,true);
    assert.equal(h.requests.length,0); assert.equal(h.ui.app.bulkSelection.expected_revision,'reviewed-batch');
  }
});

test('Batch cancellation is harmless and creation retains the reviewed plant and revision', async () => {
  const h = harness(); h.ui.wire();
  h.ui.app.state.rows = [{plant:'plant2',certificate_status:'ready'}];
  h.controls.get('#plant-filter').value = 'plant2'; h.ui.app.workspaceRevision = 'reviewed-batch';
  h.ui.openBulkGeneration(); h.controls.get('#bulk-generation-modal').close();
  await h.ui.startBulkGeneration(); assert.equal(h.requests.length,0);
  h.ui.openBulkGeneration(); h.controls.get('#plant-filter').value = 'plant1'; h.ui.app.workspaceRevision = 'changed-batch';
  vm.runInContext('submitGeneration = async payload => { app.submittedBatch = payload; };',h.context);
  await h.ui.startBulkGeneration();
  assert.deepEqual(JSON.parse(JSON.stringify(h.ui.app.submittedBatch)),{plants:['plant2'],expected_revision:'reviewed-batch'});
});

test('Material entry shows name and weight first while quantity fields retain their values', () => {
  const h = harness(); h.ui.renderMaterialRows([{item:'Laptops',weight:'21',quantity:'3',per_piece_weight:'7'}]);
  const row = h.ui.app.materialRows[0];
  assert.equal(row.element.children[1].children.length,2);
  assert.match(text(row.element.children[1]),/Material or asset.*Weight \(kg\)/);
  const calculation = row.element.children[2]; assert.equal(calculation.tag,'details'); assert.equal(calculation.open,false);
  assert.equal(row.inputs.quantity.value,'3'); assert.equal(row.inputs.per_piece_weight.value,'7');
  assert.equal(h.ui.materialWeightTotal().total,21);
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

test('Alert shortcut count respects plant selection without making Blank or N/A eligible for creation', () => {
  const h = harness(); h.controls.get('#plant-filter').value = 'plant2';
  h.ui.app.state.rows = [{plant:'plant1',certificate_status:'alert'}, {plant:'plant2',certificate_status:'alert'},
    {plant:'plant2',certificate_status:'blank'}, {plant:'plant2',certificate_status:'na'}];
  h.ui.summarize(); assert.equal(h.controls.get('#count-attention').textContent,2);
  assert.equal(h.controls.get('#count-all').textContent,3);
  assert.equal(h.controls.get('#open-bulk').disabled,true); assert.equal(h.controls.get('#generate-all').disabled,true);
});

test('First issuance opens a preview with one edit control; PDF corrections leave Excel untouched', async () => {
  for (const kind of ['existing','new','copy']) {
    const h = harness(); const form = editing(h); h.controls.get('#details-modal').open = false;
    vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {}; globalThis.previews = []; previewDetails = async () => previews.push(detailsBody(document.querySelector("#details-form")));',h.context);
    const row = {plant:'plant2',sn:'2001',supplier:'Northstar',weight:12,inward_date:'2026-10-01',items:[{item:'Laptops',weight:12}]};
    await h.ui.openDetails(kind === 'new' ? null : row,kind === 'copy');
    assert.equal(h.controls.get('#editor-fields').hidden,false);
    assert.equal(h.controls.get('#excel-detail-fields').hidden,kind === 'existing');
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
      assert.equal(h.ui.detailsBody(h.controls.get('#details-form')).pdf_only_edits,true);
      assert.equal(h.controls.get('#material-weight-total').textContent,'Total inward weight: 12 kg');
    }
  }
});

test('An explicit Edit details action starts an existing record in the form', async () => {
  const h = harness(); editing(h); h.controls.get('#details-modal').open = false;
  vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {}; globalThis.previews = []; previewDetails = async () => previews.push(detailsBody(document.querySelector("#details-form")));',h.context);
  const row = {plant:'plant2',sn:'2001',supplier:'Northstar',has_issued_certificate:true,weight:12,inward_date:'2026-10-01',items:[{item:'Laptops',weight:12}]};
  await h.ui.openDetails(row,false,false,true);
  assert.equal(h.controls.get('#editor-fields').hidden,false);
  assert.equal(h.controls.get('#editor-summary').hidden,true);
  assert.match(h.controls.get('#details-title').textContent,/Edit certificate for Northstar/);
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
  h.ui.app.state.rows = [{plant:'plant2',certificate_status:'ready'}];
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

test('Clearing record search and filters retains the plant and does not broaden batch generation', () => {
  const h = harness(); h.controls.get('#plant-filter').value = 'plant2';
  h.controls.get('#search').value = 'old'; h.controls.get('#status-filter').value = 'generated';
  h.controls.get('#template-filter').value = 'sita'; h.ui.resetRecordFilters();
  assert.equal(h.controls.get('#plant-filter').value,'plant2');
  assert.equal(h.controls.get('#search').value,''); assert.equal(h.controls.get('#status-filter').value,'all');
  assert.deepEqual(JSON.parse(JSON.stringify(h.ui.bulkGenerationPayload())),{plants:['plant2']});
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
  const row = {plant:'plant2',certificate_status:'ready',reminder_due:true,age_days:25};
  assert.equal(h.ui.matches(row),true); assert.equal(h.ui.stateLabel(row),'Ready');
  assert.match(text(h.ui.makeStatus(row)),/Certificate pending · received 25 days ago/);
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
  assert.equal(correction[0].textContent,'Review details');
  assert.ok(correction.some(node => node.textContent === 'Open in Excel'));
  const readyWithIssue = {...future,certificate_status:'ready',warnings:['inward date is in the future']};
  const readyActions = all(h.ui.makeActions(readyWithIssue,0), node => node.tag === 'button');
  assert.equal(readyActions[0].textContent,'Review details');
  assert.equal(readyActions[0].dataset.action,'details','First generation still opens its safe preview.');
  assert.ok(readyActions.some(node => node.textContent === 'Open in Excel'));
  const issued = {certificate_status:'generated',reminder_due:true,age_days:50,pdf_filename:'0001_Acme.pdf'};
  assert.match(text(h.ui.makeStatus(issued)),/PDF ready.*mark Completed/);
  const missing = {...issued,pdf_filename:''};
  assert.match(text(h.ui.makeStatus(missing)),/Saved PDF is missing/);
  assert.equal(all(h.ui.makeActions(missing,0), node => node.dataset.action === 'complete').length,0);
  assert.equal(all(h.ui.makeActions(missing,0), node => node.dataset.action === 'edit-details').length,1);
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
  h.ui.app.state.rows = [{plant:'plant1',sn:'1001',certificate_status:'na'}];
  await h.ui.saveDetails({preventDefault(){},currentTarget:form});
  assert.equal(h.requests.length,0); assert.equal(h.controls.get('#record-options').open,true);
  assert.match(h.controls.get('#details-error').textContent,/already exists/);
  assert.equal(form.elements.namedItem('case_id').focused,true);
});

test('Moving an existing custom certificate to Excel is create-only, never an unguarded upsert', async () => {
  const h = harness(); const form = editing(h);
  h.ui.app.editor.existing = true; h.ui.app.editor.recordId = 'custom-1'; h.ui.app.editor.sourceSn = null;
  h.ui.app.state.rows = [{plant:'plant1',sn:'1001',certificate_status:'na'}];
  assert.equal(h.ui.detailsBody(form).create_new,true);
  await h.ui.saveDetails({preventDefault(){},currentTarget:form});
  assert.equal(h.requests.length,0); assert.match(h.controls.get('#details-error').textContent,/already exists/);
  h.ui.app.editor.touched.clear(); form.values.case_id = ''; h.ui.suggestWorkbookId();
  assert.equal(form.values.case_id,'1002');
});

test('Attachment refresh binds new rows without blessing old cached data and then previews automatically', async () => {
  const previewUrl = `/previews/${'a'.repeat(32)}.pdf`;
  const h = harness({fetch:async endpoint => ({ok:true,json:async () => endpoint === '/api/pick-annexure'
    ? {workspace_revision:'attachment-only',annexure:{id:'a',name:'List.pdf'}}
    : endpoint === '/api/check' ? {workspace_revision:'fresh-rows',state:{rows:[{plant:'plant1',sn:'1001',supplier:'Outside edit'}]}}
    : {preview_url:previewUrl,certificate_no:'0002',warnings:[]}})});
  editing(h); h.ui.renderMaterialRows([{item:'Laptops',weight:10}]);
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

test('Issued certificates expose Edit certificate directly with no Details expansion', () => {
  const h = harness();
  const row = {plant:'plant1',sn:'1001',certificate_status:'generated',pdf_filename:'0001_Acme_Mumbai_1.pdf'};
  const buttons = all(h.ui.makeActions(row,0), node => node.tag === 'button');
  assert.ok(buttons.some(button => button.textContent === 'Edit certificate' && button.dataset.action === 'edit-details'));
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
  for (const key of ['supplier','reference','cod_no','cod_date','certificate_weight','materials']) {
    h.ui.startInlineEdit(key,key); assert.equal(h.ui.app.editor.inlineEdit.key,key);
    assert.equal(JSON.stringify(h.ui.detailsBody(form)),original); h.ui.finishInlineEdit();
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
  const secondWeight = inputs.find(input => input.attributes['aria-label'] === 'Weight (kg) for material 2');
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

test('Reopening a PDF-only correction restores last issued details and its weight override', async () => {
  const h = harness(); editing(h); h.controls.get('#details-modal').open = false;
  vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {}; previewDetails = async () => {};',h.context);
  const row = {plant:'plant1',sn:'1001',supplier:'Excel company',location:'Pune',has_issued_certificate:true,
    items:[{item:'Asset',quantity:1,per_piece_weight:10,weight:10}],certificate_weight:8,
    last_certificate_details:{supplier:'PDF correction',location:'Mumbai',certificate_type:'siemens',certificate_weight:'7.25',
      certificate_fields:{model_name:'Corrected scanner'},cod_no:'0041',cod_date:'2026-10-01',
      items:[{item:'Asset',quantity:'1',per_piece_weight:'10',weight:'10'}]}};
  await h.ui.openDetails(row);
  const form = h.controls.get('#details-form');
  assert.equal(form.values.supplier,'PDF correction'); assert.equal(form.values.location,'Mumbai');
  assert.equal(form.values.cod_no,'0041'); assert.equal(form.values.use_certificate_weight,true);
  assert.equal(h.ui.app.editor.excelCertificateWeight,8);
  assert.equal(h.ui.app.editor.firstGeneration,false);
  assert.equal(h.controls.get('#excel-detail-fields').hidden,true);
});

test('Corrections wait for Yes or No before submitting and preview refresh does not invalidate the choice', async () => {
  for (const yes of [false,true]) {
    const h = harness(), form = editing(h);
    Object.assign(h.ui.app.editor,{existing:true,firstGeneration:false});
    form.values.supplier = 'Corrected company';
    vm.runInContext('submitGeneration = async payload => {app.correctionPayload = payload;};',h.context);
    await h.ui.saveDetails({preventDefault(){},currentTarget:form});
    assert.equal(h.controls.get('#excel-save-modal').open,true);
    assert.equal(h.ui.app.correctionPayload,undefined);
    h.ui.app.previewRevision++; // Same data, newer raster preview.
    await h.ui.saveExcelChoice(yes);
    assert.equal(h.ui.app.correctionPayload.edit_scope,yes ? 'update_excel' : 'certificate_only');
    assert.equal(h.ui.app.correctionPayload.details.supplier,'Corrected company');
    assert.equal(h.controls.get('#excel-save-modal').open,false);
  }
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
    location:'Pune',inward_date:'2026-10-01',items:[{item:'Asset',quantity:1,per_piece_weight:10,weight:10}]}];
  vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {};',h.context);
  await h.ui.restoreWorkingDraft();
  assert.equal(form.values.supplier,'Excel company');
  assert.equal(h.ui.app.editor.firstGeneration,true);
  assert.equal(h.controls.get('#excel-detail-fields').hidden,false);
  assert.equal(h.ui.app.editor.fields.standard.company_display,undefined);
  assert.equal(h.ui.app.editor.fields.standard.weight_display,undefined);
  assert.equal(h.requests.length,0);
});

function clientEntry(h) {
  const form = h.controls.get('#add-client-form');
  const defaults = {plant:'plant1',certificate_type:'standard',supplier:'',location:'',inward_date:'',form6:'',reference:''};
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
  form.values.certificate_type = 'sita'; h.ui.updateClientType();
  assert.equal(inputs.form6.closest().hidden,true);
  form.values.certificate_type = 'standard'; h.ui.updateClientType();
  assert.equal(inputs.form6.closest().hidden,false);
});

test('Add client saves every material once and shows the new Excel entry in the list', async () => {
  let resolveSave;
  const h = harness({fetch:() => new Promise(resolve => { resolveSave = resolve; })});
  const {form} = clientEntry(h); h.ui.openAddClient();
  form.values.supplier = 'New company';
  h.ui.renderClientMaterials([{item:'Laptop',quantity:'2',per_piece_weight:'3.5'},{item:'Printer',weight:'8'}]);
  const pending = h.ui.saveClient();
  await h.ui.saveClient();
  assert.equal(h.requests.length,1);
  assert.equal(h.requests[0].endpoint,'/api/clients/add');
  assert.equal(h.requests[0].payload.expected_revision,'revision-1');
  assert.equal(h.requests[0].payload.details.items.length,2);
  assert.equal(h.requests[0].payload.details.items[1].weight,'8');
  resolveSave({ok:true,json:async () => ({client_added:true,client:{sn:'1002',plant:'plant1',supplier:'New company',status:'ready'},state:{rows:[],scan_complete:true},workspace_revision:'revision-2'})});
  await pending;
  assert.equal(h.controls.get('#add-client-modal').open,false);
  assert.equal(h.controls.get('#search').value,'1002');
  assert.equal(h.controls.get('#status-filter').value,'all');
  assert.equal(h.ui.app.workspaceRevision,'revision-2');
  assert.equal(h.ui.app.clientRequestId,null);
  assert.match(h.controls.get('#toast-message').textContent,/Ready to create/);
});

test('Failed Add client retains the form and stable request ID for a safe retry', async () => {
  const h = harness({fetch:async () => ({ok:false,status:400,json:async () => ({error:'Save and close Plant 1.xlsx in Excel, then try again.'})})});
  const {form} = clientEntry(h); h.ui.openAddClient(); form.values.supplier = 'New company';
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
