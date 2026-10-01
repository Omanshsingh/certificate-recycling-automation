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
  append(...nodes) { this.children.push(...nodes); }
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
}
function all(node, predicate) {
  return [ ...(predicate(node) ? [node] : []), ...node.children.flatMap(child => all(child, predicate)) ];
}
function text(node) { return [node.textContent, ...node.children.map(text)].join(' '); }
function harness(options = {}) {
  const controls = new Map();
  for (const id of ['#plant-filter', '#status-filter', '#template-filter']) { const select = new Node('select'); select.value = 'all'; controls.set(id, select); }
  controls.set('#search', new Node('input'));
  for (const id of ['#folder-inventory', '#folder-inventory-summary', '#folder-inventory-note', '#folder-inventory-warnings', '#folder-inventory-files', '#saved-file-pages', '#saved-file-page-count', '#saved-files-previous', '#saved-files-next', '#show-past-numbers', '#past-numbers-modal', '#past-number-list', '#past-number-count', '#saved-file-name', '#saved-file-facts', '#saved-file-modal']) controls.set(id, new Node('div'));
  controls.set('#past-number-search', new Node('input'));
  controls.set('#template-editor', new Node('div'));
  const html = fs.readFileSync(path.join(__dirname, '../web/index.html'), 'utf8');
  for (const match of html.matchAll(/<([\w-]+)[^>]*\bid="([^"]+)"/g)) {
    if (!controls.has(`#${match[2]}`)) controls.set(`#${match[2]}`, new Node(match[1]));
    controls.get(`#${match[2]}`).id = match[2];
  }
  const closeButtons = [...html.matchAll(/<button[^>]*data-close="([^"]+)"[^>]*>/g)].map(match => { const button = new Node('button'); button.dataset.close = match[1]; return button; });
  const document = { createElement: tag => new Node(tag), querySelector: key => controls.get(key), querySelectorAll: selector => selector === 'dialog' ? [...controls.values()].filter(node => /-modal$/.test(node.id)) : selector === '[data-close]' ? closeButtons : [] };
  const source = fs.readFileSync(path.join(__dirname, '../web/app.js'), 'utf8').replace(/\ninit\(\);\s*$/, '');
  class FormData {
    constructor(form) { this.values = form.values; }
    entries() { return Object.entries(this.values); }
    get(key) { return this.values[key]; }
  }
  const storage = new Map(), requests = [], window = new Node('window');
  window.location = {href:'http://127.0.0.1:8000/',origin:'http://127.0.0.1:8000'};
  window.confirm = () => { throw new Error('Native confirmations are unavailable in the in-app browser.'); };
  const context = vm.createContext({document, console, FormData, URL, window, setTimeout: () => 1, clearTimeout() {},
    localStorage: {getItem:key => storage.get(key) || null, setItem:(key,value) => storage.set(key,value), removeItem:key => storage.delete(key)},
    fetch: async (endpoint, init) => {
      requests.push({endpoint, ...init, payload:init.body ? JSON.parse(init.body) : undefined});
      return options.fetch ? options.fetch(endpoint, init) : {ok:true,json:async () => ({workspace_revision:'revision-2'})};
    }});
  const ui = vm.runInContext(`${source}\n({app, stateLabel, statusClass, nextStep, matches, makeActions, makeClientCell, makeDetailRow, renderFolderInventory, simpleWarning, readableNumber, detailsBody, visiblePage, pastCertificateNumbers, certificateFileTitle, showSavedFile, renderPastNumbers, api, fieldDifferences, renderChangeSummary, editorDirty, saveWorkingDraft, readWorkingDraft, clearWorkingDraft, offerDraftRecovery, restoreWorkingDraft, requestDetailsClose, finishDraftClose, certificateUrl, showGenerationSuccess, submitGeneration, bulkGenerationPayload, showWarnings, confirmWarnings, pickAnnexure, previewDetails, toggleFullPdfPreview, backupAction, openBackupRestore, restoreBackup, historyDifferences, loadHistory, showHistoryDetail, invalidatePreview, confirmAction, settleActionConfirmation, handleAction, wire, stopApp})`, context);
  return {ui, controls, storage, requests, window, context, closeButtons};
}

function editing(h) {
  const form = h.controls.get('#details-form');
  form.values = {plant:'plant1',source_kind:'workbook',edit_scope:'certificate_only',certificate_type:'standard',supplier:'Acme',weight:'100.2',case_id:'1001'};
  const inputs = {};
  form.elements = {namedItem: name => inputs[name] ||= {get value() {return form.values[name] || '';}, set value(value) {form.values[name] = value;}}};
  form.reset = () => {};
  form.reportValidity = () => true;
  h.ui.app.editor = {template:'standard',fields:{standard:{}},annexures:{},touched:new Set(),rawNumeric:{weight:'100.19999999999999'},sourceSn:'1001',sourcePlant:'plant1',workspaceRevision:'revision-1'};
  h.ui.app.templatesLoaded = true;
  h.ui.app.templates = [{id:'standard',fields:[]}];
  h.controls.get('#details-modal').open = true;
  h.ui.app.editor.baseline = h.ui.detailsBody(form);
  return form;
}
test('Overdue does not conceal canonical certificate statuses', () => {
  const {ui} = harness();
  for (const [state, label] of [['ready','Ready'],['alert','Alert'],['generated','Generated'],['generated_manual','Generated Manual'],['completed','Completed'],['blank','Blank'],['na','N/A']]) {
    assert.equal(ui.stateLabel({certificate_status:state,overdue:true}), label);
    const expectedClass = ['alert','blank'].includes(state) ? 'status-needs_attention' : `status-${state}`;
    assert.equal(ui.statusClass({certificate_status:state,overdue:true}), expectedClass);
  }
  for (const [legacy, canonical] of [['Created','Generated'],['Sent','Completed'],['Check Details','Alert'],['Working Draft','Alert']]) {
    assert.equal(ui.stateLabel({certificate_status:legacy}), canonical);
  }
});
test('Ready card filter excludes warnings and keeps blank/N/A separate', () => {
  const {ui,controls} = harness(); controls.get('#status-filter').value = 'ready';
  assert.equal(ui.matches({certificate_status:'ready'}), true);
  for (const row of [{certificate_status:'ready',warnings:['Missing weight']},{certificate_status:'blank'},{certificate_status:'na'}]) assert.equal(ui.matches(row), false);
  assert.match(ui.nextStep({certificate_status:'blank'}), /Automatic creation skips/);
  assert.match(ui.nextStep({certificate_status:'na'}), /skipped/);
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
  assert.equal(buttons[1].attributes['aria-expanded'], 'false');
  assert.equal(buttons[1].dataset.expand, '4');
});
test('Generated Manual has its own filter and can be marked Completed', () => {
  const {ui,controls} = harness(); controls.get('#status-filter').value = 'generated_manual';
  const row = {certificate_status:'generated_manual',supplier:'Acme'};
  assert.equal(ui.matches(row), true);
  const action = ui.makeActions(row,4);
  assert.equal(all(action, node => node.tag === 'button')[0].textContent, 'Mark Completed');
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
test('Floating-point display artifacts are readable without changing source data', () => {
  const {ui} = harness();
  assert.equal(ui.readableNumber(100.19999999999999),'100.2');
  assert.equal(ui.readableNumber(0),'0');
  assert.equal(ui.readableNumber(null),'');
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
  const h = harness(); h.ui.app.workspaceRevision = 'revision-1';
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
  assert.match(h.controls.get('#change-scope').textContent,/PDF and Excel/);
  form.values.source_kind = 'manual'; h.ui.renderChangeSummary();
  assert.match(h.controls.get('#change-scope').textContent,/local certificate tracking.*No Excel row/);
});

test('Working drafts survive closing locally without any backend write', () => {
  const h = harness(); const form = editing(h); form.values.supplier = 'Draft company';
  assert.equal(h.ui.editorDirty(),true); assert.equal(h.ui.saveWorkingDraft(),true);
  const saved = h.ui.readWorkingDraft();
  assert.equal(saved.values.supplier,'Draft company'); assert.equal(saved.editor.workspaceRevision,'revision-1');
  assert.equal(saved.editor.baseline.details.supplier,'Acme');
  assert.equal(saved.values.plant,'plant1'); assert.equal(h.requests.length,0);
  assert.equal(h.ui.requestDetailsClose(),false); assert.equal(h.controls.get('#unsaved-modal').open,true);
  h.ui.finishDraftClose(true);
  assert.equal(h.controls.get('#details-modal').open,false); assert.equal(h.ui.readWorkingDraft().values.supplier,'Draft company');
  assert.equal(h.controls.get('#toast').textContent,'Working draft saved on this computer. Excel is unchanged.');
  assert.equal(h.controls.get('#toast').classList.contains('visible'),true);
  assert.equal(h.controls.get('#toast').classList.contains('error'),false);
  assert.equal(h.ui.offerDraftRecovery(),true); assert.equal(h.controls.get('#draft-recovery-modal').open,true);
  assert.equal(h.requests.length,0);
});

test('Draft restoration restores data and revision with zero network mutations', async () => {
  const h = harness(); const form = editing(h); form.values.supplier = 'Recovered company';
  h.ui.app.editor.touched.add('supplier'); h.ui.saveWorkingDraft();
  h.controls.get('#details-modal').close(); h.ui.offerDraftRecovery();
  vm.runInContext('updateSourceControls = () => {}; renderTemplateEditor = () => {}; invalidatePreview = () => {};',h.context);
  form.values.supplier = 'Other company'; h.ui.app.workspaceRevision = 'revision-999';
  await h.ui.restoreWorkingDraft();
  assert.equal(form.values.supplier,'Recovered company');
  assert.equal(h.ui.app.editor.workspaceRevision,'revision-1');
  assert.equal(h.ui.app.editor.touched.has('supplier'),true);
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

test('Discarding recovery prompted by New, Review or Copy resumes exactly the requested editor', async () => {
  for (const action of ['new','details','copy']) {
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
    assert.equal(h.ui.readWorkingDraft(),null); assert.equal(h.requests.length,0);
  }
});

test('Discard removes only the local draft; unavailable storage does not close edited form', () => {
  const h = harness(); const form = editing(h); form.values.supplier = 'Discard me';
  h.ui.saveWorkingDraft(); h.ui.requestDetailsClose(); h.ui.finishDraftClose(false);
  assert.equal(h.ui.readWorkingDraft(),null); assert.equal(h.requests.length,0);
  h.controls.get('#details-modal').open = true;
  vm.runInContext('localStorage.setItem = () => { throw new Error("quota"); };',h.context);
  h.ui.requestDetailsClose(); h.ui.finishDraftClose(true);
  assert.equal(h.controls.get('#details-modal').open,true);
  assert.match(h.controls.get('#working-draft-status').textContent,/could not save/);
  assert.equal(h.controls.get('#toast').classList.contains('visible'),false);
  assert.doesNotMatch(h.controls.get('#toast').textContent,/Working draft saved/);
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
  h.ui.finishDraftClose(true);
  assert.equal(h.requests.length,0); assert.equal(h.controls.get('#action-confirmation-modal').open,true);
  h.controls.get('#confirm-action').dispatch('click');
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(h.requests[0].endpoint,'/api/stop');
  assert.equal(h.ui.readWorkingDraft().values.supplier,'Unsaved');
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

test('Warning continuation uses the response revision, token and reason and keeps names/consequences', async () => {
  let attempt = 0;
  const h = harness({fetch:async () => ({ok:true,json:async () => ++attempt === 1
    ? {requires_confirmation:true,workspace_revision:'review-revision',review_token:'review-81',warnings:[{plant:'plant2',supplier:'Acme',warnings:['Replacement weight will be used. Verify before sending.']}]}
    : {workspace_revision:'generated-revision',created:1,certificates:[{filename:'actual.pdf',supplier:'Acme',certificate_no:'0082',plant:'plant2'}]}})});
  vm.runInContext('setBusy = () => {};',h.context);
  h.ui.app.workspaceRevision = 'old-revision';
  await h.ui.submitGeneration({sn:'1001'});
  assert.equal(h.requests[0].payload.expected_revision,'old-revision');
  assert.match(text(h.controls.get('#warning-list')),/Plant 2.*Acme.*Replacement weight will be used.*Verify before sending/);
  await h.ui.confirmWarnings(); assert.equal(h.requests.length,1);
  assert.equal(h.controls.get('#warning-reason-error').hidden,false);
  h.controls.get('#override-reason').value = 'Customer confirmed the net weight';
  await h.ui.confirmWarnings();
  const payload = h.requests[1].payload;
  assert.equal(payload.expected_revision,'review-revision'); assert.equal(payload.review_token,'review-81');
  assert.equal(payload.override_reason,'Customer confirmed the net weight'); assert.equal(payload.confirmed,true);
  assert.equal(h.controls.get('#success-modal').open,true);
});

test('Empty warning confirmations keep reason optional and supply a backend-compatible default', async () => {
  const h = harness(); vm.runInContext('setBusy = () => {};',h.context);
  h.ui.showWarnings([],{endpoint:'/api/generate',payload:{expected_revision:'revision-1'},reviewToken:'no-warnings',fromDetails:false});
  assert.equal(h.controls.get('#override-reason').required,false);
  await h.ui.confirmWarnings();
  assert.equal(h.requests.length,1); assert.equal(h.requests[0].payload.override_reason,'No warnings were reported; confirmed creation.');
});

test('Stale workspace errors keep the edited form and recoverable draft open', async () => {
  const h = harness({fetch:async () => ({ok:false,status:409,json:async () => ({error:'Workspace changed. Refresh and review.'})})});
  const form = editing(h); form.values.supplier = 'Keep me'; h.ui.saveWorkingDraft();
  vm.runInContext('setBusy = () => {}; setDetailsBusy = on => { app.detailsBusy = on; };',h.context);
  await h.ui.submitGeneration(h.ui.detailsBody(form),false,'/api/create',true);
  assert.equal(h.controls.get('#details-modal').open,true); assert.equal(h.ui.readWorkingDraft().values.supplier,'Keep me');
  assert.match(h.controls.get('#toast').textContent,/Refresh.*reopen/);
  assert.equal(h.controls.get('#details-error').hidden,false);
});

test('Success shows actual server-issued certificates and safely encoded local PDF links', () => {
  const h = harness();
  h.ui.showGenerationSuccess({created:2,certificates:[{filename:'0082_Acme_Long_Recovery_Services_Mumbai_India_1.pdf',certificate_no:'0082',supplier:'Acme',plant:'plant1'},
    {filename:'CUSTOM-84_Second_Company_Pune_2.pdf',certificate_no:'CUSTOM-84',supplier:'Second',plant:'plant2',pdf_url:'/api/certificates/CUSTOM-84_Second_Company_Pune_2.pdf'}]});
  const list = h.controls.get('#created-certificates');
  assert.match(text(list),/0082.*CUSTOM-84/); assert.doesNotMatch(text(list),/Suggested COD/);
  assert.equal(all(list,node => node.tag === 'li').length,2);
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

test('Backup destination is chosen server-side and restoration sends explicit confirmation', async () => {
  const h = harness({fetch:async endpoint => ({ok:true,json:async () => ({workspace_revision:'backup-revision',...(endpoint === '/api/backup/list'
    ? {backup_destination:'D:\\Recovery',backups:[{backup_id:'snapshot-81',created_at:'2026-10-01 10:30',valid:true}]}
    : endpoint === '/api/backup/restore' ? {settings:{},state:{rows:[]},message:'Restored with recovery snapshot',restored:{warnings:['Verify certificate continuity before issuing.']}}
    : endpoint === '/api/backup/create' ? {backup:{backup_id:'snapshot-81',path:'D:\\Recovery\\snapshot.zip'},backup_destination:'D:\\Recovery',health:'Backup verified.'}
    : {backup_destination:'D:\\Recovery',settings:{backup_destination:'D:\\Recovery'}})})})});
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

test('Warning reason limit accepts 2000 characters and rejects longer text without a request', async () => {
  const h = harness(); vm.runInContext('setBusy = () => {};',h.context);
  h.ui.showWarnings([{supplier:'Acme',warnings:['Check weight.']}],{endpoint:'/api/generate',payload:{},reviewToken:'review-token'});
  h.controls.get('#override-reason').value = 'x'.repeat(2001); await h.ui.confirmWarnings();
  assert.equal(h.requests.length,0); assert.equal(h.controls.get('#warning-modal').open,true);
  h.controls.get('#override-reason').value = 'x'.repeat(2000); await h.ui.confirmWarnings();
  assert.equal(h.requests[0].payload.override_reason.length,2000);
});

test('Guarded PDF attachment selection sends the editor snapshot and adopts its new revision', async () => {
  const h = harness({fetch:async () => ({ok:true,json:async () => ({workspace_revision:'with-attachment',annexure:{id:'attachment-1',name:'List.pdf',pages:2}})})});
  const form = editing(h);
  vm.runInContext('setBusy = () => {}; setDetailsBusy = on => { app.detailsBusy = on; };',h.context);
  await h.ui.pickAnnexure();
  assert.equal(h.requests[0].payload.expected_revision,'revision-1');
  assert.equal(h.ui.detailsBody(form).expected_revision,'with-attachment');
  assert.match(h.controls.get('#annexure-name').textContent,/List.pdf.*2 pages/);
  assert.equal(h.ui.readWorkingDraft().editor.workspaceRevision,'with-attachment');
});

test('Raster first-page preview is default; full embedded PDF loads only on request', async () => {
  const id = 'a'.repeat(32);
  const h = harness({fetch:async () => ({ok:true,json:async () => ({preview_url:`/previews/${id}.pdf`,preview_image_url:`/previews/${id}.png`,certificate_no:'0091',warnings:[]})})});
  editing(h); await h.ui.previewDetails();
  assert.equal(h.controls.get('#preview-image').src,`http://127.0.0.1:8000/previews/${id}.png`);
  assert.equal(h.controls.get('#preview-image').hidden,false); assert.equal(h.controls.get('#preview-frame').hidden,true);
  assert.equal(h.controls.get('#preview-frame').src,undefined);
  assert.match(h.controls.get('#preview-status').textContent,/First-page preview.*full PDF.*attached list/);
  assert.match(h.controls.get('#preview-number').textContent,/Suggested COD: 0091.*does not reserve/);
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
