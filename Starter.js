/**
 * STUDENT WORK ANONYMIZER: STARTER  (Code.gs. Paste this once. It doesn't change.)
 * ---------------------------------------------------------------------------
 * The tool itself lives in a file in your school Drive:
 *   "Student Work Anonymizer / Tool code (do not edit) / Anonymizer code vN.js"
 * This starter builds the Anonymizer menu, loads that file whenever the tool
 * runs, and installs new versions from Rowan:
 *   - Anonymizer > Update the tool (from Rowan), or Install on the app's banner,
 *     downloads the newest code from github.com/ericdtroutman/anonymizer (code
 *     only, no student data), checks it is complete and has no mistakes, and
 *     saves it as a new file. The version before it is kept.
 *   - Anonymizer > Undo the last update goes back to the version before.
 * Nothing is sent anywhere: the starter only downloads. Student data stays in
 * this school Google account and Canvas (NY Ed Law 2-d).
 *
 * Needs the manifest (appsscript.json) from Rowan's copy page, which lists the
 * permissions the tool uses, and the Drive API service.
 */

const SHELL_VERSION = 1;
const SHELL_BASE = 'https://raw.githubusercontent.com/ericdtroutman/anonymizer/main/';
const SHELL_END = '/* END OF ANONYMIZER CODE */';
const SHELL_FOLDER = 'Tool code (do not edit)';
const SHELL_CHUNK = 30000;
const SHELL_STUBS = {};
let SHELL_LOADED = false;

/* ============================== MENU ============================== */

function onOpen() {
  try {   // remembers this Sheet for the web app
    const ss = SpreadsheetApp.getActive();
    const sp = PropertiesService.getScriptProperties();
    if (ss && sp.getProperty('SHEET_ID') !== ss.getId()) sp.setProperty('SHEET_ID', ss.getId());
  } catch (e) { /* ignore */ }
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Anonymizer')
    .addItem('Open the app (web page)', 'openWebApp')
    .addItem('Update the tool (from Rowan)', 'updateFromMenu')
    .addItem('Undo the last update', 'undoLastUpdate')
    .addSeparator()
    .addItem('1. Set up (run once)', 'setup')
    .addSeparator()
    .addItem('2. Process the Inbox', 'processInbox')
    .addItem('3. Bring in Claude\'s grades', 'bringInClaudeGrades')
    .addItem('4. Review what Rowan was not sure about', 'openReview')
    .addSeparator()
    .addSubMenu(ui.createMenu('Paper tests')
      .addItem('Make a scan folder for a test', 'makeScanFolder')
      .addItem('Match names and remove them', 'openNameMatcher'))
    .addSubMenu(ui.createMenu('Canvas')
      .addItem('Connect to Canvas', 'connectCanvas')
      .addItem('Pull submissions (new and late)', 'pullSubmissions')
      .addItem('Send grades and comments to Canvas', 'sendGradesToCanvas')
      .addItem('Update roster from Canvas', 'updateRosterFromCanvas')
      .addItem('Disconnect Canvas', 'disconnectCanvas'))
    .addSeparator()
    .addItem('Turn auto-run on or off', 'toggleAutoRun')
    .addItem('Make an Inbox folder for an assignment', 'makeAssignmentInbox')
    .addItem('Open the folders', 'showFolders')
    .addItem('Finish an assignment (delete its files)', 'finishAssignment')
    .addItem('Cancel the current run', 'cancelRun')
    .addToUi();
}

/* ============================ LOADING ============================= */

function SHELL_props_() { return PropertiesService.getScriptProperties(); }

/** Loads the tool's code (once per run). The first time, it installs the newest version. */
function SHELL_load_() {
  if (SHELL_LOADED) return;
  let id = SHELL_props_().getProperty('CODE_FILE_ID');
  if (!id) { SHELL_install_(true); id = SHELL_props_().getProperty('CODE_FILE_ID'); }
  const src = SHELL_readCode_(id);
  try { (0, eval)(src); }
  catch (e) {
    throw new Error('The tool\'s code did not load (' + e.message + '). In the Sheet, click Anonymizer > Undo the last update, and tell Rowan.');
  }
  SHELL_LOADED = true;
}

/** The code file's text, from the cache (fast) or from Drive. */
function SHELL_readCode_(id) {
  const cache = CacheService.getScriptCache();
  try {
    const n = Number(cache.get('code:' + id + ':n'));
    if (n > 0) {
      const keys = [];
      for (let i = 0; i < n; i++) keys.push('code:' + id + ':' + i);
      const got = cache.getAll(keys);
      if (keys.every(k => typeof got[k] === 'string')) return keys.map(k => got[k]).join('');
    }
  } catch (e) { /* read it from Drive */ }
  let src;
  try { src = DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8'); }
  catch (e) { throw new Error('The tool\'s code file could not be opened (' + e.message + '). Click Anonymizer > Update the tool (from Rowan) to get it again.'); }
  try {
    const put = {};
    let n = 0;
    for (let i = 0; i < src.length; i += SHELL_CHUNK) put['code:' + id + ':' + (n++)] = src.slice(i, i + SHELL_CHUNK);
    put['code:' + id + ':n'] = String(n);
    cache.putAll(put, 21600);
  } catch (e) { /* fine without the cache */ }
  return src;
}

/** Every menu item, button and timer comes through here: load the code, then run the real function. */
function SHELL_run_(name, args) {
  SHELL_load_();
  const fn = globalThis[name];
  if (typeof fn !== 'function' || fn === SHELL_STUBS[name]) {
    throw new Error('This version of the tool has nothing called "' + name + '". Tell Rowan.');
  }
  return fn.apply(globalThis, args);
}

/** For buttons added in later versions: rpc('name', [args]). Private helpers (ending in _) can't be called. */
function rpc(name, args) {
  name = String(name || '');
  if (!/^[A-Za-z][A-Za-z0-9]*$/.test(name) || /^SHELL/.test(name)) throw new Error('Unknown action.');
  return SHELL_run_(name, args || []);
}

/* ============================ UPDATES ============================= */

/** The newest version's info from Rowan: { version, date, notes, starter }. Remembered 10 minutes unless fresh. */
function SHELL_info_(fresh) {
  const cache = CacheService.getUserCache();
  if (!fresh) { try { const hit = cache.get('updateInfo'); if (hit) return JSON.parse(hit); } catch (e) { /* ignore */ } }
  const r = UrlFetchApp.fetch(SHELL_BASE + 'version.json?t=' + Date.now(), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Could not check for updates (error ' + r.getResponseCode() + '). Try again in a few minutes.');
  const info = JSON.parse(r.getContentText());
  try { cache.put('updateInfo', JSON.stringify(info), 600); } catch (e) { /* ignore */ }
  return info;
}

function SHELL_installedVersion_() { return Number(SHELL_props_().getProperty('CODE_VERSION')) || 0; }

/** The folder that keeps the code files: "Tool code (do not edit)" in the Student Work Anonymizer folder. */
function SHELL_folder_() {
  let parent = null;
  try { parent = DriveApp.getFolderById(PropertiesService.getDocumentProperties().getProperty('ROOT')); } catch (e) { parent = null; }
  if (!parent) {
    try { const ss = SpreadsheetApp.getActive(); const ps = DriveApp.getFileById(ss.getId()).getParents(); if (ps.hasNext()) parent = ps.next(); } catch (e) { parent = null; }
  }
  if (!parent) parent = DriveApp.getRootFolder();
  const it = parent.getFoldersByName(SHELL_FOLDER);
  while (it.hasNext()) { const f = it.next(); if (!f.isTrashed()) return f; }
  return parent.createFolder(SHELL_FOLDER);
}

/** Installs the newest version (first = the very first install, no "already newest" check). Returns a message. */
function SHELL_install_(first) {
  const p = SHELL_props_();
  const info = SHELL_info_(true);
  const ver = Number(info.version) || 0;
  const cur = SHELL_installedVersion_();
  if (!first && ver <= cur) return 'You already have the newest version (' + cur + ').';
  if (Number(info.starter || 1) > SHELL_VERSION) {
    throw new Error('Version ' + ver + ' needs a new starter. Open Rowan\'s copy page, copy the starter, and paste it over everything in Code.gs (Ctrl + A, Ctrl + V), then Save. Nothing was changed.');
  }
  const dl = UrlFetchApp.fetch(SHELL_BASE + 'Code.js?v=' + ver + '&t=' + Date.now(), { muteHttpExceptions: true });
  if (dl.getResponseCode() !== 200) throw new Error('Could not download the update (error ' + dl.getResponseCode() + '). Try again in a few minutes. Nothing was changed.');
  const src = dl.getContentText('UTF-8');
  if (src.indexOf('STUDENT WORK ANONYMIZER') < 0 || src.trim().slice(-SHELL_END.length) !== SHELL_END || src.indexOf('const APP_VERSION = ' + ver + ';') < 0) {
    throw new Error('The update did not download completely (or is still being posted). Try again in a few minutes. Nothing was changed.');
  }
  try { new Function(src); }   // checks the code can be read, without running it
  catch (e) { throw new Error('The update has a mistake in it, so it was not installed (' + e.message + '). Tell Rowan. Nothing was changed.'); }

  const file = SHELL_folder_().createFile('Anonymizer code v' + ver + '.js', src, MimeType.PLAIN_TEXT);
  const oldId = p.getProperty('CODE_FILE_ID'), oldVer = p.getProperty('CODE_VERSION') || '', older = p.getProperty('CODE_PREV_ID');
  if (older && older !== oldId) { try { DriveApp.getFileById(older).setTrashed(true); } catch (e) { /* already gone */ } }
  const next = { CODE_FILE_ID: file.getId(), CODE_VERSION: String(ver) };
  if (oldId) { next.CODE_PREV_ID = oldId; next.CODE_PREV_VERSION = oldVer; }
  p.setProperties(next);
  try { CacheService.getUserCache().remove('updateInfo'); } catch (e) { /* ignore */ }
  if (first) return 'Installed version ' + ver + '.';
  return 'Installed version ' + ver + '.' + (info.notes ? ' ' + info.notes : '') +
    '\n\nReload the app page (and the Sheet, if it is open) to start using it.' +
    '\n\nThe version before is kept. If anything looks wrong: Anonymizer > Undo the last update.';
}

/** Goes back to the version before the last update. */
function SHELL_undo_() {
  const p = SHELL_props_();
  const prev = p.getProperty('CODE_PREV_ID');
  if (!prev) throw new Error('There is no earlier version to go back to.');
  const cur = p.getProperty('CODE_FILE_ID'), curVer = p.getProperty('CODE_VERSION') || '';
  const prevVer = p.getProperty('CODE_PREV_VERSION') || '';
  try { DriveApp.getFileById(prev).getName(); } catch (e) { throw new Error('The earlier version\'s file is gone, so it can\'t be put back.'); }
  p.setProperties({ CODE_FILE_ID: prev, CODE_VERSION: prevVer, CODE_PREV_ID: cur || '', CODE_PREV_VERSION: curVer });
  return 'Back on version ' + prevVer + '. Reload the app page (and the Sheet, if it is open).';
}

/** For the app's banner: { current, latest, notes, ready }. */
function webUpdateInfo() {
  const current = SHELL_installedVersion_();
  try {
    const v = SHELL_info_(false);
    const latest = Number(v.version) || 0;
    return { current: current, latest: latest, date: v.date || '', notes: v.notes || '', ready: latest > current };
  } catch (e) {
    return { current: current, error: e.message };
  }
}

/** The app's Install button. */
function webInstallUpdate() {
  try { return { ok: true, text: SHELL_install_(false) }; }
  catch (e) { return { ok: false, error: (e && e.message) ? e.message : String(e) }; }
}

/** Sheet menu: Update the tool (from Rowan). */
function updateFromMenu() {
  const ui = SpreadsheetApp.getUi();
  let info;
  try { info = SHELL_info_(true); } catch (e) { ui.alert(e.message); return; }
  const cur = SHELL_installedVersion_();
  if (cur && !(Number(info.version) > cur)) { ui.alert('You have the newest version (' + cur + ').'); return; }
  const go = ui.alert('Update ready: version ' + info.version, (info.notes || '') + '\n\nInstall it now?' + (cur ? ' (You have version ' + cur + '.)' : ''), ui.ButtonSet.YES_NO);
  if (go !== ui.Button.YES) return;
  try { ui.alert(SHELL_install_(!cur)); } catch (e) { ui.alert('Update not installed', e.message, ui.ButtonSet.OK); }
}

/** Sheet menu: Undo the last update. */
function undoLastUpdate() {
  const ui = SpreadsheetApp.getUi();
  const p = SHELL_props_();
  if (!p.getProperty('CODE_PREV_ID')) { ui.alert('There is no earlier version to go back to.'); return; }
  const go = ui.alert('Undo the last update?', 'This goes back from version ' + (p.getProperty('CODE_VERSION') || '?') + ' to version ' + (p.getProperty('CODE_PREV_VERSION') || '?') + '.', ui.ButtonSet.YES_NO);
  if (go !== ui.Button.YES) return;
  try { ui.alert(SHELL_undo_()); } catch (e) { ui.alert(e.message); }
}

/* ===================== EVERY OTHER MENU ITEM AND BUTTON ===================== */
/* Each one loads the tool's code and runs the real function of the same name. */

function addMarks() { return SHELL_run_('addMarks', arguments); }
function alertMsg() { return SHELL_run_('alertMsg', arguments); }
function autoProcess() { return SHELL_run_('autoProcess', arguments); }
function bar() { return SHELL_run_('bar', arguments); }
function bringIn() { return SHELL_run_('bringIn', arguments); }
function bringInClaudeGrades() { return SHELL_run_('bringInClaudeGrades', arguments); }
function build() { return SHELL_run_('build', arguments); }
function busy() { return SHELL_run_('busy', arguments); }
function buttons() { return SHELL_run_('buttons', arguments); }
function call() { return SHELL_run_('call', arguments); }
function cancelRun() { return SHELL_run_('cancelRun', arguments); }
function card() { return SHELL_run_('card', arguments); }
function checkUpdate() { return SHELL_run_('checkUpdate', arguments); }
function clean() { return SHELL_run_('clean', arguments); }
function closeDialog() { return SHELL_run_('closeDialog', arguments); }
function closeView() { return SHELL_run_('closeView', arguments); }
function confirmable() { return SHELL_run_('confirmable', arguments); }
function connectCanvas() { return SHELL_run_('connectCanvas', arguments); }
function decide() { return SHELL_run_('decide', arguments); }
function disconnectCanvas() { return SHELL_run_('disconnectCanvas', arguments); }
function doGet() { return SHELL_run_('doGet', arguments); }
function doc() { return SHELL_run_('doc', arguments); }
function doneScreen() { return SHELL_run_('doneScreen', arguments); }
function draw() { return SHELL_run_('draw', arguments); }
function drawPdf() { return SHELL_run_('drawPdf', arguments); }
function el() { return SHELL_run_('el', arguments); }
function errText() { return SHELL_run_('errText', arguments); }
function fail() { return SHELL_run_('fail', arguments); }
function finish() { return SHELL_run_('finish', arguments); }
function finishAssignment() { return SHELL_run_('finishAssignment', arguments); }
function flagged() { return SHELL_run_('flagged', arguments); }
function frame() { return SHELL_run_('frame', arguments); }
function getFile() { return SHELL_run_('getFile', arguments); }
function go() { return SHELL_run_('go', arguments); }
function heldBox() { return SHELL_run_('heldBox', arguments); }
function keyDialog() { return SHELL_run_('keyDialog', arguments); }
function load() { return SHELL_run_('load', arguments); }
function makeAssignmentInbox() { return SHELL_run_('makeAssignmentInbox', arguments); }
function makeScanFolder() { return SHELL_run_('makeScanFolder', arguments); }
function matchBuild() { return SHELL_run_('matchBuild', arguments); }
function matchData() { return SHELL_run_('matchData', arguments); }
function matchList() { return SHELL_run_('matchList', arguments); }
function matchPage() { return SHELL_run_('matchPage', arguments); }
function matchSave() { return SHELL_run_('matchSave', arguments); }
function nameOf() { return SHELL_run_('nameOf', arguments); }
function openDialog() { return SHELL_run_('openDialog', arguments); }
function openNameMatcher() { return SHELL_run_('openNameMatcher', arguments); }
function openReview() { return SHELL_run_('openReview', arguments); }
function openView() { return SHELL_run_('openView', arguments); }
function openWebApp() { return SHELL_run_('openWebApp', arguments); }
function override() { return SHELL_run_('override', arguments); }
function paper() { return SHELL_run_('paper', arguments); }
function paperTestChose() { return SHELL_run_('paperTestChose', arguments); }
function pickAssignment() { return SHELL_run_('pickAssignment', arguments); }
function pickerChose() { return SHELL_run_('pickerChose', arguments); }
function poll() { return SHELL_run_('poll', arguments); }
function prefetch() { return SHELL_run_('prefetch', arguments); }
function processInbox() { return SHELL_run_('processInbox', arguments); }
function progress() { return SHELL_run_('progress', arguments); }
function pts() { return SHELL_run_('pts', arguments); }
function pull() { return SHELL_run_('pull', arguments); }
function pullSubmissions() { return SHELL_run_('pullSubmissions', arguments); }
function queue() { return SHELL_run_('queue', arguments); }
function refresh() { return SHELL_run_('refresh', arguments); }
function refreshNames() { return SHELL_run_('refreshNames', arguments); }
function render() { return SHELL_run_('render', arguments); }
function renderList() { return SHELL_run_('renderList', arguments); }
function reviewData() { return SHELL_run_('reviewData', arguments); }
function reviewDecide() { return SHELL_run_('reviewDecide', arguments); }
function reviewFile() { return SHELL_run_('reviewFile', arguments); }
function rosterDialog() { return SHELL_run_('rosterDialog', arguments); }
function say() { return SHELL_run_('say', arguments); }
function send() { return SHELL_run_('send', arguments); }
function sendGradesToCanvas() { return SHELL_run_('sendGradesToCanvas', arguments); }
function settings() { return SHELL_run_('settings', arguments); }
function setup() { return SHELL_run_('setup', arguments); }
function show() { return SHELL_run_('show', arguments); }
function showFolders() { return SHELL_run_('showFolders', arguments); }
function showResult() { return SHELL_run_('showResult', arguments); }
function side() { return SHELL_run_('side', arguments); }
function start() { return SHELL_run_('start', arguments); }
function stepSave() { return SHELL_run_('stepSave', arguments); }
function stepSplit() { return SHELL_run_('stepSplit', arguments); }
function stepState() { return SHELL_run_('stepState', arguments); }
function stepStrip() { return SHELL_run_('stepStrip', arguments); }
function stepText() { return SHELL_run_('stepText', arguments); }
function stepWho() { return SHELL_run_('stepWho', arguments); }
function steps() { return SHELL_run_('steps', arguments); }
function tests() { return SHELL_run_('tests', arguments); }
function toB64() { return SHELL_run_('toB64', arguments); }
function toggleAutoRun() { return SHELL_run_('toggleAutoRun', arguments); }
function updateRosterFromCanvas() { return SHELL_run_('updateRosterFromCanvas', arguments); }
function uploadFile() { return SHELL_run_('uploadFile', arguments); }
function uploadScan() { return SHELL_run_('uploadScan', arguments); }
function viewerData() { return SHELL_run_('viewerData', arguments); }
function viewerFile() { return SHELL_run_('viewerFile', arguments); }
function waiting() { return SHELL_run_('waiting', arguments); }
function webAutoRun() { return SHELL_run_('webAutoRun', arguments); }
function webBringIn() { return SHELL_run_('webBringIn', arguments); }
function webCanvasOptions() { return SHELL_run_('webCanvasOptions', arguments); }
function webConnectCanvas() { return SHELL_run_('webConnectCanvas', arguments); }
function webDisconnectCanvas() { return SHELL_run_('webDisconnectCanvas', arguments); }
function webDismissCheck() { return SHELL_run_('webDismissCheck', arguments); }
function webFinish() { return SHELL_run_('webFinish', arguments); }
function webLabelUnmatched() { return SHELL_run_('webLabelUnmatched', arguments); }
function webNameOptions() { return SHELL_run_('webNameOptions', arguments); }
function webOverrideMark() { return SHELL_run_('webOverrideMark', arguments); }
function webProcess() { return SHELL_run_('webProcess', arguments); }
function webPull() { return SHELL_run_('webPull', arguments); }
function webReleaseHeld() { return SHELL_run_('webReleaseHeld', arguments); }
function webRemoveKey() { return SHELL_run_('webRemoveKey', arguments); }
function webResult() { return SHELL_run_('webResult', arguments); }
function webRosterList() { return SHELL_run_('webRosterList', arguments); }
function webRosterSave() { return SHELL_run_('webRosterSave', arguments); }
function webSaveGrade() { return SHELL_run_('webSaveGrade', arguments); }
function webSaveKeyText() { return SHELL_run_('webSaveKeyText', arguments); }
function webScanFolder() { return SHELL_run_('webScanFolder', arguments); }
function webSend() { return SHELL_run_('webSend', arguments); }
function webStatus() { return SHELL_run_('webStatus', arguments); }
function webUpdateRoster() { return SHELL_run_('webUpdateRoster', arguments); }
function webUploadChunk() { return SHELL_run_('webUploadChunk', arguments); }
function webUploadStart() { return SHELL_run_('webUploadStart', arguments); }
function webView() { return SHELL_run_('webView', arguments); }

['addMarks', 'alertMsg', 'autoProcess', 'bar', 'bringIn', 'bringInClaudeGrades', 'build', 'busy', 'buttons', 'call', 'cancelRun', 'card', 'checkUpdate', 'clean', 'closeDialog', 'closeView', 'confirmable', 'connectCanvas', 'decide', 'disconnectCanvas', 'doGet', 'doc', 'doneScreen', 'draw', 'drawPdf', 'el', 'errText', 'fail', 'finish', 'finishAssignment', 'flagged', 'frame', 'getFile', 'go', 'heldBox', 'keyDialog', 'load', 'makeAssignmentInbox', 'makeScanFolder', 'matchBuild', 'matchData', 'matchList', 'matchPage', 'matchSave', 'nameOf', 'openDialog', 'openNameMatcher', 'openReview', 'openView', 'openWebApp', 'override', 'paper', 'paperTestChose', 'pickAssignment', 'pickerChose', 'poll', 'prefetch', 'processInbox', 'progress', 'pts', 'pull', 'pullSubmissions', 'queue', 'refresh', 'refreshNames', 'render', 'renderList', 'reviewData', 'reviewDecide', 'reviewFile', 'rosterDialog', 'say', 'send', 'sendGradesToCanvas', 'settings', 'setup', 'show', 'showFolders', 'showResult', 'side', 'start', 'stepSave', 'stepSplit', 'stepState', 'stepStrip', 'stepText', 'stepWho', 'steps', 'tests', 'toB64', 'toggleAutoRun', 'updateRosterFromCanvas', 'uploadFile', 'uploadScan', 'viewerData', 'viewerFile', 'waiting', 'webAutoRun', 'webBringIn', 'webCanvasOptions', 'webConnectCanvas', 'webDisconnectCanvas', 'webDismissCheck', 'webFinish', 'webLabelUnmatched', 'webNameOptions', 'webOverrideMark', 'webProcess', 'webPull', 'webReleaseHeld', 'webRemoveKey', 'webResult', 'webRosterList', 'webRosterSave', 'webSaveGrade', 'webSaveKeyText', 'webScanFolder', 'webSend', 'webStatus', 'webUpdateRoster', 'webUploadChunk', 'webUploadStart', 'webView'].forEach(n => { SHELL_STUBS[n] = globalThis[n]; });
