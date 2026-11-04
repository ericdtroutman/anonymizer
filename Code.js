/**
 * STUDENT WORK ANONYMIZER  (Google Apps Script, bound to a Google Sheet)
 * ---------------------------------------------------------------------
 * Takes the zip from Canvas "Download Submissions", gives every student a
 * random code (S-4821), removes names, emails and Canvas IDs from INSIDE each
 * file, and saves clean PDFs named by code. The code-to-name key stays in
 * this Sheet and never leaves your Google account.
 *
 * Handles: Word, PowerPoint, Excel, PDF, Google Docs/Slides/Sheets (incl.
 * shortcuts and .gslides/.gdoc files), Canvas text entries and Google-link
 * submissions, and zips inside the Canvas zip. Pictures (a Canva deck saved
 * as 1.png, 2.png...) are combined into ONE PDF per student, in slide order.
 * Flags anything it can't clean (video, Apple files, Canva links, pictures)
 * so you can check it by eye.
 *
 * Version 2 (2026-10-29): student zips of pictures, .gslides/.gdoc files,
 * whole-assignment counts in the Done box, wider Log code column.
 * Version 3 (2026-10-29): PDFs keep their original look. Each page is checked
 * for a name; only pages with a name are swapped for a plain version with the
 * code. Uses the free pdf-lib add-on (loaded from cdnjs), so Google asks once
 * for permission to "connect to an external service".
 * Version 4 (2026-10-29): late work. Process the Inbox asks which assignment
 * the files are for and lists earlier ones by number. Late files keep the
 * student's original code and go in an "Added <date> (send these)" folder
 * inside that assignment's folder. New assignment names no longer get a date.
 * Version 5 (2026-10-29): Canvas connection (pull submissions, send grades and
 * comments, update roster) using your own Canvas token; auto-run every 10
 * minutes; Claude writes grades into a separate codes-only Sheet and
 * "Bring in Claude's grades" copies them into Results; "Finish an assignment"
 * permanently deletes that assignment's files and rows. Originals are now kept
 * in a folder per assignment.
 * Version 6 (2026-10-29): fixes "Canvas said the token is not valid (401)" in
 * Pull submissions. Canvas hands file downloads to a separate file server, and
 * version 5 sent your Canvas token there too, which that server rejects. Now
 * the token only goes to Canvas itself. Also: a file that won't download is
 * skipped (and tried again next pull) instead of stopping the whole pull, and
 * students already pulled are always recorded, even if the pull stops early.
 * Version 7 (2026-10-29): the "Claude Grades (codes only)" Sheet can live in
 * the top Student Work Anonymizer folder OR inside "2 Ready for Claude" (where
 * Claude can reach it). Bring in Claude's grades and Send grades look in both.
 * Version 8 (2026-10-29): dropdowns instead of typing a number. Pull
 * submissions, Send grades, Process the Inbox (loose files), Make an Inbox
 * folder and Finish an assignment open a small window with a list to pick
 * from. For Pull and Send the list also shows the other assignments in your
 * Canvas courses, so you no longer have to paste a link the first time.
 * Pasting a link or typing a new name still works.
 * Version 9 (2026-11-02): Send grades never lowers a grade. Before sending,
 * it checks what Canvas already has for each student. A grade only goes to
 * Canvas when the student has no grade yet or the new score is higher. Equal
 * or higher grades already in Canvas (and excused students) are left alone,
 * with no new comment, and marked "Kept Canvas grade" on the Results tab.
 * Version 10 (2026-11-03): a Review window (4. Review what Rowan was not
 *  sure about): each answer Rowan flagged, with the student's page. Click Full
 *  credit, Rowan's points or No credit (or type points); the total updates in
 *  the Claude Grades Sheet. Send grades warns if anything is still waiting.
 * Version 11 (2026-11-03): paper tests WITHOUT cover sheets (they replace the
 *  version 10 cover sheets). Scan the whole class into the test's scan folder.
 *  The tool splits it into pages and reads them. Then Paper tests > Match
 *  names and remove them walks you through it: check where each test starts,
 *  drag a line under the name area at the top, and pick whose test each one
 *  is (the tool fills in the names it could read). The name area is painted
 *  white and the page saved as a flat picture, so the name is really gone.
 *  A test where a name still shows up is held back, NOT sent to Claude.
 * Version 24 (2026-11-04): your classes. Settings > Your classes lists every
 *  Canvas course you teach in; tick the ones you use. Only ticked classes show
 *  in the assignment lists (Pull, Send, Paper test), and Update roster adds
 *  the students from every ticked class (how a second class, like Earth
 *  Science, gets on the class list).
 * Version 23 (2026-11-04): notes for Rowan on tests. The Answer key window
 *  (and the Paper test upload) now has "Notes for Rowan" (how you want it
 *  graded: lenient, points per part, follow-through credit, other answers to
 *  accept) and an optional blank copy of the test. Notes are saved as
 *  "Grading notes - <test>.txt" and the blank as "Blank test - <test>" next
 *  to the key in "2 Ready for Claude/<test>". Rowan reads them before grading;
 *  your notes win over the key. Change them any time and ask Rowan to regrade.
 * Version 22 (2026-11-04): cover a name the tool missed. In See the work,
 *  "Cover a name or face" lets you drag a box over anything on any page
 *  (a name in a picture, a face, a signature). Save paints those spots white
 *  and saves each of those pages as a flat picture, so it's really gone; the
 *  other pages stay as they were. Big files (over 15 MB) now open in See the
 *  work too, so they can be covered.
 * Version 21 (2026-11-04): fixing a match. In See the work, every file has
 *  "Wrong student?": pick the right student and the file moves to them (its
 *  grade, review and marks go with it when nothing else is under that code).
 *  The student list always shows everyone, with a button to update the class
 *  list from Canvas when someone is missing. Matching no longer leaves the
 *  page stuck on Loading. Pull updates the class list from Canvas first, so
 *  new students' work matches by itself.
 * Version 20 (2026-11-04): one slow file can't block the Inbox any more.
 *  Google stops any run after 6 minutes. Before, a file too big to clean in
 *  that time stayed in the Inbox and every Clean started over on it. Now each
 *  file's tries are counted: on the 3rd try a PDF is rebuilt from its text
 *  in one step (faster), and a file that still stops Google is set aside in
 *  Processed originals and flagged "No file made", so the rest get cleaned.
 *  The app also waits up to 6 minutes for an answer and says plainly when
 *  Google stopped a run.
 * Version 19 (2026-11-04): runs from the starter. Code.gs in Apps Script is
 *  now a small starter you paste once (Rowan's copy page). It builds the
 *  Anonymizer menu and keeps this code in a file in your school Drive
 *  ("Student Work Anonymizer/Tool code (do not edit)"). Install puts a new
 *  version in that file; Anonymizer > Undo the last update puts the old one
 *  back. The menu, Install and Undo live in the starter, so they work even if
 *  an update has a problem. (Google doesn't let school accounts use the Apps
 *  Script API that version 17 used.)
 * Version 18 (2026-11-04): the app shows which version you have (next to
 *  the title). The first update installed with the Install button.
 * Version 17 (2026-11-04): faster, and the tool updates itself.
 *  - The app opens right away with what it showed last time, then fills in
 *    the newest numbers. Refresh reads your folders in a few big requests
 *    instead of one folder at a time, and opens each Sheet once.
 *  - See the work loads the next student while you look at this one, and
 *    keeps the ones you've seen, so flipping back and forth is quick.
 *  - Your Canvas assignment list is remembered for 10 minutes.
 *  - Updates: when Rowan makes a new version, the app shows "Update ready"
 *    with an Install button. Install saves a copy of the current code in
 *    Apps Script's Project history (so it can be undone) and puts the new
 *    code in place. No more copying and pasting. Needs a one-time setup: the
 *    Apps Script API switched on, and the manifest from Rowan's copy page.
 * Version 16 (2026-11-04): Send to Canvas always works. The Send step on
 *  each card (and a "Send to Canvas again" button) can be clicked even after
 *  everything was sent. The tool then checks Canvas and re-sends only the
 *  grades Canvas is missing (or has lower), with the feedback comment if
 *  Canvas doesn't have it yet. Canvas grades are still never lowered. Send
 *  and Pull also say how many students turned in work in Canvas that has no
 *  grade from Rowan, and Pull no longer marks a student as pulled when there
 *  was nothing it could download (Google Assignments, recordings, tool links).
 *  See the work shows each student's name (with the code beside it), sorted
 *  by name. The names come from the Key tab and only show in your app.
 * Version 15 (2026-11-03): everything in the app. Each card has See the work:
 *  flip through every student's coded file in the app, with the score and
 *  feedback beside it (change them right there), Rowan's review items, and
 *  anything the tool flagged (Looks fine) or couldn't match (Who is it?). Tests
 *  show Rowan's red-pen marks (checks, X's, points, short notes, the score).
 *  Click a test's check, X or points to change it (full credit, no credit, or
 *  points); the score moves with it. Settings has the class list (nicknames).
 * Version 14 (2026-11-03): no Sheet or folders needed for paper tests. In the
 *  app: upload the class scan, add the answer key (upload a file or type the
 *  answers), and see held tests with a button to send one to Rowan once you've
 *  checked it has no name. An answer key no longer makes the first scan look
 *  like late work.
 * Version 13 (2026-11-03): web app fixes. Coded work is counted by student
 *  (a second file from the same student no longer looks like one more to
 *  grade). Steps after the one you're on no longer show green. An old run
 *  with nothing left in the folders says so and offers Finish to clear it.
 * Version 12 (2026-11-03): a web app. The same tool as one page with its own
 *  link (Anonymizer > Open the app shows how to get it). Each assignment has a
 *  card with where it stands and the next button to press; Review and Match
 *  names open inside the page; Canvas, roster and auto-run are in Settings.
 *  The Sheet menu still works exactly as before.
 *
 * STUDENT DATA (NY Ed Law 2-d): names, IDs and original work only move between
 * Canvas and this school Google account. Claude only ever sees the coded PDFs
 * in "2 Ready for Claude" and the codes-only grades Sheet (including its
 * Review tab). Scans and their single pages have names, so they stay in the
 * PRIVATE folder.
 *
 * ONE-TIME SETUP
 *  1. Use your SCHOOL Google account. New Google Sheet -> name it
 *     "Student Work Anonymizer".
 *  2. Extensions > Apps Script. Delete what's there, paste this whole file, Save.
 *  3. In Apps Script, left side: Services (+) > Drive API > Add.
 *  4. Reload the Sheet. Menu "Anonymizer" > "1. Set up (run once)".
 *     Approve the permissions (Google says "unverified" because you wrote
 *     it: Advanced > Go to Student Work Anonymizer).
 *  5. Roster tab: in Canvas, Grades > Export > CSV. Copy the "Student" and
 *     "ID" columns into Roster columns A and B. (Optional: "SIS Login ID"
 *     into column E if it's their school email.)
 *
 * EACH PROJECT
 *  1. Canvas assignment > Download Submissions. Drag the zip into the
 *     "1 Inbox" folder (Anonymizer > Open the folders).
 *     For Google Docs/Slides submitted as links, you can also drop a
 *     shortcut or copy of the file into the Inbox.
 *  2. Anonymizer > "2. Process the Inbox". Name the assignment.
 *     Big batches stop around 5 minutes: just run it again to continue.
 *  3. Open "2 Ready for Claude/<assignment>". Skim anything the Log tab
 *     flags, then download that folder and give it to Claude with the rubric.
 *  4. Paste Claude's Code | Score | Feedback table into the Results tab
 *     at A2. Student names fill in automatically.
 *
 * EACH PAPER TEST OR QUIZ (version 11)
 *  1. Anonymizer > Paper tests > Make a scan folder for a test. Pick the test.
 *  2. Scan the class stack (one PDF is fine) into that folder.
 *  3. Auto-run (or Process the Inbox) reads the pages.
 *  4. Anonymizer > Paper tests > Match names and remove them. Follow the steps.
 *  5. Ask Claude to grade, then Anonymizer > 4. Review, then Send grades.
 */

const APP = 'Anonymizer';
const APP_VERSION = 24;   // Version 17: the app compares this with the newest version from Rowan
const TAB_ROSTER = 'Roster';
const TAB_KEY = 'Key';
const TAB_RESULTS = 'Results';
const TAB_LOG = 'Log';
const TIME_LIMIT_MS = 4 * 60 * 1000;      // stop starting new files after 4 minutes
const HARD_STOP_MS = 5.2 * 60 * 1000;     // stop mid-PDF after 5.2 minutes (Google's limit is 6)
const OUT_OF_TIME = 'OUT_OF_TIME';
const PDFLIB_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js';

const DOC_EXT = ['docx', 'doc', 'odt', 'rtf', 'txt'];
const SLIDE_EXT = ['pptx', 'ppt', 'odp'];
const SHEET_EXT = ['xlsx', 'xls', 'ods', 'csv'];
const IMG_EXT = ['jpg', 'jpeg', 'png', 'gif', 'bmp'];

const MIME_BY_EXT = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  doc: 'application/msword',
  odt: 'application/vnd.oasis.opendocument.text',
  rtf: 'application/rtf',
  txt: 'text/plain',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  ppt: 'application/vnd.ms-powerpoint',
  odp: 'application/vnd.oasis.opendocument.presentation',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xls: 'application/vnd.ms-excel',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
  csv: 'text/csv',
  pdf: 'application/pdf',
  html: 'text/html',
  htm: 'text/html',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  png: 'image/png',
  gif: 'image/gif',
  bmp: 'image/bmp',
  heic: 'image/heic',
  webp: 'image/webp',
  mp4: 'video/mp4',
  mov: 'video/quicktime',
  zip: 'application/zip'
};

/* ============================== MENU ============================== */
/* Version 19: the Anonymizer menu (onOpen) is built by the starter in Code.gs. */

/* ============================== SETUP ============================= */

function setup() {
  const ss = ss_();
  const ui = SpreadsheetApp.getUi();

  if (typeof Drive === 'undefined') {
    ui.alert('One more step first',
      'In Extensions > Apps Script, click Services (+) on the left, choose "Drive API", click Add, then run Set up again.',
      ui.ButtonSet.OK);
    return;
  }

  const roster = ensureSheet_(TAB_ROSTER, ['Student (Last, First)', 'Canvas ID',
    'Other names to hide (comma-separated nicknames)', 'Full name only? (Y)', 'Email (optional)']);
  roster.getRange('B:B').setNumberFormat('@');
  roster.getRange('A1').setNote('Paste the "Student" column from the Canvas gradebook export here. ' +
    'Column D: put Y for a student whose first or last name is also a common word (Moon, Star, Will, Hope) ' +
    'so only their full name gets hidden.');

  const key = ensureSheet_(TAB_KEY, ['Assignment', 'Code', 'Student', 'Canvas ID', 'Files', 'Notes']);
  key.getRange('D:D').setNumberFormat('@');
  key.getRange('A1').setNote('PRIVATE. This is the only place codes connect to names. Never share this Sheet.');

  const res = ensureSheet_(TAB_RESULTS, ['Code', 'Score', 'Feedback', 'Student', 'Canvas ID', 'Sent to Canvas', 'sent copy (do not edit)']);
  res.getRange('D2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(VLOOKUP(A2:A,Key!B:D,2,FALSE),"code not found")))');
  res.getRange('E2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(VLOOKUP(A2:A,Key!B:D,3,FALSE),"")))');
  res.getRange('A1').setNote('Filled by Anonymizer > Bring in Claude\'s grades (or paste Claude\'s Code | Score | Feedback table under the last row).');

  ensureSheet_(TAB_LOG, ['Time', 'Assignment', 'Code', 'Result', 'Note']);

  const blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);

  const props = PropertiesService.getDocumentProperties();
  let root;
  try { root = DriveApp.getFolderById(props.getProperty('ROOT')); } catch (e) { root = null; }
  if (!root) {
    root = DriveApp.createFolder('Student Work Anonymizer');
    props.setProperties({
      ROOT: root.getId(),
      INBOX: root.createFolder('1 Inbox - drop the Canvas zip here').getId(),
      READY: root.createFolder('2 Ready for Claude').getId(),
      DONE: root.createFolder('3 Processed originals (PRIVATE)').getId()
    });
    const done = DriveApp.getFolderById(props.getProperty('DONE'));
    props.setProperty('TEMP', done.createFolder('_temp').getId());
    try { DriveApp.getFileById(ss.getId()).moveTo(root); } catch (e) { /* fine if it can't move */ }
  }

  ss.setActiveSheet(roster);
  ui.alert('Set up done',
    'Next: paste your class roster into the Roster tab (Canvas Grades > Export > CSV, then copy the "Student" and "ID" columns).\n\n' +
    'Then drop a Canvas submissions zip into the "1 Inbox" folder and run Anonymizer > Process the Inbox.',
    ui.ButtonSet.OK);
}

function showFolders() {
  const f = folders_();
  const link = (folder, label) => '<p><a target="_blank" href="' + folder.getUrl() + '">' + label + '</a></p>';
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;font-size:14px">' +
    link(f.inbox, '1 Inbox: drop the Canvas zip here') +
    link(f.ready, '2 Ready for Claude: clean, coded PDFs') +
    link(f.done, '3 Processed originals (PRIVATE, has names)') +
    '</div>').setWidth(380).setHeight(160);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Anonymizer folders');
}

function cancelRun() {
  PropertiesService.getDocumentProperties().deleteProperty('ACTIVE_RUN');
  PropertiesService.getDocumentProperties().deleteProperty('ACTIVE_RUN_FOLDER');
  SpreadsheetApp.getUi().alert('Cleared. The next "Process the Inbox" will ask for a new assignment name.');
}

/* ============================ MAIN RUN ============================ */
/*
 * Two ways files come in:
 *  - Loose files dropped straight into "1 Inbox": you're asked which assignment.
 *  - Files inside an assignment folder in the Inbox ("1 Inbox/Apollo 13 Project"),
 *    which is where "Pull submissions from Canvas" puts them. No question asked:
 *    the folder name is the assignment. Auto-run handles these on its own.
 */

async function processInbox() {
  const start = Date.now();
  const ui = SpreadsheetApp.getUi();
  const f = folders_();
  const roster = readRoster_();
  if (!roster.length) {
    ui.alert('The Roster tab is empty. Use Anonymizer > Canvas > Update roster from Canvas, or paste your class list (see the note on cell A1).');
    return;
  }
  const props = PropertiesService.getDocumentProperties();
  const results = [];

  // A) Loose files in the Inbox
  let run = props.getProperty('ACTIVE_RUN');
  if (hasFiles_(f.inbox) || run) {
    if (!run) {
      // Ask with a dropdown. Picking one sets ACTIVE_RUN and runs this again (see pickerChose).
      showNamePicker_('loose', 'Which assignment are these files for?',
        'There are loose files in the Inbox. Pick the assignment they belong to (late work goes with the earlier one).', true);
      return;
    }
    const r = await processBatch_(run, f.inbox, f, roster, start);
    results.push(r);
    if (r.timedOut) return notDoneAlert_(ui, r);
    props.deleteProperty('ACTIVE_RUN');
  }

  // B) Assignment folders inside the Inbox (from Canvas, or ones you made)
  const subs = assignmentInboxes_(f.inbox);
  for (let i = 0; i < subs.length; i++) {
    const r = await batchFor_(subs[i], f, roster, start);
    results.push(r);
    if (r.timedOut) return notDoneAlert_(ui, r);
  }

  // C) Paper tests saved in the Match window: make their coded PDFs (version 11)
  const built = await buildScanPlans_(f, start);
  if (built.done) results.push({ text: built.text });

  if (!results.length) {
    ui.alert('Nothing to process. The Inbox is empty.\n\nUse Anonymizer > Canvas > Pull submissions, or drop files into the Inbox.');
    return;
  }
  ui.alert('Done', results.map(doneLine_).join('\n\n') +
    '\n\nSkim the PDFs before Claude grades them. The tool catches names in text, but not in pictures, handwriting or video.',
    ui.ButtonSet.OK);
}

/** Runs on a timer when auto-run is on. No pop-ups; results go to the Log tab. */
async function autoProcess() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) return;           // a run is already going
  try {
    const start = Date.now();
    const f = folders_();
    const roster = readRoster_();
    if (!roster.length) return;
    const subs = assignmentInboxes_(f.inbox);
    for (let i = 0; i < subs.length; i++) {
      const r = await batchFor_(subs[i], f, roster, start);
      if (r.timedOut) break;
      if (r.finished && r.scansRead === undefined) log_(r.run, '-', 'Auto-run finished', doneLine_(r));
    }
    if (Date.now() - start < TIME_LIMIT_MS) {
      const built = await buildScanPlans_(f, start);
      if (built.done) log_('-', '-', 'Auto-run finished', built.text);
    }
  } finally {
    lock.releaseLock();
  }
}

/** An Inbox folder named "<test> (scans)" holds scanned paper tests; any other folder is regular work. */
async function batchFor_(folder, f, roster, start) {
  const name = folder.getName();
  if (isScanFolder_(name)) return await processScanBatch_(runFromScanFolder_(name), folder, f, roster, start);
  return await processBatch_(name, folder, f, roster, start);
}

function notDoneAlert_(ui, r) {
  ui.alert('Not done yet',
    'Google limits how long a script can run. ' + (r.progress || (r.made + ' file(s) finished so far')) + ' for ' + r.run + '.\n\n' +
    'Run Anonymizer > Process the Inbox again to keep going' + (autoRunIsOn_() ? ', or just wait: auto-run will pick it up.' : '.'),
    ui.ButtonSet.OK);
}

function doneLine_(r) {
  if (r.text) return r.text;
  if (r.scansRead !== undefined) return r.run + ': ' + r.progress + ' from the scans. Next: Match names (the button on the test\'s card in the app, or Anonymizer > Paper tests > Match names and remove them).';
  return r.run + ': ' + r.made + ' clean file(s) in "2 Ready for Claude/' + r.run +
    (r.added ? '/' + r.outName + '" (late or extra work; same codes as before)' : '"') +
    (r.flagged ? '. ' + r.flagged + ' need a quick look (see the Log tab).' : '.') +
    '';
}

/* ===================== ASSIGNMENT PICKER (dropdown) ================ */
/*
 * Version 8. Google's plain pop-up boxes can only take typed text, so the
 * tool opens a small window with a dropdown instead. When you click the
 * button, the window calls pickerChose(action, value), which does the job.
 * Only assignment names are in the list (no student information).
 */

const PICKER_ACTIONS = {
  pull:   { button: 'Pull submissions', working: 'Pulling from Canvas. This can take a few minutes...' },
  send:   { button: 'Send grades', working: 'Getting the grades ready...' },
  loose:  { button: 'Clean these files', working: 'Cleaning the files. This can take a few minutes...' },
  inbox:  { button: 'Make the folder', working: 'Making the folder...' },
  finish: { button: 'Next', working: 'One moment...' },
  // Version 10. This returns a message with links for the window to show (fn = the function the window calls).
  scans:  { button: 'Make the folder', working: 'One moment...', fn: 'paperTestChose' }
};

/** Dropdown of assignment names already on the Key tab (newest first), plus an optional box for a new name. */
function showNamePicker_(action, title, intro, allowNew) {
  const options = pastAssignments_().map(a => ({ group: 'Assignments you have used', label: a, value: 'name:' + a }));
  if (!options.length && !allowNew) {
    SpreadsheetApp.getUi().alert('There are no assignments on the Key tab yet.');
    return;
  }
  showPicker_(action, title, intro, options,
    allowNew ? { label: 'Or type a NEW assignment name:', placeholder: 'for example: Hidden Figures Project' } : null);
}

/** Dropdown of Canvas assignments: the saved ones first, then the rest of your current Canvas courses. */
function showCanvasPicker_(action, title, intro) {
  canvasSettings_();   // stops here with a clear message if Canvas is not connected
  showPicker_(action, title, intro, canvasPickerOptions_(),
    { label: 'Not in the list? Paste the assignment link from Canvas:', placeholder: 'https://.../courses/123/assignments/456' });
}

/** The saved Canvas assignments, then the rest of your current Canvas courses: [{group, label, value}]. Version 12: shared with the web app. */
function canvasPickerOptions_() {
  const sh = ensureSheet_(TAB_CANVAS, ['Assignment', 'Canvas link', 'Course ID', 'Assignment ID', 'Last pulled']);
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0] && r[3] !== '');
  const options = rows.map(r => ({ group: 'Saved (used before)', label: String(r[0]), value: 'saved:' + String(r[3]) }));
  const savedIds = {};
  rows.forEach(r => { savedIds[String(r[3])] = true; });
  canvasCourseAssignments_(rows.map(r => String(r[2]))).forEach(c => {
    c.assignments.forEach(a => {
      if (savedIds[String(a.id)]) return;
      options.push({ group: 'In Canvas: ' + c.name, label: String(a.name), value: 'canvas:' + c.id + ':' + a.id });
    });
  });
  return options;
}

/* ---------- version 24: which Canvas classes are yours ---------- */
/** The class ids you ticked in Settings > Your classes, or null if you haven't picked yet (then all show). */
function shownCourses_() {
  try {
    const a = JSON.parse(PropertiesService.getDocumentProperties().getProperty('COURSES_SHOWN') || 'null');
    return Array.isArray(a) ? a.map(String) : null;
  } catch (e) { return null; }
}
/** Every Canvas course you teach in, with whether it's ticked. Through rpc. */
function webCourseList() {
  try { canvasSettings_(); } catch (e) { return { ok: false, error: 'Connect to Canvas first (above).' }; }
  const shown = shownCourses_();
  const seen = {};
  const courses = canvasGetAll_('/api/v1/courses?enrollment_type=teacher&enrollment_state=active&include[]=term&include[]=total_students')
    .filter(c => c && c.id && !seen[c.id] && (seen[c.id] = true))
    .map(c => ({ id: String(c.id), name: c.name || c.course_code || ('Course ' + c.id), term: (c.term && c.term.name && c.term.name !== 'Default Term') ? c.term.name : '',
                 students: (typeof c.total_students === 'number') ? c.total_students : null, shown: !shown || shown.indexOf(String(c.id)) > -1 }));
  return { ok: true, courses: courses, chosen: !!shown };
}
/** Saves which classes are ticked. Through rpc. */
function webSaveCourses(ids) {
  const a = (ids || []).map(String).filter(x => /^\d+$/.test(x));
  PropertiesService.getDocumentProperties().setProperty('COURSES_SHOWN', JSON.stringify(a));
  return { ok: true, text: a.length + ' class' + (a.length === 1 ? '' : 'es') + ' on. Only these show in the assignment lists, and Update roster adds their students.' };
}

/** Version 17: the Canvas list is remembered for 10 minutes, so Pull and Send open quickly. */
function canvasCourseAssignments_(savedCourseIds) {
  const cache = CacheService.getUserCache();
  const key = 'canvasList:' + (savedCourseIds || []).slice().sort().join(',') + '|' + JSON.stringify(shownCourses_());   // version 24
  try { const hit = cache.get(key); if (hit) return JSON.parse(hit); } catch (e) { /* ignore */ }
  const out = canvasCourseAssignmentsFresh_(savedCourseIds);
  try {
    const slim = out.map(c => ({ id: c.id, name: c.name, assignments: c.assignments.map(a => ({ id: a.id, name: a.name })) }));
    const text = JSON.stringify(slim);
    if (out.length && text.length < 90000) cache.put(key, text, 600);
  } catch (e) { /* ignore */ }
  return out;
}

/** Your current Canvas courses (and any saved ones) with their assignments, newest due date first. Never stops the picker. */
function canvasCourseAssignmentsFresh_(savedCourseIds) {
  const out = [];
  try {
    const courses = {};
    const order = [];
    canvasGetAll_('/api/v1/courses?enrollment_type=teacher&enrollment_state=active').forEach(c => {
      if (c && c.id && !courses[c.id]) { courses[c.id] = c.name || c.course_code || ('Course ' + c.id); order.push(String(c.id)); }
    });
    (savedCourseIds || []).forEach(id => {
      if (id && !courses[id]) {
        try { courses[id] = JSON.parse(canvasFetch_('/api/v1/courses/' + id).getContentText()).name || ('Course ' + id); order.push(id); }
        catch (e) { /* course closed or not reachable */ }
      }
    });
    const shown = shownCourses_();   // version 24: only the classes you ticked
    order.filter(id => !shown || shown.indexOf(String(id)) > -1).forEach(id => {
      try {
        const list = canvasGetAll_('/api/v1/courses/' + id + '/assignments?order_by=due_at')
          .filter(a => a && a.id && a.name && a.published !== false).reverse();
        if (list.length) out.push({ id: id, name: courses[id], assignments: list });
      } catch (e) { /* skip this course */ }
    });
  } catch (e) { /* Canvas list not available; saved ones and the link box still work */ }
  return out;
}

function escHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Builds the picker window. options: [{group, label, value}]. other: {label, placeholder} or null. */
function pickerHtml_(action, intro, options, other) {
  const a = PICKER_ACTIONS[action];
  let opts = '<option value="">Choose an assignment...</option>';
  let group = null;
  options.forEach(o => {
    if (o.group !== group) {
      if (group !== null) opts += '</optgroup>';
      group = o.group;
      opts += '<optgroup label="' + escHtml_(group) + '">';
    }
    opts += '<option value="' + escHtml_(o.value) + '">' + escHtml_(o.label) + '</option>';
  });
  if (group !== null) opts += '</optgroup>';
  return '<div id="box" style="font-family:Arial,sans-serif;font-size:14px;line-height:1.4">' +
    '<p style="margin:0 0 10px">' + escHtml_(intro) + '</p>' +
    '<select id="pick" style="width:100%;font-size:14px;padding:6px">' + opts + '</select>' +
    (other ? '<p style="margin:14px 0 4px">' + escHtml_(other.label) + '</p>' +
      '<input id="other" type="text" style="width:100%;box-sizing:border-box;font-size:14px;padding:6px" placeholder="' + escHtml_(other.placeholder) + '">' : '') +
    '<p id="msg" style="color:#b00020;min-height:18px;margin:10px 0"></p>' +
    '<button id="go" style="font-size:14px;padding:6px 14px;background:#1a73e8;color:#fff;border:0;border-radius:4px;cursor:pointer">' + escHtml_(a.button) + '</button> ' +
    '<button id="cancel" style="font-size:14px;padding:6px 14px;margin-left:6px">Cancel</button>' +
    '</div>' +
    '<script>' +
    'var pick=document.getElementById("pick"),other=document.getElementById("other"),msg=document.getElementById("msg"),go=document.getElementById("go");' +
    'pick.onchange=function(){if(other&&pick.value)other.value="";};' +
    'if(other)other.oninput=function(){if(other.value.trim())pick.value="";};' +
    'document.getElementById("cancel").onclick=function(){google.script.host.close();};' +
    'go.onclick=function(){' +
    ' var v=(other&&other.value.trim())?"typed:"+other.value.trim():pick.value;' +
    ' if(!v){msg.textContent="Choose an assignment from the list first.";return;}' +
    ' go.disabled=true;pick.disabled=true;if(other)other.disabled=true;' +
    ' msg.style.color="#444";msg.textContent=' + JSON.stringify(a.working) + ';' +
    ' google.script.run.withSuccessHandler(function(r){if(r&&r.message){showDone(r);}else{google.script.host.close();}})' +
    '  .withFailureHandler(function(e){msg.style.color="#b00020";msg.textContent=(e&&e.message)?e.message:String(e);' +
    '   go.disabled=false;pick.disabled=false;if(other)other.disabled=false;})' +
    '  [' + JSON.stringify(a.fn || 'pickerChose') + '](' + JSON.stringify(action) + ',v);' +
    '};' +
    // Version 10: a finished job can hand back a message and links to show here
    'function showDone(r){var box=document.getElementById("box");box.innerHTML="";' +
    ' String(r.message).split("\\n").forEach(function(t){if(!t)return;var p=document.createElement("p");p.style.margin="0 0 10px";p.textContent=t;box.appendChild(p);});' +
    ' (r.links||[]).forEach(function(l){var p=document.createElement("p");p.style.margin="0 0 10px";var a=document.createElement("a");a.href=l.url;a.target="_blank";a.textContent=l.label;a.style.fontWeight="bold";p.appendChild(a);box.appendChild(p);});' +
    ' var c=document.createElement("button");c.textContent="Close";c.style.cssText="font-size:14px;padding:6px 14px;margin-top:6px";c.onclick=function(){google.script.host.close();};box.appendChild(c);}' +
    '</script>';
}

function showPicker_(action, title, intro, options, other) {
  const html = HtmlService.createHtmlOutput(pickerHtml_(action, intro, options, other))
    .setWidth(460).setHeight(other ? 300 : 230);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

/** Called by the picker window. action is one of PICKER_ACTIONS; value is what was picked or typed. */
async function pickerChose(action, value) {
  value = String(value || '').trim();
  if (!PICKER_ACTIONS[action]) throw new Error('Unknown choice. Close this window and try again.');
  if (!value) throw new Error('Choose an assignment from the list first.');
  if (action === 'pull' || action === 'send') {
    const asg = resolveCanvasAssignment_(value);   // throws a plain message if the link is wrong
    if (action === 'pull') return pullSubmissionsFor_(asg);
    return sendGradesFor_(asg);
  }
  const run = resolveAssignmentName_(value);
  if (!run) throw new Error('Choose an assignment from the list, or type a name.');
  if (action === 'loose') {
    PropertiesService.getDocumentProperties().setProperty('ACTIVE_RUN', run);
    return await processInbox();
  }
  if (action === 'inbox') return makeAssignmentInboxFor_(run);
  if (action === 'finish') {
    if (pastAssignments_().indexOf(run) < 0) throw new Error('Pick the assignment from the list.');
    return finishAssignmentFor_(run);
  }
}

/** "name:Apollo 13 Project" or "typed:apollo 13 project" -> the assignment name (matches an earlier one regardless of capitals). */
function resolveAssignmentName_(value) {
  const m = String(value).match(/^(name|typed):([\s\S]*)$/);
  const answer = (m ? m[2] : String(value)).trim().replace(/[\/\\]/g, '-');
  if (!answer) return '';
  const earlier = pastAssignments_();
  return earlier.filter(a => a.toLowerCase() === answer.toLowerCase())[0] || answer;
}

function hasFiles_(folder) { return folder.getFiles().hasNext(); }

function assignmentInboxes_(inbox) {
  const out = [];
  const it = inbox.getFolders();
  while (it.hasNext()) { const fo = it.next(); if (hasFiles_(fo)) out.push(fo); }
  return out;
}

/**
 * Cleans every file in one Inbox folder for one assignment.
 * Returns { run, made, flagged, timedOut, finished, added, outName }.
 */
async function processBatch_(run, inbox, f, roster, start) {
  const props = PropertiesService.getDocumentProperties();
  const outKey = 'OUT::' + run, madeKey = 'MADE::' + run, flagKey = 'FLAG::' + run;
  try { ss_().getSheetByName(TAB_LOG).setColumnWidth(3, 90); } catch (e) { /* ignore */ }

  // Where this batch's clean files go
  let outFolder = null;
  try { outFolder = DriveApp.getFolderById(props.getProperty(outKey)); if (outFolder.isTrashed()) outFolder = null; } catch (e) { outFolder = null; }
  if (!outFolder) {
    const main = folderByName_(f.ready, run);
    const already = hasWork_(main);
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    outFolder = already ? main.createFolder('Added ' + today + ' (send these)') : main;
    props.setProperties({ [outKey]: outFolder.getId(), [madeKey]: '0', [flagKey]: '0' });
  }
  const doneFolder = folderByName_(f.done, run);   // originals, kept per assignment so cleanup is easy

  unzipAll_(inbox, doneFolder);
  const codeMap = ensureCodes_(run, roster);
  const ctx = {
    run: run, roster: roster, codeMap: codeMap,
    rules: buildRules_(roster, codeMap),
    outFolder: outFolder, temp: f.temp,
    fuzzy: buildFuzzy_(roster), start: start,
    added: outFolder.getName().indexOf('Added ') === 0
  };

  let made = 0, flagged = 0, timedOut = false;
  const allIds = [];
  const it = inbox.getFiles();
  while (it.hasNext()) allIds.push(it.next().getId());

  // Pictures (a Canva deck saved as 1.png, 2.png...) become ONE PDF per student
  const groups = {}, groupOrder = [], ids = [];
  allIds.forEach(id => {
    const file = DriveApp.getFileById(id);
    const ext = extOf_(file.getName()).toLowerCase();
    if (IMG_EXT.indexOf(ext) === -1) { ids.push(id); return; }
    const s = matchStudent_(file.getName(), roster, '');
    const key = s ? 'S:' + studentKey_(s) : 'U:' + groupStem_(file.getName());
    if (!groups[key]) { groups[key] = { student: s, files: [] }; groupOrder.push(key); }
    groups[key].files.push(file);
  });

  for (let g = 0; g < groupOrder.length; g++) {
    if (Date.now() - start > TIME_LIMIT_MS) { timedOut = true; break; }
    const grp = groups[groupOrder[g]];
    const code = grp.student ? codeMap[studentKey_(grp.student)] : unmatchedCode_(run, grp.files[0].getName());
    let outName = null;
    const notes = [];
    const tryId = grp.files[0].getId();
    if (bumpTries_(tryId) > MAX_TRIES) {   // version 20: set aside a group that keeps stopping Google
      flagged++;
      log_(run, code, 'No file made', TOO_SLOW + grp.files.map(x => x.getName()).join(', '));
      grp.files.forEach(file => { try { file.moveTo(doneFolder); } catch (e) { /* leave it */ } });
      clearTries_(tryId);
      continue;
    }
    if (!grp.student) notes.push('Could not match these pictures to a student. First file: ' + grp.files[0].getName());
    try {
      const r = combineImages_(grp.files, code, ctx);
      outName = r.outName;
      notes.push(grp.files.length + ' picture(s) combined into one PDF. Check them by eye for names, faces or handwriting.');
      if (r.skipped.length) notes.push('Could not add: ' + r.skipped.join(', '));
    } catch (e) {
      notes.push('Error combining pictures: ' + e.message);
    }
    if (outName) made++;
    flagged++;
    if (bumpKey_(run, code, outName, notes.join(' ')) > 0 && ctx.added) notes.push(ALREADY_HAD);
    log_(run, code, outName ? 'Saved ' + outName : 'No file made', notes.join(' '));
    grp.files.forEach(file => { try { file.moveTo(doneFolder); } catch (e) { /* leave it */ } });
    clearTries_(tryId);
  }

  // Every other file, one at a time
  for (let i = 0; i < ids.length && !timedOut; i++) {
    if (Date.now() - start > TIME_LIMIT_MS) { timedOut = true; break; }
    const file = DriveApp.getFileById(ids[i]);
    const startedAt = Date.now();
    const tries = bumpTries_(ids[i]);   // version 20
    if (tries > MAX_TRIES) {
      const s = matchStudent_(file.getName(), roster, '');
      const code = s ? codeMap[studentKey_(s)] : unmatchedCode_(run, file.getName());
      flagged++;
      bumpKey_(run, code, null, 'Set aside: too slow to clean.');
      log_(run, code, 'No file made', TOO_SLOW + file.getName());
      try { file.moveTo(doneFolder); } catch (e) { /* leave it */ }
      clearTries_(ids[i]);
      continue;
    }
    ctx.lastTry = tries >= MAX_TRIES;   // last try: PDFs are rebuilt from their text in one step
    let result;
    try {
      result = await processFile_(file, ctx);
    } catch (e) {
      if (e && e.message === OUT_OF_TIME) {
        if (startedAt - start > 60 * 1000) undoTry_(ids[i]);   // it only got the end of a run: not its fault
        timedOut = true; break;   // file stays in the Inbox for the next round
      }
      result = { code: '', outName: null,
        note: 'Error: ' + e.message + ' (original is in Processed originals: ' + file.getName() + ')' };
    }
    if (result.outName) made++;
    if (result.note) flagged++;
    if (result.code && bumpKey_(run, result.code, result.outName, result.note) > 0 && ctx.added) {
      result.note = (result.note ? result.note + ' ' : '') + ALREADY_HAD;
    }
    log_(run, result.code || '?', result.outName ? 'Saved ' + result.outName : 'No file made', result.note || '');
    try { file.moveTo(doneFolder); } catch (e) { /* leave it */ }
    clearTries_(ids[i]);
  }

  const totalMade = (Number(props.getProperty(madeKey)) || 0) + made;
  const totalFlagged = (Number(props.getProperty(flagKey)) || 0) + flagged;
  const finished = !timedOut && !hasFiles_(inbox);
  if (finished) {
    [outKey, madeKey, flagKey].forEach(k => props.deleteProperty(k));
    if (inbox.getId() !== f.inbox.getId()) { try { inbox.setTrashed(true); } catch (e) { /* ignore */ } }   // empty assignment folder
  } else {
    props.setProperties({ [madeKey]: String(totalMade), [flagKey]: String(totalFlagged) });
  }
  return { run: run, made: totalMade, flagged: totalFlagged, timedOut: timedOut, finished: finished,
           added: ctx.added, outName: outFolder.getName() };
}

/* ========================= PAPER TESTS (version 11) ================ */
/*
 * Scan the whole class stack into "1 Inbox/<test> (scans)". The tool:
 *  1. (auto-run or Process the Inbox) splits the scan into single pages,
 *     kept in "3 Processed originals/<test>/_pages (has names)", and reads
 *     each page with Google's text reading. Every page read is saved on the
 *     hidden "Scan Pages" tab, so a stopped run picks up where it left off.
 *  2. (Paper tests > Match names and remove them) a step-by-step window in
 *     your school account: check where each test starts, drag a line under
 *     the name area at the top, pick who each test belongs to. Your browser
 *     then paints the name area white and saves the page as a flat picture,
 *     so the name is really gone (not just covered).
 *  3. (auto-run or Process the Inbox) builds one PDF per student, named by
 *     code, in "2 Ready for Claude/<test>". Before saving, each cleaned page
 *     is read again; if a name still shows up, that test is held back in
 *     "3 Processed originals/<test>/Check these first" and NOT sent to Claude.
 */

const SCAN_SUFFIX = ' (scans)';
const TAB_SCANS = 'Scan Pages';
const TAB_PLAN = 'Scan Plan';
const HELD_FOLDER = 'Check these first (NOT sent to Claude)';
const PAGES_FOLDER = '_pages (has names)';
const SCAN_IMG_EXT = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'heic', 'webp'];
const SCAN_HEADERS = ['Assignment', 'Seq', 'Page file ID', 'From file', 'From page', 'Words', 'Guess (Canvas ID)', 'Guess score', 'Name found', 'Read at'];
const PLAN_HEADERS = ['Assignment', 'Test', 'Code', 'Canvas ID', 'Pages', 'Clean pages', 'Status', 'Saved at'];

function isScanFolder_(name) { return String(name).slice(-SCAN_SUFFIX.length) === SCAN_SUFFIX; }
function runFromScanFolder_(name) { return String(name).slice(0, -SCAN_SUFFIX.length).trim(); }

function makeScanFolder() {
  const title = 'Scan folder for a paper test';
  try {
    showCanvasPicker_('scans', title, 'Which test or quiz? Pick it from your Canvas list so the grades can go back to the right assignment later.');
  } catch (e) {
    showNamePicker_('scans', title, 'Which test or quiz? Pick an earlier one or type a new name (use the same name as the Canvas assignment).', true);
  }
}

/** Called by the picker window. Returns { message, links } for the window to show. */
function paperTestChose(action, value) {
  value = String(value || '').trim();
  if (!value) throw new Error('Choose a test from the list first.');
  if (action !== 'scans') throw new Error('Unknown choice. Close this window and try again.');
  let run;
  if (/^(saved|canvas):/.test(value) || /courses\/\d+\/assignments\/\d+/.test(value)) run = resolveCanvasAssignment_(value).name;
  else run = resolveAssignmentName_(value);
  if (!run) throw new Error('Choose a test from the list, or type a name.');
  return scanFolderMessage_(run);
}

function scanFolderMessage_(run) {
  const fo = folderByName_(folders_().inbox, run + SCAN_SUFFIX);
  return {
    message: 'Put the scan of the whole class for "' + run + '" in this folder. One big PDF is fine (several PDFs or phone pictures work too).\n' +
      (autoRunIsOn_() ? 'Auto-run reads it within about 10 minutes.' : 'Then click Anonymizer > 2. Process the Inbox.') +
      ' When the Log says the pages are read, open Anonymizer > Paper tests > Match names and remove them.',
    links: [{ label: 'Open the scan folder', url: fo.getUrl() }]
  };
}

function ensureScanTab_() {
  const sh = ensureSheet_(TAB_SCANS, SCAN_HEADERS);
  try { if (!sh.isSheetHidden()) sh.hideSheet(); } catch (e) { /* ignore */ }
  return sh;
}
function ensurePlanTab_() {
  const sh = ensureSheet_(TAB_PLAN, PLAN_HEADERS);
  try { if (!sh.isSheetHidden()) sh.hideSheet(); } catch (e) { /* ignore */ }
  return sh;
}

/** Google's reading of one page. Returns the text, or null if Google could not read it. */
function ocrBytes_(blob, ctx) {
  let id = null;
  try {
    id = importAs_(blob, MimeType.GOOGLE_DOCS, ctx.temp.getId(), true);
    return DocumentApp.openById(id).getBody().getText();
  } catch (e) {
    return null;
  } finally {
    if (id) trashQuietly_(id);
  }
}

/** The distinct words (4+ letters) on a page, used to tell page 1 of a test from the others. */
function pageWords_(text) {
  const seen = {}, out = [];
  String(text || '').toLowerCase().split(/[^a-z]+/).forEach(w => {
    if (w.length >= 4 && !seen[w]) { seen[w] = true; out.push(w); }
  });
  return out.slice(0, 400);
}

/**
 * Best guess at whose handwriting-read name is on a page. Returns { id, score }.
 * Full name = 3, last name = 2, first name = 1; a name one letter off (5+ letters) counts half.
 * Only a clear winner (1.5 or more, ahead of the next by 1+) counts as a guess. First name alone never does.
 */
function guessStudent_(text, roster) {
  const toks = String(text || '').split(/[^A-Za-zÀ-ɏ'-]+/).map(letters_).filter(t => t.length >= 2);
  const set = {};
  toks.forEach(t => { set[t] = true; });
  const near = (w) => {
    if (!w) return 0;
    if (set[w]) return 1;
    if (w.length < 5) return 0;
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i];
      if (t[0] === w[0] && Math.abs(t.length - w.length) <= 1 && editDistance_(t, w, 1) <= 1) return 0.5;
    }
    return 0;
  };
  const scored = roster.map(s => {
    const first = letters_(s.first.split(/\s+/)[0]), last = letters_(s.last);
    let score = 2 * near(last) + 1 * near(first);
    for (let i = 0; i + 1 < toks.length; i++) if (toks[i] === first && toks[i + 1] === last) { score = 3; break; }
    s.extras.forEach(x => { if (near(letters_(x)) === 1) score = Math.max(score, 1); });
    return { id: s.id, score: score };
  }).sort((a, b) => b.score - a.score);
  if (!scored.length || scored[0].score < 1.5) return { id: '', score: scored.length ? scored[0].score : 0 };
  if (scored.length > 1 && scored[0].score - scored[1].score < 1) return { id: '', score: scored[0].score };
  return scored[0];
}

/** Puts a picture on a letter-size page of a new one-page PDF. */
async function imagePagePdf_(PDFLib, blob, ext) {
  const out = await PDFLib.PDFDocument.create();
  let img;
  if (ext === 'jpg' || ext === 'jpeg') img = await out.embedJpg(new Uint8Array(blob.getBytes()));
  else if (ext === 'png') img = await out.embedPng(new Uint8Array(blob.getBytes()));
  else img = await out.embedPng(new Uint8Array(blob.getAs('image/png').getBytes()));
  const landscape = img.width > img.height;
  const W = landscape ? 792 : 612, H = landscape ? 612 : 792;
  const s = Math.min(W / img.width, H / img.height);
  const page = out.addPage([W, H]);
  page.drawImage(img, { x: (W - img.width * s) / 2, y: (H - img.height * s) / 2, width: img.width * s, height: img.height * s });
  return out.save();
}

/**
 * Step 1: splits one test's scans into single pages and reads each one.
 * Returns { run, made, flagged, held, timedOut, finished, added, outName, progress }.
 */
async function processScanBatch_(run, inbox, f, roster, start) {
  try { ss_().getSheetByName(TAB_LOG).setColumnWidth(3, 90); } catch (e) { /* ignore */ }
  const doneFolder = folderByName_(f.done, run);
  const pagesFolder = folderByName_(doneFolder, PAGES_FOLDER);
  const base = { run: run, made: 0, flagged: 0, held: 0, added: false, outName: run, timedOut: false, finished: false };
  unzipAll_(inbox, doneFolder);
  const codeMap = ensureCodes_(run, roster);
  const ctx = { rules: buildRules_(roster, codeMap), fuzzy: { singles: [], pairs: buildFuzzy_(roster).pairs }, temp: f.temp, start: start };

  const list = [];
  const it = inbox.getFiles();
  while (it.hasNext()) {
    const file = it.next();
    const ext = extOf_(file.getName()).toLowerCase();
    if (ext === 'pdf' || SCAN_IMG_EXT.indexOf(ext) > -1) { list.push(file); continue; }
    log_(run, '-', 'Not a scan', 'Skipped ' + file.getName() + ' (only PDFs and pictures are read). Moved to Processed originals.');
    try { file.moveTo(doneFolder); } catch (e) { /* leave it */ }
  }
  list.sort((a, b) => naturalCompare_(a.getName(), b.getName()));
  if (!list.length) return Object.assign(base, { finished: true });

  const PDFLib = loadPdfLib_();
  const sh = ensureScanTab_();
  const seen = {};
  let seq = 0;
  sh.getDataRange().getValues().slice(1).forEach(r => {
    if (r[0] !== run) return;
    seen[String(r[3]) + ':' + r[4]] = true;
    seq = Math.max(seq, Number(r[1]) || 0);
  });
  let read = 0;
  for (let i = 0; i < list.length; i++) {
    const file = list[i];
    const ext = extOf_(file.getName()).toLowerCase();
    let src = null, count = 1;
    if (ext === 'pdf') {
      try {
        src = await PDFLib.PDFDocument.load(new Uint8Array(file.getBlob().getBytes()), { ignoreEncryption: true, updateMetadata: false });
        count = src.getPageCount();
      } catch (e) {
        log_(run, '-', 'Scan could not be opened', file.getName() + ': ' + e.message + '. Moved to Processed originals. Scan it again.');
        try { file.moveTo(doneFolder); } catch (e2) { /* leave it */ }
        continue;
      }
    }
    for (let p = 1; p <= count; p++) {
      if (seen[file.getId() + ':' + p]) continue;
      if (Date.now() - start > TIME_LIMIT_MS) {
        return Object.assign(base, { timedOut: true, progress: (seq) + ' page(s) read so far' });
      }
      let bytes;
      if (src) {
        const one = await PDFLib.PDFDocument.create();
        const [pg] = await one.copyPages(src, [p - 1]);
        pg.node.delete(PDFLib.PDFName.of('Annots'));
        one.addPage(pg);
        bytes = await one.save();
      } else {
        bytes = await imagePagePdf_(PDFLib, file.getBlob(), ext);
      }
      seq++;
      const pageFile = pagesFolder.createFile(u8ToBlob_(bytes, seq + '.pdf'));
      const text = ocrBytes_(u8ToBlob_(bytes, 'page.pdf'), ctx);
      const g = guessStudent_(text, roster);
      sh.appendRow([run, seq, pageFile.getId(), file.getId(), p, pageWords_(text).join(' '), g.id, g.score,
        text === null ? '?' : (hasIdentifier_(text, ctx) ? 'Y' : ''), new Date()]);
      seen[file.getId() + ':' + p] = true;
      read++;
    }
    try { file.moveTo(doneFolder); } catch (e) { /* leave it */ }   // the scan has names: PRIVATE folder
  }
  log_(run, '-', 'Scans read', seq + ' page(s) read. Next: Anonymizer > Paper tests > Match names and remove them.');
  return Object.assign(base, { finished: true, progress: seq + ' page(s) read', scansRead: seq });
}

/* ---------- the Match window (all of this runs in your school account) ---------- */

function openNameMatcher() {
  const html = HtmlService.createHtmlOutput(matchHtml_()).setWidth(1180).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Match names and remove them');
}

/** Pages read for this test that are not part of a saved test yet. */
function unplannedPages_(run) {
  const planned = {};
  const plan = ss_().getSheetByName(TAB_PLAN);
  if (plan) plan.getDataRange().getValues().slice(1).forEach(r => {
    if (r[0] === run) String(r[4]).split(',').forEach(s => { if (s) planned[String(s).trim()] = true; });
  });
  const sh = ss_().getSheetByName(TAB_SCANS);
  if (!sh) return [];
  return sh.getDataRange().getValues().slice(1).filter(r => r[0] === run && !planned[String(r[1])])
    .map(r => ({ seq: Number(r[1]), fileId: String(r[2]), words: String(r[5] || ''), guess: String(r[6] || ''),
                 score: Number(r[7]) || 0, name: String(r[8] || '') }))
    .sort((a, b) => a.seq - b.seq);
}

/** Tests with pages ready to match (and none still waiting in the scan folder). */
function matchList() {
  const sh = ss_().getSheetByName(TAB_SCANS);
  if (!sh) return [];
  const runs = {};
  sh.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) runs[String(r[0])] = true; });
  const inbox = folders_().inbox;
  return Object.keys(runs).map(run => {
    let waiting = false;
    const it = inbox.getFoldersByName(run + SCAN_SUFFIX);
    while (it.hasNext()) if (hasFiles_(it.next())) waiting = true;
    return { run: run, pages: unplannedPages_(run).length, waiting: waiting };
  }).filter(x => x.pages > 0);
}

/** Similarity of two pages' word lists (shared words / all words). */
function wordSim_(a, b) {
  const A = {}, B = {};
  a.split(' ').forEach(w => { if (w) A[w] = true; });
  b.split(' ').forEach(w => { if (w) B[w] = true; });
  let both = 0, all = 0;
  Object.keys(A).forEach(w => { all++; if (B[w]) both++; });
  Object.keys(B).forEach(w => { if (!A[w]) all++; });
  return all ? both / all : 0;
}

/**
 * Where each test starts. A page that looks like the stack's first page (the printed words of page 1)
 * starts a new test. The cut-off is the biggest gap in how alike the pages are.
 * Returns the indexes (into pages) where a test starts.
 */
function suggestStarts_(pages) {
  if (!pages.length) return [];
  const sims = pages.map((p, i) => (i === 0 ? 1 : wordSim_(pages[0].words, p.words)));
  const vals = sims.slice(1).filter((s, i) => pages[i + 1].words.split(' ').length >= 5).sort((a, b) => a - b);
  let cut = null;
  if (vals.length) {
    let gap = 0;
    for (let i = 1; i < vals.length; i++) if (vals[i] - vals[i - 1] > gap) { gap = vals[i] - vals[i - 1]; cut = (vals[i] + vals[i - 1]) / 2; }
    if (gap < 0.1) cut = vals[0] >= 0.3 ? 0 : null;   // all alike: a one-page quiz (every page starts one); none alike: no guess
  }
  if (cut === null) return [0];
  const starts = [];
  sims.forEach((s, i) => { if (i === 0 || (s > cut && pages[i].words.split(' ').length >= 5)) starts.push(i); });
  return starts;
}

function courseIdFor_(run) {
  const sh = ss_().getSheetByName(TAB_CANVAS);
  if (!sh) return '';
  const row = sh.getDataRange().getValues().slice(1).filter(r => String(r[0]) === run)[0];
  return row ? String(row[2]) : '';
}

/** Everything the window needs for one test: pages, a suggested split, and the class list. */
function matchData(run) {
  const pages = unplannedPages_(run);
  const roster = readRoster_();
  let students = roster;
  const course = courseIdFor_(run);
  if (course) {
    try {
      const ids = {};
      canvasGetAll_('/api/v1/courses/' + course + '/users?enrollment_type[]=student').forEach(u => { ids[String(u.id)] = true; });
      const inCourse = roster.filter(s => ids[s.id]);
      if (inCourse.length) students = inCourse;
    } catch (e) { /* whole roster */ }
  }
  const taken = {};
  const plan = ss_().getSheetByName(TAB_PLAN);
  if (plan) plan.getDataRange().getValues().slice(1).forEach(r => { if (r[0] === run && r[3]) taken[String(r[3])] = true; });
  return {
    run: run,
    pages: pages.map(p => ({ seq: p.seq, fileId: p.fileId, guess: p.guess, name: p.name, blank: p.words.split(' ').filter(Boolean).length < 5 })),
    starts: suggestStarts_(pages),
    students: students.slice().sort((a, b) => a.raw.localeCompare(b.raw))
      .map(s => ({ id: s.id, label: s.raw, taken: !!taken[s.id] }))
  };
}

/** One page's PDF, for the window to draw. Only pages in a "_pages" folder of this tool. */
function matchPage(fileId) {
  const file = DriveApp.getFileById(String(fileId));
  const parents = file.getParents();
  let ok = false;
  while (parents.hasNext()) if (parents.next().getName() === PAGES_FOLDER) ok = true;
  if (!ok || !/^\d+(-clean)?\.(pdf|jpg)$/.test(file.getName())) throw new Error('That page is not part of a scanned test.');
  return Utilities.base64Encode(file.getBlob().getBytes());
}

/**
 * Saves one test from the window: who it is, its pages, and the cleaned pictures of the pages with the name area blanked.
 * test: { pages: [seq...], studentId: '<Canvas ID>' | 'hold', clean: [{ seq, jpg: base64 }] }
 */
function matchSave(run, test) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const pages = unplannedPages_(run);
    const bySeq = {};
    pages.forEach(p => { bySeq[p.seq] = p; });
    const seqs = (test.pages || []).map(Number);
    if (!seqs.length || seqs.some(s => !bySeq[s])) throw new Error('Some of these pages were already saved. Close the window and open it again.');
    let code = '', canvasId = '';
    if (test.studentId && test.studentId !== 'hold') {
      const roster = readRoster_();
      const s = roster.filter(x => x.id === String(test.studentId))[0];
      if (!s) throw new Error('That student is not on the Roster.');
      code = ensureCodes_(run, roster)[studentKey_(s)];
      canvasId = s.id;
    }
    const pagesFolder = folderByName_(folderByName_(folders_().done, run), PAGES_FOLDER);
    const clean = [];
    (test.clean || []).forEach(c => {
      if (seqs.indexOf(Number(c.seq)) < 0) return;
      const blob = Utilities.newBlob(Utilities.base64Decode(String(c.jpg)), 'image/jpeg', c.seq + '-clean.jpg');
      clean.push(c.seq + '=' + pagesFolder.createFile(blob).getId());
    });
    const sh = ensurePlanTab_();
    const n = sh.getDataRange().getValues().slice(1).filter(r => r[0] === run).length + 1;
    sh.appendRow([run, n, code, canvasId, seqs.join(','), clean.join(';'), 'Ready to save', '']);
    return { ok: true, code: code };
  } finally {
    lock.releaseLock();
  }
}

/** Last step of the window: try to build the PDFs now. (If this window can't, auto-run or Process the Inbox will.) */
async function matchBuild() {
  const f = folders_();
  const r = await buildScanPlans_(f, Date.now());
  return { message: r.text, done: r.done };
}

/**
 * Step 3: builds one PDF per saved test. Cleaned pages go in as the flat pictures; the rest as scanned.
 * Each cleaned picture is read again: a name still there holds the test back.
 */
async function buildScanPlans_(f, start) {
  const plan = ss_().getSheetByName(TAB_PLAN);
  if (!plan) return { done: 0, text: '' };
  const rows = plan.getDataRange().getValues();
  const todo = [];
  for (let i = 1; i < rows.length; i++) if (rows[i][6] === 'Ready to save') todo.push(i + 1);
  if (!todo.length) return { done: 0, text: '' };
  const PDFLib = loadPdfLib_();
  const roster = readRoster_();
  const scanRows = ss_().getSheetByName(TAB_SCANS).getDataRange().getValues().slice(1);
  const props = PropertiesService.getDocumentProperties();
  const tz = Session.getScriptTimeZone();
  const counts = {};
  let done = 0, stopped = false;
  for (let k = 0; k < todo.length; k++) {
    if (Date.now() - start > TIME_LIMIT_MS) { stopped = true; break; }
    const rowNo = todo[k];
    const r = plan.getRange(rowNo, 1, 1, PLAN_HEADERS.length).getValues()[0];
    const run = String(r[0]), code = String(r[2]);
    const c = counts[run] || (counts[run] = { made: 0, held: 0 });
    const doneFolder = folderByName_(f.done, run);

    // Where this test's clean file goes: Ready/<run>, or an Added folder for late/make-up tests
    const outKey = 'SOUT::' + run;
    let outFolder = null;
    try { outFolder = DriveApp.getFolderById(props.getProperty(outKey)); if (outFolder.isTrashed()) outFolder = null; } catch (e) { outFolder = null; }
    if (!outFolder) {
      const main = folderByName_(f.ready, run);
      const already = hasWork_(main);
      outFolder = already ? main.createFolder('Added ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd') + ' (send these)') : main;
      props.setProperty(outKey, outFolder.getId());
    }
    const added = outFolder.getName().indexOf('Added ') === 0;
    const codeMap = ensureCodes_(run, roster);
    const ctx = { rules: buildRules_(roster, codeMap), fuzzy: { singles: [], pairs: buildFuzzy_(roster).pairs }, temp: f.temp };

    const pageFile = {};
    const nameOn = {};
    scanRows.forEach(s => { if (s[0] === run) { pageFile[String(s[1])] = String(s[2]); nameOn[String(s[1])] = String(s[8]); } });
    const clean = {};
    String(r[5] || '').split(';').forEach(x => { const m = x.split('='); if (m[1]) clean[m[0]] = m[1]; });
    const seqs = String(r[4]).split(',').filter(Boolean);

    const notes = [];
    let hold = !code;
    const out = await PDFLib.PDFDocument.create();
    for (let i = 0; i < seqs.length; i++) {
      const s = seqs[i];
      if (code && clean[s]) {
        const jpg = DriveApp.getFileById(clean[s]).getBlob();
        const again = ocrBytes_(jpg.setContentType('image/jpeg'), ctx);
        if (again === null || hasIdentifier_(again, ctx)) { hold = true; notes.push('Page ' + (i + 1) + ': a name may still be there after blanking (it may be written lower on the page).'); }
        const img = await out.embedJpg(new Uint8Array(jpg.getBytes()));
        const landscape = img.width > img.height;
        const W = landscape ? 792 : 612, H = landscape ? 612 : 792;
        const sc = Math.min(W / img.width, H / img.height);
        out.addPage([W, H]).drawImage(img, { x: (W - img.width * sc) / 2, y: (H - img.height * sc) / 2, width: img.width * sc, height: img.height * sc });
      } else {
        if (code && nameOn[s] === 'Y') { hold = true; notes.push('Page ' + (i + 1) + ' has a name and was not blanked.'); }
        if (code && nameOn[s] === '?') { hold = true; notes.push('Google could not read page ' + (i + 1) + ', so it could not be checked for a name.'); }
        const src = await PDFLib.PDFDocument.load(new Uint8Array(DriveApp.getFileById(pageFile[s]).getBlob().getBytes()));
        const [p] = await out.copyPages(src, [0]);
        out.addPage(p);
      }
    }
    out.setTitle(code || 'Scanned test'); out.setAuthor(''); out.setSubject(''); out.setKeywords([]);
    out.setCreator('Student Work Anonymizer'); out.setProducer('Student Work Anonymizer');
    const bytes = await out.save();
    const target = hold ? folderByName_(doneFolder, HELD_FOLDER) : outFolder;
    const outName = nextOutName_(code || ('Not matched (test ' + r[1] + ')'), '.pdf', target);
    target.createFile(u8ToBlob_(bytes, outName));
    if (!code) notes.push('You chose "not sure" for this test. Grade it by hand (it still has the name).');
    else if (hold) notes.push('Held back, NOT sent to Claude. Open it in "' + HELD_FOLDER + '": if no name shows, move it into "2 Ready for Claude/' + run + '".');
    if (code && bumpKey_(run, code, outName, notes.join(' ')) > 0 && added) notes.push(ALREADY_HAD);
    log_(run, code || '?', (hold ? 'Held ' : 'Saved ') + outName, (seqs.length + ' page(s). ' + notes.join(' ')).trim());
    plan.getRange(rowNo, 7, 1, 2).setValues([[hold ? 'Held' : 'Saved', Utilities.formatDate(new Date(), tz, 'MMM d, h:mm a')]]);
    if (hold) c.held++; else c.made++;
    done++;
  }

  // A test whose pages are all saved: remove the single pages (they have names) and the working rows
  Object.keys(counts).forEach(run => {
    const left = unplannedPages_(run).length;
    const waiting = plan.getDataRange().getValues().slice(1).some(x => x[0] === run && x[6] === 'Ready to save');
    if (left || waiting) return;
    const doneFolder = folderByName_(f.done, run);
    const it = doneFolder.getFoldersByName(PAGES_FOLDER);
    while (it.hasNext()) deleteForever_(it.next().getId());
    deleteRows_(ss_().getSheetByName(TAB_SCANS), x => x[0] === run, 2);
    deleteRows_(plan, x => x[0] === run, 2);
    props.deleteProperty('SOUT::' + run);
  });
  const text = Object.keys(counts).map(run => run + ': ' + counts[run].made + ' test(s) saved in "2 Ready for Claude"' +
    (counts[run].held ? ', ' + counts[run].held + ' held back in "3 Processed originals/' + run + '/' + HELD_FOLDER + '" (the Log tab says why)' : '') + '.').join('\n') +
    (stopped ? '\nNot finished (Google time limit). Auto-run or Process the Inbox will finish the rest.' : '');
  return { done: done, text: text, stopped: stopped, counts: counts };
}

function matchHtml_() {
  return '<!doctype html><html><head><meta charset="utf-8">' +
    '<script src="' + PDFJS_URL + '"></scr' + 'ipt>' +
    '<script>var PDFJS_WORKER=' + JSON.stringify(PDFJS_WORKER_URL) + ';</scr' + 'ipt></head>' +
    MATCH_PAGE_HTML + '</html>';
}

const MATCH_PAGE_HTML = `
<body>
<style>
  :root{--fg:#1f2933;--muted:#5f6b7a;--line:#dde3ea;--soft:#f4f6f9;--blue:#1a73e8;--blue-soft:#e8f0fe;--green:#137333;--green-soft:#e6f4ea;--red:#b3261e;--amber-soft:#fef3e0}
  *{box-sizing:border-box}
  body{margin:0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:var(--fg);background:#fff;height:100vh;display:flex;flex-direction:column}
  .steps{display:flex;gap:6px;padding:10px 16px;border-bottom:1px solid var(--line);align-items:center;flex-wrap:wrap}
  .steps span{padding:5px 10px;border-radius:14px;background:var(--soft);color:var(--muted);font-size:13px}
  .steps span.on{background:var(--blue);color:#fff;font-weight:bold}
  .steps span.done{background:var(--green-soft);color:var(--green)}
  .steps .run{margin-left:auto;font-weight:bold;color:var(--fg);background:none}
  .body{flex:1;overflow:auto;padding:16px 20px;min-height:0}
  .foot{display:flex;gap:10px;align-items:center;padding:10px 16px;border-top:1px solid var(--line)}
  .foot .msg{flex:1;color:var(--muted)}
  .foot .msg.err{color:var(--red);font-weight:bold}
  button{font-size:14px;padding:8px 14px;border-radius:6px;border:1px solid var(--line);background:#fff;cursor:pointer}
  button.go{background:var(--blue);color:#fff;border-color:var(--blue);font-weight:bold}
  button:disabled{opacity:.5;cursor:default}
  h2{font-size:18px;margin:0 0 6px}
  p.lead{margin:0 0 14px;color:var(--muted);max-width:820px;line-height:1.45}
  .grid{display:flex;flex-wrap:wrap;gap:10px}
  .thumb{width:124px;border:2px solid var(--line);border-radius:6px;padding:4px;cursor:pointer;background:#fff;position:relative;text-align:center}
  .thumb canvas{width:112px;display:block;margin:0 auto;background:var(--soft);min-height:140px}
  .thumb .cap{font-size:11px;color:var(--muted);margin-top:3px}
  .thumb.start{border-color:var(--blue);background:var(--blue-soft)}
  .thumb.start .cap{color:var(--blue);font-weight:bold}
  .thumb .flag{position:absolute;top:6px;right:6px;background:#fff3cd;color:#6b4e00;font-size:10px;padding:1px 4px;border-radius:3px}
  .row{display:flex;gap:16px;align-items:flex-start;flex-wrap:wrap}
  .opts{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:14px}
  .opts input[type=number]{width:60px;font-size:14px;padding:5px}
  .big{position:relative;width:560px;max-width:100%;border:1px solid var(--line);background:var(--soft)}
  .big canvas{width:100%;display:block}
  .shade{position:absolute;left:0;right:0;top:0;background:rgba(255,255,255,.88);border-bottom:3px solid var(--red);pointer-events:none}
  .handle{position:absolute;left:0;right:0;height:22px;margin-top:-11px;cursor:ns-resize}
  .handle b{position:absolute;right:6px;top:2px;background:var(--red);color:#fff;font-size:11px;padding:2px 6px;border-radius:3px}
  .side{flex:1;min-width:260px}
  .mini{display:flex;flex-wrap:wrap;gap:8px}
  .mini>div{position:relative;width:150px;border:1px solid var(--line)}
  .mini canvas{width:100%;display:block}
  .who{display:grid;grid-template-columns:70px 1fr 300px;gap:10px 14px;align-items:center}
  .who .crop{border:1px solid var(--line);background:var(--soft);min-height:40px}
  .who .crop canvas{width:100%;display:block}
  .who select{font-size:14px;padding:6px;width:100%}
  .who .n{font-weight:bold}
  .who .n small{display:block;font-weight:normal;color:var(--muted)}
  .warn{color:var(--red);font-size:12px;margin-top:3px}
  .guess{color:var(--green);font-size:12px;margin-top:3px}
  .bar{height:10px;background:var(--soft);border-radius:5px;overflow:hidden;max-width:520px;margin:12px 0}
  .bar i{display:block;height:100%;width:0;background:var(--green);transition:width .3s}
  .box{background:var(--soft);border-radius:8px;padding:14px 16px;max-width:760px;line-height:1.5;white-space:pre-wrap}
  .center{padding:60px;text-align:center;color:var(--muted)}
</style>
<div class="steps" id="steps"></div>
<div class="body" id="body"><div class="center">Loading...</div></div>
<div class="foot"><span class="msg" id="msg"></span><button id="back" type="button">Back</button><button id="next" class="go" type="button">Next</button></div>
<script>
var $ = function(id){ return document.getElementById(id); };
function el(tag, cls, text){ var e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; }
function say(t, bad){ $('msg').textContent = t || ''; $('msg').className = 'msg' + (bad ? ' err' : ''); }
function call(fn){ var args = Array.prototype.slice.call(arguments, 1); return new Promise(function(ok, no){ var r = google.script.run.withSuccessHandler(ok).withFailureHandler(no); r[fn].apply(r, args); }); }
if (window.pdfjsLib) { try { pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER; } catch (e) {} }

var STEPS = ['Where tests start', 'Name area', 'Who is it', 'Save'];
var S = { step: 0, run: '', pages: [], starts: {}, students: [], strip: 0.12, blankAll: false, who: {}, docs: {} };

function steps(){
  var b = $('steps'); b.innerHTML = '';
  STEPS.forEach(function(t, i){ b.appendChild(el('span', i === S.step ? 'on' : (i < S.step ? 'done' : ''), (i + 1) + '. ' + t)); });
  if (S.run) b.appendChild(el('span', 'run', S.run));
  $('back').disabled = S.step === 0;
}

/* ---------- pages ---------- */
function doc(p){
  if (S.docs[p.fileId]) return S.docs[p.fileId];
  S.docs[p.fileId] = call('matchPage', p.fileId).then(function(b64){
    var raw = atob(b64), bytes = new Uint8Array(raw.length);
    for (var i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i);
    return pdfjsLib.getDocument({ data: bytes }).promise.then(function(d){ return d.getPage(1); });
  });
  return S.docs[p.fileId];
}
/* Draws page p into canvas cv, widthPx wide. top/bottom = the part of the page to show (0 to 1). white = blank the top this much. */
function draw(p, cv, widthPx, opts){
  opts = opts || {};
  return doc(p).then(function(pg){
    var base = pg.getViewport({ scale: 1 });
    var vp = pg.getViewport({ scale: widthPx / base.width });
    var full = document.createElement('canvas'); full.width = Math.round(vp.width); full.height = Math.round(vp.height);
    var fx = full.getContext('2d'); fx.fillStyle = '#fff'; fx.fillRect(0, 0, full.width, full.height);
    return pg.render({ canvasContext: fx, viewport: vp }).promise.then(function(){
      if (opts.white) { fx.fillStyle = '#fff'; fx.fillRect(0, 0, full.width, Math.ceil(full.height * opts.white)); }
      var t = opts.top || 0, btm = opts.bottom || 1;
      var y0 = Math.floor(full.height * t), h = Math.max(1, Math.ceil(full.height * (btm - t)));
      cv.width = full.width; cv.height = h;
      cv.getContext('2d').drawImage(full, 0, y0, full.width, h, 0, 0, full.width, h);
      return full;
    });
  });
}
function queue(jobs){ var i = 0; function next(){ if (i >= jobs.length) return Promise.resolve(); var j = jobs[i++]; return j().catch(function(){}).then(next); } return Promise.all([next(), next(), next()]); }

/* ---------- tests from the starts ---------- */
function tests(){
  var out = [], cur = null;
  S.pages.forEach(function(p, i){
    if (S.starts[i] || !cur) { cur = { first: i, pages: [] }; out.push(cur); }
    cur.pages.push(i);
  });
  return out;
}

/* ---------- step 0: pick the test, then load ---------- */
call('matchList').then(function(list){
  if (!list.length) { $('body').innerHTML = ''; $('body').appendChild(el('div', 'center', 'No scanned tests are ready yet. Put the class scan in the test’s scan folder (Anonymizer > Paper tests > Make a scan folder), wait for auto-run or click Process the Inbox, then open this again.')); $('next').disabled = true; $('back').disabled = true; return; }
  if (list.length === 1) return load(list[0].run);
  var b = $('body'); b.innerHTML = '';
  b.appendChild(el('h2', '', 'Which test?'));
  var sel = el('select'); sel.style.cssText = 'font-size:15px;padding:6px;min-width:320px';
  list.forEach(function(x){ var o = el('option', '', x.run + ' (' + x.pages + ' pages)' + (x.waiting ? ', more scans still being read' : '')); o.value = x.run; sel.appendChild(o); });
  b.appendChild(sel);
  $('next').onclick = function(){ load(sel.value); };
}).catch(function(e){ say(e.message || String(e), true); });

function load(run){
  S.run = run; say('Loading the pages...');
  call('matchData', run).then(function(d){
    S.pages = d.pages; S.students = d.students; S.starts = {};
    (d.starts || [0]).forEach(function(i){ S.starts[i] = true; });
    S.guessStarts = JSON.parse(JSON.stringify(S.starts));
    say(''); go(0);
  }).catch(function(e){ say(e.message || String(e), true); });
}

function go(n){
  S.step = n; steps(); say('');
  $('next').disabled = false; $('next').textContent = n === 3 ? 'Remove names and save' : 'Next';
  [stepSplit, stepStrip, stepWho, stepSave][n]();
}
$('back').onclick = function(){ if (S.step > 0) go(S.step - 1); };

/* ---------- step 1: where each test starts ---------- */
function stepSplit(){
  var b = $('body'); b.innerHTML = '';
  b.appendChild(el('h2', '', 'Where does each student’s test start?'));
  b.appendChild(el('p', 'lead', 'The blue pages start a new test. Click a page to change it. Check that every first page is blue and nothing else is.'));
  var o = el('div', 'opts');
  var auto = el('button', '', 'Use the tool’s guess'); auto.type = 'button';
  var n = el('input'); n.type = 'number'; n.min = '1'; n.value = String(Math.max(1, Math.round(S.pages.length / Math.max(1, tests().length))));
  var every = el('button', '', 'Every'); every.type = 'button';
  o.appendChild(auto); o.appendChild(el('span', '', 'or a new test')); o.appendChild(every); o.appendChild(n); o.appendChild(el('span', '', 'pages'));
  var sum = el('strong'); sum.style.marginLeft = '12px'; o.appendChild(sum);
  b.appendChild(o);
  var g = el('div', 'grid'); b.appendChild(g);
  var cells = [];
  function paint(){
    var t = tests(), counts = {};
    t.forEach(function(x){ counts[x.pages.length] = (counts[x.pages.length] || 0) + 1; });
    var most = Object.keys(counts).sort(function(a, c){ return counts[c] - counts[a]; })[0];
    sum.textContent = t.length + ' test' + (t.length === 1 ? '' : 's') + ', most have ' + most + ' page' + (most === '1' ? '' : 's') + '.';
    var ti = 0;
    cells.forEach(function(c, i){
      if (S.starts[i] || i === 0) ti++;
      c.className = 'thumb' + (S.starts[i] || i === 0 ? ' start' : '');
      c.cap.textContent = (S.starts[i] || i === 0) ? 'Test ' + ti + ' starts' : 'page ' + (i + 1);
    });
  }
  auto.onclick = function(){ S.starts = JSON.parse(JSON.stringify(S.guessStarts)); paint(); };
  every.onclick = function(){ var k = Math.max(1, parseInt(n.value, 10) || 1); S.starts = {}; for (var i = 0; i < S.pages.length; i += k) S.starts[i] = true; paint(); };
  var jobs = [];
  S.pages.forEach(function(p, i){
    var c = el('div', 'thumb'); var cv = el('canvas'); c.appendChild(cv); c.cap = el('div', 'cap'); c.appendChild(c.cap);
    if (p.name === 'Y') c.appendChild(el('span', 'flag', 'name'));
    if (p.blank) c.appendChild(el('span', 'flag', 'blank'));
    c.onclick = function(){ if (i === 0) return; if (S.starts[i]) delete S.starts[i]; else S.starts[i] = true; paint(); };
    g.appendChild(c); cells.push(c);
    jobs.push(function(){ return draw(p, cv, 224); });
  });
  paint(); queue(jobs);
  $('next').onclick = function(){ S.starts[0] = true; go(1); };
}

/* ---------- step 2: the name area ---------- */
function stepStrip(){
  var b = $('body'); b.innerHTML = '';
  var t = tests();
  b.appendChild(el('h2', '', 'Drag the red line just below where the name is written'));
  b.appendChild(el('p', 'lead', 'Everything above the line is blanked out on the first page of every test. Leave the first question below the line. The small pages on the right show other students’ tests, so you can check their names are covered too.'));
  var row = el('div', 'row'); b.appendChild(row);
  var big = el('div', 'big'); var cv = el('canvas'); big.appendChild(cv);
  var shade = el('div', 'shade'); big.appendChild(shade);
  var handle = el('div', 'handle'); handle.appendChild(el('b', '', 'drag')); big.appendChild(handle);
  row.appendChild(big);
  var side = el('div', 'side'); row.appendChild(side);
  var lab = el('label'); var cb = el('input'); cb.type = 'checkbox'; cb.checked = S.blankAll;
  lab.appendChild(cb); lab.appendChild(el('span', '', ' Also blank the top of every other page (use this if students write their name on every page)'));
  side.appendChild(lab);
  side.appendChild(el('p', 'lead', 'Pages where the tool read a name are blanked the same way automatically. If a name is still found after blanking, that test is held back for you and not sent to Rowan.'));
  var mini = el('div', 'mini'); side.appendChild(mini);
  var minis = [];
  t.slice(1, 7).forEach(function(x){ var d = el('div'); var c2 = el('canvas'); d.appendChild(c2); var s2 = el('div', 'shade'); d.appendChild(s2); mini.appendChild(d); minis.push({ p: S.pages[x.first], cv: c2, shade: s2 }); });
  function place(){
    var pct = (S.strip * 100) + '%';
    shade.style.height = pct; handle.style.top = pct;
    minis.forEach(function(m){ m.shade.style.height = pct; });
  }
  cb.onchange = function(){ S.blankAll = cb.checked; };
  draw(S.pages[t[0].first], cv, 1100).then(place);
  queue(minis.map(function(m){ return function(){ return draw(m.p, m.cv, 300); }; }));
  place();
  var dragging = false;
  function move(e){
    if (!dragging) return;
    var r = big.getBoundingClientRect();
    var y = (e.touches ? e.touches[0].clientY : e.clientY) - r.top;
    S.strip = Math.min(0.6, Math.max(0.02, y / r.height)); place();
    e.preventDefault();
  }
  handle.onmousedown = function(e){ dragging = true; e.preventDefault(); };
  big.onmousedown = function(e){ dragging = true; move(e); };
  document.onmousemove = move; document.onmouseup = function(){ dragging = false; };
  $('next').onclick = function(){ go(2); };
}

/* ---------- step 3: who is each test ---------- */
function stepWho(){
  var b = $('body'); b.innerHTML = '';
  var t = tests();
  b.appendChild(el('h2', '', 'Whose test is each one?'));
  b.appendChild(el('p', 'lead', 'The tool read the names it could and filled those in (green). Check each one and fill in the rest. Pick “Not sure” to keep a test out and grade it by hand.'));
  var grid = el('div', 'who'); b.appendChild(grid);
  var selects = [];
  var jobs = [];
  t.forEach(function(x, k){
    var p = S.pages[x.first];
    var n = el('div', 'n', 'Test ' + (k + 1)); n.appendChild(el('small', '', x.pages.length + ' page' + (x.pages.length === 1 ? '' : 's')));
    var crop = el('div', 'crop'); var cv = el('canvas'); crop.appendChild(cv);
    var cell = el('div'); var sel = el('select');
    var o0 = el('option', '', 'Pick the student...'); o0.value = ''; sel.appendChild(o0);
    var oh = el('option', '', 'Not sure (keep it out, I’ll grade it by hand)'); oh.value = 'hold'; sel.appendChild(oh);
    S.students.forEach(function(s){ var o = el('option', '', s.label + (s.taken ? '  (already saved)' : '')); o.value = s.id; sel.appendChild(o); });
    var key = x.first;
    if (S.who[key] !== undefined) sel.value = S.who[key];
    else if (p.guess && S.students.some(function(s){ return s.id === p.guess; })) { sel.value = p.guess; S.who[key] = p.guess; }
    var note = el('div', ''); cell.appendChild(sel); cell.appendChild(note);
    sel.onchange = function(){ S.who[key] = sel.value; check(); };
    selects.push({ sel: sel, note: note, key: key, guessed: p.guess && sel.value === p.guess });
    grid.appendChild(n); grid.appendChild(crop); grid.appendChild(cell);
    jobs.push(function(){ return draw(p, cv, 900, { top: 0, bottom: Math.min(1, S.strip + 0.05) }); });
  });
  function check(){
    var used = {};
    selects.forEach(function(x){ if (x.sel.value && x.sel.value !== 'hold') used[x.sel.value] = (used[x.sel.value] || 0) + 1; });
    var missing = 0, dup = 0;
    selects.forEach(function(x){
      x.note.className = ''; x.note.textContent = '';
      if (!x.sel.value) missing++;
      else if (x.sel.value !== 'hold' && used[x.sel.value] > 1) { dup++; x.note.className = 'warn'; x.note.textContent = 'This student is picked for more than one test.'; }
      else if (x.guessed && x.sel.value === S.pages[x.key].guess) { x.note.className = 'guess'; x.note.textContent = 'Filled in from the name the tool read. Check it.'; }
    });
    $('next').disabled = missing > 0 || dup > 0;
    say(missing ? missing + ' test(s) still need a student.' : (dup ? 'One student is picked twice.' : 'All set. Next removes the names and saves.'), dup > 0);
  }
  check(); queue(jobs);
  $('next').onclick = function(){ go(3); };
}

/* ---------- step 4: blank the names and save ---------- */
function stepSave(){
  var b = $('body'); b.innerHTML = '';
  var t = tests();
  var keep = t.filter(function(x){ return S.who[x.first] && S.who[x.first] !== 'hold'; }).length;
  b.appendChild(el('h2', '', 'Ready to remove the names'));
  b.appendChild(el('p', 'lead', keep + ' test(s) will have the name area blanked and go to “2 Ready for Claude.” ' + (t.length - keep) + ' marked “Not sure” stay in your private folder. Your browser paints the name area white and saves each of those pages as a flat picture, so the name is gone for good.'));
  var bar = el('div', 'bar'); var fill = el('i'); bar.appendChild(fill); b.appendChild(bar);
  var out = el('div', 'box', ''); out.style.display = 'none'; b.appendChild(out);
  $('next').onclick = function(){
    $('next').disabled = true; $('back').disabled = true;
    var i = 0;
    function one(){
      if (i >= t.length) return finish();
      var x = t[i];
      var who = S.who[x.first];
      say('Saving test ' + (i + 1) + ' of ' + t.length + '...');
      var need = who === 'hold' ? [] : x.pages.filter(function(pi, j){ return j === 0 || S.blankAll || S.pages[pi].name === 'Y'; });
      var clean = [];
      var chain = Promise.resolve();
      need.forEach(function(pi){
        chain = chain.then(function(){
          var cv = document.createElement('canvas');
          return draw(S.pages[pi], cv, 1700, { white: S.strip }).then(function(full){
            clean.push({ seq: S.pages[pi].seq, jpg: full.toDataURL('image/jpeg', 0.85).split(',')[1] });
          });
        });
      });
      return chain.then(function(){
        return call('matchSave', S.run, { pages: x.pages.map(function(pi){ return S.pages[pi].seq; }), studentId: who, clean: clean });
      }).then(function(){ i++; fill.style.width = (100 * i / t.length) + '%'; return one(); });
    }
    function finish(){
      say('Making the coded PDFs...');
      return call('matchBuild').then(function(r){
        out.style.display = 'block';
        out.textContent = (r && r.message) ? r.message + '\\n\\nNext: tell Rowan “grade.” You can close this window.'
          : 'Saved. The coded PDFs will be made by auto-run within about 10 minutes (or click Anonymizer > 2. Process the Inbox now). Then tell Rowan “grade.”';
        say(''); $('next').textContent = 'Close'; $('next').disabled = false;
        $('next').onclick = function(){ google.script.host.close(); };
      }).catch(function(){
        out.style.display = 'block';
        out.textContent = 'Saved. Click Anonymizer > 2. Process the Inbox to make the coded PDFs (or wait for auto-run). Then tell Rowan “grade.”';
        say(''); $('next').textContent = 'Close'; $('next').disabled = false;
        $('next').onclick = function(){ google.script.host.close(); };
      });
    }
    one().catch(function(e){ say('Stopped at test ' + (i + 1) + ': ' + ((e && e.message) || e) + '. Tests before it are saved; open this window again to finish.', true); });
  };
}
</script>
</body>`;

/* ============================ AUTO-RUN ============================ */

function autoRunIsOn_() {
  return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'autoProcess');
}

function toggleAutoRun() {
  const ui = SpreadsheetApp.getUi();
  if (autoRunIsOn_()) {
    ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'autoProcess').forEach(t => ScriptApp.deleteTrigger(t));
    ui.alert('Auto-run is OFF. Use Process the Inbox when you want files cleaned.');
    return;
  }
  ScriptApp.newTrigger('autoProcess').timeBased().everyMinutes(10).create();
  ui.alert('Auto-run is ON',
    'Every 10 minutes the tool cleans any files sitting in an assignment folder inside the Inbox ' +
    '(that is where Pull submissions puts them). No clicking and no "Not done yet."\n\n' +
    'Loose files dropped straight into the Inbox still need Process the Inbox, because the tool has to ask which assignment they belong to.',
    ui.ButtonSet.OK);
}

function makeAssignmentInbox() {
  showNamePicker_('inbox', 'Make an Inbox folder',
    'Which assignment is the folder for? Pick an earlier one or type a new name.', true);
}

function makeAssignmentInboxFor_(run) {
  const ui = SpreadsheetApp.getUi();
  const fo = folderByName_(folders_().inbox, run);
  const html = HtmlService.createHtmlOutput('<div style="font-family:Arial;font-size:14px"><p>Drop files for <b>' +
    run.replace(/</g, '&lt;') + '</b> here:</p><p><a target="_blank" href="' + fo.getUrl() + '">Open the folder</a></p></div>')
    .setWidth(360).setHeight(130);
  ui.showModelessDialog(html, 'Inbox folder ready');
}

/* ============================== CANVAS ============================ */
/*
 * Optional. Uses YOUR Canvas access token, stored only in your Google account
 * (this script's private user settings). Student work goes Canvas -> your
 * school Google Drive and grades go your Google Sheet -> Canvas. Nothing here
 * sends names to Claude.
 */

const TAB_CANVAS = 'Canvas Assignments';
const TAB_PULLS = 'Canvas Pulls';

function canvasSettings_() {
  const base = PropertiesService.getDocumentProperties().getProperty('CANVAS_BASE');
  const token = PropertiesService.getUserProperties().getProperty('CANVAS_TOKEN');
  if (!base || !token) throw new Error('Canvas is not connected yet. Use Anonymizer > Canvas > Connect to Canvas first.');
  return { base: base, token: token };
}

function canvasFetch_(pathOrUrl, opts) {
  const s = canvasSettings_();
  const url = /^https?:/.test(pathOrUrl) ? pathOrUrl : s.base + pathOrUrl;
  const o = Object.assign({ muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + s.token } }, opts || {});
  const resp = UrlFetchApp.fetch(url, o);
  const code = resp.getResponseCode();
  if (code === 401) throw new Error('Canvas said the token is not valid (401) for ' + url.replace(s.base, '').split('?')[0] + '. Make a new token and use Connect to Canvas again.');
  if (code >= 400) throw new Error('Canvas error ' + code + ' for ' + url.replace(s.base, ''));
  return resp;
}

/** The "ktufsd.instructure.com" part of a web address, lowercase. */
function hostOf_(url) {
  const m = String(url || '').match(/^https?:\/\/([^\/?#]+)/i);
  return m ? m[1].toLowerCase() : '';
}

/**
 * Downloads one submitted file and returns it as a blob.
 * A Canvas file link sends you on to Canvas's file server, a different web
 * address. The token is sent ONLY to the Canvas address you connected with;
 * the file server gets the link Canvas made for it, without the token
 * (sending the token there is what caused the 401 in version 5).
 */
function canvasDownload_(url) {
  const s = canvasSettings_();
  const baseHost = hostOf_(s.base);
  let next = url;
  for (let hop = 0; hop < 8; hop++) {
    const o = { muteHttpExceptions: true, followRedirects: false };
    if (hostOf_(next) === baseHost) o.headers = { Authorization: 'Bearer ' + s.token };
    let resp = UrlFetchApp.fetch(next, o);
    let code = resp.getResponseCode();
    if ((code === 401 || code === 403) && o.headers) {   // the link may carry its own permission; try it without the token
      delete o.headers;
      resp = UrlFetchApp.fetch(next, o);
      code = resp.getResponseCode();
    }
    if (code >= 300 && code < 400) {
      const h = resp.getHeaders() || {};
      const loc = h['Location'] || h['location'];
      if (!loc) throw new Error('Canvas sent the download somewhere without saying where');
      next = /^https?:/i.test(loc) ? loc
        : (String(next).match(/^https?:\/\/[^\/?#]+/i)[0] + (loc.charAt(0) === '/' ? '' : '/') + loc);
      continue;
    }
    if (code >= 400) throw new Error('the file would not download (error ' + code + ' from ' + hostOf_(next) + ')');
    return resp.getBlob();
  }
  throw new Error('the file download was sent on too many times');
}

/** GET every page of a Canvas list. */
function canvasGetAll_(path) {
  let url = canvasSettings_().base + path + (path.indexOf('?') > -1 ? '&' : '?') + 'per_page=100';
  const out = [];
  while (url) {
    const resp = canvasFetch_(url);
    JSON.parse(resp.getContentText()).forEach(x => out.push(x));
    const link = String(resp.getHeaders()['Link'] || resp.getHeaders()['link'] || '');
    const next = link.split(',').filter(p => /rel="next"/.test(p))[0];
    url = next ? next.match(/<([^>]+)>/)[1] : null;
  }
  return out;
}

function connectCanvas() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Connect to Canvas (1 of 2)',
    'Paste the web address you use for Canvas, for example https://ktufsd.instructure.com',
    ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const m = r1.getResponseText().trim().match(/^(https?:\/\/[^\/\s]+)/i);
  if (!m) { ui.alert('That does not look like a web address. Try again.'); return; }
  const r2 = ui.prompt('Connect to Canvas (2 of 2)',
    'Paste your Canvas access token.\n\nTo make one: in Canvas click Account > Settings, scroll to "Approved Integrations," ' +
    'click "+ New Access Token," name it "Student Work Anonymizer," set an expiration date (end of the school year), and copy the token.\n\n' +
    'The token is saved only in your own Google account. Never paste it anywhere else, including chats with Claude.',
    ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  PropertiesService.getDocumentProperties().setProperty('CANVAS_BASE', m[1].toLowerCase());
  PropertiesService.getUserProperties().setProperty('CANVAS_TOKEN', r2.getResponseText().trim());
  try {
    canvasFetch_('/api/v1/users/self');
    ensureSheet_(TAB_CANVAS, ['Assignment', 'Canvas link', 'Course ID', 'Assignment ID', 'Last pulled']);
    ensureSheet_(TAB_PULLS, ['Assignment', 'Canvas ID', 'Attempt', 'Pulled']);
    ui.alert('Connected to Canvas. You can now use Pull submissions, Send grades, and Update roster.');
  } catch (e) {
    PropertiesService.getUserProperties().deleteProperty('CANVAS_TOKEN');
    ui.alert('Could not connect', e.message + '\n\nIf your district does not allow access tokens, keep using Download Submissions and the gradebook import.', ui.ButtonSet.OK);
  }
}

function disconnectCanvas() {
  PropertiesService.getUserProperties().deleteProperty('CANVAS_TOKEN');
  SpreadsheetApp.getUi().alert('Disconnected. The token was removed from your Google account. You can also delete it in Canvas (Account > Settings > Approved Integrations).');
}

/**
 * Turns the picker's choice into a Canvas assignment, saving it on the Canvas Assignments tab the first time.
 * value: "saved:<assignment id>", "canvas:<course id>:<assignment id>", or "typed:<pasted link>".
 * Returns { name, course, assignment, row }. Throws a plain message if it can't.
 */
function resolveCanvasAssignment_(value) {
  const sh = ensureSheet_(TAB_CANVAS, ['Assignment', 'Canvas link', 'Course ID', 'Assignment ID', 'Last pulled']);
  const rows = sh.getDataRange().getValues().slice(1);
  const rowOfId = id => rows.findIndex(r => r[0] && String(r[3]) === String(id));
  value = String(value || '').trim();

  let course, assignment, link = '';
  let m = value.match(/^saved:(\d+)$/);
  if (m) {
    const at = rowOfId(m[1]);
    if (at < 0) throw new Error('That assignment is no longer on the Canvas Assignments tab. Close this window and try again.');
    return { name: String(rows[at][0]), course: String(rows[at][2]), assignment: String(rows[at][3]), row: at + 2 };
  }
  m = value.match(/^canvas:(\d+):(\d+)$/);
  if (m) {
    course = m[1]; assignment = m[2];
  } else {
    link = value.replace(/^typed:/, '').trim();
    m = link.match(/courses\/(\d+)\/assignments\/(\d+)/);
    if (!m) throw new Error('That is not a Canvas assignment link. Open the assignment in Canvas and copy the whole address bar. It should look like .../courses/123/assignments/456');
    course = m[1]; assignment = m[2];
  }
  const existing = rowOfId(assignment);
  if (existing > -1) return { name: String(rows[existing][0]), course: course, assignment: assignment, row: existing + 2 };

  const info = JSON.parse(canvasFetch_('/api/v1/courses/' + course + '/assignments/' + assignment).getContentText());
  const name = String(info.name || ('Assignment ' + assignment)).replace(/[\/\\]/g, '-').trim();
  if (!link) link = canvasSettings_().base + '/courses/' + course + '/assignments/' + assignment;
  sh.appendRow([name, link, course, assignment, '']);
  sh.getRange(sh.getLastRow(), 3, 1, 2).setNumberFormat('@');
  return { name: name, course: course, assignment: assignment, row: sh.getLastRow() };
}

/** Downloads new submissions (and new attempts) for one assignment into "1 Inbox/<assignment>". */
function pullSubmissions() {
  const ui = SpreadsheetApp.getUi();
  try {
    showCanvasPicker_('pull', 'Pull submissions from Canvas', 'Which assignment do you want to pull?');
  } catch (e) {
    ui.alert('Pull submissions stopped', e.message, ui.ButtonSet.OK);
  }
}

function pullSubmissionsFor_(asg) {
  const ui = ui_();
  try {
    pullSubmissions_(ui, asg);
  } catch (e) {
    ui.alert('Pull submissions stopped', e.message, ui.ButtonSet.OK);
  }
}

function pullSubmissions_(ui, asg) {
  const start = Date.now();
  if (!asg) return;
  try { updateRoster_('courses/' + asg.course); } catch (e) { /* pull anyway; names may not match */ }   // version 21
  const f = folders_();
  const inbox = folderByName_(f.inbox, asg.name);
  const pulls = ensureSheet_(TAB_PULLS, ['Assignment', 'Canvas ID', 'Attempt', 'Pulled']);
  const seen = {};
  pulls.getDataRange().getValues().slice(1).forEach(r => { if (r[0] === asg.name) seen[r[1] + ':' + r[2]] = true; });

  const subs = canvasGetAll_('/api/v1/courses/' + asg.course + '/assignments/' + asg.assignment + '/submissions');
  let students = 0, files = 0, skipped = 0, stopped = false, noFile = 0;
  const newRows = [], problems = [];
  try {
    for (let i = 0; i < subs.length; i++) {
      const s = subs[i];
      if (!s.submitted_at || !s.user_id) continue;
      // Version 16: Google Assignments, recordings and tool links have no file Pull can download.
      // Count them (every pull) and don't mark them as pulled.
      if (!hasDownload_(s)) { noFile++; continue; }
      const key = s.user_id + ':' + (s.attempt || 1);
      if (seen[key]) { skipped++; continue; }
      if (Date.now() - start > TIME_LIMIT_MS) { stopped = true; break; }
      const tag = s.user_id + '_' + (s.attempt || 1) + '_';

      // Download everything for this student first. If one file fails, nothing is saved
      // for them and they are not marked as pulled, so the next pull tries again.
      const blobs = [];
      try {
        (s.attachments || []).forEach(att => {
          const blob = canvasDownload_(att.url).setName(tag + att.filename);
          const ext = extOf_(att.filename).toLowerCase();
          if (MIME_BY_EXT[ext]) blob.setContentType(MIME_BY_EXT[ext]);
          blobs.push(blob);
        });
      } catch (e) {
        problems.push('Canvas ID ' + s.user_id + ': ' + e.message);
        continue;
      }
      if (s.submission_type === 'online_text_entry' && s.body) {
        blobs.push(Utilities.newBlob('<html><body>' + s.body + '</body></html>', 'text/html', tag + 'text.html'));
      }
      if (s.submission_type === 'online_url' && s.url) {
        blobs.push(Utilities.newBlob('<html><body><a href="' + s.url + '">' + s.url + '</a></body></html>', 'text/html', tag + 'link.html'));
      }
      blobs.forEach(b => { inbox.createFile(b); files++; });
      newRows.push([asg.name, String(s.user_id), s.attempt || 1, new Date()]);
      students++;
    }
  } finally {
    // Always record who was pulled, even if something above stopped the loop
    if (newRows.length) pulls.getRange(pulls.getLastRow() + 1, 1, newRows.length, 4).setValues(newRows);
  }
  ss_().getSheetByName(TAB_CANVAS).getRange(asg.row, 5).setValue(new Date());

  ui.alert('Pulled from Canvas: ' + asg.name,
    students + ' student(s), ' + files + ' file(s) are in "1 Inbox/' + asg.name + '".' +
    (skipped ? '\n' + skipped + ' submission(s) were already pulled before, so they were skipped.' : '') +
    (problems.length ? '\n\n' + problems.length + ' student(s) could not be downloaded this time. Run Pull submissions again later to retry them:\n' +
      problems.slice(0, 10).join('\n') + (problems.length > 10 ? '\n...and ' + (problems.length - 10) + ' more.' : '') : '') +
    (stopped ? '\n\nNot finished (Google time limit). Run Pull submissions again to get the rest.' : '') +
    '\n\n' + (autoRunIsOn_() ? 'Auto-run will clean them within about 10 minutes.' : 'Next: click "Clean the Inbox now" in the app (or Anonymizer > 2. Process the Inbox).') +
    (noFile ? '\n\n' + noFile + ' student(s) turned in something Pull can\'t download (usually Google Assignments, a recording, or a link from another tool). ' +
      'They are not marked as pulled. In Canvas, open each one, download it as a PDF with the student\'s name in the file name, and drop it in Settings > 1 Inbox.' : ''),
    ui.ButtonSet.OK);
}

/** Version 16: true if a submission has something Pull can download (a file, a text entry or a web link). */
function hasDownload_(s) {
  if ((s.attachments || []).length) return true;
  if (s.submission_type === 'online_text_entry' && s.body) return true;
  if (s.submission_type === 'online_url' && s.url) return true;
  return false;
}

/** Fills the Roster from every course you've pulled from, or a course link. Keeps your notes in columns C to E. */
function updateRosterFromCanvas() {
  const ui = SpreadsheetApp.getUi();
  try {
    const resp = ui.prompt('Update roster from Canvas',
      'Paste a link to any page in each course (separate several with spaces), or leave blank to use the courses on the Canvas Assignments tab.',
      ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    ui.alert(updateRoster_(resp.getResponseText()));
  } catch (e) {
    ui.alert(e.message);
  }
}

/** Adds Canvas students who aren't on the Roster yet. links: course links (or blank for the saved courses). Returns a message. */
function updateRoster_(links) {
  const sh = ensureSheet_(TAB_CANVAS, ['Assignment', 'Canvas link', 'Course ID', 'Assignment ID', 'Last pulled']);
  let courses = shownCourses_() || sh.getDataRange().getValues().slice(1).map(r => String(r[2])).filter(Boolean);   // version 24: your ticked classes
  courses = courses.slice();
  (String(links || '').match(/courses\/(\d+)/g) || []).forEach(m => courses.push(m.replace('courses/', '')));
  courses = courses.filter((c, i) => courses.indexOf(c) === i);
  if (!courses.length) return 'No classes yet. Tick your classes under Your classes (or paste a link from a Canvas course).';
  const roster = ss_().getSheetByName(TAB_ROSTER);
  const data = roster.getDataRange().getValues();
  const byId = {};
  data.slice(1).forEach((r, i) => { if (r[1] !== '') byId[String(r[1])] = i + 2; });
  let added = 0;
  courses.forEach(c => {
    canvasGetAll_('/api/v1/courses/' + c + '/users?enrollment_type[]=student').forEach(u => {
      const id = String(u.id);
      if (byId[id]) return;
      roster.appendRow([u.sortable_name || u.name, id, '', '', (u.email || u.login_id || '')]);
      byId[id] = roster.getLastRow();
      added++;
    });
  });
  roster.getRange('B:B').setNumberFormat('@');
  return added + ' new student(s) added to the Roster. Students already there were left as they are.';
}

/* ===================== CLAUDE'S GRADES (codes only) ================ */
/*
 * Claude writes Code | Score | Feedback into a separate Sheet called
 * "Claude Grades (codes only)" in the Student Work Anonymizer folder.
 * It has no names in it. This menu copies new grades into the Results tab.
 */

const CLAUDE_SHEET_NAME = 'Claude Grades (codes only)';

/** Finds the Claude Grades Sheet in the top folder or in "2 Ready for Claude". Returns its first tab, or null. */
function claudeGradesSheet_() {
  const book = claudeGradesBook_();
  return book ? book.getSheets()[0] : null;
}

/** The Claude Grades Sheet itself (all its tabs), or null. Version 10: the Review tab lives here too.
 *  Version 17: found once per run (opening a Sheet takes about a second). */
let CG_BOOK = undefined;
function claudeGradesBook_() {
  if (CG_BOOK !== undefined) return CG_BOOK;
  CG_BOOK = findClaudeGradesBook_();
  return CG_BOOK;
}
function findClaudeGradesBook_() {
  const props = PropertiesService.getDocumentProperties();
  const places = [props.getProperty('ROOT'), props.getProperty('READY')];
  for (let i = 0; i < places.length; i++) {
    let folder;
    try { folder = DriveApp.getFolderById(places[i]); } catch (e) { continue; }
    const it = folder.getFilesByName(CLAUDE_SHEET_NAME);
    while (it.hasNext()) {
      const file = it.next();
      if (!file.isTrashed() && file.getMimeType() === MimeType.GOOGLE_SHEETS) return SpreadsheetApp.open(file);
    }
  }
  return null;
}

/** Copies Claude's grades into Results. New codes are added; changed scores or feedback are updated. Returns how many changed. */
function bringInClaudeGrades_() {
  const src = claudeGradesSheet_();
  if (!src) return -1;
  const rows = src.getDataRange().getValues().slice(1).filter(r => String(r[1]).trim());
  const res = ss_().getSheetByName(TAB_RESULTS);
  const data = res.getDataRange().getValues();
  const rowOf = {};
  data.slice(1).forEach((r, i) => { if (r[0]) rowOf[String(r[0]).trim()] = i + 2; });
  let changed = 0;
  rows.forEach(r => {
    const code = String(r[1]).trim(), score = r[2], fb = String(r[3] || '');
    const at = rowOf[code];
    if (at) {
      const cur = data[at - 1];
      if (String(cur[1]) !== String(score) || String(cur[2]) !== fb) { res.getRange(at, 2, 1, 2).setValues([[score, fb]]); changed++; }
    } else {
      let next = 2;
      const colA = res.getRange('A:A').getValues();
      while (next <= colA.length && colA[next - 1][0] !== '') next++;
      res.getRange(next, 1, 1, 3).setValues([[code, score, fb]]);
      rowOf[code] = next;
      changed++;
    }
  });
  return changed;
}

function bringInClaudeGrades() {
  const n = bringInClaudeGrades_();
  SpreadsheetApp.getUi().alert(n < 0
    ? 'There is no "' + CLAUDE_SHEET_NAME + '" Sheet in the Student Work Anonymizer folder or in "2 Ready for Claude" yet. Ask Claude to grade an assignment and it will write the grades there.'
    : n + ' grade(s) added or updated in the Results tab.');
}

/** Sends scores (and feedback as a comment) for one assignment to Canvas. Skips anything already sent unchanged. */
function sendGradesToCanvas() {
  const ui = SpreadsheetApp.getUi();
  try {
    showCanvasPicker_('send', 'Send grades to Canvas', 'Which assignment are these grades for? Pick the same one you pulled from.');
  } catch (e) {
    ui.alert(e.message);
  }
}

function sendGradesFor_(asg) {
  const ui = ui_();
  if (!asg) return;
  bringInClaudeGrades_();

  // Version 10: answers Rowan flagged that are still waiting for you
  const waiting = undecidedReviewCount_(asg.name);
  if (waiting > 0) {
    const go = ui.alert('Some answers are still waiting for you',
      waiting + ' answer(s) for "' + asg.name + '" are still waiting in Review (the Review button on the card in the app).\n\n' +
      'Send anyway, using Rowan\'s suggested points for those?', ui.ButtonSet.YES_NO);
    if (go !== ui.Button.YES) return;
  }

  const key = ss_().getSheetByName(TAB_KEY).getDataRange().getValues().slice(1);
  const codeInfo = {};
  key.forEach(r => { codeInfo[String(r[1])] = { run: String(r[0]), canvasId: String(r[3]) }; });
  const res = ss_().getSheetByName(TAB_RESULTS);
  if (res.getRange(1, 6).getValue() === '') res.getRange(1, 6, 1, 2).setValues([['Sent to Canvas', 'sent copy (do not edit)']]).setFontWeight('bold');
  const data = res.getDataRange().getValues();

  // Assignment names on the Key are what this tool called it; match by the Canvas name, with or without an old date ending
  const sameRun = run => run === asg.name || run.replace(/ \d{4}-\d{2}-\d{2}$/, '') === asg.name;
  let todo = [];
  const all = [];
  const gradedIds = {};
  for (let i = 1; i < data.length; i++) {
    const code = String(data[i][0]).trim();
    const info = codeInfo[code];
    if (!code || !info || !sameRun(info.run) || data[i][1] === '' || !info.canvasId) continue;
    const sig = String(data[i][1]) + '||' + String(data[i][2]);
    const t = { row: i + 1, canvasId: info.canvasId, score: data[i][1], fb: String(data[i][2]),
                oldFb: String(data[i][6]).split('||').slice(1).join('||'), sig: sig };
    all.push(t);
    gradedIds[info.canvasId] = true;
    if (String(data[i][6]) !== sig) todo.push(t);
  }
  const matched = all.length;
  if (!matched) {
    ui.alert('No graded students found for "' + asg.name + '".\n\nThe tool matches the Canvas assignment name to the assignment name on the Key tab. ' +
      'If they are different, change the name in the first column of the Canvas Assignments tab to match the Key tab exactly, then try again.');
    return;
  }
  // Version 16: everything was sent before. Check them all again: Canvas may be missing some (it gets
  // only the ones it doesn't have, or has lower). You are asked once Canvas has been checked.
  const again = !todo.length;
  if (again) todo = all;

  // Version 9: never lower a grade. Look up what Canvas has now; only send when the new score is higher
  // (or Canvas has no grade yet). Anything already equal or higher in Canvas, or excused, stays as it is.
  let current;
  try {
    current = canvasCurrentScores_(asg);
  } catch (e) {
    ui.alert('Send grades stopped', 'Could not read the grades already in Canvas, so nothing was sent (this keeps existing grades safe).\n\n' + e.message, ui.ButtonSet.OK);
    return;
  }
  const plan = splitByCanvasScore_(todo, current);
  const tz = Session.getScriptTimeZone();
  const stamp = () => Utilities.formatDate(new Date(), tz, 'MMM d, h:mm a');
  // Record the kept ones so they aren't asked about again (a later change to that score is checked again).
  // On a send-again they were already recorded as sent, so they are left as they are.
  if (!again) plan.keep.forEach(t => {
    res.getRange(t.row, 6, 1, 2).setValues([['Kept Canvas grade (' + t.canvasLabel + ') ' + stamp(), t.sig]]);
  });
  try { res.hideColumns(7); } catch (e) { /* ignore */ }
  const keptLine = plan.keep.length
    ? (again
      ? plan.keep.length + ' student(s) already have this grade (or a higher one) in Canvas, so theirs were left alone.'
      : plan.keep.length + ' student(s) already have an equal or higher grade in Canvas, so theirs were left alone (marked "Kept Canvas grade" on the Results tab).')
    : '';
  const strayLine = ungradedInCanvasLine_(current, gradedIds);
  const where = canvasSettings_().base + '/courses/' + asg.course + '/assignments/' + asg.assignment;
  if (!plan.send.length) {
    ui.alert('Nothing to change in Canvas for ' + asg.name + '.' +
      (again ? '\n\nCanvas already has a grade for all ' + all.length + ' student(s) Rowan graded, on this assignment:\n' + where +
        '\nIf Canvas looks ungraded to you, check that it is the same assignment (another section or a copy of the assignment has its own grades).' : '') +
      (keptLine && !again ? '\n\n' + keptLine : '') +
      (strayLine ? '\n\n' + strayLine : ''));
    return;
  }
  const ok = ui.alert('Send to Canvas',
    (again
      ? 'These grades were sent before, but Canvas is missing ' + plan.send.length + ' of them (no grade, or a lower one) for "' + asg.name + '". Send those ' + plan.send.length + ' again now?\n\n' +
        'Feedback comments only go to students who don\'t already have them. Canvas grades are never lowered.'
      : 'Send ' + plan.send.length + ' grade(s) for "' + asg.name + '" to Canvas now?\n\n' +
        'These are only students with no grade yet or a lower one. Canvas grades are never lowered.') +
    (keptLine ? '\n\n' + keptLine : ''), ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  let sent = 0;
  const errors = [];
  plan.send.forEach(t => {
    try {
      const payload = { 'submission[posted_grade]': String(t.score) };
      if (needsComment_(t, current[String(t.canvasId)])) payload['comment[text_comment]'] = t.fb;
      canvasFetch_('/api/v1/courses/' + asg.course + '/assignments/' + asg.assignment + '/submissions/' + t.canvasId,
        { method: 'put', payload: payload });
      res.getRange(t.row, 6, 1, 2).setValues([['Sent ' + stamp(), t.sig]]);
      sent++;
    } catch (e) {
      errors.push(res.getRange(t.row, 1).getValue() + ': ' + e.message);
    }
  });
  try { res.hideColumns(7); } catch (e) { /* ignore */ }
  ui.alert(sent + ' grade(s) sent to Canvas for ' + asg.name + '.' +
    (keptLine ? '\n\n' + keptLine : '') +
    (errors.length ? '\n\nThese did not go through:\n' + errors.join('\n') : '') +
    (strayLine ? '\n\n' + strayLine : '') +
    '\n\nIf your course hides grades until you post them, post them in Canvas when you are ready.');
}

/**
 * What Canvas has now for each student on this assignment:
 * { canvasId: { score: number|null, excused: bool, submitted: bool, comments: [text]|null } }.
 * Version 16: also whether they turned something in, and the comments already on it.
 */
function canvasCurrentScores_(asg) {
  const out = {};
  canvasGetAll_('/api/v1/courses/' + asg.course + '/assignments/' + asg.assignment + '/submissions?include[]=submission_comments').forEach(s => {
    if (!s || !s.user_id) return;
    const n = (s.score === null || s.score === undefined || s.score === '') ? null : Number(s.score);
    out[String(s.user_id)] = { score: (n === null || isNaN(n)) ? null : n, excused: !!s.excused,
      submitted: !!s.submitted_at,
      comments: Array.isArray(s.submission_comments) ? s.submission_comments.map(c => String((c && c.comment) || '').trim()) : null };
  });
  return out;
}

/** Version 16: send the feedback comment only when Canvas doesn't already have it (or, if Canvas didn't say, when it changed). */
function needsComment_(t, c) {
  const fb = String(t.fb || '').trim();
  if (!fb) return false;
  if (c && Array.isArray(c.comments)) return c.comments.indexOf(fb) < 0;
  return t.fb !== t.oldFb;
}

/** Version 16: students who turned something in, have no grade in Canvas, and have no grade from Rowan. Counts only. */
function ungradedInCanvasLine_(current, gradedIds) {
  const n = Object.keys(current).filter(id => {
    const c = current[id];
    return c.submitted && c.score === null && !c.excused && !gradedIds[id];
  }).length;
  if (!n) return '';
  return 'Canvas also shows ' + n + ' student(s) who turned something in with no grade yet, and Rowan has no grade for them. ' +
    'Either their work is still waiting for Rowan (say "grade"), or it never came over (Google Assignments, recordings and links from other tools can\'t be pulled; ' +
    'Pull tells you how many). Grade those in Canvas, or download them as PDFs into Settings > 1 Inbox.';
}

/**
 * Splits grades into ones to send and ones to keep as they are in Canvas.
 * Send only when Canvas has no grade yet, or the new score is higher. Excused students are always kept.
 * A score that isn't a number is only sent when Canvas has no grade yet.
 */
function splitByCanvasScore_(todo, current) {
  const send = [], keep = [];
  todo.forEach(t => {
    const c = current[String(t.canvasId)] || { score: null, excused: false };
    if (c.excused) { t.canvasLabel = 'excused'; keep.push(t); return; }
    if (c.score === null) { send.push(t); return; }
    const mine = Number(t.score);
    if (String(t.score).trim() !== '' && !isNaN(mine) && mine > c.score) { send.push(t); return; }
    t.canvasLabel = String(c.score);
    keep.push(t);
  });
  return { send: send, keep: keep };
}

/* ===================== REVIEW WINDOW (version 10) ================== */
/*
 * When Rowan isn't sure about an answer (can't read it, partly right, a fair
 * answer the key didn't expect), Rowan writes a row on the "Review" tab of
 * the Claude Grades (codes only) Sheet instead of guessing. The Score Rowan
 * writes already includes Rowan's suggested points. This window shows each
 * row with the student's page, and your choice moves the score up or down by
 * the difference. Codes only: no names are read or shown here.
 *
 * Review tab columns (Rowan fills A to I, this window fills J and K):
 * Assignment | Code | Question | Points possible | Rowan's points |
 * Student's answer (as Rowan read it) | Why Rowan was not sure | Page |
 * Where on page (top, middle, bottom) | Eric's points | Decided
 */

const REVIEW_TAB = 'Review';
const REVIEW_HEADERS = ['Assignment', 'Code', 'Question', 'Points possible', 'Rowan\'s points',
  'Student\'s answer (as Rowan read it)', 'Why Rowan was not sure', 'Page', 'Where on page',
  'Eric\'s points', 'Decided'];
const PDFJS_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
const PDFJS_WORKER_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

/** The Review tab of the Claude Grades Sheet (made if missing), or null if that Sheet can't be found. */
function ensureReviewTab_() {
  const book = claudeGradesBook_();
  if (!book) return null;
  let sh = book.getSheetByName(REVIEW_TAB);
  if (!sh) {
    sh = book.insertSheet(REVIEW_TAB, book.getSheets().length);   // after the grades tab, never first
    sh.getRange(1, 1, 1, REVIEW_HEADERS.length).setValues([REVIEW_HEADERS]).setFontWeight('bold').setBackground('#e8eef7');
    sh.setFrozenRows(1);
  }
  return sh;
}

function sameRunName_(run, name) {
  run = String(run || '').trim();
  return run === name || run.replace(/ \d{4}-\d{2}-\d{2}$/, '') === name;
}

/** How many Review rows for this assignment are still waiting for a decision. */
function undecidedReviewCount_(name) {
  try {
    const book = claudeGradesBook_();
    const sh = book ? book.getSheetByName(REVIEW_TAB) : null;
    if (!sh) return 0;
    return sh.getDataRange().getValues().slice(1)
      .filter(r => String(r[1]).trim() && sameRunName_(r[0], name) && String(r[9]).trim() === '').length;
  } catch (e) {
    return 0;
  }
}

function openReview() {
  const ui = SpreadsheetApp.getUi();
  const sh = ensureReviewTab_();
  if (!sh) {
    ui.alert('There is no "' + CLAUDE_SHEET_NAME + '" Sheet yet. Ask Rowan to grade an assignment first.');
    return;
  }
  const html = HtmlService.createHtmlOutput(reviewHtml_()).setWidth(1180).setHeight(780);
  ui.showModalDialog(html, 'Review: what Rowan was not sure about');
}

function num_(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}

/** Everything the Review window needs (codes and points only). */
function reviewData() {
  const sh = ensureReviewTab_();
  if (!sh) throw new Error('The Claude Grades Sheet could not be found.');
  const tz = Session.getScriptTimeZone();
  const items = [];
  sh.getDataRange().getValues().slice(1).forEach((r, i) => {
    if (!String(r[1]).trim()) return;
    const decided = r[10] instanceof Date ? Utilities.formatDate(r[10], tz, 'MMM d, h:mm a') : String(r[10] || '');
    items.push({
      row: i + 2, assignment: String(r[0]).trim(), code: String(r[1]).trim(), question: String(r[2]),
      possible: num_(r[3]), rowan: num_(r[4]), answer: String(r[5] || ''), why: String(r[6] || ''),
      page: num_(r[7]) || 1, where: String(r[8] || '').toLowerCase(), eric: num_(r[9]), decided: decided
    });
  });
  const scores = {};
  const grades = claudeGradesSheet_();
  if (grades) grades.getDataRange().getValues().slice(1).forEach(r => {
    if (String(r[1]).trim()) scores[String(r[0]).trim() + '|' + String(r[1]).trim()] = String(r[2]);
  });
  items.forEach(it => { it.score = scores[it.assignment + '|' + it.code] || ''; });
  return { items: items };
}

/** The coded file Rowan graded, for the window to show. Only files named S-#### in "2 Ready for Claude". */
function reviewFile(assignment, code) {
  code = String(code || '').trim();
  if (!/^S-\d{4}$/.test(code)) throw new Error('That is not a student code.');
  const ready = folders_().ready;
  const pattern = new RegExp('^' + code + '(-\\d+)?\\.(pdf|png|jpe?g)$', 'i');
  const look = folder => {
    const it = folder.getFiles();
    const found = [];
    while (it.hasNext()) { const fl = it.next(); if (!fl.isTrashed() && pattern.test(fl.getName())) found.push(fl); }
    found.sort((a, b) => naturalCompare_(a.getName(), b.getName()));
    return found[0] || null;
  };
  let file = null;
  const fit = ready.getFoldersByName(String(assignment));
  while (!file && fit.hasNext()) {
    const fo = fit.next();
    if (fo.isTrashed()) continue;
    file = look(fo);
    const sub = fo.getFolders();
    while (!file && sub.hasNext()) file = look(sub.next());
  }
  if (!file) return { found: false };
  const out = { found: true, id: file.getId(), name: file.getName(), mime: file.getMimeType() };
  if (file.getSize() <= 15 * 1024 * 1024) out.data = Utilities.base64Encode(file.getBlob().getBytes());
  return out;
}

/**
 * Saves your choice for one Review row and moves that student's score by the difference.
 * expect: { assignment, code, question } so a row that moved since the window opened isn't changed by mistake.
 */
function reviewDecide(row, expect, points) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureReviewTab_();
    const r = sh.getRange(row, 1, 1, REVIEW_HEADERS.length).getValues()[0];
    if (String(r[0]).trim() !== expect.assignment || String(r[1]).trim() !== expect.code || String(r[2]) !== expect.question) {
      throw new Error('The Review tab changed while this window was open. Close the window and open it again.');
    }
    const possible = num_(r[3]);
    const p = num_(points);
    if (p === null || p < 0 || (possible !== null && p > possible)) {
      throw new Error('Points have to be a number from 0 to ' + (possible === null ? 'the points possible' : possible) + '.');
    }
    const before = num_(r[9]) !== null ? num_(r[9]) : (num_(r[4]) || 0);
    const delta = p - before;

    let score = null, note = '';
    const grades = claudeGradesSheet_();
    const data = grades.getDataRange().getValues();
    let at = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === expect.assignment && String(data[i][1]).trim() === expect.code) { at = i; break; }
    }
    if (at < 0) {
      note = 'This code has no score in the Claude Grades Sheet yet, so only your choice was saved.';
    } else if (num_(data[at][2]) === null) {
      note = 'The score in the Claude Grades Sheet is not a number, so it was not changed. Fix it there by hand.';
    } else {
      score = Math.round((num_(data[at][2]) + delta) * 100) / 100;
      if (delta !== 0) grades.getRange(at + 1, 3).setValue(score);
    }
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, h:mm a');
    sh.getRange(row, 10, 1, 2).setValues([[p, stamp]]);
    SpreadsheetApp.flush();
    return { eric: p, decided: stamp, score: score === null ? '' : String(score), note: note };
  } finally {
    lock.releaseLock();
  }
}

function reviewHtml_() {
  return '<!doctype html><html><head><meta charset="utf-8">' +
    '<script src="' + PDFJS_URL + '"></scr' + 'ipt>' +
    '<script>var PDFJS_WORKER=' + JSON.stringify(PDFJS_WORKER_URL) + ';</scr' + 'ipt></head>' +
    REVIEW_PAGE_HTML + '</html>';
}

const REVIEW_PAGE_HTML = `
<body>
<style>
  :root{--fg:#1f2933;--muted:#5f6b7a;--line:#dde3ea;--soft:#f4f6f9;--blue:#1a73e8;--blue-soft:#e8f0fe;--green:#137333;--green-soft:#e6f4ea;--red:#b3261e;--red-soft:#fce8e6;--amber:#8a5300;--amber-soft:#fef3e0}
  *{box-sizing:border-box}
  body{margin:0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:var(--fg);background:#fff;height:100vh;overflow:hidden}
  .top{display:flex;align-items:center;gap:12px;padding:10px 16px;border-bottom:1px solid var(--line);flex-wrap:wrap}
  .top select{font-size:14px;padding:6px 8px;max-width:360px}
  .top label{color:var(--muted);display:flex;align-items:center;gap:6px;cursor:pointer}
  .bar{flex:1;min-width:140px;height:8px;background:var(--soft);border-radius:4px;overflow:hidden}
  .bar i{display:block;height:100%;background:var(--green);width:0;transition:width .3s}
  .count{color:var(--muted);white-space:nowrap}
  .main{display:grid;grid-template-columns:400px 1fr;height:calc(100vh - 54px)}
  .card{padding:18px 18px 24px;overflow:auto;border-right:1px solid var(--line)}
  .code{font-size:13px;font-weight:bold;color:var(--muted);letter-spacing:.03em}
  h2{font-size:19px;line-height:1.3;margin:6px 0 4px}
  .worth{color:var(--muted);margin-bottom:14px}
  .label{font-size:12px;font-weight:bold;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;margin:14px 0 4px}
  .answer{background:var(--soft);border-radius:8px;padding:10px 12px;white-space:pre-wrap;line-height:1.45}
  .why{background:var(--amber-soft);color:#4a2f00;border-radius:8px;padding:10px 12px;line-height:1.45;white-space:pre-wrap}
  .choices{display:grid;gap:8px;margin-top:16px}
  .choices button{font-size:15px;padding:11px 12px;border-radius:8px;border:1px solid var(--line);background:#fff;cursor:pointer;text-align:left;display:flex;justify-content:space-between;align-items:center}
  .choices button:hover{border-color:var(--blue)}
  .choices button b{font-size:16px}
  .choices .full{background:var(--green-soft);border-color:#b7dfc2}
  .choices .rowan{background:var(--blue-soft);border-color:#c6d9fb}
  .choices .none{background:var(--red-soft);border-color:#f4c7c3}
  .choices button.picked{outline:3px solid var(--fg);outline-offset:1px}
  .custom{display:flex;gap:8px;align-items:center;margin-top:4px}
  .custom input{width:90px;font-size:15px;padding:8px;border:1px solid var(--line);border-radius:8px}
  .custom button{font-size:14px;padding:9px 14px;border-radius:8px;border:0;background:var(--blue);color:#fff;cursor:pointer}
  .keys{color:var(--muted);font-size:12px;margin-top:10px}
  .nav{display:flex;justify-content:space-between;margin-top:18px}
  .nav button{font-size:13px;padding:6px 12px;border-radius:6px;border:1px solid var(--line);background:#fff;cursor:pointer}
  .msg{min-height:20px;margin-top:12px;font-weight:bold}
  .msg.ok{color:var(--green)} .msg.err{color:var(--red)}
  .decided{margin-top:10px;color:var(--green);font-weight:bold}
  .main>*{min-height:0}
  .view{display:flex;flex-direction:column;background:#eef1f5;min-width:0;overflow:hidden}
  .viewbar{display:flex;align-items:center;gap:8px;padding:8px 12px;background:#fff;border-bottom:1px solid var(--line);flex-wrap:wrap}
  .viewbar button{font-size:13px;padding:5px 10px;border-radius:6px;border:1px solid var(--line);background:#fff;cursor:pointer}
  .viewbar a{margin-left:auto;color:var(--blue)}
  .scroll{flex:1;min-height:0;overflow:auto;padding:12px;text-align:center}
  .scroll canvas,.scroll img{max-width:100%;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.18)}
  .scroll iframe{width:100%;height:100%;border:0;background:#fff}
  .center{padding:60px 30px;text-align:center;color:var(--muted);line-height:1.6}
  .done{padding:50px 40px;max-width:640px;line-height:1.6}
  .done h2{font-size:22px}
  .done li{margin-bottom:4px}
</style>
<div class="top">
  <select id="asg" aria-label="Assignment"></select>
  <label><input type="checkbox" id="showAll"> Show ones I already decided</label>
  <div class="bar" aria-hidden="true"><i id="barFill"></i></div>
  <span class="count" id="count"></span>
</div>
<div class="main">
  <div class="card" id="card"><div class="center">Loading...</div></div>
  <div class="view">
    <div class="viewbar">
      <button id="prevPage" type="button">&#8592; Page</button>
      <span id="pageLabel"></span>
      <button id="nextPage" type="button">Page &#8594;</button>
      <a id="openFile" target="_blank" href="#" hidden>Open the whole file</a>
    </div>
    <div class="scroll" id="scroll"><div class="center">The student's page shows here.</div></div>
  </div>
</div>
<script>
var all = [], list = [], idx = 0, files = {}, viewing = null, viewPage = 1, busy = false;
var $ = function(id){ return document.getElementById(id); };
function el(tag, cls, text){ var e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; }
function pts(n){ return (n === null || n === undefined || n === '') ? '?' : String(n); }

if (window.pdfjsLib) { try { pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER; } catch (e) {} }

google.script.run.withSuccessHandler(start).withFailureHandler(fail).reviewData();

function fail(e){ $('card').innerHTML = ''; $('card').appendChild(el('div', 'center', (e && e.message) ? e.message : String(e))); }

function start(d){
  all = d.items || [];
  var names = [], seen = {};
  all.forEach(function(it){ if (!seen[it.assignment]) { seen[it.assignment] = true; names.push(it.assignment); } });
  var sel = $('asg'); sel.innerHTML = '';
  if (!names.length) { sel.appendChild(el('option', '', 'Nothing to review')); sel.disabled = true; }
  names.forEach(function(n){ var o = el('option', '', n); o.value = n; sel.appendChild(o); });
  var firstOpen = names.filter(function(n){ return waiting(n) > 0; })[0];
  if (firstOpen) sel.value = firstOpen;
  refreshNames();
  sel.onchange = function(){ idx = 0; build(); };
  $('showAll').onchange = function(){ idx = 0; build(); };
  build();
}

function waiting(name){ return all.filter(function(it){ return it.assignment === name && it.eric === null; }).length; }
function refreshNames(){
  Array.prototype.forEach.call($('asg').options, function(o){ if (o.value) o.textContent = o.value + (waiting(o.value) ? '  (' + waiting(o.value) + ' to go)' : '  (done)'); });
}

function build(){
  var name = $('asg').value, showAll = $('showAll').checked;
  list = all.filter(function(it){ return it.assignment === name && (showAll || it.eric === null); });
  if (idx >= list.length) idx = Math.max(0, list.length - 1);
  progress();
  if (!list.length) return doneScreen(name);
  card(list[idx]);
}

function progress(){
  var name = $('asg').value;
  var mine = all.filter(function(it){ return it.assignment === name; });
  var done = mine.filter(function(it){ return it.eric !== null; }).length;
  $('barFill').style.width = mine.length ? (100 * done / mine.length) + '%' : '0';
  $('count').textContent = mine.length ? done + ' of ' + mine.length + ' decided' : '';
}

function doneScreen(name){
  var c = $('card'); c.innerHTML = '';
  var box = el('div', 'done');
  var mine = all.filter(function(it){ return it.assignment === name; });
  if (!mine.length) {
    box.appendChild(el('h2', '', 'Nothing to review'));
    box.appendChild(el('p', '', 'Rowan did not flag anything for this assignment.'));
  } else {
    box.appendChild(el('h2', '', 'All done with ' + name));
    box.appendChild(el('p', '', 'Every flagged answer has your decision, and the totals in the Claude Grades Sheet are updated.'));
    var ol = el('ol');
    ['Close this window.', 'Anonymizer > Canvas > Send grades and comments to Canvas.', 'Pick ' + name + '.'].forEach(function(t){ ol.appendChild(el('li', '', t)); });
    box.appendChild(el('p', '', 'Next:')); box.appendChild(ol);
    box.appendChild(el('p', '', 'Want to change one? Tick "Show ones I already decided."'));
  }
  c.appendChild(box);
  $('scroll').innerHTML = ''; $('scroll').appendChild(el('div', 'center', ''));
  $('pageLabel').textContent = ''; $('openFile').hidden = true;
}

function card(it){
  var c = $('card'); c.innerHTML = '';
  c.appendChild(el('div', 'code', it.code + (it.score !== '' ? '   ·   score now ' + it.score : '')));
  c.appendChild(el('h2', '', it.question));
  c.appendChild(el('div', 'worth', 'Worth ' + pts(it.possible) + ' point' + (it.possible === 1 ? '' : 's') + '   ·   page ' + it.page));
  c.appendChild(el('div', 'label', "Student's answer (as Rowan read it)"));
  c.appendChild(el('div', 'answer', it.answer || '(Rowan did not type it out. See the page.)'));
  c.appendChild(el('div', 'label', 'Why Rowan was not sure'));
  c.appendChild(el('div', 'why', it.why || '(no note)'));

  var ch = el('div', 'choices');
  var opts = [];
  if (it.possible !== null) opts.push({ cls: 'full', label: 'Full credit', v: it.possible, key: 'F' });
  if (it.rowan !== null && it.rowan !== it.possible && it.rowan !== 0) opts.push({ cls: 'rowan', label: "Rowan's suggestion", v: it.rowan, key: 'R' });
  opts.push({ cls: 'none', label: 'No credit', v: 0, key: 'N' });
  opts.forEach(function(o){
    var b = el('button', o.cls); b.type = 'button';
    b.appendChild(el('span', '', o.label + (o.key !== 'R' && it.rowan === o.v ? '  (Rowan\u2019s pick)' : '')));
    b.appendChild(el('b', '', String(o.v)));
    if (it.eric !== null && it.eric === o.v) b.classList.add('picked');
    b.onclick = function(){ decide(it, o.v); };
    b.dataset.key = o.key;
    ch.appendChild(b);
  });
  var cu = el('div', 'custom');
  cu.appendChild(el('span', '', 'Or give'));
  var inp = el('input'); inp.type = 'number'; inp.min = '0'; inp.step = '0.5'; if (it.possible !== null) inp.max = String(it.possible);
  inp.setAttribute('aria-label', 'Points to give');
  if (it.eric !== null) inp.value = it.eric;
  var save = el('button', '', 'Save'); save.type = 'button';
  save.onclick = function(){ if (inp.value === '') { say('Type a number of points first.', true); return; } decide(it, Number(inp.value)); };
  inp.onkeydown = function(e){ if (e.key === 'Enter') save.onclick(); };
  cu.appendChild(inp); cu.appendChild(el('span', '', 'points')); cu.appendChild(save);
  ch.appendChild(cu);
  c.appendChild(ch);
  c.appendChild(el('div', 'keys', 'Keys: F full credit, ' + (opts.some(function(o){ return o.key === 'R'; }) ? 'R Rowan’s, ' : '') + 'N no credit, arrows to move.'));
  if (it.eric !== null) c.appendChild(el('div', 'decided', 'You gave ' + it.eric + (it.decided ? ' (' + it.decided + ')' : '') + '. Click another choice to change it.'));
  var m = el('div', 'msg'); m.id = 'msg'; c.appendChild(m);

  var nav = el('div', 'nav');
  var back = el('button', '', '← Back'); back.type = 'button'; back.disabled = idx === 0;
  back.onclick = function(){ if (idx > 0) { idx--; build(); } };
  var skip = el('button', '', 'Skip for now →'); skip.type = 'button'; skip.disabled = idx >= list.length - 1;
  skip.onclick = function(){ if (idx < list.length - 1) { idx++; build(); } };
  nav.appendChild(back); nav.appendChild(el('span', 'count', (idx + 1) + ' of ' + list.length)); nav.appendChild(skip);
  c.appendChild(nav);

  show(it);
}

function say(t, bad){ var m = $('msg'); if (!m) return; m.className = 'msg ' + (bad ? 'err' : 'ok'); m.textContent = t; }

function decide(it, v){
  if (busy) return;
  busy = true;
  say('Saving...', false);
  google.script.run.withSuccessHandler(function(r){
    busy = false;
    it.eric = r.eric; it.decided = r.decided;
    if (r.score !== '') all.forEach(function(x){ if (x.assignment === it.assignment && x.code === it.code) x.score = r.score; });
    refreshNames();
    var showAll = $('showAll').checked;
    if (showAll && idx < list.length - 1) idx++;
    build();
    say(r.note ? r.note : ('Saved. ' + it.code + ' is now ' + r.score + '.'), !!r.note);
  }).withFailureHandler(function(e){
    busy = false;
    say((e && e.message) ? e.message : String(e), true);
  }).reviewDecide(it.row, { assignment: it.assignment, code: it.code, question: it.question }, v);
}

document.addEventListener('keydown', function(e){
  if (e.target && e.target.tagName === 'INPUT') return;
  var k = e.key.toUpperCase();
  if (k === 'ARROWRIGHT' && idx < list.length - 1) { idx++; build(); return; }
  if (k === 'ARROWLEFT' && idx > 0) { idx--; build(); return; }
  var b = document.querySelector('.choices button[data-key="' + k + '"]');
  if (b) b.click();
});

/* ---------- the student's page ---------- */
function show(it){
  viewing = it;
  var key = it.assignment + '|' + it.code;
  if (files[key] && files[key].ready) return render(files[key], it.page, it.where);
  $('scroll').innerHTML = ''; $('scroll').appendChild(el('div', 'center', 'Loading the page...'));
  $('pageLabel').textContent = '';
  if (files[key] && files[key].loading) return;
  files[key] = { loading: true };
  google.script.run.withSuccessHandler(function(f){
    files[key] = f; f.ready = true;
    if (f.found && f.data && /pdf/.test(f.mime) && window.pdfjsLib) {
      var raw = atob(f.data), bytes = new Uint8Array(raw.length);
      for (var i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i);
      f.data = null;
      f.doc = pdfjsLib.getDocument({ data: bytes }).promise;
    }
    if (viewing && viewing.assignment + '|' + viewing.code === key) render(f, viewing.page, viewing.where);
  }).withFailureHandler(function(e){
    files[key] = null;
    $('scroll').innerHTML = ''; $('scroll').appendChild(el('div', 'center', 'Could not load the page: ' + ((e && e.message) ? e.message : e)));
  }).reviewFile(it.assignment, it.code);
}

function render(f, page, where){
  var sc = $('scroll');
  var link = $('openFile');
  if (!f.found) { sc.innerHTML = ''; sc.appendChild(el('div', 'center', 'The coded file for this student is not in "2 Ready for Claude" anymore.')); link.hidden = true; return; }
  link.href = 'https://drive.google.com/file/d/' + f.id + '/view'; link.hidden = false;
  viewPage = page || 1;
  if (f.doc) {
    f.doc.then(function(doc){
      viewPage = Math.min(Math.max(1, viewPage), doc.numPages);
      $('pageLabel').textContent = 'Page ' + viewPage + ' of ' + doc.numPages;
      $('prevPage').onclick = function(){ if (viewPage > 1) drawPdf(doc, --viewPage, 'top'); };
      $('nextPage').onclick = function(){ if (viewPage < doc.numPages) drawPdf(doc, ++viewPage, 'top'); };
      drawPdf(doc, viewPage, where);
    }).catch(function(){ f.doc = null; frame(f, page); });
    return;
  }
  if (f.data && String(f.mime).indexOf('image/') === 0) {
    sc.innerHTML = ''; var img = el('img'); img.src = 'data:' + f.mime + ';base64,' + f.data; img.alt = 'Student work'; sc.appendChild(img);
    $('pageLabel').textContent = ''; return;
  }
  frame(f, page);
}

function drawPdf(doc, n, where){
  doc.getPage(n).then(function(pg){
    $('pageLabel').textContent = 'Page ' + n + ' of ' + doc.numPages;
    var sc = $('scroll');
    var width = Math.max(300, sc.clientWidth - 30);
    var base = pg.getViewport({ scale: 1 });
    var scale = width / base.width;
    var ratio = window.devicePixelRatio || 1;
    var vp = pg.getViewport({ scale: scale * ratio });
    var cv = document.createElement('canvas');
    cv.width = vp.width; cv.height = vp.height;
    cv.style.width = (vp.width / ratio) + 'px'; cv.style.height = (vp.height / ratio) + 'px';
    return pg.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise.then(function(){
      sc.innerHTML = ''; sc.appendChild(cv);
      var at = { top: 0, middle: 0.5, bottom: 1 }[where] || 0;
      sc.scrollTop = (sc.scrollHeight - sc.clientHeight) * at;
    });
  });
}

/** Backup when the page can't be drawn here: Google's own viewer (scroll to the page yourself). */
function frame(f, page){
  var sc = $('scroll'); sc.innerHTML = '';
  $('pageLabel').textContent = 'Scroll to page ' + (page || 1);
  $('prevPage').onclick = null; $('nextPage').onclick = null;
  var fr = document.createElement('iframe');
  fr.src = 'https://drive.google.com/file/d/' + f.id + '/preview';
  fr.title = 'Student work';
  sc.appendChild(fr);
}
</script>
</body>`;

/* ============================ WEB APP (version 12) ================== */
/*
 * The same tool as a web page with its own link (Deploy > Test deployments
 * > Web app). It runs in YOUR school Google account, as you, and only you can
 * open it. The Sheet stays behind it: every button here does exactly what the
 * Sheet menu does, on the same tabs and folders. Names never show on the home
 * page (codes and counts only); the Match window shows names to you only,
 * as it does from the Sheet.
 */

let SS_CACHE = null;
/** This Sheet, whether the code runs from the Sheet menu, a timer or the web app. */
function ss_() {
  if (SS_CACHE) return SS_CACHE;
  let ss = null;
  try { ss = SpreadsheetApp.getActive(); } catch (e) { ss = null; }
  if (ss) {
    try { if (PropertiesService.getScriptProperties().getProperty('SHEET_ID') !== ss.getId()) PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId()); } catch (e) { /* ignore */ }
    SS_CACHE = ss;
    return ss;
  }
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Open the Student Work Anonymizer Sheet once (or reload it), then try again.');
  SS_CACHE = SpreadsheetApp.openById(id);
  return SS_CACHE;
}

/* Pop-ups: in the Sheet they are Google's boxes; in the web app they are collected and shown on the page. */
let WEB_UI = null;
function ui_() { return WEB_UI || SpreadsheetApp.getUi(); }

/** A stand-in for Google's pop-ups. Says YES to the first yesCount questions, then stops at the next one. */
function webUi_(yesCount) {
  let yes = Number(yesCount) || 0;
  const ui = {
    msgs: [], asked: null,
    ButtonSet: { OK: 'OK', YES_NO: 'YES_NO', OK_CANCEL: 'OK_CANCEL' },
    Button: { OK: 'OK', YES: 'YES', NO: 'NO', CANCEL: 'CANCEL' },
    alert: function (a, b, c) {
      const title = b === undefined ? '' : String(a);
      const text = b === undefined ? String(a) : String(b);
      if (c === 'YES_NO') {
        if (yes > 0) { yes--; return 'YES'; }
        if (!ui.asked) ui.asked = { title: title, text: text };
        return 'NO';
      }
      ui.msgs.push(title ? title + '\n' + text : text);
      return 'OK';
    },
    showModalDialog: function () {}, showModelessDialog: function () {}
  };
  return ui;
}

/** Runs one job with the web stand-in. Returns { ok, text, confirm } for the page. */
function withWebUi_(yesCount, job) {
  WEB_UI = webUi_(yesCount);
  try {
    job();
    return { ok: true, text: WEB_UI.msgs.join('\n\n'), confirm: WEB_UI.asked };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  } finally {
    WEB_UI = null;
  }
}

function doGet() {
  // Version 17: the page starts with what it showed last time (codes and counts only), then refreshes
  let cached = 'null';
  try { cached = CacheService.getUserCache().get('status') || 'null'; } catch (e) { cached = 'null'; }
  const html = webHomeHtml_().replace('/*CACHED_STATUS*/null', cached.replace(/</g, '\\u003c'));
  return HtmlService.createHtmlOutput(html)
    .setTitle('Student Work Anonymizer')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Sheet menu: how to open the web app. */
function openWebApp() {
  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; }
  const html = HtmlService.createHtmlOutput('<div style="font-family:Arial;font-size:14px;line-height:1.5">' +
    (url ? '<p><a target="_blank" href="' + url + '"><b>Open the Anonymizer app</b></a></p><p>If that link shows old buttons, use your Test deployment link instead (below).</p>' : '') +
    '<p><b>First time only:</b></p><ol>' +
    '<li>Extensions &gt; Apps Script.</li>' +
    '<li>Blue <b>Deploy</b> button (top right) &gt; <b>Test deployments</b>.</li>' +
    '<li>Click the gear next to "Select type" and pick <b>Web app</b>.</li>' +
    '<li>Copy the <b>URL</b> (it ends in /dev) and bookmark it.</li></ol>' +
    '<p>That link always runs your newest code, and only you can open it.</p></div>').setWidth(440).setHeight(320);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Anonymizer app');
}

/* ---------- what the home page shows ---------- */

/** Coded work in a Ready folder (and its Added folders): { students, files }. S-1234.pdf and S-1234-2.pdf are one student. */
function codedIn_(folder) {
  const codes = {};
  let files = 0;
  const count = fo => { const it = fo.getFiles(); while (it.hasNext()) { const m = it.next().getName().match(/^(S-\d{4})(-\d+)?\.(pdf|png|jpe?g)$/i); if (m) { files++; codes[m[1].toUpperCase()] = true; } } };
  count(folder);
  const sub = folder.getFolders();
  while (sub.hasNext()) count(sub.next());
  return { students: Object.keys(codes).length, files: files, codes: codes };
}
function filesIn_(parent, name) {
  let n = 0;
  const it = parent.getFoldersByName(name);
  while (it.hasNext()) { const fo = it.next(); if (fo.isTrashed()) continue; const f2 = fo.getFiles(); while (f2.hasNext()) { f2.next(); n++; } }
  return n;
}

/**
 * Version 17: children of several folders in a few big Drive requests (not one folder at a time).
 * Returns [{ id, name, folder, parents: [id] }], trashed files left out.
 */
function driveKids_(parentIds) {
  const out = [];
  const ids = (parentIds || []).filter(Boolean);
  const v3 = !!Drive.Files.create;
  for (let i = 0; i < ids.length; i += 40) {
    const q = '(' + ids.slice(i, i + 40).map(id => "'" + id + "' in parents").join(' or ') + ') and trashed = false';
    let token = '';
    do {
      const opt = v3 ? { q: q, pageSize: 1000, fields: 'nextPageToken,files(id,name,mimeType,parents)' }
                     : { q: q, maxResults: 1000, fields: 'nextPageToken,items(id,title,mimeType,parents(id))' };
      if (token) opt.pageToken = token;
      const r = Drive.Files.list(opt);
      (r.files || r.items || []).forEach(x => out.push({
        id: x.id, name: String(x.name || x.title || ''), folder: x.mimeType === 'application/vnd.google-apps.folder',
        parents: (x.parents || []).map(p => typeof p === 'string' ? p : p.id)
      }));
      token = r.nextPageToken || '';
    } while (token);
  }
  return out;
}

/** Every assignment with where it stands (codes and counts only). Version 17: fast version, with the old one as a backup. */
function webStatus() {
  let out;
  try { out = webStatusFast_(); }
  catch (e) { out = webStatusSlow_(); out.slowNote = e.message; }
  out.version = APP_VERSION;
  try { const t = JSON.stringify(out); if (t.length < 95000) CacheService.getUserCache().put('status', t, 21600); } catch (e) { /* ignore */ }
  return out;
}

function webStatusFast_() {
  const props = PropertiesService.getDocumentProperties();
  const READY = props.getProperty('READY'), INBOX = props.getProperty('INBOX'), DONE = props.getProperty('DONE');
  if (!READY || !INBOX || !DONE) throw new Error('Folders not found. Run Anonymizer > Set up first.');
  const folderUrl = id => 'https://drive.google.com/drive/folders/' + id;
  const fileUrl = id => 'https://drive.google.com/file/d/' + id + '/view';
  let canvas = false;
  try { canvasSettings_(); canvas = true; } catch (e) { canvas = false; }

  // Sheets: each tab read once
  const key = ss_().getSheetByName(TAB_KEY).getDataRange().getValues().slice(1);
  const pastNames = [], onKey = {};
  key.slice().reverse().forEach(r => { const a = String(r[0] || '').trim(); if (a && !onKey[a]) { onKey[a] = true; pastNames.push(a); } });
  const scanSh = ss_().getSheetByName(TAB_SCANS);
  const scanRows = scanSh ? scanSh.getDataRange().getValues().slice(1) : [];
  const plan = ss_().getSheetByName(TAB_PLAN);
  const planRows = plan ? plan.getDataRange().getValues().slice(1) : [];
  const planned = {}, toBuild = {};
  planRows.forEach(r => {
    if (r[6] === 'Ready to save') toBuild[r[0]] = (toBuild[r[0]] || 0) + 1;
    String(r[4]).split(',').forEach(x => { if (x) planned[r[0] + '|' + String(x).trim()] = true; });
  });
  const toMatch = {};
  scanRows.forEach(r => { if (r[0] && !planned[r[0] + '|' + String(r[1])]) toMatch[r[0]] = (toMatch[r[0]] || 0) + 1; });

  const graded = {}, sig = {};
  const book = claudeGradesBook_();
  const grades = book ? book.getSheets()[0] : null;
  if (grades) grades.getDataRange().getValues().slice(1).forEach(r => {
    const code = String(r[1]).trim();
    if (!code || r[2] === '') return;
    const run = String(r[0]).trim();
    (graded[run] = graded[run] || []).push(code);
    sig[code] = String(r[2]) + '||' + String(r[3] || '');
  });
  const sentSig = {};
  const res = ss_().getSheetByName(TAB_RESULTS);
  if (res) res.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) sentSig[String(r[0]).trim()] = String(r[6] || ''); });
  const waitingReview = {};
  try {
    const rv = book ? book.getSheetByName(REVIEW_TAB) : null;
    if (rv) rv.getDataRange().getValues().slice(1).forEach(r => { if (String(r[1]).trim() && String(r[9]).trim() === '') waitingReview[String(r[0]).trim()] = (waitingReview[String(r[0]).trim()] || 0) + 1; });
  } catch (e) { /* no Review tab yet */ }
  const saved = {};
  const canvasTab = ss_().getSheetByName(TAB_CANVAS);
  if (canvasTab) canvasTab.getDataRange().getValues().slice(1).forEach(r => { if (r[0] && r[3] !== '') saved[String(r[0])] = 'saved:' + String(r[3]); });
  const logSh = ss_().getSheetByName(TAB_LOG);
  const logRows = logSh ? logSh.getDataRange().getValues().slice(1) : [];
  const unmatchedCount = {};
  key.forEach(r => { if (/^UNMATCHED: /.test(String(r[2]))) unmatchedCount[String(r[0])] = (unmatchedCount[String(r[0])] || 0) + 1; });

  // Drive: three levels, a few requests in all
  const L1 = driveKids_([READY, INBOX, DONE]);
  const under = (list, parentId) => list.filter(x => x.parents.indexOf(parentId) > -1);
  const byName = list => { const m = {}; list.forEach(x => { (m[x.name] = m[x.name] || []).push(x); }); return m; };
  const readyRuns = byName(under(L1, READY).filter(x => x.folder));
  const inboxRuns = byName(under(L1, INBOX).filter(x => x.folder));
  const doneRuns = byName(under(L1, DONE).filter(x => x.folder));
  const looseFiles = under(L1, INBOX).some(x => !x.folder);
  const L2 = driveKids_(L1.filter(x => x.folder).map(x => x.id));
  const level3Parents = [];
  Object.keys(readyRuns).forEach(n => readyRuns[n].forEach(fo => under(L2, fo.id).filter(x => x.folder).forEach(x => level3Parents.push(x.id))));
  Object.keys(doneRuns).forEach(n => doneRuns[n].forEach(fo => under(L2, fo.id).filter(x => x.folder && x.name === HELD_FOLDER).forEach(x => level3Parents.push(x.id))));
  const L3 = level3Parents.length ? driveKids_(level3Parents) : [];
  const filesUnder = (folders, all) => { let n = []; (folders || []).forEach(fo => { n = n.concat(under(all, fo.id).filter(x => !x.folder)); }); return n; };

  const names = [];
  const add = n => { n = String(n || '').trim(); if (n && names.indexOf(n) < 0) names.push(n); };
  pastNames.forEach(add);
  Object.keys(readyRuns).forEach(add);
  Object.keys(inboxRuns).forEach(n => add(isScanFolder_(n) ? runFromScanFolder_(n) : n));
  scanRows.forEach(r => add(r[0]));

  const list = names.map(run => {
    const g = graded[run] || [];
    const gradedSet = {};
    g.forEach(c => { gradedSet[c] = true; });
    // coded work: files in the Ready folder and its Added folders
    const codes = {};
    let codedFiles = 0, keyFile = null, notesFile = null;
    (readyRuns[run] || []).forEach(fo => {
      const top = under(L2, fo.id);
      const files = top.filter(x => !x.folder);
      top.filter(x => x.folder).forEach(sub => { under(L3, sub.id).filter(x => !x.folder).forEach(x => files.push(x)); });
      files.forEach(x => { const m = x.name.match(/^(S-\d{4})(-\d+)?\.(pdf|png|jpe?g)$/i); if (m) { codedFiles++; codes[m[1].toUpperCase()] = true; } });
      if (!keyFile) keyFile = top.filter(x => !x.folder && /answer key/i.test(x.name))[0] || null;
      if (!notesFile) notesFile = top.filter(x => !x.folder && /^grading notes/i.test(x.name))[0] || null;   // version 23
    });
    const heldList = [];
    (doneRuns[run] || []).forEach(fo => under(L2, fo.id).filter(x => x.folder && x.name === HELD_FOLDER).forEach(h => {
      under(L3, h.id).filter(x => !x.folder).forEach(x => heldList.push({ id: x.id, name: x.name, url: fileUrl(x.id), canSend: /^S-\d{4}(-\d+)?\.pdf$/i.test(x.name) }));
    }));
    const students = Object.keys(codes).length;
    const a = {
      run: run,
      paper: !!inboxRuns[run + SCAN_SUFFIX],
      inbox: filesUnder(inboxRuns[run], L2).length,
      scans: filesUnder(inboxRuns[run + SCAN_SUFFIX], L2).length,
      toMatch: toMatch[run] || 0,
      toBuild: toBuild[run] || 0,
      coded: students,
      extraFiles: codedFiles - students,
      toGrade: Object.keys(codes).filter(c => !gradedSet[c]).length,
      onKey: !!onKey[run],
      checks: checksFor_(run, logRows, dismissedFor_(run)).filter(c => c.kind === 'missing' || !gradedSet[c.code]).length + (unmatchedCount[run] || 0),
      held: heldList.length,
      graded: g.length,
      review: waitingReview[run] || 0,
      toSend: g.filter(c => sentSig[c] !== sig[c]).length,
      canvas: saved[run] || '',
      key: keyFile ? { name: keyFile.name, url: fileUrl(keyFile.id) } : null,
      notes: !!notesFile,
      heldFiles: heldList
    };
    if (a.toMatch || a.scans) a.paper = true;
    a.next = a.inbox || a.scans ? 'clean' : a.toMatch ? 'match' : a.toBuild ? 'clean' : a.toGrade ? 'rowan'
      : a.review ? 'review' : a.toSend ? 'send' : a.graded ? 'finish' : (a.paper || a.canvas || !a.onKey) ? 'get' : 'old';
    return a;
  });
  let auto = false;
  try { auto = autoRunIsOn_(); } catch (e) { auto = false; }
  return {
    canvas: canvas, canvasBase: props.getProperty('CANVAS_BASE') || '', autoRun: auto,
    looseFiles: looseFiles,
    links: { sheet: ss_().getUrl(), inbox: folderUrl(INBOX), ready: folderUrl(READY), done: folderUrl(DONE), grades: book ? book.getUrl() : '' },
    assignments: list
  };
}

/** The version 16 status (one folder at a time). Kept as a backup in case the fast one can't run. */
function webStatusSlow_() {
  const f = folders_();
  const props = PropertiesService.getDocumentProperties();
  let canvas = false;
  try { canvasSettings_(); canvas = true; } catch (e) { canvas = false; }
  const names = [];
  const add = n => { n = String(n || '').trim(); if (n && names.indexOf(n) < 0) names.push(n); };
  pastAssignments_().forEach(add);
  [f.ready, f.inbox].forEach(parent => { const it = parent.getFolders(); while (it.hasNext()) { const fo = it.next(); if (!fo.isTrashed()) add(isScanFolder_(fo.getName()) ? runFromScanFolder_(fo.getName()) : fo.getName()); } });
  const scanSh = ss_().getSheetByName(TAB_SCANS);
  if (scanSh) scanSh.getDataRange().getValues().slice(1).forEach(r => add(r[0]));

  const key = ss_().getSheetByName(TAB_KEY).getDataRange().getValues().slice(1);
  const runOf = {};
  key.forEach(r => { runOf[String(r[1])] = String(r[0]); });
  const graded = {}, sig = {};
  const grades = claudeGradesSheet_();
  if (grades) grades.getDataRange().getValues().slice(1).forEach(r => {
    const code = String(r[1]).trim();
    if (!code || r[2] === '') return;
    const run = String(r[0]).trim();
    (graded[run] = graded[run] || []).push(code);
    sig[code] = String(r[2]) + '||' + String(r[3] || '');
  });
  const sentSig = {};
  const res = ss_().getSheetByName(TAB_RESULTS);
  if (res) res.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) sentSig[String(r[0]).trim()] = String(r[6] || ''); });
  const waitingReview = {};
  try {
    const book = claudeGradesBook_();
    const rv = book ? book.getSheetByName(REVIEW_TAB) : null;
    if (rv) rv.getDataRange().getValues().slice(1).forEach(r => { if (String(r[1]).trim() && String(r[9]).trim() === '') waitingReview[String(r[0]).trim()] = (waitingReview[String(r[0]).trim()] || 0) + 1; });
  } catch (e) { /* no Review tab yet */ }
  const plan = ss_().getSheetByName(TAB_PLAN);
  const toBuild = {};
  if (plan) plan.getDataRange().getValues().slice(1).forEach(r => { if (r[6] === 'Ready to save') toBuild[r[0]] = (toBuild[r[0]] || 0) + 1; });
  const saved = {};
  const canvasTab = ss_().getSheetByName(TAB_CANVAS);
  if (canvasTab) canvasTab.getDataRange().getValues().slice(1).forEach(r => { if (r[0] && r[3] !== '') saved[String(r[0])] = 'saved:' + String(r[3]); });

  const logSh = ss_().getSheetByName(TAB_LOG);
  const logRows = logSh ? logSh.getDataRange().getValues().slice(1) : [];
  const unmatchedCount = {};
  key.forEach(r => { if (/^UNMATCHED: /.test(String(r[2]))) unmatchedCount[String(r[0])] = (unmatchedCount[String(r[0])] || 0) + 1; });
  const list = names.map(run => {
    const readyFo = (() => { const it = f.ready.getFoldersByName(run); while (it.hasNext()) { const x = it.next(); if (!x.isTrashed()) return x; } return null; })();
    const doneFo = (() => { const it = f.done.getFoldersByName(run); while (it.hasNext()) { const x = it.next(); if (!x.isTrashed()) return x; } return null; })();
    const g = graded[run] || [];
    const coded = readyFo ? codedIn_(readyFo) : { students: 0, files: 0, codes: {} };
    const gradedSet = {};
    g.forEach(c => { gradedSet[c] = true; });
    const a = {
      run: run,
      paper: f.inbox.getFoldersByName(run + SCAN_SUFFIX).hasNext(),
      inbox: filesIn_(f.inbox, run),
      scans: filesIn_(f.inbox, run + SCAN_SUFFIX),
      toMatch: unplannedPages_(run).length,
      toBuild: toBuild[run] || 0,
      coded: coded.students,
      extraFiles: coded.files - coded.students,
      toGrade: Object.keys(coded.codes).filter(c => !gradedSet[c]).length,
      onKey: pastAssignments_().indexOf(run) > -1,
      checks: checksFor_(run, logRows, dismissedFor_(run)).filter(c => c.kind === 'missing' || !gradedSet[c.code]).length + (unmatchedCount[run] || 0),
      held: doneFo ? filesIn_(doneFo, HELD_FOLDER) : 0,
      graded: g.length,
      review: waitingReview[run] || 0,
      toSend: g.filter(c => sentSig[c] !== sig[c]).length,
      canvas: saved[run] || '',
      key: (() => { if (!readyFo) return null; const k = keyFiles_(readyFo)[0]; return k ? { name: k.getName(), url: k.getUrl() } : null; })(),
      heldFiles: (() => {
        if (!doneFo) return [];
        const out = [];
        const it = doneFo.getFoldersByName(HELD_FOLDER);
        while (it.hasNext()) { const fo = it.next(); if (fo.isTrashed()) continue; const fi = fo.getFiles(); while (fi.hasNext()) { const x = fi.next(); out.push({ id: x.getId(), name: x.getName(), url: x.getUrl(), canSend: /^S-\d{4}(-\d+)?\.pdf$/i.test(x.getName()) }); } }
        return out;
      })()
    };
    if (a.toMatch || a.scans) a.paper = true;
    a.next = a.inbox || a.scans ? 'clean' : a.toMatch ? 'match' : a.toBuild ? 'clean' : a.toGrade ? 'rowan'
      : a.review ? 'review' : a.toSend ? 'send' : a.graded ? 'finish' : (a.paper || a.canvas || !a.onKey) ? 'get' : 'old';
    return a;
  });
  let auto = false;
  try { auto = autoRunIsOn_(); } catch (e) { auto = false; }
  return {
    canvas: canvas, canvasBase: props.getProperty('CANVAS_BASE') || '', autoRun: auto,
    looseFiles: hasFiles_(f.inbox),
    links: { sheet: ss_().getUrl(), inbox: f.inbox.getUrl(), ready: f.ready.getUrl(), done: f.done.getUrl(), grades: (claudeGradesBook_() || { getUrl: () => '' }).getUrl() },
    assignments: list
  };
}

/** The Canvas list for the page's pickers (same as the Sheet's dropdown). */
function webCanvasOptions() {
  try { canvasSettings_(); } catch (e) { return { connected: false, options: [] }; }
  return { connected: true, options: canvasPickerOptions_() };
}
function webNameOptions() { return pastAssignments_(); }

/* ---------- the buttons ---------- */

function webPull(value) { return withWebUi_(0, () => pullSubmissionsFor_(resolveCanvasAssignment_(String(value)))); }
function webSend(value, yesCount) { return withWebUi_(yesCount, () => sendGradesFor_(resolveCanvasAssignment_(String(value)))); }
function webFinish(run, yesCount) {
  run = String(run || '');
  if (pastAssignments_().indexOf(run) < 0) return { ok: false, error: 'That assignment is not on the Key tab.' };
  return withWebUi_(yesCount, () => finishAssignmentFor_(run));
}
function webBringIn() {
  const n = bringInClaudeGrades_();
  return { ok: true, text: n < 0 ? 'There is no Claude Grades Sheet yet.' : n + ' grade(s) added or updated in the Results tab.' };
}
function webScanFolder(value) {
  try {
    value = String(value || '').trim();
    const run = (/^(saved|canvas):/.test(value) || /courses\/\d+\/assignments\/\d+/.test(value)) ? resolveCanvasAssignment_(value).name : resolveAssignmentName_(value);
    if (!run) throw new Error('Choose a test from the list, or type a name.');
    return Object.assign({ ok: true, run: run }, scanFolderMessage_(run));
  }
  catch (e) { return { ok: false, error: e.message }; }
}
function webAutoRun(on) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'autoProcess').forEach(t => ScriptApp.deleteTrigger(t));
  if (on) ScriptApp.newTrigger('autoProcess').timeBased().everyMinutes(10).create();
  return { ok: true, text: on ? 'Auto-run is on. Every 10 minutes the tool cleans what is waiting in the Inbox folders.' : 'Auto-run is off. Use "Clean the Inbox now" when you want files cleaned.' };
}
function webConnectCanvas(base, token) {
  const m = String(base || '').trim().match(/^(https?:\/\/[^\/\s]+)/i);
  if (!m) return { ok: false, error: 'That does not look like a web address. It should look like https://ktufsd.instructure.com' };
  if (!String(token || '').trim()) return { ok: false, error: 'Paste your Canvas token.' };
  PropertiesService.getDocumentProperties().setProperty('CANVAS_BASE', m[1].toLowerCase());
  PropertiesService.getUserProperties().setProperty('CANVAS_TOKEN', String(token).trim());
  try {
    canvasFetch_('/api/v1/users/self');
    ensureSheet_(TAB_CANVAS, ['Assignment', 'Canvas link', 'Course ID', 'Assignment ID', 'Last pulled']);
    ensureSheet_(TAB_PULLS, ['Assignment', 'Canvas ID', 'Attempt', 'Pulled']);
    return { ok: true, text: 'Connected to Canvas.' };
  } catch (e) {
    PropertiesService.getUserProperties().deleteProperty('CANVAS_TOKEN');
    return { ok: false, error: e.message };
  }
}
function webDisconnectCanvas() {
  PropertiesService.getUserProperties().deleteProperty('CANVAS_TOKEN');
  return { ok: true, text: 'Disconnected. You can also delete the token in Canvas (Account > Settings > Approved Integrations).' };
}
function webUpdateRoster(links) {
  try { return { ok: true, text: updateRoster_(links) }; }
  catch (e) { return { ok: false, error: e.message }; }
}

/**
 * Clean the Inbox (and save matched paper tests), the same as 2. Process the Inbox.
 * Long jobs stop at Google's time limit and say more: true; the page then runs it again.
 * The answer is also kept for 10 minutes under token, in case the page doesn't get it directly.
 */
async function webProcess(token, looseRun) {
  let out;
  try {
    out = await webProcess_(looseRun);
  } catch (e) {
    out = { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
  try { CacheService.getUserCache().put('web:' + token, JSON.stringify(out), 600); } catch (e) { /* ignore */ }
  return out;
}
function webResult(token) {
  const v = CacheService.getUserCache().get('web:' + token);
  return v ? JSON.parse(v) : null;
}

async function webProcess_(looseRun) {
  const start = Date.now();
  const f = folders_();
  const roster = readRoster_();
  if (!roster.length) return { ok: false, error: 'The Roster tab is empty. Open Settings > Update roster from Canvas first.' };
  const props = PropertiesService.getDocumentProperties();
  const results = [];
  let run = props.getProperty('ACTIVE_RUN') || (looseRun ? resolveAssignmentName_('typed:' + looseRun) : '');
  if (hasFiles_(f.inbox) || run) {
    if (!run) return { ok: true, needLoose: true, assignments: pastAssignments_() };
    props.setProperty('ACTIVE_RUN', run);
    const r = await processBatch_(run, f.inbox, f, roster, start);
    results.push(r);
    if (r.timedOut) return { ok: true, more: true, text: (r.progress || (r.made + ' file(s) cleaned so far')) + ' for ' + r.run + '. Still going...' };
    props.deleteProperty('ACTIVE_RUN');
  }
  const subs = assignmentInboxes_(f.inbox);
  for (let i = 0; i < subs.length; i++) {
    const r = await batchFor_(subs[i], f, roster, start);
    results.push(r);
    if (r.timedOut) return { ok: true, more: true, text: (r.progress || (r.made + ' file(s) cleaned so far')) + ' for ' + r.run + '. Still going...' };
  }
  const built = await buildScanPlans_(f, start);
  if (built.done) results.push({ text: built.text });
  if (built.stopped) return { ok: true, more: true, text: built.text + ' Still going...' };
  if (!results.length) return { ok: true, text: 'Nothing to clean. The Inbox is empty.' };
  const pictures = results.some(r => r.flagged && r.scansRead === undefined && !r.text);
  return { ok: true, text: results.map(doneLine_).join('\n\n') + (pictures ? '\n\nSome files need a quick look (pictures like Canva slides can still show a name). Open Settings > 2 Ready for Claude to skim them before Rowan grades.' : '') };
}

/** The Review or Match window, to show inside the web page. */
function webView(name, run) {
  const html = name === 'review' ? reviewHtml_() : name === 'match' ? matchHtml_() : name === 'viewer' ? viewerHtml_(run) : '';
  if (!html) throw new Error('Unknown page.');
  if (name === 'review') ensureReviewTab_();
  if (name === 'viewer') ensureMarksTab_();
  // Inside the app, the window borrows the page's connection to the server, and "close" goes back home
  const shim = '<script>window.google={script:{host:{close:function(){parent.closeView();}}}};' +
    'Object.defineProperty(window.google.script,"run",{get:function(){return parent.google.script.run;}});</scr' + 'ipt>';
  return html.replace('<meta charset="utf-8">', '<meta charset="utf-8">' + shim);
}

/* ---------- version 14: uploads, answer keys and held tests, all from the app ---------- */

const KEY_PREFIX = 'Answer key - ';
const UPLOAD_CHUNK = 4 * 1024 * 1024;   // the page sends files in 4 MB pieces (Google's limit per call)

/** True if a Ready folder already has coded work (an answer key doesn't count). */
function hasWork_(folder) {
  const it = folder.getFiles();
  while (it.hasNext()) if (/^S-\d{4}/i.test(it.next().getName())) return true;
  return folder.getFolders().hasNext();
}

function keyFiles_(folder) {
  const out = [];
  const it = folder.getFiles();
  while (it.hasNext()) { const fl = it.next(); if (!fl.isTrashed() && /answer key/i.test(fl.getName())) out.push(fl); }
  return out;
}

function readyFolderFor_(run) { return folderByName_(folders_().ready, run); }

/* ---------- version 23: notes for Rowan and a blank copy of the test ---------- */
const NOTES_PREFIX = 'Grading notes - ';
const BLANK_PREFIX = 'Blank test - ';
function filesMatching_(folder, re) {
  const out = [];
  const it = folder.getFiles();
  while (it.hasNext()) { const fl = it.next(); if (!fl.isTrashed() && re.test(fl.getName())) out.push(fl); }
  return out;
}
/** The test's notes for Rowan (text) and its blank copy, if any. Through rpc. */
function webGetNotes(run) {
  const fo = readyFolderFor_(String(run || '').trim());
  const n = filesMatching_(fo, /^grading notes/i)[0];
  const b = filesMatching_(fo, /^blank test/i)[0];
  return { text: n ? n.getBlob().getDataAsString('UTF-8') : '', blank: b ? { name: b.getName(), url: b.getUrl() } : null };
}
/** Saves (or clears) the notes for Rowan. Through rpc. */
function webSaveNotes(run, text) {
  run = String(run || '').trim();
  if (!run) return { ok: false, error: 'Which test are these notes for?' };
  const fo = readyFolderFor_(run);
  filesMatching_(fo, /^grading notes/i).forEach(fl => fl.setTrashed(true));
  text = String(text || '').trim();
  if (!text) { log_(run, '-', 'Grading notes removed', 'In the app.'); return { ok: true, text: 'Notes removed.' }; }
  fo.createFile(NOTES_PREFIX + run.replace(/[\/\\]/g, '-') + '.txt', text, MimeType.PLAIN_TEXT);
  log_(run, '-', 'Grading notes saved', 'In the app.');
  return { ok: true, text: 'Notes saved. Rowan reads them before grading, and they win over the key.' };
}

/** Saves typed answers as the test's answer key (replacing any earlier key). */
function webSaveKeyText(run, text) {
  run = String(run || '').trim();
  text = String(text || '').trim();
  if (!run) return { ok: false, error: 'Which test is this key for?' };
  if (!text) return { ok: false, error: 'Type the answers first.' };
  const folder = readyFolderFor_(run);
  keyFiles_(folder).forEach(fl => fl.setTrashed(true));
  folder.createFile(Utilities.newBlob(text, 'text/plain', KEY_PREFIX + run.replace(/[\/\\]/g, '-') + '.txt'));
  log_(run, '-', 'Answer key saved', 'Typed in the app.');
  return { ok: true, text: 'Answer key saved for "' + run + '". Rowan will use it when you say "grade."' };
}

function webRemoveKey(run) {
  keyFiles_(readyFolderFor_(String(run))).forEach(fl => fl.setTrashed(true));
  return { ok: true, text: 'Answer key removed. Rowan will write one from the test and show it to you.' };
}

/**
 * Starts an upload from the page straight into Drive (Google's resumable upload).
 * kind: 'scan' (into "1 Inbox/<test> (scans)") or 'key' (into "2 Ready for Claude/<test>", replacing any earlier key).
 */
function webUploadStart(kind, run, name, mime, size) {
  run = String(run || '').trim();
  if (!run) throw new Error('Which test is this for?');
  if (kind !== 'scan' && kind !== 'key' && kind !== 'blank') throw new Error('Unknown upload.');
  const f = folders_();
  const parent = kind === 'scan' ? folderByName_(f.inbox, run + SCAN_SUFFIX) : folderByName_(f.ready, run);
  const ext = (String(name).match(/\.[A-Za-z0-9]+$/) || [''])[0].toLowerCase();
  const fileName = kind === 'key' ? KEY_PREFIX + run.replace(/[\/\\]/g, '-') + ext
    : kind === 'blank' ? BLANK_PREFIX + run.replace(/[\/\\]/g, '-') + ext
    : String(name || 'scan.pdf').replace(/[\/\\]/g, '-');
  const resp = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id', {
    method: 'post', contentType: 'application/json; charset=UTF-8', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Upload-Content-Type': mime || 'application/octet-stream', 'X-Upload-Content-Length': String(size) },
    payload: JSON.stringify({ name: fileName, parents: [parent.getId()], mimeType: mime || 'application/octet-stream' })
  });
  const h = resp.getHeaders() || {};
  const url = h['Location'] || h['location'];
  if (resp.getResponseCode() >= 300 || !url) throw new Error('Google Drive would not start the upload (error ' + resp.getResponseCode() + ').');
  const token = Utilities.getUuid();
  CacheService.getUserCache().put('up:' + token, JSON.stringify({ url: url, kind: kind, run: run, parent: parent.getId(), size: Number(size) }), 21600);
  return { token: token, chunk: UPLOAD_CHUNK };
}

/** One 4 MB piece of an upload. Returns { done } (and the file when the last piece is in). */
function webUploadChunk(token, start, b64) {
  const s = JSON.parse(CacheService.getUserCache().get('up:' + token) || 'null');
  if (!s) throw new Error('The upload timed out. Try again.');
  const bytes = Utilities.base64Decode(String(b64));
  const end = Number(start) + bytes.length - 1;
  const resp = UrlFetchApp.fetch(s.url, { method: 'put', payload: bytes, muteHttpExceptions: true,
    headers: { 'Content-Range': 'bytes ' + start + '-' + end + '/' + s.size } });
  const code = resp.getResponseCode();
  if (code === 308) return { done: false };
  if (code !== 200 && code !== 201) throw new Error('Google Drive stopped the upload (error ' + code + '). Try again.');
  const id = JSON.parse(resp.getContentText() || '{}').id;
  if (s.kind === 'key') {
    keyFiles_(DriveApp.getFolderById(s.parent)).forEach(fl => { if (fl.getId() !== id) fl.setTrashed(true); });
    log_(s.run, '-', 'Answer key saved', 'Uploaded in the app.');
  }
  if (s.kind === 'blank') {   // version 23: one blank test at a time
    filesMatching_(DriveApp.getFolderById(s.parent), /^blank test/i).forEach(fl => { if (fl.getId() !== id) fl.setTrashed(true); });
    log_(s.run, '-', 'Blank test saved', 'Uploaded in the app.');
  }
  return { done: true };
}

/** A held test you checked: it has no name, so it goes to "2 Ready for Claude/<test>" for Rowan. */
function webReleaseHeld(run, fileId) {
  run = String(run || '');
  const file = DriveApp.getFileById(String(fileId));
  const parents = file.getParents();
  let ok = false;
  while (parents.hasNext()) {
    const p = parents.next();
    if (p.getName() !== HELD_FOLDER) continue;
    const gp = p.getParents();
    while (gp.hasNext()) if (gp.next().getName() === run) ok = true;
  }
  if (!ok || !/^S-\d{4}(-\d+)?\.pdf$/i.test(file.getName())) return { ok: false, error: 'Only a held test named by its code can be sent to Rowan. Grade the others by hand.' };
  const target = readyFolderFor_(run);
  const code = file.getName().match(/^(S-\d{4})/i)[1];
  if (target.getFilesByName(file.getName()).hasNext()) file.setName(nextOutName_(code, '.pdf', target));
  file.moveTo(target);
  log_(run, code, 'Sent to Rowan', 'You checked this held test and moved it to Ready.');
  return { ok: true, text: code + ' is now in "2 Ready for Claude/' + run + '". Say "grade" to Rowan.' };
}

/* ---------- version 15: see the work, marked-up tests, grades and checks, all in the app ---------- */
/*
 * The viewer shows every coded file for an assignment with its grade and feedback (you can change them),
 * Rowan's review items, anything the tool flagged, and for tests Rowan's red-pen marks.
 * Marks live on the "Marks" tab of the Claude Grades Sheet (codes only), written by Rowan:
 * Assignment | Code | Page | X (0 left to 1 right) | Y (0 top to 1 bottom) | Mark (check, x, partial, note, score) | Text | Question | Points | Possible | Changed by Eric
 * Click a question's mark in the viewer to change it (full credit, no credit, or points); the score moves by the difference.
 */

const MARKS_TAB = 'Marks';
const MARKS_HEADERS = ['Assignment', 'Code', 'Page', 'X (0 left, 1 right)', 'Y (0 top, 1 bottom)', 'Mark (check, x, partial, note, score)', 'Text', 'Question', 'Points', 'Possible', 'Changed by Eric'];

function ensureMarksTab_() {
  const book = claudeGradesBook_();
  if (!book) return null;
  let sh = book.getSheetByName(MARKS_TAB);
  if (!sh) {
    sh = book.insertSheet(MARKS_TAB, book.getSheets().length);
    sh.getRange(1, 1, 1, MARKS_HEADERS.length).setValues([MARKS_HEADERS]).setFontWeight('bold').setBackground('#e8eef7');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Coded files in a test's Ready folder and its Added folders: [{ code, id, name, mime, size, added }]. */
function codedFiles_(readyFo) {
  const out = [];
  const add = (fo, added) => {
    const it = fo.getFiles();
    while (it.hasNext()) {
      const fl = it.next();
      const m = fl.getName().match(/^(S-\d{4})(-\d+)?\.(pdf|png|jpe?g)$/i);
      if (m && !fl.isTrashed()) out.push({ code: m[1].toUpperCase(), id: fl.getId(), name: fl.getName(), mime: fl.getMimeType(), size: fl.getSize(), added: added });
    }
  };
  add(readyFo, '');
  const sub = readyFo.getFolders();
  while (sub.hasNext()) { const fo = sub.next(); if (!fo.isTrashed()) add(fo, fo.getName()); }
  out.sort((a, b) => naturalCompare_(a.name, b.name));
  return out;
}

/** Plain-words reason for a flagged file, from the Log note. */
function checkReason_(note) {
  const n = String(note || '');
  if (/picture|Image:/i.test(n)) return 'Made from pictures. Look for names, faces or handwriting.';
  if (/swapped for a plain version/i.test(n)) return 'A page with a name was replaced by a plain text version. Check it still reads right.';
  if (/rebuilt from its text/i.test(n)) return 'Rebuilt from its text, so the layout may look different.';
  if (/already had a file/i.test(n)) return 'A second file from this student (a redo, or an extra file).';
  if (/Could not match/i.test(n)) return 'The tool could not tell whose this is.';
  return n;
}

/** Things to look over for one assignment, from the Log (minus the ones you marked fine): [{ code, file, reason, note, kind }]. */
function checksFor_(run, logRows, dismissed) {
  const out = [];
  logRows.forEach(r => {
    if (String(r[1]) !== run) return;
    const code = String(r[2]), result = String(r[3] || ''), note = String(r[4] || '');
    const cleanNote = note.replace(/^\d+ page\(s\)\.\s*/, '').trim();
    if (!/^S-\d{4}$/.test(code)) return;
    if (result.indexOf('Saved ') === 0 && cleanNote) {
      const file = result.slice(6).trim();
      if (dismissed[code + '|' + file]) return;
      out.push({ code: code, file: file, reason: checkReason_(cleanNote), note: cleanNote, kind: 'look' });
    } else if (result === 'No file made') {
      if (dismissed[code + '|-']) return;
      out.push({ code: code, file: '-', reason: 'No clean file was made, so Rowan can’t grade it. ' + cleanNote, note: cleanNote, kind: 'missing' });
    }
  });
  return out;
}
function dismissedFor_(run) {
  const d = {};
  String(PropertiesService.getDocumentProperties().getProperty('CHK::' + run) || '').split('\n').forEach(x => { if (x) d[x] = true; });
  return d;
}

/** Everything the viewer needs for one assignment. Runs in your school account, shown only to you. */
function viewerData(run) {
  run = String(run || '');
  const f = folders_();
  let readyFo = null;
  const it = f.ready.getFoldersByName(run);
  while (it.hasNext()) { const x = it.next(); if (!x.isTrashed()) { readyFo = x; break; } }
  const files = readyFo ? codedFiles_(readyFo) : [];

  const grades = {};
  const gs = claudeGradesSheet_();
  if (gs) gs.getDataRange().getValues().slice(1).forEach(r => {
    if (String(r[0]).trim() === run && String(r[1]).trim()) grades[String(r[1]).trim()] = { score: String(r[2]), fb: String(r[3] || ''), rubric: String(r[4] || '') };
  });
  const sent = {};
  const res = ss_().getSheetByName(TAB_RESULTS);
  if (res) res.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) sent[String(r[0]).trim()] = { status: String(r[5] || ''), sig: String(r[6] || '') }; });
  Object.keys(grades).forEach(c => {
    const g = grades[c], s = sent[c];
    g.sent = !s || !s.sig ? 'Not sent yet' : (s.sig === g.score + '||' + g.fb ? s.status : 'Changed since it was sent (Send again)');
  });

  const review = {};
  const book = claudeGradesBook_();
  const rv = book ? book.getSheetByName(REVIEW_TAB) : null;
  if (rv) rv.getDataRange().getValues().slice(1).forEach((r, i) => {
    if (String(r[0]).trim() !== run || !String(r[1]).trim()) return;
    (review[String(r[1]).trim()] = review[String(r[1]).trim()] || []).push({ row: i + 2, question: String(r[2]), possible: num_(r[3]), rowan: num_(r[4]), eric: num_(r[9]) });
  });
  const marks = {};
  const mk = book ? book.getSheetByName(MARKS_TAB) : null;
  if (mk) mk.getDataRange().getValues().slice(1).forEach((r, i) => {
    if (String(r[0]).trim() !== run || !String(r[1]).trim()) return;
    (marks[String(r[1]).trim()] = marks[String(r[1]).trim()] || []).push({ row: i + 2, page: num_(r[2]) || 1, x: num_(r[3]), y: num_(r[4]), mark: String(r[5] || 'note').toLowerCase().trim(), text: String(r[6] || ''), question: String(r[7] || ''),
      points: num_(r[8]), possible: num_(r[9]), changed: String(r[10] || '') });
  });

  const logSh = ss_().getSheetByName(TAB_LOG);
  const checks = {};
  checksFor_(run, logSh ? logSh.getDataRange().getValues().slice(1) : [], dismissedFor_(run)).forEach(c => { (checks[c.code] = checks[c.code] || []).push(c); });

  const unmatched = {};
  const names = {};   // version 16: code -> student name, shown only to you in this app
  ss_().getSheetByName(TAB_KEY).getDataRange().getValues().slice(1).forEach(r => {
    if (String(r[0]) !== run) return;
    if (/^UNMATCHED: /.test(String(r[2]))) unmatched[String(r[1])] = String(r[2]).replace(/^UNMATCHED: /, '');
    else if (String(r[2]).trim()) names[String(r[1])] = String(r[2]).trim();
  });
  const students = readRoster_().slice().sort((a, b) => a.raw.localeCompare(b.raw)).map(s => ({ id: s.id, label: s.raw }));
  const paper = f.inbox.getFoldersByName(run + SCAN_SUFFIX).hasNext();
  return { run: run, paper: paper, files: files, grades: grades, review: review, marks: marks, checks: checks, unmatched: unmatched, names: names,
           students: students };
}

/** One coded file for the viewer (only S-#### files in "2 Ready for Claude"). Big files come back as a Drive viewer link instead. */
function viewerFile(fileId) {
  const file = viewerCheck_(fileId);
  const out = { id: file.getId(), mime: file.getMimeType(), name: file.getName(), size: file.getSize() };
  if (file.getSize() <= 15 * 1024 * 1024) out.data = Utilities.base64Encode(file.getBlob().getBytes());
  return out;
}

/** Version 22: one piece of a big coded file (Drive sends only those bytes), for See the work. Through rpc. */
function viewerFileChunk(fileId, start, len) {
  const file = viewerCheck_(fileId);
  start = Math.max(0, Number(start) || 0);
  const end = Math.min(file.getSize(), start + (Number(len) || 4718592)) - 1;
  if (end < start) return '';
  const r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + file.getId() + '?alt=media', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), Range: 'bytes=' + start + '-' + end }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 206 && r.getResponseCode() !== 200) throw new Error('Drive would not send the file (error ' + r.getResponseCode() + ').');
  let bytes = r.getContent();
  if (r.getResponseCode() === 200 && bytes.length > end - start + 1) bytes = bytes.slice(start, end + 1);   // Drive ignored the range
  return Utilities.base64Encode(bytes);
}

/**
 * Version 22: covers what you boxed. pages: [{ page: 1-based, jpg: base64 of the whole page with white boxes painted }].
 * Those pages become flat pictures (the name is really gone); the other pages are copied as they were.
 * The coded file is replaced. Its original is still in Processed originals. Through rpc.
 */
async function webCoverFile(run, fileId, pages) {
  run = String(run || '');
  const file = viewerCheck_(fileId);
  const m = file.getName().match(/^(S-\d{4})(-\d+)?\.(pdf|png|jpe?g)$/i);
  if (!pages || !pages.length) return { ok: false, error: 'Draw a box over the name first.' };
  const covered = {};
  pages.forEach(p => { if (p && p.jpg) covered[Number(p.page)] = String(p.jpg); });
  const PDFLib = loadPdfLib_();
  const out = await PDFLib.PDFDocument.create();
  const addJpg = async (b64, W, H) => {
    const img = await out.embedJpg(new Uint8Array(Utilities.base64Decode(b64)));
    if (!W || !H) { const landscape = img.width > img.height; W = landscape ? 792 : 612; H = landscape ? 612 : 792; }
    const s = Math.min(W / img.width, H / img.height);
    out.addPage([W, H]).drawImage(img, { x: (W - img.width * s) / 2, y: (H - img.height * s) / 2, width: img.width * s, height: img.height * s });
  };
  const isPdf = /pdf/i.test(m[3]);
  let total = 1;
  if (isPdf) {
    const src = await PDFLib.PDFDocument.load(new Uint8Array(file.getBlob().getBytes()), { ignoreEncryption: true, updateMetadata: false });
    total = src.getPageCount();
    for (let i = 0; i < total; i++) {
      if (covered[i + 1]) {
        const pg = src.getPage(i);
        const sz = pg.getSize();
        const turned = Math.abs(((pg.getRotation().angle || 0) % 180)) === 90;
        await addJpg(covered[i + 1], turned ? sz.height : sz.width, turned ? sz.width : sz.height);
      } else {
        const [cp] = await out.copyPages(src, [i]);
        out.addPage(cp);
      }
    }
  } else {
    if (!covered[1]) return { ok: false, error: 'Draw a box over the name first.' };
    await addJpg(covered[1], 0, 0);
  }
  out.setTitle(m[1]); out.setAuthor(''); out.setSubject(''); out.setKeywords([]);
  out.setCreator('Student Work Anonymizer'); out.setProducer('Student Work Anonymizer');
  const bytes = await out.save();
  const parent = file.getParents().next();
  const name = isPdf ? file.getName() : nextOutName_(m[1] + (m[2] || ''), '.pdf', parent);
  const nf = parent.createFile(u8ToBlob_(bytes, name));
  deleteForever_(file.getId());
  const n = Object.keys(covered).length;
  log_(run, m[1], 'Covered in the app', n + ' page(s) painted over and saved as pictures (' + name + ').');
  return { ok: true, id: nf.getId(), code: m[1], text: 'Covered. ' + n + ' page' + (n === 1 ? '' : 's') + ' saved as a flat picture, so what you boxed is gone for good.' };
}

/** A coded file (S-####) inside "2 Ready for Claude" (or one of its folders), or an error. */
function viewerCheck_(fileId) {
  const file = DriveApp.getFileById(String(fileId));
  if (!/^S-\d{4}(-\d+)?\.(pdf|png|jpe?g)$/i.test(file.getName())) throw new Error('That is not a coded file.');
  const ready = folders_().ready.getId();
  let ok = false;
  const p = file.getParents();
  while (p.hasNext()) {
    const fo = p.next();
    const gp = fo.getParents();
    while (gp.hasNext()) { const g = gp.next(); if (g.getId() === ready) ok = true; else { const g2 = g.getParents(); while (g2.hasNext()) if (g2.next().getId() === ready) ok = true; } }
  }
  if (!ok) throw new Error('That file is not in "2 Ready for Claude".');
  return file;
}

/** Saves a score and feedback you changed in the viewer (into the Claude Grades Sheet, where Send picks it up). */
function webSaveGrade(run, code, score, fb) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const gs = claudeGradesSheet_();
    if (!gs) return { ok: false, error: 'The Claude Grades Sheet could not be found.' };
    const data = gs.getDataRange().getValues();
    const s = String(score).trim();
    const value = s !== '' && !isNaN(Number(s)) ? Number(s) : s;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === run && String(data[i][1]).trim() === code) {
        gs.getRange(i + 1, 3, 1, 2).setValues([[value, String(fb || '')]]);
        return { ok: true, text: 'Saved. It goes to Canvas the next time you click Send.' };
      }
    }
    gs.appendRow([run, code, value, String(fb || ''), 'Added by you in the app']);
    return { ok: true, text: 'Saved. It goes to Canvas the next time you click Send.' };
  } finally {
    lock.releaseLock();
  }
}

/** First word of a question label, for matching marks and Review rows: "Q4 Explain tides" -> "q4". */
function qKey_(q) { return String(q || '').trim().split(/[ .:)]/)[0].toLowerCase(); }

/** A mark's points and points possible, from its columns (or its text like "1/2"). */
function markPoints_(r) {
  const t = String(r[6] || '').match(/([0-9.]+)\s*\/\s*([0-9.]+)/);
  let possible = num_(r[9]);
  if (possible === null && t) possible = Number(t[2]);
  let points = num_(r[8]);
  const kind = String(r[5] || '').toLowerCase().trim();
  if (points === null) points = kind === 'check' ? possible : kind === 'x' ? 0 : (t ? Number(t[1]) : null);
  return { points: points, possible: possible };
}

/**
 * You changed a question's mark in the viewer: full credit, no credit, or points.
 * The mark changes and the student's score moves by the difference. If the question is also on the
 * Review tab, it counts as your Review decision (so nothing is counted twice).
 */
function webOverrideMark(run, code, row, question, points) {
  const book = claudeGradesBook_();
  const sh = book ? book.getSheetByName(MARKS_TAB) : null;
  if (!sh) return { ok: false, error: 'No marks for this test.' };
  const r = sh.getRange(Number(row), 1, 1, MARKS_HEADERS.length).getValues()[0];
  if (String(r[0]).trim() !== run || String(r[1]).trim() !== code || String(r[7]) !== String(question)) return { ok: false, error: 'The marks changed. Go back home and open See the work again.' };
  const mp = markPoints_(r);
  const p = num_(points);
  if (p === null || p < 0 || (mp.possible !== null && p > mp.possible)) return { ok: false, error: 'Points have to be a number from 0 to ' + (mp.possible === null ? 'the points possible' : mp.possible) + '.' };
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, h:mm a');
  let score = '';
  // Is this question waiting (or decided) on the Review tab? Then this is the Review decision.
  const rv = book.getSheetByName(REVIEW_TAB);
  let reviewRow = -1;
  if (rv) rv.getDataRange().getValues().forEach((x, i) => { if (i && String(x[0]).trim() === run && String(x[1]).trim() === code && qKey_(x[2]) === qKey_(question)) reviewRow = i + 1; });
  if (reviewRow > 0) {
    const x = rv.getRange(reviewRow, 1, 1, 3).getValues()[0];
    const out = reviewDecide(reviewRow, { assignment: run, code: code, question: String(x[2]) }, p);
    score = out.score;
  } else {
    const before = mp.points === null ? 0 : mp.points;
    const grades = claudeGradesSheet_();
    const data = grades.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === run && String(data[i][1]).trim() === code && num_(data[i][2]) !== null) {
        score = String(Math.round((num_(data[i][2]) + p - before) * 100) / 100);
        grades.getRange(i + 1, 3).setValue(Number(score));
        break;
      }
    }
  }
  const possible = mp.possible;
  const kind = possible !== null && p === possible ? 'check' : p === 0 ? 'x' : 'partial';
  sh.getRange(Number(row), 6, 1, 6).setValues([[kind, kind === 'partial' ? p + '/' + possible : '', String(question), p, possible === null ? '' : possible, stamp]]);
  return { ok: true, mark: kind, text: kind === 'partial' ? p + '/' + possible : '', points: p, possible: possible, score: score, changed: stamp };
}

function webDismissCheck(run, code, file) {
  const props = PropertiesService.getDocumentProperties();
  const k = 'CHK::' + run;
  const cur = String(props.getProperty(k) || '');
  props.setProperty(k, cur + (cur ? '\n' : '') + code + '|' + (file || '-'));
  return { ok: true };
}

/**
 * A file the tool couldn't match: you pick the student. The file is renamed to that student's code,
 * and any grade, review or marks Rowan wrote under the old code move with it.
 */
function webLabelUnmatched(run, code, studentId) {
  const roster = readRoster_();
  const s = roster.filter(x => x.id === String(studentId))[0];
  if (!s) return { ok: false, error: 'Pick a student from the list.' };
  const real = ensureCodes_(run, roster)[studentKey_(s)];
  const f = folders_();
  let readyFo = null;
  const it = f.ready.getFoldersByName(run);
  while (it.hasNext()) { const x = it.next(); if (!x.isTrashed()) readyFo = x; }
  if (readyFo) codedFiles_(readyFo).filter(x => x.code === code).forEach(x => {
    const fl = DriveApp.getFileById(x.id);
    const parent = fl.getParents().next();
    fl.setName(nextOutName_(real, x.name.slice(x.name.lastIndexOf('.')), parent));
  });
  const book = claudeGradesBook_();
  if (book) [book.getSheets()[0], book.getSheetByName(REVIEW_TAB), book.getSheetByName(MARKS_TAB)].forEach(sh => {
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) if (String(data[i][0]).trim() === run && String(data[i][1]).trim() === code) sh.getRange(i + 1, 2).setValue(real);
  });
  deleteRows_(ss_().getSheetByName(TAB_KEY), r => r[0] === run && String(r[1]) === code && /^UNMATCHED: /.test(String(r[2])), 2);
  bumpKey_(run, real, 'renamed', 'Was unmatched file ' + code + ', matched in the app.');
  log_(run, real, 'Matched in the app', 'Was ' + code + '.');
  return { ok: true, text: 'Matched. That file is now ' + real + '.' };
}

/**
 * Version 21: moves one coded file to the student you pick (fixes a wrong or missing match).
 * Its grade, Review and Marks rows go with it when the old code has no other files and the new
 * student has no grade yet. Called from the page through the starter's rpc().
 */
function webMoveFile(run, fileId, studentId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    run = String(run || '');
    const roster = readRoster_();
    const s = roster.filter(x => x.id === String(studentId))[0];
    if (!s) return { ok: false, error: 'Pick a student from the list.' };
    const real = ensureCodes_(run, roster)[studentKey_(s)];
    const f = folders_();
    let readyFo = null;
    const it = f.ready.getFoldersByName(run);
    while (it.hasNext()) { const x = it.next(); if (!x.isTrashed()) readyFo = x; }
    if (!readyFo) return { ok: false, error: 'This assignment has no folder in "2 Ready for Claude".' };
    const files = codedFiles_(readyFo);
    const me = files.filter(x => x.id === String(fileId))[0];
    if (!me) return { ok: false, error: 'That file is not in "2 Ready for Claude/' + run + '" any more. Close See the work and open it again.' };
    const from = me.code;
    if (from === real) return { ok: true, text: 'That file already belongs to ' + s.raw + '.' };
    const file = DriveApp.getFileById(me.id);
    const parent = file.getParents().next();
    const newName = nextOutName_(real, me.name.slice(me.name.lastIndexOf('.')), parent);
    file.setName(newName);
    const others = files.filter(x => x.code === from && x.id !== me.id).length;
    let note = '';
    const book = claudeGradesBook_();
    if (others) {
      note = ' ' + from + ' still has other files, so its grade stayed where it was. Check both scores.';
    } else if (book) {
      const gs = book.getSheets()[0];
      const g = gs.getDataRange().getValues();
      const has = c => g.some((r, i) => i && String(r[0]).trim() === run && String(r[1]).trim() === c && String(r[2]) !== '');
      if (has(real) && has(from)) {
        note = ' ' + s.raw + ' already had a grade, so the grade from this file was left under ' + from + '. Check the score.';
      } else {
        [gs, book.getSheetByName(REVIEW_TAB), book.getSheetByName(MARKS_TAB)].forEach(sh => {
          if (!sh) return;
          const data = sh.getDataRange().getValues();
          for (let i = 1; i < data.length; i++) if (String(data[i][0]).trim() === run && String(data[i][1]).trim() === from) sh.getRange(i + 1, 2).setValue(real);
        });
      }
    }
    if (!others) deleteRows_(ss_().getSheetByName(TAB_KEY), r => r[0] === run && String(r[1]) === from && /^UNMATCHED: /.test(String(r[2])), 2);
    bumpKey_(run, real, newName, 'File moved here from ' + from + ' in the app.');
    log_(run, real, 'Moved in the app', 'Was ' + from + '. Now ' + newName + '.');
    return { ok: true, code: real, text: 'Moved. That file now belongs to ' + s.raw + ' (' + real + ').' + note };
  } finally {
    lock.releaseLock();
  }
}

/** The class list for Settings: names, nicknames to hide, and "full name only". */
function webRosterList() {
  const sh = ss_().getSheetByName(TAB_ROSTER);
  return sh.getDataRange().getValues().slice(1).map((r, i) => ({ row: i + 2, name: String(r[0] || ''), id: String(r[1] || ''), nick: String(r[2] || ''), full: /^y/i.test(String(r[3] || '')) }))
    .filter(x => x.name);
}
function webRosterSave(row, id, nick, full) {
  const sh = ss_().getSheetByName(TAB_ROSTER);
  const r = sh.getRange(Number(row), 1, 1, 4).getValues()[0];
  if (String(r[1]) !== String(id)) return { ok: false, error: 'The class list changed. Close Settings and open it again.' };
  sh.getRange(Number(row), 3, 1, 2).setValues([[String(nick || ''), full ? 'Y' : '']]);
  return { ok: true };
}

function viewerHtml_(run) {
  return '<!doctype html><html><head><meta charset="utf-8">' +
    '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Caveat:wght@600;700&display=swap">' +
    '<script src="' + PDFJS_URL + '"></scr' + 'ipt>' +
    '<script>var PDFJS_WORKER=' + JSON.stringify(PDFJS_WORKER_URL) + ';var RUN=' + JSON.stringify(String(run)) + ';</scr' + 'ipt></head>' +
    VIEWER_PAGE_HTML + '</html>';
}

const VIEWER_PAGE_HTML = `
<body>
<style>
  :root{--fg:#1f2933;--muted:#5f6b7a;--line:#dde3ea;--soft:#f4f6f9;--blue:#1a73e8;--blue-soft:#e8f0fe;--green:#137333;--green-soft:#e6f4ea;--red:#c5221f;--amber:#b06000;--amber-soft:#fef3e0}
  *{box-sizing:border-box}
  body{margin:0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:var(--fg);height:100vh;display:flex;flex-direction:column;background:#fff}
  .top{display:flex;align-items:center;gap:10px;padding:8px 14px;border-bottom:1px solid var(--line);flex-wrap:wrap}
  .top select{font-size:14px;padding:5px}
  .top .who{font-weight:bold;font-size:15px}
  button{font-size:13px;padding:6px 11px;border-radius:7px;border:1px solid var(--line);background:#fff;cursor:pointer;color:var(--fg)}
  button.go{background:var(--blue);border-color:var(--blue);color:#fff;font-weight:bold}
  button:disabled{opacity:.45;cursor:default}
  .grid{flex:1;display:grid;grid-template-columns:220px 1fr 330px;min-height:0}
  .grid>*{min-height:0;overflow:auto}
  .list{border-right:1px solid var(--line);background:var(--soft)}
  .list div{padding:8px 12px;border-bottom:1px solid var(--line);cursor:pointer;display:flex;justify-content:space-between;gap:6px}
  .list div.on{background:var(--blue-soft);font-weight:bold}
  .list .s{color:var(--muted);font-weight:normal}
  .list .dot{color:var(--amber);font-weight:bold}
  .pages{background:#e9edf2;padding:14px}
  .pg{position:relative;margin:0 auto 14px;max-width:860px;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.2)}
  .pg canvas,.pg img{width:100%;display:block}
  .pg .no{position:absolute;left:8px;top:6px;font-size:11px;color:var(--muted);background:rgba(255,255,255,.85);padding:1px 5px;border-radius:3px}
  .coverbar{position:sticky;top:0;z-index:8;margin:0 auto 12px;max-width:860px;background:var(--amber-soft);color:#4a2f00;border:1px solid #f3d9a7;border-radius:10px;padding:10px 12px;display:flex;gap:10px;align-items:center;flex-wrap:wrap}
  .coverbar span{flex:1;min-width:200px}
  .covering .pg{cursor:crosshair;user-select:none}
  .covering .mk{pointer-events:none}
  .cbox{position:absolute;background:rgba(255,255,255,.92);outline:2px solid var(--red);z-index:4}
  .cbox b{position:absolute;right:-11px;top:-11px;width:22px;height:22px;line-height:22px;text-align:center;border-radius:50%;background:var(--red);color:#fff;font-size:14px;cursor:pointer}
  .mk{position:absolute;transform:translate(-50%,-50%);font-family:Caveat,'Comic Sans MS',cursive;color:var(--red);font-weight:700;line-height:1;white-space:nowrap;pointer-events:auto}
  .mk.check,.mk.x{font-size:44px}
  .mk.partial{font-size:34px}
  .mk.note{font-size:26px;transform:translate(0,-50%);white-space:normal;max-width:34%}
  .mk.score{font-size:46px;border:3px solid var(--red);border-radius:50%;padding:6px 14px}
  .mk.wait{color:var(--amber)}
  .mk.click{cursor:pointer;border-radius:8px;padding:0 4px}
  .mk.click:hover{background:rgba(197,34,31,.08);outline:2px dashed rgba(197,34,31,.5)}
  .mk.mine{text-decoration:underline;text-decoration-thickness:2px}
  .pop{position:absolute;z-index:5;background:#fff;border:1px solid var(--line);box-shadow:0 6px 24px rgba(0,0,0,.2);border-radius:10px;padding:10px 12px;width:300px;font-size:13px}
  .pop .r{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px}
  .pop input{width:80px;font-size:14px;padding:5px;border:1px solid var(--line);border-radius:6px}
  .pop .close{position:absolute;right:6px;top:4px;border:0;font-size:18px;padding:0 6px;background:none}
  .pop .pmsg{color:var(--red);margin-top:6px;min-height:14px}
  .side{border-left:1px solid var(--line);padding:14px 16px}
  .side h3{margin:0 0 4px;font-size:18px}
  .side label{display:block;font-weight:bold;font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin:14px 0 4px}
  .side input,.side textarea,.side select{width:100%;font-size:14px;padding:7px;border:1px solid var(--line);border-radius:7px;font-family:inherit}
  .side input.score{width:90px;font-size:18px;font-weight:bold}
  .muted{color:var(--muted);font-size:13px}
  .box{border-radius:8px;padding:9px 11px;margin-top:8px;line-height:1.4}
  .box.amber{background:var(--amber-soft);color:#4a2f00}
  .box.soft{background:var(--soft)}
  .box button{margin-top:6px}
  .msg{min-height:18px;margin-top:8px;font-weight:bold;color:var(--green)}
  .msg.err{color:var(--red)}
  .center{padding:50px;text-align:center;color:var(--muted)}
  .files{display:flex;gap:6px;flex-wrap:wrap;margin-top:6px}
  .files button.on{background:var(--blue-soft);border-color:#c6d9fb}
  iframe{width:100%;height:80vh;border:0;background:#fff}
</style>
<div class="top">
  <select id="filter" aria-label="Show">
    <option value="all">Everyone</option><option value="flag">Flagged to look over</option><option value="review">Waiting for your review</option><option value="ungraded">Not graded yet</option>
  </select>
  <button id="prev" type="button">&#8592; Previous student</button>
  <span class="who" id="who"></span>
  <button id="next" type="button">Next student &#8594;</button>
  <span class="muted" id="count"></span>
</div>
<div class="grid">
  <div class="list" id="list"></div>
  <div class="pages" id="pages"><div class="center">Loading...</div></div>
  <div class="side" id="side"></div>
</div>
<script>
var $ = function(id){ return document.getElementById(id); };
function el(tag, cls, text){ var e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; }
function call(fn){ var args = Array.prototype.slice.call(arguments, 1); return new Promise(function(ok, no){ var r = google.script.run.withSuccessHandler(ok).withFailureHandler(no); r[fn].apply(r, args); }); }
function errText(e){ return (e && e.message) ? e.message : String(e); }
if (window.pdfjsLib) { try { pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER; } catch (e) {} }
var D = null, people = [], shown = [], at = 0, fileAt = 0, drawing = 0;

/* Version 21: the data can be loaded again (after a file is moved) without reloading the page */
function loadData(showCode){
  return call('viewerData', RUN).then(function(d){
    D = d;
    var by = {};
    d.files.forEach(function(f){ (by[f.code] = by[f.code] || []).push(f); });
    people = Object.keys(by).map(function(c){ return { code: c, files: by[c] }; });
    Object.keys(d.grades).forEach(function(c){ if (!by[c]) people.push({ code: c, files: [] }); });
    people.sort(function(a, b){ return nameOf(a.code).localeCompare(nameOf(b.code)); });
    $('filter').onchange = function(){ at = 0; build(); };
    if (showCode) { $('filter').value = 'all'; shown = people; var k = people.map(function(p){ return p.code; }).indexOf(showCode); at = k > -1 ? k : 0; fileAt = 0; }
    build();
  }).catch(function(e){ $('pages').innerHTML = ''; $('pages').appendChild(el('div', 'center', errText(e))); });
}
loadData();

/* Version 21: pick a student for a file (a wrong match, or one the tool couldn't match) */
function studentPicker(box, file, label){
  var sel = el('select'); sel.appendChild(el('option', '', 'Pick the student...')); sel.options[0].value = '';
  D.students.forEach(function(st){ var o = el('option', '', st.label); o.value = st.id; sel.appendChild(o); });
  var mb = el('button', 'go', label); mb.type = 'button';
  var msg = el('div', 'msg');
  mb.onclick = function(){
    if (!sel.value) { msg.className = 'msg err'; msg.textContent = 'Pick the student first.'; return; }
    mb.disabled = true; msg.className = 'msg'; msg.textContent = 'Moving...';
    call('rpc', 'webMoveFile', [RUN, file.id, sel.value]).then(function(r){
      if (!r || !r.ok) { mb.disabled = false; msg.className = 'msg err'; msg.textContent = (r && r.error) || 'That did not work. Try again.'; return; }
      FILES = {}; FILE_ORDER = [];
      loadData(r.code).then(function(){ alertMsg(r.text); });
    }).catch(function(e){ mb.disabled = false; msg.className = 'msg err'; msg.textContent = errText(e); });
  };
  var up = el('button', '', 'Student not in the list? Update the class list'); up.type = 'button'; up.style.marginTop = '6px';
  up.onclick = function(){
    up.disabled = true; up.textContent = 'Updating from Canvas...';
    call('webUpdateRoster', '').then(function(r){
      up.textContent = (r && r.ok) ? r.text : ((r && r.error) || 'Could not update.');
      return call('viewerData', RUN).then(function(d){ D.students = d.students; sel.innerHTML = ''; sel.appendChild(el('option', '', 'Pick the student...')); sel.options[0].value = '';
        D.students.forEach(function(st){ var o = el('option', '', st.label); o.value = st.id; sel.appendChild(o); }); });
    }).catch(function(e){ up.textContent = errText(e); });
  };
  box.appendChild(sel); box.appendChild(mb); box.appendChild(el('br')); box.appendChild(up); box.appendChild(msg);
}

function nameOf(c){ var n = D && D.names && D.names[c]; return n || (D && D.unmatched[c] !== undefined ? 'Not matched yet' : c); }
function waiting(c){ return (D.review[c] || []).filter(function(r){ return r.eric === null; }).length; }
function flagged(c){ return (D.checks[c] || []).length || D.unmatched[c] !== undefined; }
function build(){
  var f = $('filter').value;
  shown = people.filter(function(p){ return f === 'all' || (f === 'flag' && flagged(p.code)) || (f === 'review' && waiting(p.code)) || (f === 'ungraded' && !D.grades[p.code]); });
  var l = $('list'); l.innerHTML = '';
  if (!shown.length) { l.appendChild(el('div', 'muted', 'Nobody here.')); $('pages').innerHTML = ''; $('pages').appendChild(el('div', 'center', people.length ? 'Nothing matches this filter.' : 'No coded files yet.')); $('side').innerHTML = ''; $('who').textContent = ''; $('count').textContent = ''; return; }
  if (at >= shown.length) at = shown.length - 1;
  renderList();
  show();
}
function renderList(){
  var l = $('list'); l.innerHTML = '';
  shown.forEach(function(p, i){
    var r = el('div', i === at ? 'on' : '');
    var left = el('span', '', nameOf(p.code)); if (flagged(p.code) || waiting(p.code)) left.appendChild(el('span', 'dot', ' •'));
    r.appendChild(left); r.appendChild(el('span', 's', D.grades[p.code] ? D.grades[p.code].score : '–'));
    r.onclick = function(){ at = i; fileAt = 0; build(); };
    l.appendChild(r);
  });
}

/* Version 17: files are kept once loaded, and the next and previous students load in the background */
var FILES = {}, FILE_ORDER = [];
function getFile(id){
  if (!FILES[id]) {
    FILES[id] = call('viewerFile', id).then(function(f){
      if (f.data || !f.size || f.size > 60 * 1024 * 1024) return f;   // version 22: big files come in 4.5 MB pieces
      var CH = 4718592, parts = [], start = 0;
      function next(){
        if (start >= f.size) { f.data = parts.join(''); return f; }
        var s0 = start; start += CH;
        return call('rpc', 'viewerFileChunk', [id, s0, CH]).then(function(b){ parts.push(b); return next(); });
      }
      return next();
    }).catch(function(e){ delete FILES[id]; throw e; });
    FILE_ORDER.push(id);
    while (FILE_ORDER.length > 12) { var old = FILE_ORDER.shift(); if (old !== id) delete FILES[old]; }
  }
  return FILES[id];
}
function prefetch(){
  [at + 1, at - 1, at + 2].forEach(function(i){ var q = shown[i]; if (q && q.files[0]) getFile(q.files[0].id).catch(function(){}); });
}
var PAGES = {}, COVER = null;
function show(){
  var p = shown[at];
  PAGES = {}; COVER = null;
  $('who').textContent = nameOf(p.code) + (nameOf(p.code) !== p.code ? '  (' + p.code + ')' : ''); $('count').textContent = (at + 1) + ' of ' + shown.length;
  $('prev').disabled = at === 0; $('next').disabled = at >= shown.length - 1;
  side(p);
  var pages = $('pages'); pages.innerHTML = '';
  var file = p.files[fileAt];
  if (!file) { pages.appendChild(el('div', 'center', 'No coded file for this student (it may have been moved or finished).')); return; }
  pages.appendChild(el('div', 'center', 'Loading the pages...'));
  var mine = ++drawing;
  setTimeout(prefetch, 400);
  getFile(file.id).then(function(f){
    if (mine !== drawing) return;
    pages.innerHTML = '';
    if (!f.data) { var fr = el('iframe'); fr.src = 'https://drive.google.com/file/d/' + f.id + '/preview'; fr.title = 'Student work'; pages.appendChild(fr); return; }
    if (f.mime.indexOf('image/') === 0) { var pg = el('div', 'pg'); pg.dataset.num = '1'; var im = el('img'); im.src = 'data:' + f.mime + ';base64,' + f.data; pg.appendChild(im); pages.appendChild(pg); PAGES[1] = { img: im, pg: pg }; addMarks(pg, p.code, 1); return; }
    var raw = atob(f.data), bytes = new Uint8Array(raw.length);
    for (var i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i);
    pdfjsLib.getDocument({ data: bytes }).promise.then(function(doc){
      var n = 0;
      function one(){
        if (mine !== drawing || n >= doc.numPages) return;
        n++;
        var num = n;
        return doc.getPage(num).then(function(page){
          var w = Math.max(320, Math.min(860, pages.clientWidth - 28));
          var base = page.getViewport({ scale: 1 });
          var ratio = window.devicePixelRatio || 1;
          var vp = page.getViewport({ scale: (w / base.width) * ratio });
          var pg = el('div', 'pg'); pg.dataset.num = String(num); PAGES[num] = { page: page, pg: pg }; var cv = el('canvas'); cv.width = vp.width; cv.height = vp.height; pg.appendChild(cv);
          pg.appendChild(el('span', 'no', 'page ' + num + ' of ' + doc.numPages));
          pages.appendChild(pg);
          return page.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise.then(function(){ addMarks(pg, p.code, num); return one(); });
        });
      }
      return one();
    }).catch(function(){ pages.innerHTML = ''; var fr = el('iframe'); fr.src = 'https://drive.google.com/file/d/' + f.id + '/preview'; pages.appendChild(fr); });
  }).catch(function(e){ if (mine === drawing) { pages.innerHTML = ''; pages.appendChild(el('div', 'center', 'Could not load: ' + errText(e))); } });
}

/* ---------- version 22: cover a name or face ---------- */
function startCover(p){
  var file = p.files[fileAt];
  if (!file || !Object.keys(PAGES).length) { alertMsg('Wait for the pages to finish loading, then try again.', true); return; }
  if (COVER) return;
  COVER = { code: p.code, file: file, boxes: [] };
  var pages = $('pages'); pages.classList.add('covering');
  var bar = el('div', 'coverbar'); bar.id = 'coverbar';
  bar.appendChild(el('span', '', 'Drag a box over each thing to hide (name, face, signature). Click the x on a box to remove it.'));
  var save = el('button', 'go', 'Paint them out and save'); save.type = 'button'; save.onclick = saveCover;
  var cancel = el('button', '', 'Cancel'); cancel.type = 'button'; cancel.onclick = stopCover;
  bar.appendChild(save); bar.appendChild(cancel);
  pages.insertBefore(bar, pages.firstChild);
  pages.onmousedown = function(e){
    if (!COVER || e.button !== 0) return;
    var pg = e.target.closest ? e.target.closest('.pg') : null;
    if (!pg || e.target.tagName === 'B') return;
    e.preventDefault();
    var r = pg.getBoundingClientRect();
    var x0 = (e.clientX - r.left) / r.width, y0 = (e.clientY - r.top) / r.height;
    var box = { num: Number(pg.dataset.num), x: x0, y: y0, w: 0, h: 0 };
    var d = el('div', 'cbox'); pg.appendChild(d);
    function place(){ d.style.left = (box.x * 100) + '%'; d.style.top = (box.y * 100) + '%'; d.style.width = (box.w * 100) + '%'; d.style.height = (box.h * 100) + '%'; }
    function move(ev){
      var x1 = Math.min(1, Math.max(0, (ev.clientX - r.left) / r.width)), y1 = Math.min(1, Math.max(0, (ev.clientY - r.top) / r.height));
      box.x = Math.min(x0, x1); box.y = Math.min(y0, y1); box.w = Math.abs(x1 - x0); box.h = Math.abs(y1 - y0); place();
    }
    function up(){
      document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up);
      if (box.w < 0.005 || box.h < 0.003) { d.remove(); return; }
      COVER.boxes.push(box);
      var x = el('b', '', '\u00d7'); x.title = 'Remove this box';
      x.onmousedown = function(ev){ ev.stopPropagation(); };
      x.onclick = function(ev){ ev.stopPropagation(); COVER.boxes = COVER.boxes.filter(function(b){ return b !== box; }); d.remove(); };
      d.appendChild(x);
    }
    document.addEventListener('mousemove', move); document.addEventListener('mouseup', up);
  };
}
function stopCover(){
  COVER = null;
  var pages = $('pages'); pages.classList.remove('covering'); pages.onmousedown = null;
  var bar = $('coverbar'); if (bar) bar.remove();
  Array.prototype.slice.call(pages.querySelectorAll('.cbox')).forEach(function(b){ b.remove(); });
}
/* Draws each boxed page big, paints the boxes white, and sends the pictures to be saved */
function saveCover(){
  if (!COVER || !COVER.boxes.length) { alertMsg('Drag a box over the name first.', true); return; }
  var c = COVER, nums = {};
  c.boxes.forEach(function(b){ (nums[b.num] = nums[b.num] || []).push(b); });
  var bar = $('coverbar'); bar.querySelectorAll('button').forEach(function(b){ b.disabled = true; }); bar.firstChild.textContent = 'Painting them out and saving...';
  var out = [];
  var chain = Promise.resolve();
  Object.keys(nums).forEach(function(k){
    chain = chain.then(function(){
      var P = PAGES[k], cv = document.createElement('canvas');
      var paint = function(){
        var x = cv.getContext('2d'); x.fillStyle = '#fff';
        nums[k].forEach(function(b){ x.fillRect(Math.floor(b.x * cv.width) - 2, Math.floor(b.y * cv.height) - 2, Math.ceil(b.w * cv.width) + 4, Math.ceil(b.h * cv.height) + 4); });
        out.push({ page: Number(k), jpg: cv.toDataURL('image/jpeg', 0.88).split(',')[1] });
      };
      if (P.img) { cv.width = P.img.naturalWidth; cv.height = P.img.naturalHeight; var x0 = cv.getContext('2d'); x0.fillStyle = '#fff'; x0.fillRect(0, 0, cv.width, cv.height); x0.drawImage(P.img, 0, 0); paint(); return; }
      var base = P.page.getViewport({ scale: 1 });
      var vp = P.page.getViewport({ scale: Math.min(3, 1700 / base.width) });
      cv.width = Math.round(vp.width); cv.height = Math.round(vp.height);
      var ctx = cv.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cv.width, cv.height);
      return P.page.render({ canvasContext: ctx, viewport: vp }).promise.then(paint);
    });
  });
  chain.then(function(){ return call('rpc', 'webCoverFile', [RUN, c.file.id, out]); }).then(function(r){
    if (!r || !r.ok) throw new Error((r && r.error) || 'That did not save.');
    delete FILES[c.file.id];
    return loadData(c.code).then(function(){ alertMsg(r.text); });
  }).catch(function(e){
    if (!$('coverbar')) return;
    bar.querySelectorAll('button').forEach(function(b){ b.disabled = false; });
    bar.firstChild.textContent = 'Not saved: ' + errText(e) + ' Try again.';
  });
}

/* Rowan's red-pen marks. A mark tied to a question you decided in Review shows your points. */
function addMarks(pg, code, num){
  (D.marks[code] || []).filter(function(m){ return m.page === num && m.x !== null && m.y !== null; }).forEach(function(m){
    var kind = ['check', 'x', 'partial', 'note', 'score'].indexOf(m.mark) > -1 ? m.mark : 'note';
    var text = kind === 'check' ? '✓' : kind === 'x' ? '✗' : m.text;
    var cls = 'mk ' + kind;
    if (kind === 'score' && D.grades[code]) { var tot = String(m.text).match(/[/] *([0-9.]+)/); text = D.grades[code].score + (tot ? '/' + tot[1] : ''); }
    if (m.question) {
      var q = String(m.question).trim().split(/[ .:)]/)[0].toLowerCase();
      var r = (D.review[code] || []).filter(function(x){ return String(x.question).trim().split(/[ .:)]/)[0].toLowerCase() === q; })[0];
      if (r) {
        if (r.eric === null) { cls += ' wait'; text = '? / ' + (r.possible === null ? '' : r.possible); }
        else { text = r.eric + ' / ' + r.possible; cls = 'mk partial'; }
      }
    }
    var d = el('div', cls, text);
    d.style.left = (m.x * 100) + '%'; d.style.top = (m.y * 100) + '%';
    if (m.question && kind !== 'note' && kind !== 'score') { d.classList.add('click'); d.title = 'Click to change ' + m.question; d.onclick = function(e){ e.stopPropagation(); override(pg, code, num, m, d); }; }
    if (m.changed && kind !== 'note' && kind !== 'score') d.classList.add('mine');
    if (kind !== 'note' && kind !== 'check' && kind !== 'x' && m.text && text !== m.text) d.title = m.text;
    pg.appendChild(d);
  });
}

/* Change a question's mark: full credit, no credit, or points */
function override(pg, code, num, m, d){
  var old = document.getElementById('pop'); if (old) old.remove();
  var possible = m.possible;
  if (possible === null) { var t = String(m.text).match(/[/] *([0-9.]+)/); if (t) possible = Number(t[1]); }
  var pop = el('div', 'pop'); pop.id = 'pop';
  pop.appendChild(el('b', '', m.question + (possible !== null ? ' (worth ' + possible + ')' : '')));
  var row = el('div', 'r');
  var mk = function(label, v, cls){ var b = el('button', cls || '', label); b.type = 'button'; b.onclick = function(){ save(v); }; row.appendChild(b); };
  if (possible !== null) mk('Right: full credit (' + possible + ')', possible, 'go');
  mk('Wrong: no credit (0)', 0);
  pop.appendChild(row);
  var r2 = el('div', 'r'); var inp = el('input'); inp.type = 'number'; inp.step = '0.5'; inp.min = '0'; if (possible !== null) inp.max = String(possible); inp.placeholder = 'points';
  var sb = el('button', '', 'Give these points'); sb.type = 'button'; sb.onclick = function(){ if (inp.value !== '') save(Number(inp.value)); };
  inp.onkeydown = function(e){ if (e.key === 'Enter') sb.click(); };
  r2.appendChild(inp); r2.appendChild(sb); pop.appendChild(r2);
  var note = el('div', 'pmsg'); pop.appendChild(note);
  var x = el('button', 'close', '\u00d7'); x.type = 'button'; x.onclick = function(){ pop.remove(); }; pop.appendChild(x);
  if (m.x > 0.5) pop.style.right = Math.max(2, 100 - m.x * 100 - 4) + '%'; else pop.style.left = Math.max(2, m.x * 100 - 4) + '%';
  pop.style.top = (m.y * 100 + 3) + '%';
  pg.appendChild(pop);
  function save(v){
    note.textContent = 'Saving...';
    call('webOverrideMark', RUN, code, m.row, m.question, v).then(function(r){
      if (!r.ok) { note.textContent = r.error; return; }
      m.mark = r.mark; m.text = r.text; m.points = r.points; m.possible = r.possible; m.changed = r.changed;
      (D.review[code] || []).forEach(function(q){ if (String(q.question).trim().split(/[ .:)]/)[0].toLowerCase() === String(m.question).trim().split(/[ .:)]/)[0].toLowerCase()) q.eric = r.points; });
      if (r.score !== '' && D.grades[code]) { D.grades[code].score = r.score; D.grades[code].sent = 'Changed since it was sent (Send again)'; }
      pop.remove();
      Array.prototype.slice.call(pg.querySelectorAll('.mk')).forEach(function(x){ x.remove(); });
      addMarks(pg, code, num);
      renderList(); side(shown[at]);
      alertMsg(m.question + ' changed. ' + code + ' is now ' + (D.grades[code] ? D.grades[code].score : '') + '.');
    }).catch(function(e){ note.textContent = errText(e); });
  }
}

function side(p){
  var s = $('side'); s.innerHTML = '';
  var g = D.grades[p.code];
  s.appendChild(el('h3', '', nameOf(p.code)));
  if (nameOf(p.code) !== p.code) s.appendChild(el('div', 'muted', p.code));
  if (p.files.length > 1) {
    s.appendChild(el('div', 'muted', p.files.length + ' files from this student:'));
    var fl = el('div', 'files');
    p.files.forEach(function(f, i){ var b = el('button', i === fileAt ? 'on' : '', f.name + (f.added ? ' (late)' : '')); b.type = 'button'; b.onclick = function(){ fileAt = i; show(); }; fl.appendChild(b); });
    s.appendChild(fl);
  }
  (D.checks[p.code] || []).forEach(function(c){
    var b = el('div', 'box amber', c.reason);
    if (c.kind === 'look') { var ok = el('button', '', 'Looks fine'); ok.type = 'button'; ok.onclick = function(){ call('webDismissCheck', RUN, c.code, c.file).then(function(){ D.checks[p.code] = D.checks[p.code].filter(function(x){ return x !== c; }); build(); }); }; b.appendChild(el('br')); b.appendChild(ok); }
    s.appendChild(b);
  });
  if (D.unmatched[p.code] !== undefined && p.files[fileAt]) {
    var u = el('div', 'box amber', 'The tool could not tell whose this is (file: ' + D.unmatched[p.code] + '). Who is it? ');
    studentPicker(u, p.files[fileAt], 'Match');
    s.appendChild(u);
  } else if (p.files[fileAt]) {
    var w = el('details'); w.style.marginTop = '10px';
    w.appendChild(el('summary', 'muted', 'Wrong student? Move this file'));
    var wb = el('div', 'box soft'); studentPicker(wb, p.files[fileAt], 'Move it'); w.appendChild(wb);
    s.appendChild(w);
  }
  if (p.files[fileAt]) {   // version 22
    var cvb = el('button', '', 'Cover a name or face'); cvb.type = 'button'; cvb.style.marginTop = '10px';
    cvb.title = 'Drag a box over anything the tool missed; it gets painted out for good';
    cvb.onclick = function(){ startCover(p); };
    s.appendChild(cvb);
  }
  s.appendChild(el('label', '', 'Score'));
  var sc = el('input', 'score'); sc.value = g ? g.score : ''; sc.placeholder = '–'; s.appendChild(sc);
  s.appendChild(el('label', '', 'Feedback (goes to the student)'));
  var fb = el('textarea'); fb.rows = 5; fb.value = g ? g.fb : ''; s.appendChild(fb);
  var save = el('button', 'go', 'Save changes'); save.type = 'button'; save.style.marginTop = '8px'; save.disabled = true;
  sc.oninput = fb.oninput = function(){ save.disabled = (sc.value === (g ? g.score : '') && fb.value === (g ? g.fb : '')); };
  var msg = el('div', 'msg'); msg.id = 'msg';
  save.onclick = function(){
    save.disabled = true; msg.className = 'msg'; msg.textContent = 'Saving...';
    call('webSaveGrade', RUN, p.code, sc.value, fb.value).then(function(r){
      if (!r.ok) { msg.className = 'msg err'; msg.textContent = r.error; return; }
      D.grades[p.code] = { score: sc.value, fb: fb.value, rubric: g ? g.rubric : '', sent: 'Changed since it was sent (Send again)' };
      g = D.grades[p.code];
      msg.textContent = r.text; renderList();
      if (sentLine) sentLine.textContent = 'Canvas: ' + g.sent;
    }).catch(function(e){ msg.className = 'msg err'; msg.textContent = errText(e); });
  };
  s.appendChild(save); s.appendChild(msg);
  var sentLine = el('div', 'muted', g ? 'Canvas: ' + (g.sent || 'Not sent yet') : ''); s.appendChild(sentLine);
  if (!g) s.appendChild(el('div', 'muted', 'Rowan hasn’t graded this one yet.'));
  if (g && g.rubric) { s.appendChild(el('label', '', 'Breakdown')); s.appendChild(el('div', 'box soft', g.rubric)); }
  var rv = D.review[p.code] || [];
  if (rv.length) {
    s.appendChild(el('label', '', 'Rowan wasn’t sure about'));
    rv.forEach(function(r){ s.appendChild(el('div', 'box ' + (r.eric === null ? 'amber' : 'soft'), r.question + ': ' + (r.eric === null ? 'waiting for you (Rowan suggests ' + r.rowan + ' of ' + r.possible + ')' : 'you gave ' + r.eric + ' of ' + r.possible))); });
    if (rv.some(function(r){ return r.eric === null; })) { var go = el('button', '', 'Decide these in Review'); go.type = 'button'; go.style.marginTop = '8px'; go.onclick = function(){ if (parent && parent.openView) parent.openView('review', 'Review what Rowan was not sure about'); }; s.appendChild(go); }
  }
}
function alertMsg(t, bad){ var m = $('msg'); if (m) { m.className = 'msg' + (bad ? ' err' : ''); m.textContent = t; } }

$('prev').onclick = function(){ if (at > 0) { at--; fileAt = 0; build(); } };
$('next').onclick = function(){ if (at < shown.length - 1) { at++; fileAt = 0; build(); } };
document.addEventListener('keydown', function(e){
  if (e.target && /INPUT|TEXTAREA|SELECT/.test(e.target.tagName)) return;
  if (e.key === 'ArrowRight') $('next').click();
  if (e.key === 'ArrowLeft') $('prev').click();
});
</script>
</body>`;

function webHomeHtml_() { return WEB_HOME_HTML; }

const WEB_HOME_HTML = `<!doctype html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Student Work Anonymizer</title>
<style>
  :root{--fg:#1f2933;--muted:#5f6b7a;--line:#e1e6ec;--soft:#f5f7fa;--bg:#f0f3f7;--card:#fff;--blue:#1a73e8;--blue-soft:#e8f0fe;--green:#137333;--green-soft:#e6f4ea;--amber:#8a5300;--amber-soft:#fef3e0;--red:#b3261e;--red-soft:#fce8e6}
  *{box-sizing:border-box}
  body{margin:0;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:var(--fg);background:var(--bg)}
  header{background:#fff;border-bottom:1px solid var(--line);padding:12px 20px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;position:sticky;top:0;z-index:5}
  header h1{font-size:18px;margin:0}
  header .ver{color:var(--muted);font-size:12px;margin-right:auto}
  .chip{display:inline-flex;align-items:center;gap:6px;padding:5px 10px;border-radius:14px;font-size:13px;background:var(--soft);color:var(--muted);border:1px solid var(--line);cursor:pointer}
  .chip.on{background:var(--green-soft);color:var(--green);border-color:#b7dfc2}
  .chip.off{background:var(--amber-soft);color:var(--amber);border-color:#f3d9a7}
  main{max-width:1060px;margin:0 auto;padding:20px 16px 60px}
  .start{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:12px;margin-bottom:22px}
  .tile{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;cursor:pointer;text-align:left;font:inherit;color:inherit}
  .tile:hover{border-color:var(--blue);box-shadow:0 1px 4px rgba(26,115,232,.15)}
  .tile b{display:block;font-size:15px;margin-bottom:4px}
  .tile span{color:var(--muted);line-height:1.4}
  .banner{background:var(--amber-soft);border:1px solid #f3d9a7;color:#4a2f00;border-radius:10px;padding:10px 14px;margin-bottom:16px;display:flex;gap:10px;align-items:center;flex-wrap:wrap}
  h2{font-size:16px;margin:0 0 10px;display:flex;align-items:center;gap:10px}
  h2 button{font-size:12px;padding:4px 10px}
  .card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px;margin-bottom:12px}
  .card .top{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
  .card .name{font-size:16px;font-weight:bold;margin-right:auto}
  .badge{font-size:11px;padding:2px 8px;border-radius:10px;background:var(--soft);color:var(--muted);text-transform:uppercase;letter-spacing:.04em}
  .badge.paper{background:#efedff;color:#4b3fb5}
  .stepper{display:flex;gap:4px;margin:12px 0 10px;flex-wrap:wrap}
  .stepper div{flex:1;min-width:92px;padding:7px 8px;border-radius:8px;background:var(--soft);color:var(--muted);font-size:12px;border:1px solid transparent}
  .stepper div b{display:block;font-size:12px;color:inherit;margin-bottom:2px}
  .stepper .done{background:var(--green-soft);color:var(--green)}
  .stepper .now{background:var(--blue-soft);color:var(--blue);border-color:#c6d9fb}
  .stepper .click{cursor:pointer}
  .stepper .click:hover,.stepper .click:focus-visible{border-color:var(--blue);outline:none}
  .facts{color:var(--muted);font-size:13px;margin-bottom:10px}
  .actions{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
  button{font-size:14px;padding:8px 14px;border-radius:8px;border:1px solid var(--line);background:#fff;cursor:pointer;color:var(--fg)}
  button.go{background:var(--blue);color:#fff;border-color:var(--blue);font-weight:bold}
  button.small{font-size:13px;padding:6px 10px}
  button.danger{color:var(--red)}
  button:disabled{opacity:.5;cursor:default}
  .wait{background:var(--amber-soft);color:#4a2f00;border-radius:8px;padding:8px 12px;font-size:13px}
  .empty{color:var(--muted);padding:30px;text-align:center;background:#fff;border-radius:12px;border:1px dashed var(--line)}
  .modal{position:fixed;inset:0;background:rgba(15,23,42,.45);display:flex;align-items:center;justify-content:center;z-index:20;padding:16px}
  .modal[hidden]{display:none}
  .dialog{background:#fff;border-radius:14px;width:560px;max-width:100%;max-height:90vh;overflow:auto;padding:20px 22px;box-shadow:0 10px 40px rgba(0,0,0,.2)}
  .dialog h3{margin:0 0 10px;font-size:17px}
  .dialog p{margin:0 0 10px;line-height:1.5;white-space:pre-wrap}
  .dialog label{display:block;font-weight:bold;margin:12px 0 4px;font-size:13px}
  .dialog select,.dialog input,.dialog textarea{width:100%;font-size:14px;padding:8px;border:1px solid var(--line);border-radius:8px;font-family:inherit}
  .dialog .row{display:flex;gap:8px;justify-content:flex-end;margin-top:16px;flex-wrap:wrap}
  .dialog .err{color:var(--red);font-weight:bold;min-height:18px;margin-top:8px}
  .dialog a{color:var(--blue);font-weight:bold}
  .dialog hr{border:0;border-top:1px solid var(--line);margin:16px 0}
  .spin{display:inline-block;width:16px;height:16px;border:3px solid var(--blue-soft);border-top-color:var(--blue);border-radius:50%;animation:s 1s linear infinite;vertical-align:-3px;margin-right:8px}
  @keyframes s{to{transform:rotate(360deg)}}
  .view{position:fixed;inset:0;background:#fff;z-index:15;display:flex;flex-direction:column}
  .view[hidden]{display:none}
  .view .bar{display:flex;align-items:center;gap:12px;padding:8px 14px;border-bottom:1px solid var(--line)}
  .view iframe{flex:1;border:0;width:100%}
</style></head>
<body>
<header>
  <h1>Student Work Anonymizer</h1>
  <span class="ver" id="ver"></span>
  <span class="chip" id="canvasChip">Canvas</span>
  <span class="chip" id="autoChip">Auto-run</span>
  <button class="small" id="settingsBtn" type="button">Settings</button>
</header>
<main>
  <div class="start">
    <button class="tile" id="tPull" type="button"><b>Get projects from Canvas</b><span>Pull new and late submissions for an assignment.</span></button>
    <button class="tile" id="tPaper" type="button"><b>Paper test or quiz</b><span>Upload the class scan and the answer key, then match names.</span></button>
    <button class="tile" id="tClean" type="button"><b>Clean the Inbox now</b><span>Remove names and make the coded files (auto-run does this every 10 minutes).</span></button>
  </div>
  <div id="update"></div>
  <div id="loose"></div>
  <h2>Your assignments <button class="small" id="refresh" type="button">Refresh</button></h2>
  <div id="list"><div class="empty"><span class="spin"></span>Loading...</div></div>
</main>

<div class="modal" id="modal" hidden><div class="dialog" id="dialog"></div></div>
<div class="view" id="view" hidden><div class="bar"><button id="viewBack" type="button">&#8592; Back to home</button><b id="viewTitle"></b></div><iframe id="frame" title="Window"></iframe></div>

<script>
var $ = function(id){ return document.getElementById(id); };
function el(tag, cls, text){ var e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; }
function call(fn){ var args = Array.prototype.slice.call(arguments, 1); return new Promise(function(ok, no){ var r = google.script.run.withSuccessHandler(ok).withFailureHandler(no); r[fn].apply(r, args); }); }
function errText(e){ return (e && e.message) ? e.message : String(e); }
var ST = null;
var CACHED = /*CACHED_STATUS*/null;   // version 17: what the page showed last time, so it opens right away

/* ---------- dialogs ---------- */
function openDialog(title){ var d = $('dialog'); d.innerHTML = ''; d.appendChild(el('h3', '', title)); $('modal').hidden = false; return d; }
function closeDialog(){ $('modal').hidden = true; }
function buttons(d, list){ var r = el('div', 'row'); list.forEach(function(b){ var x = el('button', b.cls || '', b.label); x.type = 'button'; x.onclick = b.fn; r.appendChild(x); }); d.appendChild(r); return r; }
function busy(title, text){ var d = openDialog(title); var p = el('p'); var s = el('span', 'spin'); p.appendChild(s); p.appendChild(document.createTextNode(text || 'Working...')); d.appendChild(p); return p; }
function showResult(title, r, links){
  var d = openDialog(title);
  if (r && r.ok === false) { var e = el('p'); e.style.color = 'var(--red)'; e.textContent = r.error; d.appendChild(e); }
  else { d.appendChild(el('p', '', (r && (r.text || r.message)) || 'Done.')); }
  (links || (r && r.links) || []).forEach(function(l){ var p = el('p'); var a = el('a', '', l.label); a.href = l.url; a.target = '_blank'; p.appendChild(a); d.appendChild(p); });
  buttons(d, [{ label: 'OK', cls: 'go', fn: function(){ closeDialog(); refresh(); } }]);
}
function fail(title){ return function(e){ showResult(title, { ok: false, error: errText(e) }); }; }

/* A dropdown of assignments: the Canvas list, or names already used */
function pickAssignment(title, intro, okLabel, useCanvas, done){
  var d = busy(title, 'Getting your assignments...');
  var load = useCanvas ? call('webCanvasOptions') : Promise.resolve({ connected: false, options: [] });
  load.then(function(c){
    return (c.connected ? Promise.resolve([]) : call('webNameOptions')).then(function(names){
      var d2 = openDialog(title);
      d2.appendChild(el('p', '', c.connected ? intro : intro + (useCanvas ? ' (Canvas is not connected, so pick a name or type a new one.)' : '')));
      var sel = el('select'); sel.appendChild(el('option', '', 'Choose an assignment...')); sel.options[0].value = '';
      if (c.connected) {
        var groups = {};
        c.options.forEach(function(o){ if (!groups[o.group]) { groups[o.group] = el('optgroup'); groups[o.group].label = o.group; sel.appendChild(groups[o.group]); } var op = el('option', '', o.label); op.value = o.value; groups[o.group].appendChild(op); });
      } else {
        names.forEach(function(n){ var op = el('option', '', n); op.value = 'name:' + n; sel.appendChild(op); });
      }
      d2.appendChild(sel);
      d2.appendChild(el('label', '', c.connected ? 'Not in the list? Paste the assignment link from Canvas:' : 'Or type a new name:'));
      var other = el('input'); other.type = 'text'; d2.appendChild(other);
      var err = el('div', 'err'); d2.appendChild(err);
      buttons(d2, [{ label: 'Cancel', fn: closeDialog }, { label: okLabel, cls: 'go', fn: function(){
        var v = other.value.trim() ? 'typed:' + other.value.trim() : sel.value;
        if (!v) { err.textContent = 'Choose an assignment first.'; return; }
        done(v);
      } }]);
    });
  }).catch(fail(title));
}

/* Runs a job that may ask "Are you sure?" once or twice */
function confirmable(title, fn, arg, yes){
  busy(title, 'Working...');
  call(fn, arg, yes || 0).then(function(r){
    if (r && r.ok && r.confirm) {
      var d = openDialog(r.confirm.title || title);
      d.appendChild(el('p', '', r.confirm.text));
      // "No": whatever already happened before this question is shown (for example "finished" before "also delete older files?")
      buttons(d, [{ label: 'No', fn: function(){ if (r.text) showResult(title, r); else { closeDialog(); refresh(); } } }, { label: 'Yes', cls: 'go', fn: function(){ confirmable(title, fn, arg, (yes || 0) + 1); } }]);
      return;
    }
    showResult(title, r);
  }).catch(fail(title));
}

/* ---------- the buttons ---------- */
function pull(value){
  var go = function(v){ busy('Get projects from Canvas', 'Pulling from Canvas. This can take a few minutes...'); call('webPull', v).then(function(r){ showResult('Get projects from Canvas', r); }).catch(fail('Get projects from Canvas')); };
  if (value) return go(value);
  if (ST && !ST.canvas) return settings('Connect to Canvas first, then pull.');
  pickAssignment('Get projects from Canvas', 'Which assignment do you want to pull?', 'Pull submissions', true, go);
}
function paper(){
  pickAssignment('Paper test or quiz', 'Which test or quiz? Pick it from Canvas so the grades can go back to the right assignment.', 'Next', true, function(v){
    busy('Paper test or quiz');
    call('webScanFolder', v).then(function(r){ if (r && r.ok) uploadScan(r.run); else showResult('Paper test or quiz', r); }).catch(fail('Paper test or quiz'));
  });
}
function clean(looseRun, rounds){
  rounds = rounds || 0;
  var token = String(Date.now()) + Math.random().toString(36).slice(2);
  busy('Clean the Inbox', rounds ? 'Still working (Google makes the tool take breaks on big batches)...' : 'Removing names and making the coded files. This can take a few minutes...');
  var handle = function(r){
    if (!r || r.ok === undefined) return poll(token, 0).then(handle);
    if (r.needLoose) {
      var d = openDialog('Which assignment are these files for?');
      d.appendChild(el('p', '', 'There are files sitting loose in the Inbox. Pick the assignment they belong to (late work goes with the earlier one), or type a new name.'));
      var sel = el('select'); sel.appendChild(el('option', '', 'Choose...')); sel.options[0].value = '';
      (r.assignments || []).forEach(function(n){ var o = el('option', '', n); o.value = n; sel.appendChild(o); });
      d.appendChild(sel); d.appendChild(el('label', '', 'Or a new name:')); var t = el('input'); d.appendChild(t);
      var err = el('div', 'err'); d.appendChild(err);
      buttons(d, [{ label: 'Cancel', fn: closeDialog }, { label: 'Clean them', cls: 'go', fn: function(){ var v = t.value.trim() || sel.value; if (!v) { err.textContent = 'Choose an assignment first.'; return; } clean(v); } }]);
      return;
    }
    if (r.more && rounds < 15) return clean(looseRun, rounds + 1);
    showResult('Clean the Inbox', r);
  };
  call('webProcess', token, looseRun || '').then(handle).catch(function(e){ poll(token, 0).then(function(r){ if (r) handle(r); else fail('Clean the Inbox')(e); }); });
}
function poll(token, n){
  return new Promise(function(res){ setTimeout(res, 3000); }).then(function(){ return call('webResult', token); }).then(function(r){ if (r || n >= 120) return r || { ok: false, error: 'Google stopped this run before it finished (each run can only take 6 minutes). Click Clean the Inbox again to keep going. A file that stops Google 3 times is set aside for you, so the rest still get cleaned.' }; return poll(token, n + 1); });
}
function send(a){
  var go = function(v){ confirmable('Send grades to Canvas', 'webSend', v, 0); };
  if (ST && !ST.canvas) return settings('Connect to Canvas first, then send.');
  if (a && a.canvas) return go(a.canvas);
  pickAssignment('Send grades to Canvas', 'Which Canvas assignment are the grades for' + (a ? ' (' + a.run + ')' : '') + '?', 'Next', true, go);
}
function finish(a){ confirmable('Finish ' + a.run, 'webFinish', a.run, 0); }
function bringIn(){ busy('Bring in Rowan’s grades'); call('webBringIn').then(function(r){ showResult('Bring in Rowan’s grades', r); }).catch(fail('Bring in')); }

/* ---------- uploads (version 14) ---------- */
function toB64(u8){ var s = '', step = 32768; for (var i = 0; i < u8.length; i += step) s += String.fromCharCode.apply(null, u8.subarray(i, i + step)); return btoa(s); }
function uploadFile(kind, run, file, progress){
  return call('webUploadStart', kind, run, file.name, file.type || '', file.size).then(function(st){
    var at = 0;
    function next(){
      if (at >= file.size && file.size > 0) return Promise.resolve();
      return file.slice(at, Math.min(file.size, at + st.chunk)).arrayBuffer().then(function(buf){
        var u8 = new Uint8Array(buf), start = at;
        return call('webUploadChunk', st.token, start, toB64(u8)).then(function(r){
          at = start + u8.length; progress(at);
          if (r && r.done) return;
          return next();
        });
      });
    }
    return next();
  });
}
function bar(d){ var b = el('div'); b.style.cssText = 'height:10px;background:var(--soft);border-radius:5px;overflow:hidden;margin:10px 0'; var i = el('i'); i.style.cssText = 'display:block;height:100%;width:0;background:var(--green);transition:width .2s'; b.appendChild(i); d.appendChild(b); return i; }

/* Upload the class scan (and, if you like, the answer key) for a paper test */
/* Version 23: notes for Rowan and a blank copy of the test, used by both windows */
var NL = String.fromCharCode(10);
function notesFields(d, run){
  d.appendChild(el('label', '', 'Blank test (optional): helps Rowan see how many parts each question has'));
  var bf = el('input'); bf.type = 'file'; bf.accept = '.pdf,.doc,.docx,image/*'; d.appendChild(bf);
  var saved = el('p', ''); saved.style.cssText = 'margin:4px 0 0;color:var(--muted);font-size:13px'; d.appendChild(saved);
  d.appendChild(el('label', '', 'Notes for Rowan: how you want it graded (these win over the key)'));
  var nt = el('textarea'); nt.rows = 5;
  nt.placeholder = ['Grade leniently.', 'Q4 has two parts, 1 point each.', 'Q6: if part a is wrong but b uses it right, b gets full credit.', 'Q9: accept either answer.'].join(NL);
  d.appendChild(nt);
  var loaded = '';
  call('rpc', 'webGetNotes', [run]).then(function(r){
    if (!r) return;
    loaded = r.text || ''; if (!nt.value) nt.value = loaded;
    if (r.blank) { saved.textContent = 'Saved: '; var l = el('a', '', r.blank.name); l.href = r.blank.url; l.target = '_blank'; saved.appendChild(l); saved.appendChild(document.createTextNode(' (a new one replaces it)')); }
  }).catch(function(){});
  return {
    blankFile: function(){ return bf.files[0] || null; },
    /* uploads the blank (if picked) and saves the notes (if changed); returns what was done */
    save: function(onStep){
      var done = [], chain = Promise.resolve();
      if (bf.files[0]) chain = chain.then(function(){ onStep && onStep('Uploading the blank test...'); return uploadFile('blank', run, bf.files[0], function(){}).then(function(){ done.push('blank test saved'); }); });
      if (nt.value.trim() !== loaded.trim()) chain = chain.then(function(){ onStep && onStep('Saving your notes...'); return call('rpc', 'webSaveNotes', [run, nt.value]).then(function(r){ if (!r || !r.ok) throw new Error((r && r.error) || 'Notes not saved.'); done.push(nt.value.trim() ? 'notes saved' : 'notes removed'); }); });
      return chain.then(function(){ return done; });
    }
  };
}

function uploadScan(run){
  var d = openDialog('Upload the class scan: ' + run);
  d.appendChild(el('p', '', 'Pick the scan of the whole class. One big PDF is fine; several PDFs or phone pictures work too.'));
  var f = el('input'); f.type = 'file'; f.multiple = true; f.accept = '.pdf,image/*'; d.appendChild(f);
  d.appendChild(el('label', '', 'Answer key (optional, you can add it later)'));
  var kf = el('input'); kf.type = 'file'; kf.accept = '.pdf,.doc,.docx,.txt,image/*'; d.appendChild(kf);
  d.appendChild(el('p', '', 'or type it:')).style.cssText = 'margin:8px 0 4px;color:var(--muted)';
  var kt = el('textarea'); kt.rows = 4; kt.placeholder = '1 B\\n2 C\\n3 The Moon’s gravity pulls the ocean (2 points)'; d.appendChild(kt);
  var nf = notesFields(d, run);
  var err = el('div', 'err'); d.appendChild(err);
  var fill = bar(d); fill.parentNode.style.display = 'none';
  var row = buttons(d, [{ label: 'Cancel', fn: closeDialog }, { label: 'Upload and read the pages', cls: 'go', fn: function(){
    var files = Array.prototype.slice.call(f.files || []);
    if (!files.length) { err.textContent = 'Pick the scan file first.'; return; }
    row.querySelectorAll('button').forEach(function(b){ b.disabled = true; });
    fill.parentNode.style.display = 'block'; err.style.color = 'var(--muted)';
    var total = files.reduce(function(n, x){ return n + x.size; }, 0) + (kf.files[0] ? kf.files[0].size : 0), doneBytes = 0;
    var chain = Promise.resolve();
    files.forEach(function(x, i){ chain = chain.then(function(){ err.textContent = 'Uploading ' + x.name + ' (' + (i + 1) + ' of ' + files.length + ')...'; var base = doneBytes; return uploadFile('scan', run, x, function(n){ fill.style.width = (100 * (base + n) / total) + '%'; }).then(function(){ doneBytes = base + x.size; }); }); });
    if (kf.files[0]) chain = chain.then(function(){ err.textContent = 'Uploading the answer key...'; var base = doneBytes; return uploadFile('key', run, kf.files[0], function(n){ fill.style.width = (100 * (base + n) / total) + '%'; }); });
    else if (kt.value.trim()) chain = chain.then(function(){ return call('webSaveKeyText', run, kt.value); });
    chain = chain.then(function(){ return nf.save(function(t){ err.textContent = t; }); });
    chain.then(function(){ clean(); }).catch(function(e){ err.style.color = ''; err.textContent = 'Upload stopped: ' + errText(e); row.querySelectorAll('button').forEach(function(b){ b.disabled = false; }); });
  } }]);
}

/* See, add, replace or remove a test's answer key */
function keyDialog(a){
  var d = openDialog('Answer key and notes: ' + a.run);
  if (a.key) {
    var p = el('p'); p.appendChild(document.createTextNode('Saved: ')); var l = el('a', '', a.key.name); l.href = a.key.url; l.target = '_blank'; p.appendChild(l); d.appendChild(p);
    d.appendChild(el('p', '', 'Adding a new one below replaces it.')).style.color = 'var(--muted)';
  } else {
    d.appendChild(el('p', '', 'No key yet. Without one, Rowan writes a key from the test and shows it to you to check. Put the points next to each answer if they are not all 1 point.'));
  }
  d.appendChild(el('label', '', 'Upload a file (PDF, Word, Google Doc saved as PDF, or a photo)'));
  var kf = el('input'); kf.type = 'file'; kf.accept = '.pdf,.doc,.docx,.txt,image/*'; d.appendChild(kf);
  d.appendChild(el('label', '', 'or type the answers'));
  var kt = el('textarea'); kt.rows = 6; kt.placeholder = '1 B\\n2 C\\n3 The Moon’s gravity pulls the ocean (2 points)'; d.appendChild(kt);
  d.appendChild(el('hr'));
  var nf = notesFields(d, a.run);
  var err = el('div', 'err'); d.appendChild(err);
  var list = [{ label: 'Cancel', fn: closeDialog }];
  if (a.key) list.unshift({ label: 'Remove key', cls: 'danger', fn: function(){ call('webRemoveKey', a.run).then(function(r){ showResult('Answer key', r); }).catch(fail('Answer key')); } });
  list.push({ label: 'Save', cls: 'go', fn: function(){
    err.style.color = 'var(--muted)';
    var done = [], chain = Promise.resolve();
    if (kf.files[0]) chain = chain.then(function(){ err.textContent = 'Uploading the key...'; return uploadFile('key', a.run, kf.files[0], function(){}).then(function(){ done.push('answer key saved'); }); });
    else if (kt.value.trim()) chain = chain.then(function(){ return call('webSaveKeyText', a.run, kt.value).then(function(r){ if (!r || !r.ok) throw new Error((r && r.error) || 'Key not saved.'); done.push('answer key saved'); }); });
    chain.then(function(){ return nf.save(function(t){ err.textContent = t; }); }).then(function(more){
      done = done.concat(more || []);
      if (!done.length) { err.style.color = ''; err.textContent = 'Nothing changed. Pick a file, type the answers, or write notes first.'; return; }
      var t = done.join(', '); showResult('Answer key and notes', { ok: true, text: t.charAt(0).toUpperCase() + t.slice(1) + ' for "' + a.run + '". Rowan reads all of it when you say "grade."' });
    }).catch(function(e){ err.style.color = ''; err.textContent = errText(e); });
  } });
  buttons(d, list);
}

/* Tests held back because a name may still be on them */
function heldBox(a){
  var box = el('div'); box.style.cssText = 'margin-top:10px;background:var(--amber-soft);border-radius:8px;padding:10px 12px';
  box.appendChild(el('div', '', a.heldFiles.length + ' test' + (a.heldFiles.length === 1 ? ' was' : 's were') + ' held back (a name may still be on it). Open each one. No name: send it to Rowan. A name: grade it by hand.')).style.cssText = 'margin-bottom:6px;color:#4a2f00';
  a.heldFiles.forEach(function(h){
    var r = el('div'); r.style.cssText = 'display:flex;gap:10px;align-items:center;margin-top:6px;flex-wrap:wrap';
    r.appendChild(el('b', '', h.name));
    var o = el('a', '', 'Open'); o.href = h.url; o.target = '_blank'; o.style.color = 'var(--blue)'; r.appendChild(o);
    if (h.canSend) { var s = el('button', 'small', 'No name on it: send to Rowan'); s.type = 'button'; s.onclick = function(){ busy('Send to Rowan'); call('webReleaseHeld', a.run, h.id).then(function(x){ showResult('Send to Rowan', x); }).catch(fail('Send to Rowan')); }; r.appendChild(s); }
    else r.appendChild(el('span', '', '(not matched to a student: grade by hand)')).style.color = 'var(--muted)';
    box.appendChild(r);
  });
  return box;
}

/* ---------- Review and Match, inside this page ---------- */
function openView(name, title, run){
  $('viewTitle').textContent = title; $('view').hidden = false;
  var fr = $('frame'); fr.srcdoc = '<p style="font-family:Arial;padding:30px;color:#5f6b7a">Loading...</p>';
  call('webView', name, run || '').then(function(html){ fr.srcdoc = html; }).catch(function(e){ fr.srcdoc = '<p style="font-family:Arial;padding:30px;color:#b3261e">' + errText(e).replace(/</g, '&lt;') + '</p>'; });
}
function closeView(){ $('view').hidden = true; $('frame').srcdoc = ''; refresh(); }
$('viewBack').onclick = closeView;

/* ---------- settings ---------- */
function settings(note){
  var d = openDialog('Settings');
  if (note) { var n = el('p'); n.style.color = 'var(--amber)'; n.style.fontWeight = 'bold'; n.textContent = note; d.appendChild(n); }
  d.appendChild(el('p', '', ST && ST.canvas ? 'Canvas is connected (' + ST.canvasBase + ').' : 'Canvas is not connected.'));
  d.appendChild(el('label', '', 'Canvas web address'));
  var base = el('input'); base.value = (ST && ST.canvasBase) || 'https://ktufsd.instructure.com'; d.appendChild(base);
  d.appendChild(el('label', '', 'Canvas access token'));
  var tok = el('input'); tok.type = 'password'; tok.placeholder = 'In Canvas: Account > Settings > + New Access Token'; d.appendChild(tok);
  var msg = el('div', 'err'); d.appendChild(msg);
  var r1 = buttons(d, [{ label: 'Disconnect', cls: 'danger', fn: function(){ call('webDisconnectCanvas').then(function(r){ msg.style.color = 'var(--green)'; msg.textContent = r.text; refresh(); }); } },
    { label: 'Connect', cls: 'go', fn: function(){ msg.style.color = ''; msg.textContent = 'Checking...'; call('webConnectCanvas', base.value, tok.value).then(function(r){ tok.value = ''; msg.style.color = r.ok ? 'var(--green)' : ''; msg.textContent = r.ok ? r.text : r.error; refresh(); }).catch(function(e){ msg.textContent = errText(e); }); } }]);
  r1.style.justifyContent = 'flex-start';
  d.appendChild(el('p', '', 'The token is saved only in your own Google account. Never paste it anywhere else, including chats with Rowan.')).style.cssText = 'color:var(--muted);font-size:12px;margin-top:8px';
  d.appendChild(el('hr'));
  var lab = el('label'); var cb = el('input'); cb.type = 'checkbox'; cb.style.width = 'auto'; cb.checked = !!(ST && ST.autoRun);
  lab.appendChild(cb); lab.appendChild(document.createTextNode(' Auto-run: clean the Inbox every 10 minutes'));
  d.appendChild(lab);
  var am = el('div', 'err'); am.style.color = 'var(--green)'; d.appendChild(am);
  cb.onchange = function(){ am.textContent = 'Saving...'; call('webAutoRun', cb.checked).then(function(r){ am.textContent = r.text; refresh(); }); };
  d.appendChild(el('hr'));
  /* version 24: your classes */
  d.appendChild(el('label', '', 'Your classes'));
  d.appendChild(el('p', '', 'Tick the classes you teach. Only these show in the assignment lists, and Update roster adds their students.')).style.cssText = 'color:var(--muted);font-size:13px;margin:0 0 6px';
  var cbox = el('div'); cbox.style.cssText = 'border:1px solid var(--line);border-radius:8px;padding:6px 10px;max-height:220px;overflow:auto';
  cbox.appendChild(el('div', '', 'Loading your Canvas classes...')).style.color = 'var(--muted)';
  d.appendChild(cbox);
  var cmsg = el('div', 'err'); d.appendChild(cmsg);
  var cchecks = [];
  call('rpc', 'webCourseList', []).then(function(r){
    cbox.innerHTML = '';
    if (!r || !r.ok) { cbox.appendChild(el('div', '', (r && r.error) || 'Could not load your classes.')).style.color = 'var(--muted)'; return; }
    if (!r.courses.length) { cbox.appendChild(el('div', '', 'Canvas shows no classes where you are a teacher.')).style.color = 'var(--muted)'; return; }
    r.courses.forEach(function(c){
      var lab = el('label'); lab.style.cssText = 'display:flex;gap:8px;align-items:center;font-weight:normal;margin:6px 0';
      var cb = el('input'); cb.type = 'checkbox'; cb.style.width = 'auto'; cb.checked = c.shown; cb.value = c.id;
      lab.appendChild(cb);
      lab.appendChild(document.createTextNode(c.name + (c.term ? ' (' + c.term + ')' : '') + (c.students !== null ? ', ' + c.students + ' students' : '')));
      cbox.appendChild(lab); cchecks.push(cb);
    });
    if (!r.chosen) { cmsg.style.color = 'var(--amber)'; cmsg.textContent = 'All your Canvas classes show right now. Untick the ones that aren\u2019t yours, then Save classes.'; }
  }).catch(function(e){ cbox.innerHTML = ''; cbox.appendChild(el('div', '', errText(e))); });
  var rc = buttons(d, [{ label: 'Save classes', fn: function(){
    var ids = cchecks.filter(function(x){ return x.checked; }).map(function(x){ return x.value; });
    if (!cchecks.length) return;
    if (!ids.length) { cmsg.style.color = ''; cmsg.textContent = 'Tick at least one class.'; return; }
    cmsg.style.color = ''; cmsg.textContent = 'Saving...';
    call('rpc', 'webSaveCourses', [ids]).then(function(r){ cmsg.style.color = 'var(--green)'; cmsg.textContent = r.text; }).catch(function(e){ cmsg.textContent = errText(e); });
  } }]);
  rc.style.justifyContent = 'flex-start';
  d.appendChild(el('label', '', 'Update roster from Canvas'));
  d.appendChild(el('p', '', 'Adds new students from the classes ticked above. To add another course just once, paste a link to any page in it.')).style.cssText = 'color:var(--muted);font-size:13px';
  var links = el('textarea'); links.rows = 2; d.appendChild(links);
  var rm = el('div', 'err'); d.appendChild(rm);
  var r2 = buttons(d, [{ label: 'Update roster', fn: function(){ rm.style.color = ''; rm.textContent = 'Updating...'; call('webUpdateRoster', links.value).then(function(r){ rm.style.color = r.ok ? 'var(--green)' : ''; rm.textContent = r.ok ? r.text : r.error; }).catch(function(e){ rm.textContent = errText(e); }); } }]);
  r2.style.justifyContent = 'flex-start';
  var cl = buttons(d, [{ label: 'Class list (nicknames to hide)', fn: rosterDialog }]); cl.style.justifyContent = 'flex-start';
  d.appendChild(el('hr'));
  if (ST) {
    var lk = el('p'); [['The Sheet', ST.links.sheet], ['1 Inbox', ST.links.inbox], ['2 Ready for Claude', ST.links.ready], ['3 Processed originals (PRIVATE)', ST.links.done], ['Claude Grades Sheet', ST.links.grades]].forEach(function(x){ if (!x[1]) return; var a = el('a', '', x[0]); a.href = x[1]; a.target = '_blank'; a.style.marginRight = '14px'; lk.appendChild(a); });
    d.appendChild(lk);
  }
  buttons(d, [{ label: 'Close', cls: 'go', fn: function(){ closeDialog(); refresh(); } }]);
}

/* ---------- class list (version 15) ---------- */
function rosterDialog(){
  var d = busy('Class list', 'Loading...');
  call('webRosterList').then(function(list){
    var d2 = openDialog('Class list');
    d2.style.width = '720px';
    d2.appendChild(el('p', '', 'Add nicknames a student goes by (separate several with commas) so the tool hides those too. Tick "full name only" for a student whose first or last name is also a common word (Moon, Star, Will, Hope). Changes save on their own.'));
    var q = el('input'); q.placeholder = 'Search'; d2.appendChild(q);
    var t = el('div'); t.style.cssText = 'margin-top:10px;max-height:50vh;overflow:auto;border:1px solid var(--line);border-radius:8px';
    d2.appendChild(t);
    var rows = [];
    list.forEach(function(s){
      var r = el('div'); r.style.cssText = 'display:grid;grid-template-columns:1fr 1fr 110px;gap:8px;align-items:center;padding:6px 10px;border-bottom:1px solid var(--line)';
      r.appendChild(el('span', '', s.name));
      var n = el('input'); n.value = s.nick; n.placeholder = 'nicknames'; r.appendChild(n);
      var l = el('label'); l.style.cssText = 'font-weight:normal;margin:0;font-size:13px'; var c = el('input'); c.type = 'checkbox'; c.style.width = 'auto'; c.checked = s.full; l.appendChild(c); l.appendChild(document.createTextNode(' full name only')); r.appendChild(l);
      var saveRow = function(){ call('webRosterSave', s.row, s.id, n.value, c.checked).then(function(x){ r.style.background = x.ok ? 'var(--green-soft)' : 'var(--red-soft)'; setTimeout(function(){ r.style.background = ''; }, 900); }); };
      n.onchange = saveRow; c.onchange = saveRow;
      t.appendChild(r); rows.push({ r: r, name: s.name.toLowerCase() });
    });
    q.oninput = function(){ var v = q.value.toLowerCase(); rows.forEach(function(x){ x.r.style.display = x.name.indexOf(v) > -1 ? '' : 'none'; }); };
    buttons(d2, [{ label: 'Done', cls: 'go', fn: function(){ settings(); } }]);
  }).catch(fail('Class list'));
}

/* ---------- the list ---------- */
var STEP_NAMES = [['get', 'Get the work'], ['clean', 'Remove names'], ['rowan', 'Rowan grades'], ['review', 'Your review'], ['send', 'Send to Canvas'], ['finish', 'Finish']];
function stepState(a){
  var s = {};
  if (a.next === 'old') return s;
  var hasWork = a.coded > 0 || a.graded > 0 || a.inbox || a.scans || a.toMatch || a.toBuild;
  s.get = hasWork ? 'done' : 'now';
  s.clean = (a.inbox || a.scans || a.toMatch || a.toBuild) ? 'now' : (a.coded || a.graded ? 'done' : '');
  s.rowan = a.next === 'rowan' ? 'now' : (a.graded && a.graded >= a.coded ? 'done' : '');
  s.review = a.review ? (a.next === 'review' ? 'now' : '') : (a.graded ? 'done' : '');
  s.send = a.toSend ? (a.next === 'send' ? 'now' : '') : (a.graded ? 'done' : '');
  s.finish = a.next === 'finish' ? 'now' : '';
  // nothing after the step you're on shows as done yet
  var seen = false;
  STEP_NAMES.forEach(function(x){ if (seen && s[x[0]] === 'done') s[x[0]] = ''; if (s[x[0]] === 'now') seen = true; });
  return s;
}
function stepText(k, a){
  if (k === 'get') return a.next === 'old' ? 'nothing left' : a.coded ? a.coded + ' student' + (a.coded === 1 ? '' : 's') + (a.extraFiles ? ' (+' + a.extraFiles + ' extra file' + (a.extraFiles === 1 ? '' : 's') + ')' : '') : (a.paper ? 'scan it' : 'pull it');
  if (k === 'clean') return a.toMatch ? a.toMatch + ' pages to match' : (a.inbox || a.scans) ? 'waiting' : a.toBuild ? 'saving' : (a.held ? a.held + ' held' : '');
  if (k === 'rowan') return a.toGrade ? a.toGrade + ' to grade' : (a.graded ? a.graded + ' graded' : '');
  if (k === 'review') return a.review ? a.review + ' waiting' : '';
  if (k === 'send') return a.toSend ? a.toSend + ' to send' : (a.graded ? 'all sent (click to send again)' : '');
  return '';
}
function render(){
  $('ver').textContent = ST && ST.version ? 'Version ' + ST.version : '';
  $('canvasChip').textContent = ST.canvas ? 'Canvas connected' : 'Canvas not connected';
  $('canvasChip').className = 'chip ' + (ST.canvas ? 'on' : 'off');
  $('autoChip').textContent = 'Auto-run ' + (ST.autoRun ? 'on' : 'off');
  $('autoChip').className = 'chip ' + (ST.autoRun ? 'on' : '');
  var lo = $('loose'); lo.innerHTML = '';
  if (ST.looseFiles) { var b = el('div', 'banner'); b.appendChild(el('span', '', 'There are files sitting loose in the Inbox.')); var x = el('button', 'small go', 'Clean them'); x.onclick = function(){ clean(); }; b.appendChild(x); lo.appendChild(b); }
  var list = $('list'); list.innerHTML = '';
  if (!ST.assignments.length) { list.appendChild(el('div', 'empty', 'Nothing here yet. Start with "Get projects from Canvas" or "Paper test or quiz".')); return; }
  ST.assignments.forEach(function(a){
    var c = el('div', 'card');
    var top = el('div', 'top'); top.appendChild(el('span', 'name', a.run)); top.appendChild(el('span', 'badge' + (a.paper ? ' paper' : ''), a.paper ? 'Paper test' : 'Assignment'));
    c.appendChild(top);
    var st = stepState(a), sp = el('div', 'stepper');
    STEP_NAMES.forEach(function(s){ var d = el('div', st[s[0]] || ''); d.appendChild(el('b', '', s[1])); d.appendChild(document.createTextNode(stepText(s[0], a) || ' ')); sp.appendChild(d);
      if (s[0] === 'send' && (a.graded || a.toSend)) { d.className += ' click'; d.title = 'Send to Canvas'; d.setAttribute('role', 'button'); d.tabIndex = 0; d.onclick = function(){ send(a); }; d.onkeydown = function(e){ if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); send(a); } }; } });
    c.appendChild(sp);
    var act = el('div', 'actions');
    var main = null;
    if (a.next === 'clean') main = ['Remove names now', function(){ clean(); }];
    else if (a.next === 'match') main = ['Match names (' + a.toMatch + ' pages)', function(){ openView('match', 'Match names and remove them'); }];
    else if (a.next === 'rowan') act.appendChild(el('span', 'wait', 'Waiting for Rowan: open your Rowan chat and say “grade.”'));
    else if (a.next === 'review') main = ['Review ' + a.review + ' answer' + (a.review === 1 ? '' : 's'), function(){ openView('review', 'Review what Rowan was not sure about'); }];
    else if (a.next === 'send') main = ['Send ' + a.toSend + ' grade' + (a.toSend === 1 ? '' : 's') + ' to Canvas', function(){ send(a); }];
    else if (a.next === 'get') main = a.paper ? ['Upload the class scan', function(){ uploadScan(a.run); }] : ['Pull from Canvas', function(){ pull(a.canvas); }];
    if (main) { var m = el('button', 'go', main[0]); m.type = 'button'; m.onclick = main[1]; act.appendChild(m); }
    var extra = [];
    if (a.coded || a.graded) extra.push(['See the work', function(){ openView('viewer', a.run, a.run); }]);
    if (a.checks) extra.push(['Look over ' + a.checks + ' flagged', function(){ openView('viewer', a.run, a.run); }]);
    if (a.paper) extra.push([(a.key || a.notes) ? 'Answer key and notes \u2713' : 'Add answer key and notes', function(){ keyDialog(a); }]);
    if (a.paper && a.next !== 'get') extra.push(['Upload more scans (make-ups)', function(){ uploadScan(a.run); }]);
    if (a.canvas && !a.paper && a.next !== 'get') extra.push(['Pull late work', function(){ pull(a.canvas); }]);
    if (a.review && a.next !== 'review') extra.push(['Review (' + a.review + ')', function(){ openView('review', 'Review what Rowan was not sure about'); }]);
    if ((a.toSend || a.graded) && a.next !== 'send') extra.push([a.toSend ? 'Send to Canvas' : 'Send to Canvas again', function(){ send(a); }]);
    if (a.graded || a.next === 'old') extra.push([a.next === 'old' ? 'Finish (clear it out)' : 'Finish (delete its files)', function(){ finish(a); }, 'danger']);
    if (a.next === 'old') act.appendChild(el('span', 'facts', 'Nothing for this one is left in the folders and Rowan has no grades under this name. It is probably an old run. Finish clears its leftover codes and files. '));
    extra.forEach(function(x){ var b2 = el('button', 'small' + (x[2] ? ' ' + x[2] : ''), x[0]); b2.type = 'button'; b2.onclick = x[1]; act.appendChild(b2); });
    if (a.next === 'finish') act.insertBefore(el('span', 'facts', 'All sent. Finish it once no more late work is coming. '), act.firstChild);
    c.appendChild(act);
    if (a.heldFiles && a.heldFiles.length) c.appendChild(heldBox(a));
    list.appendChild(c);
  });
}
function refresh(){
  $('refresh').textContent = 'Checking...'; $('refresh').disabled = true;
  return call('webStatus').then(function(s){ ST = s; render(); }).catch(function(e){ if (!ST) { $('list').innerHTML = ''; $('list').appendChild(el('div', 'empty', 'Could not load: ' + errText(e))); } })
    .then(function(){ $('refresh').textContent = 'Refresh'; $('refresh').disabled = false; });
}

/* ---------- updates (version 17) ---------- */
function checkUpdate(){
  call('webUpdateInfo').then(function(u){
    var box = $('update'); box.innerHTML = '';
    if (!u || !u.ready) return;
    var b = el('div', 'banner');
    b.appendChild(el('span', '', 'Update ready from Rowan (version ' + u.latest + '). ' + (u.notes || '')));
    var x = el('button', 'small go', 'Install'); x.type = 'button';
    x.onclick = function(){
      busy('Install the update', 'Installing. This takes about 10 seconds...');
      call('webInstallUpdate').then(function(r){
        var d = openDialog('Install the update');
        if (r.ok) { d.appendChild(el('p', '', r.text)); buttons(d, [{ label: 'Reload now', cls: 'go', fn: function(){ try { window.top.location.reload(); } catch (e) { location.reload(); } } }]); }
        else { var p = el('p', '', r.error); p.style.color = 'var(--red)'; d.appendChild(p); buttons(d, [{ label: 'OK', cls: 'go', fn: closeDialog }]); }
      }).catch(fail('Install the update'));
    };
    b.appendChild(x); box.appendChild(b);
  }).catch(function(){});
}
$('tPull').onclick = function(){ pull(); };
$('tPaper').onclick = paper;
$('tClean').onclick = function(){ clean(); };
$('refresh').onclick = refresh;
$('settingsBtn').onclick = function(){ settings(); };
$('canvasChip').onclick = function(){ settings(); };
$('autoChip').onclick = function(){ settings(); };
if (CACHED) { ST = CACHED; render(); }
refresh().then(checkUpdate);
</script>
</body></html>`;

/* Version 19: checking for updates, Install and Undo are in the starter (Code.gs). */

/* ============================== CLEANUP =========================== */

/** Permanently deletes one assignment's files and rows once you are done with it (including late work). */
function finishAssignment() {
  showNamePicker_('finish', 'Finish an assignment',
    'Which assignment is finished? (Grades are in Canvas and no more late work is coming.) You will be asked to confirm before anything is deleted.', false);
}

function finishAssignmentFor_(run) {
  const ui = ui_();
  if (!run) return;
  const sure = ui.alert('Finish "' + run + '"?',
    'This PERMANENTLY deletes, for this assignment only:\n' +
    '- the original files (with names) and the coded PDFs\n' +
    '- for a paper test: the scans and its pages\n' +
    '- its rows on the Key, Results, Log and Canvas Pulls tabs, and in Claude\'s grades Sheet (including Review)\n\n' +
    'Make sure the grades are in Canvas first. This cannot be undone.', ui.ButtonSet.YES_NO);
  if (sure !== ui.Button.YES) return;

  const f = folders_();
  const codes = {};
  ss_().getSheetByName(TAB_KEY).getDataRange().getValues().slice(1)
    .forEach(r => { if (r[0] === run) codes[String(r[1])] = true; });

  [f.done, f.ready, f.inbox].forEach(parent => {
    const it = parent.getFoldersByName(run);
    while (it.hasNext()) deleteForever_(it.next().getId());
  });
  const scans = f.inbox.getFoldersByName(run + SCAN_SUFFIX);   // version 10: paper test scans
  while (scans.hasNext()) deleteForever_(scans.next().getId());
  emptyFolder_(f.temp);

  deleteRows_(ss_().getSheetByName(TAB_RESULTS), r => codes[String(r[0]).trim()], 2);
  deleteRows_(ss_().getSheetByName(TAB_LOG), r => r[1] === run, 2);
  deleteRows_(ss_().getSheetByName(TAB_PULLS), r => r[0] === run, 2);
  deleteRows_(ss_().getSheetByName(TAB_KEY), r => r[0] === run, 2);
  const cg = claudeGradesSheet_();
  if (cg) deleteRows_(cg, r => codes[String(r[1]).trim()] || r[0] === run, 2);
  const book = claudeGradesBook_();
  const rv = book ? book.getSheetByName(REVIEW_TAB) : null;
  if (rv) deleteRows_(rv, r => codes[String(r[1]).trim()] || r[0] === run, 2);
  deleteRows_(ss_().getSheetByName(TAB_SCANS), r => r[0] === run, 2);
  deleteRows_(ss_().getSheetByName(TAB_PLAN), r => r[0] === run, 2);
  PropertiesService.getDocumentProperties().deleteProperty('SOUT::' + run);

  // Originals left loose by versions 1 to 4 (before they were kept per assignment)
  const loose = [];
  const it = f.done.getFiles();
  while (it.hasNext()) loose.push(it.next().getId());
  let extra = '';
  if (loose.length && ui.alert('Older originals',
      loose.length + ' original file(s) from before version 5 are loose in "3 Processed originals." They have names in them. Delete them too?',
      ui.ButtonSet.YES_NO) === ui.Button.YES) {
    loose.forEach(deleteForever_);
    extra = ' Also deleted ' + loose.length + ' older original(s).';
  }
  ui.alert('"' + run + '" is finished. Its files and rows are deleted.' + extra +
    '\n\nAlso delete any copies you downloaded to your computer (submissions.zip, Grades csv files).');
}

function deleteForever_(id) {
  try { Drive.Files.remove ? Drive.Files.remove(id) : Drive.Files.delete(id); }
  catch (e) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e2) { try { DriveApp.getFolderById(id).setTrashed(true); } catch (e3) { /* ignore */ } } }
}

function emptyFolder_(folder) {
  const it = folder.getFiles();
  while (it.hasNext()) deleteForever_(it.next().getId());
}

function deleteRows_(sh, test, firstRow) {
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= firstRow - 1; i--) {
    if (test(data[i])) sh.deleteRow(i + 1);
  }
}

/* ========================= FILE HANDLING ========================== */

async function processFile_(file, ctx) {
  let target = file;
  let mime = file.getMimeType();
  const name = file.getName();

  if (mime === 'application/vnd.google-apps.shortcut') {
    target = DriveApp.getFileById(file.getTargetId());
    mime = target.getMimeType();
  }

  let owner = '';
  try { owner = target.getOwner() ? target.getOwner().getEmail().toLowerCase() : ''; } catch (e) { owner = ''; }

  const student = matchStudent_(name, ctx.roster, owner) ||
                  (target !== file ? matchStudent_(target.getName(), ctx.roster, owner) : null);
  const code = student ? ctx.codeMap[studentKey_(student)] : unmatchedCode_(ctx.run, name);
  const notes = student ? [] : ['Could not match this file to a student. Original file: ' + name];

  const ext = (extOf_(name) || extOf_(target.getName())).toLowerCase();
  let outName = null;

  if (mime === MimeType.GOOGLE_DOCS) {
    outName = copyScrubExport_(target, 'doc', code, ctx);
  } else if (mime === MimeType.GOOGLE_SLIDES) {
    outName = copyScrubExport_(target, 'slides', code, ctx);
  } else if (mime === MimeType.GOOGLE_SHEETS) {
    outName = copyScrubExport_(target, 'sheet', code, ctx);
  } else if (mime.indexOf('application/vnd.google-apps') === 0) {
    notes.push('Google file type not supported (' + mime.replace('application/vnd.google-apps.', '') + '). Check it yourself.');
  } else if (['gdoc', 'gslides', 'gsheet'].indexOf(ext) > -1) {
    // Google Drive for desktop shortcut files: a small text file pointing at the real Google file
    const r = handleGoogleShortcut_(target, code, ctx);
    outName = r.outName;
    if (r.note) notes.push(r.note);
  } else if (['key', 'pages', 'numbers'].indexOf(ext) > -1) {
    notes.push('Apple ' + ext + ' file. Google cannot open these. Ask the student for a PDF (File > Export To > PDF).');
  } else if (DOC_EXT.indexOf(ext) > -1) {
    outName = importScrubExport_(target, MimeType.GOOGLE_DOCS, 'doc', code, ctx);
  } else if (SLIDE_EXT.indexOf(ext) > -1) {
    outName = importScrubExport_(target, MimeType.GOOGLE_SLIDES, 'slides', code, ctx);
  } else if (SHEET_EXT.indexOf(ext) > -1) {
    outName = importScrubExport_(target, MimeType.GOOGLE_SHEETS, 'sheet', code, ctx);
  } else if (ext === 'pdf' && ctx.lastTry) {   // version 20: too slow page by page twice; one-step rebuild
    outName = importScrubExport_(target, MimeType.GOOGLE_DOCS, 'doc', code, ctx);
    notes.push('This PDF was too slow to check page by page, so it was rebuilt from its text in one step. Layout or pictures may look different; check it by eye.');
  } else if (ext === 'pdf') {
    const r = await handlePdf_(target, code, ctx);
    outName = r.outName;
    if (r.note) notes.push(r.note);
  } else if (ext === 'html' || ext === 'htm') {
    const r = handleHtml_(target, code, ctx);
    outName = r.outName;
    if (r.note) notes.push(r.note);
  } else if (IMG_EXT.indexOf(ext) > -1) {
    try {
      const png = target.getBlob().getAs('image/png');   // re-encoding drops photo metadata (EXIF/GPS)
      outName = nextOutName_(code, '.png', ctx.outFolder);
      ctx.outFolder.createFile(png.setName(outName));
      notes.push('Image: check it by eye for names, faces or handwriting.');
    } catch (e) {
      notes.push('Image could not be copied (' + ext + '). Check the original yourself.');
    }
  } else if (ext === 'zip') {
    notes.push('A zip inside the zip. Unzip it and drop those files into the Inbox.');
  } else {
    notes.push('Not processed (' + (ext || mime) + '). Video, Apple Pages/Keynote and similar files need a manual check.');
  }

  return { code: code, outName: outName, note: notes.join(' ') };
}

function googleFileById_(id, code, ctx) {
  try {
    const gf = DriveApp.getFileById(id);
    const m = gf.getMimeType();
    const kind = m === MimeType.GOOGLE_SLIDES ? 'slides' : m === MimeType.GOOGLE_SHEETS ? 'sheet' : m === MimeType.GOOGLE_DOCS ? 'doc' : null;
    if (kind) return { outName: copyScrubExport_(gf, kind, code, ctx), note: '' };
    return { outName: null, note: 'Link to a Google file type it cannot clean. Check it yourself.' };
  } catch (e) {
    return { outName: null, note: 'Google file you do not have access to. Open it from Canvas, use File > Download > PDF, rename the PDF to the student\'s name, and drop it in the Inbox.' };
  }
}

function handleGoogleShortcut_(file, code, ctx) {
  const text = file.getBlob().getDataAsString();
  let id = '';
  try {
    const j = JSON.parse(text);
    id = j.doc_id || String(j.resource_id || '').split(':').pop() || '';
    if (!id && j.url) { const m = String(j.url).match(/\/d\/([a-zA-Z0-9_-]{20,})/); if (m) id = m[1]; }
  } catch (e) {
    const m = text.match(/\/d\/([a-zA-Z0-9_-]{20,})/) || text.match(/"doc_id"\s*:\s*"([a-zA-Z0-9_-]{20,})"/);
    if (m) id = m[1];
  }
  if (!id) return { outName: null, note: 'Google shortcut file with no link inside. Open the work from Canvas, download it as a PDF, rename it to the student\'s name, and drop it in the Inbox.' };
  return googleFileById_(id, code, ctx);
}

function handleHtml_(file, code, ctx) {
  const html = file.getBlob().getDataAsString();
  const g = html.match(/https:\/\/(?:docs|drive)\.google\.com\/[^"'\s<>]*?\/d\/([a-zA-Z0-9_-]{20,})/);
  if (g) return googleFileById_(g[1], code, ctx);
  const isLinkOnly = /http-equiv=["']?refresh/i.test(html) ||
    (html.replace(/<[^>]+>/g, ' ').trim().length < 400 && /https?:\/\//i.test(html));
  if (isLinkOnly) {
    const url = (html.match(/https?:\/\/[^"'\s<>]+/) || [''])[0];
    return { outName: null, note: 'Link submission, not a file (' + url.split('?')[0] + '). Open it and export a PDF into the Inbox if you want it graded.' };
  }
  // Canvas text-entry submission
  return { outName: importScrubExport_(file, MimeType.GOOGLE_DOCS, 'doc', code, ctx), note: '' };
}

function copyScrubExport_(googleFile, kind, code, ctx) {
  const tmp = googleFile.makeCopy('tmp_' + Utilities.getUuid(), ctx.temp);
  return scrubAndExport_(tmp.getId(), kind, code, ctx);
}

function importScrubExport_(file, targetMime, kind, code, ctx) {
  let blob = file.getBlob();
  const ext = extOf_(file.getName()).toLowerCase();
  if (MIME_BY_EXT[ext]) blob = blob.setContentType(MIME_BY_EXT[ext]);
  const tmpId = importAs_(blob, targetMime, ctx.temp.getId(), ext === 'pdf');
  return scrubAndExport_(tmpId, kind, code, ctx);
}

function scrubAndExport_(tmpId, kind, code, ctx) {
  try {
    if (kind === 'doc') scrubDocument_(tmpId, ctx.rules);
    else if (kind === 'slides') scrubSlides_(tmpId, ctx.rules);
    else if (kind === 'sheet') scrubSheet_(tmpId, ctx.rules);
    const outName = nextOutName_(code, '.pdf', ctx.outFolder);
    const tmp = DriveApp.getFileById(tmpId);
    tmp.setName(outName.replace(/\.pdf$/, ''));
    ctx.outFolder.createFile(tmp.getAs(MimeType.PDF).setName(outName));
    return outName;
  } finally {
    try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e) { /* ignore */ }
  }
}

function importAs_(blob, targetMime, folderId, isPdf) {
  const title = 'tmp_' + Utilities.getUuid();
  if (Drive.Files.create) {   // Drive API v3 (current default)
    return Drive.Files.create({ name: title, mimeType: targetMime, parents: [folderId] }, blob).id;
  }
  return Drive.Files.insert(  // Drive API v2
    { title: title, mimeType: targetMime, parents: [{ id: folderId }] }, blob,
    { convert: true, ocr: !!isPdf }).id;
}

/* ========================= PDFs, PAGE BY PAGE ===================== */
/*
 * A PDF keeps its original look. Each page is checked for a student's name
 * (or email or Canvas ID). Pages without one are copied exactly as the student
 * made them. Only a page WITH a name is swapped for a plain version of that page
 * with the code in the name's place. The new PDF has no hidden title, author,
 * comments or annotations from the original.
 */

function loadPdfLib_() {
  if (typeof globalThis.PDFLib !== 'undefined') return globalThis.PDFLib;
  if (typeof globalThis.setTimeout === 'undefined') {
    globalThis.setTimeout = function (fn) { fn(); return 0; };
    globalThis.clearTimeout = function () {};
  }
  const code = UrlFetchApp.fetch(PDFLIB_URL).getContentText();
  (0, eval)(code);
  if (typeof globalThis.PDFLib === 'undefined') throw new Error('the PDF add-on did not load');
  return globalThis.PDFLib;
}

function u8ToBlob_(u8, name) {
  const arr = new Array(u8.length);
  for (let i = 0; i < u8.length; i++) { const b = u8[i]; arr[i] = b > 127 ? b - 256 : b; }
  return Utilities.newBlob(arr, 'application/pdf', name);
}

/** Google reads a PDF's words (Google Docs conversion). Returns the temp doc id and its text. */
function ocrPdf_(u8, ctx) {
  const id = importAs_(u8ToBlob_(u8, 'tmp.pdf'), MimeType.GOOGLE_DOCS, ctx.temp.getId(), true);
  let text = '';
  try { text = DocumentApp.openById(id).getBody().getText(); } catch (e) { text = ''; }
  return { id: id, text: text };
}

function trashQuietly_(id) {
  try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* ignore */ }
}

/** True if the text holds any student's name, email or Canvas ID (also catches near-misspellings). */
function hasIdentifier_(text, ctx) {
  if (!text) return false;
  for (let i = 0; i < ctx.rules.length; i++) if (ctx.rules[i].js.test(text)) return true;
  return fuzzyHit_(text, ctx.fuzzy);
}

function buildFuzzy_(roster) {
  const singles = {}, pairs = {};
  roster.forEach(s => {
    const firstMain = letters_(s.first.split(/\s+/)[0]);
    const last = letters_(s.last);
    if (firstMain && last) pairs[firstMain + last] = true;
    if (!s.fullOnly) {
      if (firstMain.length >= 5) singles[firstMain] = true;
      if (last.length >= 5) singles[last] = true;
    }
    s.extras.forEach(x => { const l = letters_(x); if (l.length >= 5) singles[l] = true; });
  });
  return { singles: Object.keys(singles), pairs: Object.keys(pairs) };
}

/** Catches a name that Google's reading got one letter wrong (Smlth for Smith). */
function fuzzyHit_(text, fz) {
  if (!fz) return false;
  const toks = String(text).split(/[^A-Za-zÀ-ɏ'’-]+/).map(letters_).filter(t => t.length >= 2);
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    for (let k = 0; k < fz.singles.length && t.length >= 4; k++) {
      const s = fz.singles[k];
      if (t[0] === s[0] && Math.abs(t.length - s.length) <= 1 && editDistance_(t, s, 1) <= 1) return true;
    }
    if (i + 1 < toks.length) {
      const p = t + toks[i + 1];
      for (let k = 0; k < fz.pairs.length; k++) {
        const s = fz.pairs[k];
        if (p[0] === s[0] && Math.abs(p.length - s.length) <= 2 && editDistance_(p, s, 2) <= 2) return true;
      }
    }
  }
  return false;
}

function editDistance_(a, b, max) {
  let prev = [];
  for (let j = 0; j <= b.length; j++) prev[j] = j;
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    let best = i;
    for (let j = 1; j <= b.length; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      if (cur[j] < best) best = cur[j];
    }
    if (best > max) return max + 1;
    prev = cur;
  }
  return prev[b.length];
}

/** A plain, name-free version of one page: Google's reading of it, names swapped for codes, saved as PDF. */
function plainPageFromDoc_(docId, ctx, size) {
  if (size) {
    try {   // same shape as the original page (a landscape slide stays landscape)
      const d = DocumentApp.openById(docId);
      d.getBody().setPageWidth(size.width).setPageHeight(size.height)
        .setMarginTop(36).setMarginBottom(36).setMarginLeft(36).setMarginRight(36);
      d.saveAndClose();
    } catch (e) { /* keep Google's default page */ }
  }
  scrubDocument_(docId, ctx.rules);
  return new Uint8Array(DriveApp.getFileById(docId).getAs(MimeType.PDF).getBytes());
}

/**
 * Builds the clean PDF. Returns { bytes, total, kept, swapped }.
 * hooks (used for testing outside Google): { ocr(u8) -> {id, text}, plain(id) -> u8, done(id) }
 */
async function redactPdf_(PDFLib, srcBytes, ctx, hooks) {
  const src = await PDFLib.PDFDocument.load(srcBytes, { ignoreEncryption: false, updateMetadata: false });
  const total = src.getPageCount();
  const out = await PDFLib.PDFDocument.create();
  const timeUp = () => ctx.start && Date.now() - ctx.start > HARD_STOP_MS;

  // Every page is read on its own. (Reading the whole PDF at once can skip pages on long files.)
  let kept = 0, swapped = 0;
  for (let i = 0; i < total; i++) {
    if (timeUp()) throw new Error(OUT_OF_TIME);
    const one = await PDFLib.PDFDocument.create();
    const [single] = await one.copyPages(src, [i]);
    one.addPage(single);
    const pageDoc = hooks.ocr(await one.save());
    const needsSwap = hasIdentifier_(pageDoc.text, ctx);
    if (!needsSwap) {
      hooks.done(pageDoc.id);
      const [p] = await out.copyPages(src, [i]);
      p.node.delete(PDFLib.PDFName.of('Annots'));   // comments, sticky notes and links can carry names
      out.addPage(p);
      kept++;
      continue;
    }
    // Swap this page for a plain version, sized like the original page
    const size = src.getPage(i).getSize();
    const plainBytes = await hooks.plain(pageDoc.id, size);
    hooks.done(pageDoc.id);
    const plain = await PDFLib.PDFDocument.load(plainBytes);
    const embedded = await out.embedPdf(plain, plain.getPageIndices());
    embedded.forEach(ep => {
      const pg = out.addPage([size.width, size.height]);
      const s = Math.min(size.width / ep.width, size.height / ep.height);
      pg.drawPage(ep, { x: (size.width - ep.width * s) / 2, y: (size.height - ep.height * s) / 2, xScale: s, yScale: s });
    });
    swapped++;
  }
  out.setTitle(ctx.codeForTitle || 'Student work');
  out.setAuthor('');
  out.setSubject('');
  out.setKeywords([]);
  out.setCreator('Student Work Anonymizer');
  out.setProducer('Student Work Anonymizer');
  return { bytes: await out.save(), total: total, kept: kept, swapped: swapped };
}

async function handlePdf_(file, code, ctx) {
  let result;
  try {
    const PDFLib = loadPdfLib_();
    const hooks = {
      ocr: u8 => ocrPdf_(u8, ctx),
      plain: (id, size) => plainPageFromDoc_(id, ctx, size),
      done: id => trashQuietly_(id)
    };
    ctx.codeForTitle = code;
    result = await redactPdf_(PDFLib, new Uint8Array(file.getBlob().getBytes()), ctx, hooks);
  } catch (e) {
    if (e && e.message === OUT_OF_TIME) throw e;
    // Locked, damaged or unusual PDF: fall back to rebuilding it from its text (version 2 behavior)
    const outName = importScrubExport_(file, MimeType.GOOGLE_DOCS, 'doc', code, ctx);
    return { outName: outName, note: 'This PDF could not be handled page by page (' + e.message + '), so it was rebuilt from its text. Layout or pictures may look different.' };
  }
  const outName = nextOutName_(code, '.pdf', ctx.outFolder);
  ctx.outFolder.createFile(u8ToBlob_(result.bytes, outName));
  let note = '';
  if (result.swapped) {
    note = result.swapped + ' of ' + result.total + ' page(s) had a name and were swapped for a plain version with the code. The other pages are exactly as turned in.';
  }
  return { outName: outName, note: note };
}

/* =========================== SCRUBBING ============================ */

function scrubDocument_(docId, rules) {
  const doc = DocumentApp.openById(docId);
  const parts = [];
  let tabs = [];
  try { tabs = allTabs_(doc.getTabs()); } catch (e) { tabs = []; }
  if (tabs.length) {
    tabs.forEach(t => {
      try {
        const dt = t.asDocumentTab();
        parts.push(dt.getBody(), dt.getHeader(), dt.getFooter());
        dt.getFootnotes().forEach(fn => parts.push(fn.getFootnoteContents()));
      } catch (e) { /* not a document tab */ }
    });
  } else {
    parts.push(doc.getBody(), doc.getHeader(), doc.getFooter());
    doc.getFootnotes().forEach(fn => parts.push(fn.getFootnoteContents()));
  }
  parts.forEach(p => {
    if (!p) return;
    let text = p.getText();
    rules.forEach(r => {
      if (r.js.test(text)) {
        p.replaceText(r.re2, r.replacement);
        text = p.getText();
      }
    });
  });
  doc.saveAndClose();
}

function allTabs_(tabs) {
  let out = [];
  tabs.forEach(t => { out.push(t); out = out.concat(allTabs_(t.getChildTabs())); });
  return out;
}

function scrubSlides_(presId, rules) {
  const pres = SlidesApp.openById(presId);
  const ranges = [];
  const visit = el => {
    const t = el.getPageElementType();
    if (t === SlidesApp.PageElementType.SHAPE) {
      ranges.push(el.asShape().getText());
    } else if (t === SlidesApp.PageElementType.TABLE) {
      const tb = el.asTable();
      for (let r = 0; r < tb.getNumRows(); r++) {
        for (let c = 0; c < tb.getNumColumns(); c++) {
          try { ranges.push(tb.getCell(r, c).getText()); } catch (e) { /* merged cell */ }
        }
      }
    } else if (t === SlidesApp.PageElementType.GROUP) {
      el.asGroup().getChildren().forEach(visit);
    }
  };
  pres.getSlides().forEach(s => {
    s.getPageElements().forEach(visit);
    try { ranges.push(s.getNotesPage().getSpeakerNotesShape().getText()); } catch (e) { /* no notes */ }
  });
  ranges.forEach(tr => {
    let text;
    try { text = tr.asString(); } catch (e) { return; }
    rules.forEach(r => {
      if (!r.js.test(text)) return;
      const hits = tr.find(r.re2);
      for (let i = hits.length - 1; i >= 0; i--) hits[i].setText(r.replacement);
      text = tr.asString();
    });
  });
  pres.saveAndClose();
}

function scrubSheet_(ssId, rules) {
  const ss = SpreadsheetApp.openById(ssId);
  rules.forEach(r => {
    ss.createTextFinder(r.re2).useRegularExpression(true).matchCase(true).replaceAllWith(r.replacement);
  });
  SpreadsheetApp.flush();
}

/**
 * One rule for emails, then one rule per student covering every form of
 * their name plus their Canvas ID. Each student's mentions become [S-1234].
 * Patterns are case-insensitive and match whole words only.
 */
function buildRules_(roster, codeMap) {
  const rules = [{
    js: /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i,
    re2: '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}',
    replacement: '[email removed]'
  }];
  roster.forEach(s => {
    const code = codeMap[studentKey_(s)];
    const vars = variantsFor_(s).sort((a, b) => b.length - a.length);
    const re2Parts = vars.map(v => wordWrap_(v, ciPattern_(v)));
    const jsParts = vars.map(v => wordWrap_(v, escapeRe_(v).replace(/\s+/g, '\\s+')));
    if (s.id && s.id.length >= 4) {
      re2Parts.push('\\b' + s.id + '\\b');
      jsParts.push('\\b' + s.id + '\\b');
    }
    if (!re2Parts.length) return;
    rules.push({
      js: new RegExp('(?:' + jsParts.join('|') + ')', 'i'),
      re2: '(?:' + re2Parts.join('|') + ')',
      replacement: '[' + code + ']'
    });
  });
  return rules;
}

function variantsFor_(s) {
  const v = {};
  const add = x => { x = String(x || '').trim(); if (x.length >= 2) v[x] = true; };
  const firstMain = s.first.split(/\s+/)[0];
  if (s.first && s.last) {
    add(s.first + ' ' + s.last);
    add(firstMain + ' ' + s.last);
    add(s.last + ', ' + s.first);
    add(s.last + ', ' + firstMain);
    add(s.last + ' ' + firstMain);
  }
  if (!s.fullOnly) {
    add(s.first);
    add(firstMain);
    add(s.last);
    s.last.split(/[\s-]+/).forEach(p => { if (p.length > 2) add(p); });
  }
  s.extras.forEach(add);
  return Object.keys(v);
}

function ciPattern_(str) {
  return str.split('').map(ch => {
    if (/\s/.test(ch)) return '\\s+';
    const lo = ch.toLowerCase(), up = ch.toUpperCase();
    if (lo !== up) return '[' + lo + up + ']';
    return escapeRe_(ch);
  }).join('').replace(/(\\s\+)+/g, '\\s+');
}

function wordWrap_(raw, pattern) {
  return (/^\w/.test(raw) ? '\\b' : '') + pattern + (/\w$/.test(raw) ? '\\b' : '');
}

function escapeRe_(s) {
  return s.replace(/[.*+?^${}()|[\]\\\/-]/g, '\\$&');
}

/* ======================= ROSTER & MATCHING ======================== */

function readRoster_() {
  const sh = ss_().getSheetByName(TAB_ROSTER);
  if (!sh) throw new Error('Run Anonymizer > Set up first.');
  const vals = sh.getDataRange().getValues().slice(1);
  const out = [];
  vals.forEach(row => {
    const raw = String(row[0] || '').trim();
    const id = String(row[1] || '').trim().replace(/\.0+$/, '');
    if (!raw || /points possible|test student|student,\s*test/i.test(raw)) return;
    let first, last;
    if (raw.indexOf(',') > -1) {
      last = raw.split(',')[0].trim();
      first = raw.split(',').slice(1).join(',').trim();
    } else {
      const p = raw.split(/\s+/);
      first = p[0];
      last = p.slice(1).join(' ');
    }
    out.push({
      raw: raw, first: first, last: last, id: id,
      extras: String(row[2] || '').split(',').map(x => x.trim()).filter(Boolean),
      fullOnly: /^y/i.test(String(row[3] || '')),
      email: String(row[4] || '').trim().toLowerCase(),
      key: letters_(last + first),
      keyFL: letters_(first + last)
    });
  });
  return out;
}

/** Canvas names downloads like "smithjohn_LATE_123456_7890123_Project.docx". */
function matchStudent_(fileName, roster, ownerEmail) {
  const tokens = fileName.split('_');
  for (let i = 0; i < tokens.length; i++) {
    if (/^\d+$/.test(tokens[i])) {
      const byId = roster.filter(r => r.id === tokens[i]);
      if (byId.length === 1) return byId[0];
    }
  }
  const k = letters_(tokens[0]);
  const byKey = roster.filter(r => r.key === k);
  if (byKey.length === 1) return byKey[0];

  if (ownerEmail) {
    const byEmail = roster.filter(r => r.email && r.email === ownerEmail);
    if (byEmail.length === 1) return byEmail[0];
  }
  const flat = letters_(fileName);
  const loose = roster.filter(r => (r.key.length > 4 && flat.indexOf(r.key) > -1) ||
                                   (r.keyFL.length > 4 && flat.indexOf(r.keyFL) > -1));
  return loose.length === 1 ? loose[0] : null;
}

function letters_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z]/g, '');
}

function studentKey_(s) {
  return s.id ? 'id:' + s.id : 'name:' + s.raw;
}

/* ============================ KEY TAB ============================= */

function ensureCodes_(run, roster) {
  const sh = ss_().getSheetByName(TAB_KEY);
  const data = sh.getDataRange().getValues().slice(1);
  const used = {};
  const map = {};
  data.forEach(r => {
    used[String(r[1])] = true;
    if (r[0] === run) map[r[3] !== '' ? 'id:' + String(r[3]) : 'name:' + r[2]] = String(r[1]);
  });
  const rows = [];
  roster.forEach(s => {
    const k = studentKey_(s);
    if (map[k]) return;
    const code = newCode_(used);
    map[k] = code;
    rows.push([run, code, s.raw, s.id, 0, '']);
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  return map;
}

/** Assignment names already on the Key tab, newest first. */
function pastAssignments_() {
  const sh = ss_().getSheetByName(TAB_KEY);
  const seen = {}, out = [];
  sh.getDataRange().getValues().slice(1).reverse().forEach(r => {
    const a = String(r[0] || '').trim();
    if (a && !seen[a]) { seen[a] = true; out.push(a); }
  });
  return out;
}

/** Finds a folder by name inside a parent, or makes it (if the old one was trashed). */
function folderByName_(parent, name) {
  const it = parent.getFoldersByName(name);
  while (it.hasNext()) { const fo = it.next(); if (!fo.isTrashed()) return fo; }
  return parent.createFolder(name);
}

function unmatchedCode_(run, fileName) {
  const sh = ss_().getSheetByName(TAB_KEY);
  const used = {};
  sh.getDataRange().getValues().slice(1).forEach(r => { used[String(r[1])] = true; });
  const code = newCode_(used);
  sh.appendRow([run, code, 'UNMATCHED: ' + fileName, '', 0, '']);
  return code;
}

function newCode_(used) {
  let code;
  do { code = 'S-' + (1000 + Math.floor(Math.random() * 9000)); } while (used[code]);
  used[code] = true;
  return code;
}

/** Adds one to the student's file count. Returns how many files they had BEFORE this one. */
function bumpKey_(run, code, outName, note) {
  const sh = ss_().getSheetByName(TAB_KEY);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === run && String(data[i][1]) === code) {
      const before = Number(data[i][4]) || 0;
      if (outName) sh.getRange(i + 1, 5).setValue(before + 1);
      if (note) sh.getRange(i + 1, 6).setValue((data[i][5] ? data[i][5] + ' | ' : '') + note);
      return before;
    }
  }
  return 0;
}

/** Version 20: counts tries per file (saved before the work starts, so a run Google stops still counts). */
const MAX_TRIES = 3;
function bumpTries_(id) {
  const p = PropertiesService.getDocumentProperties();
  const n = (Number(p.getProperty('TRY::' + id)) || 0) + 1;
  p.setProperty('TRY::' + id, String(n));
  return n;
}
function undoTry_(id) {
  const p = PropertiesService.getDocumentProperties();
  const n = (Number(p.getProperty('TRY::' + id)) || 0) - 1;
  if (n > 0) p.setProperty('TRY::' + id, String(n)); else p.deleteProperty('TRY::' + id);
}
function clearTries_(id) { try { PropertiesService.getDocumentProperties().deleteProperty('TRY::' + id); } catch (e) { /* ignore */ } }
const TOO_SLOW = 'This file stopped Google\'s 6-minute limit ' + MAX_TRIES + ' times (too big or too slow to clean), so it was set aside in Processed originals and NOT sent to Rowan. ' +
  'To grade it: open it, save or print it as a smaller PDF (or take the name off), and drop that in the Inbox. Original: ';

/** Note added when a student already had work saved for this assignment (a redo, or a second file). */
const ALREADY_HAD = 'This student already had a file for this assignment, so this may be a redo or an extra file.';

/* ============================ HELPERS ============================= */

/**
 * Unzips every zip in the Inbox, including a student's zip inside the Canvas zip.
 * Files from a student's zip keep the student's Canvas name tag
 * (smithjohn_123456_7890_1.png) so they still match that student.
 */
function unzipAll_(inbox, done) {
  const errors = [];
  const failed = {};
  for (let pass = 0; pass < 3; pass++) {
    const zips = [];
    const it = inbox.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if ((f.getMimeType() === MimeType.ZIP || /\.zip$/i.test(f.getName())) && !failed[f.getId()]) zips.push(f);
    }
    if (!zips.length) break;
    zips.forEach(z => {
      try {
        const zipName = z.getName();
        const tag = canvasTag_(zipName) || (zipName.replace(/\.zip$/i, '') + '__');
        const blob = z.getBlob().setContentType('application/zip');
        Utilities.unzip(blob).forEach(b => {
          const full = String(b.getName() || '');
          const base = full.split('/').pop();
          if (!base || base.charAt(0) === '.' || /__MACOSX/.test(full)) return;
          const ext = extOf_(base).toLowerCase();
          if (MIME_BY_EXT[ext]) b.setContentType(MIME_BY_EXT[ext]);
          // A file that already carries a Canvas tag (from the class zip) keeps its name
          const name = canvasTag_(base) ? base : tag + base;
          inbox.createFile(b.setName(name));
        });
        z.moveTo(done);
      } catch (e) {
        failed[z.getId()] = true;
        errors.push(z.getName() + ': ' + e.message);
      }
    });
  }
  return errors;
}

/** The "smithjohn_LATE_123456_7890123_" part Canvas puts at the start of file names. */
function canvasTag_(name) {
  const m = String(name).match(/^([a-z]+_(?:late_)?\d+_(?:\d+_)?)/i);
  return m ? m[1] : '';
}

/** Groups loose pictures that came from the same place: "Deck__3.png" -> "Deck", "slide 3.png" -> "slide". */
function groupStem_(name) {
  const stem = String(name).replace(/\.[^.]+$/, '');
  if (stem.indexOf('__') > -1) return stem.slice(0, stem.lastIndexOf('__'));
  return stem.replace(/[\s_\-()]*\d+[)\s]*$/, '');
}

function naturalCompare_(a, b) {
  const ax = String(a).toLowerCase().match(/\d+|\D+/g) || [];
  const bx = String(b).toLowerCase().match(/\d+|\D+/g) || [];
  for (let i = 0; i < Math.min(ax.length, bx.length); i++) {
    const x = ax[i], y = bx[i];
    if (x === y) continue;
    const nx = /^\d/.test(x), ny = /^\d/.test(y);
    if (nx && ny) return Number(x) - Number(y);
    return x < y ? -1 : 1;
  }
  return ax.length - bx.length;
}

/** Puts a student's pictures on slides in order (1, 2, 3 ... 10) and saves them as one PDF. */
function combineImages_(files, code, ctx) {
  files = files.slice().sort((a, b) => naturalCompare_(a.getName(), b.getName()));
  const pres = SlidesApp.create('tmp_' + Utilities.getUuid());
  const presFile = DriveApp.getFileById(pres.getId());
  try { presFile.moveTo(ctx.temp); } catch (e) { /* stays in My Drive until trashed */ }
  const skipped = [];
  try {
    const W = pres.getPageWidth(), H = pres.getPageHeight();
    const first = pres.getSlides()[0];
    first.getPageElements().forEach(el => el.remove());
    let used = 0;
    files.forEach(f => {
      let blob;
      try { blob = f.getBlob().getAs('image/png'); } catch (e) { blob = f.getBlob(); }
      try {
        const slide = used === 0 ? first : pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
        const img = slide.insertImage(blob);
        const s = Math.min(W / img.getWidth(), H / img.getHeight());
        img.setWidth(img.getWidth() * s).setHeight(img.getHeight() * s);
        img.setLeft((W - img.getWidth()) / 2).setTop((H - img.getHeight()) / 2);
        used++;
      } catch (e) {
        skipped.push(f.getName());
      }
    });
    if (!used) throw new Error('none of the pictures could be opened');
    pres.saveAndClose();
    const outName = nextOutName_(code, '.pdf', ctx.outFolder);
    ctx.outFolder.createFile(presFile.getAs(MimeType.PDF).setName(outName));
    return { outName: outName, skipped: skipped };
  } finally {
    try { presFile.setTrashed(true); } catch (e) { /* ignore */ }
  }
}

function nextOutName_(code, ext, folder) {
  let name = code + ext, n = 2;
  while (folder.getFilesByName(name).hasNext()) name = code + '-' + (n++) + ext;
  return name;
}

function extOf_(name) {
  const m = String(name || '').match(/\.([A-Za-z0-9]+)$/);
  return m ? m[1] : '';
}

function folders_() {
  const p = PropertiesService.getDocumentProperties();
  try {
    return {
      inbox: DriveApp.getFolderById(p.getProperty('INBOX')),
      ready: DriveApp.getFolderById(p.getProperty('READY')),
      done: DriveApp.getFolderById(p.getProperty('DONE')),
      temp: DriveApp.getFolderById(p.getProperty('TEMP'))
    };
  } catch (e) {
    throw new Error('Folders not found. Run Anonymizer > Set up first.');
  }
}

function ensureSheet_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8eef7');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, headers.length);
  }
  return sh;
}

function log_(run, code, result, note) {
  ss_().getSheetByName(TAB_LOG).appendRow([new Date(), run, code, result, note]);
}

/* END OF ANONYMIZER CODE */
