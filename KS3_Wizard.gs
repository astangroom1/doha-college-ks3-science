/**
 * ============================================================
 *  KS3 SCIENCE — PAPER 1 FOLLOW-UP EMAIL SCRIPT
 *  Year 9 | Doha College | Andy Stangroom
 *  v2.0 — Full scheduling, reschedule, cancel, send-now
 * ============================================================
 *
 *  FLOW:
 *   1. Teacher opens menu → KS3 Science → Paper 1
 *   2. Selects class from permitted list
 *   3. Chooses: Send Now OR Schedule (date/time picker)
 *   4. Queue Status dialog shows all queued/scheduled entries
 *      with Reschedule / Cancel Schedule / Send Now buttons
 *   5. 5-min trigger processes queue:
 *      - Send Now entries → immediate
 *      - Scheduled entries → only when scheduledFor time passed
 *   6. Teacher receives summary email on completion
 *
 *  CRASH GUARDS (7):
 *   1. Atomic queue ops via LockService
 *   2. State validation before every action
 *   3. Scheduled entries fire exactly once (marked before work)
 *   4. Stuck run auto-recovery (15 min)
 *   5. Per-student try/catch (one failure never kills the class)
 *   6. Empty sheet guards on all getRange calls
 *   7. Idempotent status check (never double-sends)
 *
 * ============================================================
 */


// ═══════════════════════════════════════════════════════════
// COLUMN / ROW CONSTANTS  —  verified against setup script v4
// ═══════════════════════════════════════════════════════════

// Year Controller row numbers (values in col B)
var YC_ROW = {
  academicYear:   4,
  yearGroup:      5,
  currentTerm:    6,
  ownerEmail:     10,
  hoksName:       11,
  hoksEmail:      12,
  signOffName:    13,
  p1MaxQ:         15,
  bannerP1:       19,
  bannerP2:       20,
  bannerModelAns: 21,
  bannerAchieve:  22,
  bannerChaseUp:  23,
  bannerTeacher:  24,
  bannerReports:  25,
  masterFolderId: 34,
};

// Classes tab columns (1-based)
var CL_COL = {
  academicYear: 1,
  classCode:    2,
  yearGroup:    3,
  leadName:     4,
  leadEmail:    5,
  leadFirst:    6,
  otherName:    7,
  otherEmail:   8,
  otherFirst:   9,
  lessonSplit:  10,
};

// Student Register columns (1-based)
var SR_COL = {
  studentId:   1,
  firstName:   2,
  lastName:    3,
  fullName:    4,
  email:       5,
  gender:      6,
  yearGroup:   7,
  classCode:   8,
  tutorGroup:  9,
  status:      10,
  fatherEmail: 12,
  motherEmail: 14,
};

// Paper 1 row/column structure
var P1_ROW = {
  answerKey:     1,
  aoTag:         2,
  disciplineTag: 3,
  topicTag:      4,
  bloomsTag:     5,
  qTypeTag:      6,
  header:        7,
  dataStart:     8,
};
var P1_COL = {
  studentId: 1,
  fullName:  2,
  classCode: 3,
  q1:        4,
  qLast:     43,  // Q40 = col AQ
  score:     44,  // col AR
  pct:       45,  // col AS
  status:    46,  // col AT
};

// Follow-Up Tasks columns (1-based)
var FU_COL = {
  taskId:       1,
  paperId:      2,
  taskType:     3,
  questionNum:  4,
  variant:      6,
  ao:           7,
  discipline:   9,
  topic:        10,
  questionText: 11,
  optionA:      12,
  optionB:      13,
  optionC:      14,
  optionD:      15,
  correctAns:   16,
  explanation:  17,
};

// Processing Log columns (1-based)
var PL_COL = {
  logId:        1,
  timestamp:    2,
  scriptName:   3,
  paperId:      4,
  academicYear: 5,
  term:         6,
  classCode:    7,
  triggerType:  8,
  status:       9,
  processed:    10,
  failed:       11,
  errors:       12,
  runBy:        13,
  notes:        14,
  // Per-student log columns (cols 15 onwards — only populated on student rows)
  studentName:  15,
  studentId:    16,
  emailSentTo:  17,
  score:        18,
  wrongQCount:  19,
  wrongQList:   20,
};

// Queue / script property keys
var QUEUE_KEY      = 'P1_QUEUE';
var PROCESSING_KEY = 'P1_CURRENT';
var MAX_QUEUE_SIZE = 6;
var STUCK_MINUTES  = 15;
var TZ             = 'Asia/Qatar';


// ═══════════════════════════════════════════════════════════
// MENU
// ═══════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('KS3 Science')
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Paper 1')
        .addItem('Send Follow-Up Emails — Live',    'showP1LiveDialog')
        .addItem('Send Follow-Up Emails — Test',    'showP1TestDialog')
        .addItem('Preview Emails',                  'showP1PreviewDialog')
        .addSeparator()
        .addItem('Queue Status & Management',       'showQueueManager')
    )
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Paper 2')
        .addItem('Send Follow-Up Emails — Live',    'showP2LiveDialog')
        .addItem('Send Follow-Up Emails — Test',    'showP2TestDialog')
        .addItem('Preview Emails',                  'showP2PreviewDialog')
        .addSeparator()
        .addItem('Queue Status & Management',       'showP2QueueManager')
    )
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('End of Term')
        .addItem('Archive & Clear Paper 1',         'archiveAndClearP1')
        .addItem('Archive & Clear Paper 2',         'archiveAndClearP2')
        .addSeparator()
        .addItem('Archive & Clear Both Papers',     'archiveAndClearBoth')
    )
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Admin')
        .addItem('Install / Refresh P1 Queue Trigger', 'installQueueTrigger')
        .addItem('Install / Refresh P2 Queue Trigger', 'installP2QueueTrigger')
        .addSeparator()
        .addItem('Reset Stuck Runs (P1)',               'resetStuckRuns')
        .addItem('Reset Stuck Runs (P2)',               'resetP2StuckRuns')
        .addSeparator()
        .addItem('Process P1 Queue Now',                'processP1Queue')
        .addItem('Process P2 Queue Now',                'processP2Queue')
        .addSeparator()
        .addItem('Show Admin Tabs',                     'showAdminTabs')
        .addItem('Hide Admin Tabs',                     'hideAdminTabs')
    )
    .addToUi();
}


// ═══════════════════════════════════════════════════════════
// ADMIN TAB VISIBILITY
// ═══════════════════════════════════════════════════════════

var ADMIN_TABS = [
  'Year Controller','Classes','Form Tutors','Student Register',
  'Assessment Bank','Results Database','Computed Profiles',
  'Processing Log','Archive P1','Archive P2',
  'Paper Config','Follow-Up Tasks',
];

function showAdminTabs() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var config = getConfig_(ss);
  if (!isAdminAccount_(Session.getActiveUser().getEmail(), config)) {
    SpreadsheetApp.getUi().alert('Only the script owner can show admin tabs.');
    return;
  }
  ADMIN_TABS.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) { sheet.showSheet(); }
  });
  SpreadsheetApp.getUi().alert('✅ Admin tabs are now visible.');
}

function hideAdminTabs() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var config = getConfig_(ss);
  if (!isAdminAccount_(Session.getActiveUser().getEmail(), config)) {
    SpreadsheetApp.getUi().alert('Only the script owner can hide admin tabs.');
    return;
  }
  ADMIN_TABS.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      try { sheet.hideSheet(); } catch(e) {}
      // Google won't hide the last visible sheet — silently skip
    }
  });
  SpreadsheetApp.getUi().alert('✅ Admin tabs hidden. Teachers now see Paper 1, Paper 2 and Wayground Import only.');
}




function archiveAndClearBoth() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var config = getConfig_(ss);
  var ui     = SpreadsheetApp.getUi();

  var confirm = ui.alert(
    'Archive & Clear Both Papers',
    'This will:\n\n' +
    '✅ Copy all Paper 1 student data → Archive P1\n' +
    '✅ Copy all Paper 2 student data → Archive P2\n' +
    '🗑️ Clear student data rows in both Paper 1 and Paper 2\n\n' +
    'Tags, answer key and column structure are preserved.\n\n' +
    'Current term: ' + config.currentTerm + ' · ' + config.academicYear + '\n\n' +
    'This cannot be undone. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) { return; }

  var p1Done = doArchivePaper_(ss, config, 'P1');
  var p2Done = doArchivePaper_(ss, config, 'P2');

  ui.alert(
    '✅ Archive Complete',
    'Paper 1: ' + p1Done + ' rows archived and cleared.\n' +
    'Paper 2: ' + p2Done + ' rows archived and cleared.\n\n' +
    'Ready for ' + config.academicYear + ' next term.',
    ui.ButtonSet.OK
  );
}

function archiveAndClearP1() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var config = getConfig_(ss);
  var ui     = SpreadsheetApp.getUi();

  var confirm = ui.alert(
    'Archive & Clear Paper 1',
    'This will copy all Paper 1 student data to Archive P1, then clear the data rows.\n\n' +
    'Tags and answer key are preserved.\n\n' +
    'Current term: ' + config.currentTerm + ' · ' + config.academicYear + '\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) { return; }

  var done = doArchivePaper_(ss, config, 'P1');
  ui.alert('✅ Done', 'Paper 1: ' + done + ' rows archived and cleared.', ui.ButtonSet.OK);
}

function archiveAndClearP2() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var config = getConfig_(ss);
  var ui     = SpreadsheetApp.getUi();

  var confirm = ui.alert(
    'Archive & Clear Paper 2',
    'This will copy all Paper 2 student data to Archive P2, then clear the data rows.\n\n' +
    'Tags, sub-question labels and max marks row are preserved.\n\n' +
    'Current term: ' + config.currentTerm + ' · ' + config.academicYear + '\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) { return; }

  var done = doArchivePaper_(ss, config, 'P2');
  ui.alert('✅ Done', 'Paper 2: ' + done + ' rows archived and cleared.', ui.ButtonSet.OK);
}

/**
 * Core archive logic. Copies data rows from Paper 1/2 to Archive P1/P2,
 * stamping the term and archive date, then clears the source data rows.
 * Headers, tags, answer key and max marks rows are never touched.
 * Returns the number of rows archived.
 */
function doArchivePaper_(ss, config, paper) {
  var srcName  = paper === 'P1' ? 'Paper 1'   : 'Paper 2';
  var dstName  = paper === 'P1' ? 'Archive P1' : 'Archive P2';
  var dataStart = 8;   // rows 1-7 are always structural

  var src = ss.getSheetByName(srcName);
  var dst = ss.getSheetByName(dstName);

  if (!src || !dst) {
    SpreadsheetApp.getUi().alert('Could not find ' + srcName + ' or ' + dstName + ' tab.');
    return 0;
  }

  var srcLastRow = src.getLastRow();
  if (srcLastRow < dataStart) { return 0; }  // nothing to archive

  var numDataRows = srcLastRow - dataStart + 1;
  var numCols     = src.getLastColumn();

  // Read all student data rows from source
  var dataRange = src.getRange(dataStart, 1, numDataRows, numCols);
  var data      = dataRange.getValues();

  // Filter to rows that actually have a student ID (skip blank rows)
  var toArchive = data.filter(function(row) {
    return safeId_(row[0]) !== '';
  });

  if (toArchive.length === 0) { return 0; }

  // Find first empty row in archive (after header row 2)
  var dstLastRow = Math.max(dst.getLastRow(), 2);
  var insertRow  = dstLastRow + 1;

  // Ensure archive has enough rows and columns
  var dstMaxRows = dst.getMaxRows();
  if (insertRow + toArchive.length > dstMaxRows) {
    dst.insertRowsAfter(dstMaxRows, toArchive.length + 10);
  }
  var dstMaxCols = dst.getMaxColumns();
  // Archive has extra cols: Student ID, Full Name, Class, Term, Date Archived = 5 prefix cols
  // then original Q/SQ cols
  var archiveCols = numCols + 2;  // +2 for Term and Date Archived prepended
  if (archiveCols > dstMaxCols) {
    dst.insertColumnsAfter(dstMaxCols, archiveCols - dstMaxCols);
  }

  // Build archive rows: prepend Term and Date Archived to each row
  var today     = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy');
  var term      = config.currentTerm;
  var archiveRows = toArchive.map(function(row) {
    // Archive tab structure: cols A-C = StudentID, FullName, Class (same as source)
    // then col D = Term, col E = Date Archived, then remaining cols
    return [row[0], row[1], row[2], term, today].concat(row.slice(3));
  });

  // Write to archive
  dst.getRange(insertRow, 1, archiveRows.length, archiveRows[0].length)
     .setValues(archiveRows);
  SpreadsheetApp.flush();

  // Clear the source data rows (preserve row height, formatting)
  // Clear all rows from dataStart to srcLastRow
  src.getRange(dataStart, 1, numDataRows, numCols).clearContent();
  SpreadsheetApp.flush();

  return toArchive.length;
}




function showP1LiveDialog()    { showClassDialog_('LIVE');    }
function showP1TestDialog()    { showClassDialog_('TEST');    }
function showP1PreviewDialog() { showClassDialog_('PREVIEW'); }

function showClassDialog_(mode) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var config  = getConfig_(ss);
  var user    = Session.getActiveUser().getEmail();
  var classes = getPermittedClasses_(ss, user, config);

  if (classes.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No classes found for your account (' + user + ').\n\n' +
      'Please ask Andy to check your entry in the Classes tab.'
    );
    return;
  }

  var modeLabel    = mode === 'LIVE'    ? 'Live — emails sent to students' :
                     mode === 'TEST'    ? 'Test — all emails go to Andy only' :
                                          'Preview — no emails sent';
  var headerColour = mode === 'LIVE'    ? '#1a237e' :
                     mode === 'TEST'    ? '#bf360c' : '#1b5e20';
  var title        = mode === 'LIVE'    ? 'Paper 1 — Send Follow-Up Emails (Live)' :
                     mode === 'TEST'    ? 'Paper 1 — Send Follow-Up Emails (Test)' :
                                          'Paper 1 — Preview Emails';

  var options = classes.map(function(c) {
    return '<option value="' + escAttr_(c.classCode) + '">' +
           escHtml_(c.classCode) + ' — ' + escHtml_(c.leadName) +
           (c.lessonSplit ? ' (' + c.lessonSplit + ')' : '') +
           '</option>';
  }).join('');

  var qatarToday = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    'body{font-family:Arial,sans-serif;font-size:13px;padding:0;margin:0;color:#212121;}' +
    '.hdr{background:' + headerColour + ';color:#fff;padding:16px 20px;}' +
    '.hdr h3{margin:0;font-size:15px;}.hdr p{margin:4px 0 0;font-size:11px;opacity:0.85;}' +
    '.body{padding:20px;}' +
    'label{font-weight:bold;display:block;margin-bottom:6px;font-size:12px;color:#555;}' +
    'select{width:100%;padding:8px;font-size:13px;border:1px solid #ccc;border-radius:4px;margin-bottom:16px;box-sizing:border-box;}' +
    '.btn-row{display:flex;gap:10px;}' +
    '.btn{flex:1;padding:10px;font-size:13px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;}' +
    '.btn-primary{background:' + headerColour + ';color:#fff;}' +
    '.btn-secondary{background:#e8eaf6;color:#3949ab;}' +
    '.btn-cancel{background:#f5f5f5;color:#666;border:1px solid #ddd;}' +
    '.btn:hover{opacity:0.88;}' +
    '#sched{display:none;margin-top:16px;border-top:1px solid #e0e0e0;padding-top:16px;}' +
    '.qp{display:flex;gap:8px;margin-bottom:12px;}' +
    '.qpb{padding:6px 10px;font-size:11px;background:#f5f5f5;color:#666;border:1px solid #ddd;border-radius:4px;cursor:pointer;}' +
    'input[type=date],input[type=time]{padding:8px;border:1px solid #ccc;border-radius:4px;font-size:13px;flex:1;box-sizing:border-box;}' +
    '.dt-row{display:flex;gap:10px;margin-bottom:10px;}' +
    '#msg{display:none;padding:12px;border-radius:6px;margin-top:12px;font-size:13px;}' +
    '.msg-ok{background:#e8f5e9;color:#2e7d32;}.msg-err{background:#ffebee;color:#c62828;}' +
    '</style></head><body>' +
    '<div class="hdr"><h3>' + escHtml_(title) + '</h3><p>' + escHtml_(modeLabel) + '</p></div>' +
    '<div class="body">' +
    '<label>Select your class:</label>' +
    '<select id="cc">' + options + '</select>' +
    (mode === 'PREVIEW'
      ? '<div class="btn-row"><button class="btn btn-primary" onclick="doPreview()">Open Preview</button><button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button></div>'
      : '<div class="btn-row"><button class="btn btn-primary" onclick="doNext()">Next →</button><button class="btn btn-secondary" onclick="showSched()">Schedule</button><button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button></div>'
    ) +
    '<div id="sched" style="display:none;margin-top:16px;border-top:1px solid #e0e0e0;padding-top:16px;">' +
    '<label>Schedule delivery date &amp; time:</label>' +
    '<div class="dt-row"><input type="date" id="sd"><input type="time" id="st"></div>' +
    '<div class="qp">' +
    '<button class="qpb" onclick="qp(0,16,0)">Today 4pm</button>' +
    '<button class="qpb" onclick="qp(1,8,0)">Tomorrow 8am</button>' +
    '<button class="qpb" onclick="qp(1,16,0)">Tomorrow 4pm</button>' +
    '</div>' +
    '<div class="btn-row"><button class="btn btn-primary" onclick="doSchedule()">Confirm Schedule</button><button class="btn btn-cancel" onclick="hideSched()">Back</button></div>' +
    '</div>' +
    '<div id="msg"></div></div>' +
    '<script>' +
    'var QT="' + qatarToday + '";var MODE="' + mode + '";' +
    'function cc(){return document.getElementById("cc").value;}' +
    'function showMsg(t,ok){var m=document.getElementById("msg");m.className="msg-"+(ok?"ok":"err");m.style.display="block";m.textContent=t;}' +
    'function doPreview(){google.script.run.withSuccessHandler(function(){google.script.host.close();}).withFailureHandler(function(e){showMsg(e.message,false);}).handlePreview(cc());}' +
    'function doNext(){' +
    '  showMsg("Loading students...",true);' +
    '  google.script.run.withSuccessHandler(function(r){showMsg(r,true);setTimeout(function(){google.script.host.close();},2000);}).withFailureHandler(function(e){showMsg(e.message,false);}).openStudentSelector(cc(),MODE);' +
    '}' +
    'function showSched(){document.getElementById("sched").style.display="block";}' +
    'function hideSched(){document.getElementById("sched").style.display="none";}' +
    'function qp(days,h,m){var p=QT.split("-"),d=new Date(+p[0],+p[1]-1,+p[2]);d.setDate(d.getDate()+days);var y=d.getFullYear(),mo=d.getMonth()+1,dy=d.getDate();document.getElementById("sd").value=y+"-"+(mo<10?"0"+mo:mo)+"-"+(dy<10?"0"+dy:dy);document.getElementById("st").value=(h<10?"0"+h:h)+":"+(m<10?"0"+m:m);}' +
    'function doSchedule(){var d=document.getElementById("sd").value,t=document.getElementById("st").value;if(!d||!t){showMsg("Please select a date and time.",false);return;}showMsg("Opening student list...",true);google.script.run.withSuccessHandler(function(r){showMsg(r,true);setTimeout(function(){google.script.host.close();},2000);}).withFailureHandler(function(e){showMsg(e.message,false);}).openStudentSelectorScheduled(cc(),MODE,d,t);}' +
    '</script></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(440).setHeight(mode === 'PREVIEW' ? 210 : 280),
    title
  );
}


// ═══════════════════════════════════════════════════════════
// STEP 2 — STUDENT SELECTION DIALOG
// ═══════════════════════════════════════════════════════════

// Called from the class dialog — opens student selector
function openStudentSelector(classCode, mode) {
  showStudentSelectorDialog_(classCode, mode, null, null);
  return 'Opening student list for ' + classCode + '...';
}

function openStudentSelectorScheduled(classCode, mode, dateStr, timeStr) {
  showStudentSelectorDialog_(classCode, mode, dateStr, timeStr);
  return 'Opening student list for ' + classCode + '...';
}

function showStudentSelectorDialog_(classCode, mode, dateStr, timeStr) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var config    = getConfig_(ss);
  var students  = getStudentData_(ss, classCode);
  var classInfo = getClassInfo_(ss, classCode, config.academicYear);

  if (students.length === 0) {
    SpreadsheetApp.getUi().alert('No active students found for ' + classCode + '.');
    return;
  }

  // Get P1 data to check status for each student
  var p1Sheet   = ss.getSheetByName('Paper 1');
  var p1LastRow = p1Sheet.getLastRow();
  var p1Map     = {};

  if (p1LastRow >= P1_ROW.dataStart) {
    var p1Data = p1Sheet.getRange(
      P1_ROW.dataStart, 1,
      p1LastRow - P1_ROW.dataStart + 1,
      P1_COL.status
    ).getValues();
    p1Data.forEach(function(row) {
      var sid  = safeId_(row[P1_COL.studentId - 1]);
      var name = String(row[P1_COL.fullName - 1]).trim().toLowerCase();
      if (sid)  { p1Map[sid]  = row; }
      if (name) { p1Map['name:' + name] = row; }  // name fallback key
    });
  }

  // Build student rows
  // - Has P1 data + not yet sent → selectable (pre-ticked)
  // - Already sent → shown greyed out, not selectable
  // - No P1 data → shown with warning, not selectable
  var studentRows = '';
  var selectableCount = 0;

  students.forEach(function(student) {
    // Try ID first, then fall back to name match
    var p1Row   = p1Map[safeId_(student.studentId)] ||
                  p1Map['name:' + student.fullName.toLowerCase()];
    var status  = p1Row ? String(p1Row[P1_COL.status - 1]).trim() : '';
    var hasData = !!p1Row;
    var alreadySent = (status === 'Emails Sent' || status === 'Complete');
    var score   = hasData ? p1Row[P1_COL.score - 1] : '';
    var pct     = hasData ? p1Row[P1_COL.pct - 1]   : '';
    var pctDisplay = (pct !== '' && pct !== null && pct !== undefined)
      ? Math.round(Number(pct)) + '%' : '—';

    // In TEST mode, always show as selectable regardless of sent status
    var selectable = hasData && (mode === 'TEST' || !alreadySent);

    if (selectable) { selectableCount++; }

    var rowBg    = alreadySent && mode !== 'TEST' ? '#f5f5f5' : '#fff';
    var nameCol  = alreadySent && mode !== 'TEST' ? '#9e9e9e' : '#212121';
    var badge    = alreadySent && mode !== 'TEST'
      ? '<span style="font-size:10px;background:#e8f5e9;color:#2e7d32;border-radius:3px;padding:1px 5px;margin-left:6px;">✓ Sent</span>'
      : !hasData
      ? '<span style="font-size:10px;background:#fff3e0;color:#e65100;border-radius:3px;padding:1px 5px;margin-left:6px;">No data</span>'
      : '';

    studentRows +=
      '<tr style="background:' + rowBg + ';border-bottom:1px solid #f0f0f0;">' +
      '<td style="padding:7px 8px;text-align:center;">' +
      (selectable
        ? '<input type="checkbox" id="s_' + escAttr_(student.studentId) + '" value="' + escAttr_(student.studentId) + '" checked style="width:16px;height:16px;cursor:pointer;">'
        : '<input type="checkbox" disabled style="width:16px;height:16px;opacity:0.3;">') +
      '</td>' +
      '<td style="padding:7px 8px;font-size:12px;color:' + nameCol + ';font-weight:' + (selectable ? 'bold' : 'normal') + ';">' +
      escHtml_(student.fullName) + badge +
      '</td>' +
      '<td style="padding:7px 8px;font-size:12px;text-align:center;color:' + nameCol + ';">' +
      (hasData ? escHtml_(String(pctDisplay)) : '—') +
      '</td>' +
      '</tr>';
  });

  var headerColour = mode === 'LIVE' ? '#1a237e' : mode === 'TEST' ? '#bf360c' : '#1b5e20';
  var schedLabel   = (dateStr && timeStr) ? ' — Scheduled: ' + dateStr + ' ' + timeStr : '';
  var modeNote     = mode === 'TEST'
    ? '<div style="background:#fff3e0;border:1px solid #ffcc02;border-radius:6px;padding:8px 12px;margin-bottom:12px;font-size:11px;color:#e65100;">TEST MODE — All emails will go to Andy only. Status will NOT be updated so you can re-run freely.</div>'
    : '';

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    'body{font-family:Arial,sans-serif;font-size:13px;padding:0;margin:0;color:#212121;}' +
    '.hdr{background:' + headerColour + ';color:#fff;padding:14px 20px;}' +
    '.hdr h3{margin:0;font-size:14px;}.hdr p{margin:3px 0 0;font-size:11px;opacity:0.85;}' +
    '.body{padding:16px;}' +
    '.tbl{width:100%;border-collapse:collapse;margin-bottom:12px;}' +
    '.tbl th{background:#f5f5f5;padding:7px 8px;font-size:11px;color:#666;text-align:left;font-weight:bold;border-bottom:2px solid #e0e0e0;}' +
    '.tbl th:first-child{text-align:center;width:36px;}' +
    '.tbl th:last-child{text-align:center;width:60px;}' +
    '.tbl tr:hover td{background:#fafafa;}' +
    '.btn-row{display:flex;gap:8px;}' +
    '.btn{flex:1;padding:9px;font-size:13px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;}' +
    '.btn-primary{background:' + headerColour + ';color:#fff;}' +
    '.btn-sec{background:#e8eaf6;color:#3949ab;}' +
    '.btn-cancel{background:#f5f5f5;color:#666;border:1px solid #ddd;}' +
    '.btn:hover{opacity:0.88;}' +
    '.sel-row{display:flex;gap:8px;margin-bottom:10px;font-size:11px;align-items:center;}' +
    '.sel-btn{background:none;border:1px solid #ccc;border-radius:3px;padding:3px 8px;cursor:pointer;font-size:11px;color:#555;}' +
    '#msg{display:none;padding:10px;border-radius:6px;margin-top:10px;font-size:12px;}' +
    '.msg-ok{background:#e8f5e9;color:#2e7d32;}.msg-err{background:#ffebee;color:#c62828;}' +
    '</style></head><body>' +
    '<div class="hdr"><h3>📋 ' + escHtml_(classCode) + ' — Select Students</h3>' +
    '<p>' + escHtml_(mode) + escHtml_(schedLabel) + ' · ' + selectableCount + ' student' + (selectableCount !== 1 ? 's' : '') + ' ready to send</p></div>' +
    '<div class="body">' +
    modeNote +
    '<div class="sel-row">' +
    '<button class="sel-btn" onclick="selAll(true)">Select All</button>' +
    '<button class="sel-btn" onclick="selAll(false)">Deselect All</button>' +
    '<span style="color:#9e9e9e;margin-left:4px;">Only students with Paper 1 data can be selected.</span>' +
    '</div>' +
    '<div style="max-height:280px;overflow-y:auto;border:1px solid #e0e0e0;border-radius:6px;">' +
    '<table class="tbl"><thead><tr>' +
    '<th>✓</th><th>Student Name</th><th>Score %</th>' +
    '</tr></thead><tbody>' + studentRows + '</tbody></table>' +
    '</div>' +
    '<div id="msg"></div>' +
    '<div class="btn-row" style="margin-top:12px;">' +
    (dateStr && timeStr
      ? '<button class="btn btn-primary" onclick="doSend()">Confirm Schedule</button>'
      : '<button class="btn btn-primary" onclick="doSend()">Send Now</button>') +
    '<button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>' +
    '</div>' +
    '</div>' +
    '<script>' +
    'var MODE="' + mode + '";var CC="' + escAttr_(classCode) + '";' +
    'var DS="' + (dateStr || '') + '";var TS="' + (timeStr || '') + '";' +
    'function selAll(v){var boxes=document.querySelectorAll("input[type=checkbox]:not([disabled])");boxes.forEach(function(b){b.checked=v;});}' +
    'function getSelected(){var boxes=document.querySelectorAll("input[type=checkbox]:not([disabled]):checked");return Array.from(boxes).map(function(b){return b.value;});}' +
    'function showMsg(t,ok){var m=document.getElementById("msg");m.className="msg-"+(ok?"ok":"err");m.style.display="block";m.textContent=t;}' +
    'function doSend(){' +
    '  var sel=getSelected();' +
    '  if(sel.length===0){showMsg("Please select at least one student.",false);return;}' +
    '  showMsg("Adding to queue for "+sel.length+" student(s)...",true);' +
    '  google.script.run' +
    '  .withSuccessHandler(function(r){showMsg(r,true);setTimeout(function(){google.script.host.close();},3000);})' +
    '  .withFailureHandler(function(e){showMsg(e.message,false);})' +
    '  .handleStudentSelection(CC,MODE,sel,DS,TS);' +
    '}' +
    '</script></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(480).setHeight(520),
    'Select Students — ' + classCode
  );
}


// ═══════════════════════════════════════════════════════════
// DIALOG CALLBACKS (called via google.script.run)
// ═══════════════════════════════════════════════════════════

function handlePreview(classCode) {
  showPreviewSidebar_(classCode);
}

function handleSendNow(classCode, mode) {
  return addToQueue_(classCode, mode, null);
}

function handleSchedule(classCode, mode, dateStr, timeStr) {
  var parts     = dateStr.split('-');
  var timeParts = timeStr.split(':');
  var schedDateStr = parts[0] + '-' + pad_(parseInt(parts[1])) + '-' + pad_(parseInt(parts[2])) +
                     'T' + pad_(parseInt(timeParts[0])) + ':' + pad_(parseInt(timeParts[1])) + ':00+03:00';
  var scheduledFor = new Date(schedDateStr).getTime();

  if (isNaN(scheduledFor) || scheduledFor <= Date.now()) {
    throw new Error('Please choose a date and time in the future.');
  }

  var queueResult = addToQueue_(classCode, mode, scheduledFor);
  var isError = (queueResult.indexOf('currently being processed') !== -1) ||
                (queueResult.indexOf('already in the queue') !== -1) ||
                (queueResult.indexOf('queue is full') !== -1) ||
                (queueResult.indexOf('System is busy') !== -1);
  if (isError) { throw new Error(queueResult); }

  var formattedTime = Utilities.formatDate(new Date(scheduledFor), TZ, 'EEE d MMM \'at\' HH:mm');
  return classCode + ' scheduled for ' + formattedTime + '. You\'ll receive a summary email when sent.';
}

/**
 * Called from student selection dialog.
 * selectedIds = array of student IDs to send to.
 * dateStr/timeStr = null for Send Now, set for scheduled.
 */
function handleStudentSelection(classCode, mode, selectedIds, dateStr, timeStr) {
  if (!selectedIds || selectedIds.length === 0) {
    throw new Error('No students selected.');
  }

  var scheduledFor = null;
  if (dateStr && timeStr) {
    var parts    = dateStr.split('-');
    var tParts   = timeStr.split(':');
    var isoStr   = parts[0] + '-' + pad_(parseInt(parts[1])) + '-' + pad_(parseInt(parts[2])) +
                   'T' + pad_(parseInt(tParts[0])) + ':' + pad_(parseInt(tParts[1])) + ':00+03:00';
    scheduledFor = new Date(isoStr).getTime();
    if (isNaN(scheduledFor) || scheduledFor <= Date.now()) {
      throw new Error('Please choose a date and time in the future.');
    }
  }

  var result = addToQueue_(classCode, mode, scheduledFor, selectedIds);
  var isError = (result.indexOf('currently being processed') !== -1) ||
                (result.indexOf('already in the queue') !== -1) ||
                (result.indexOf('queue is full') !== -1) ||
                (result.indexOf('System is busy') !== -1);
  if (isError) { throw new Error(result); }

  if (scheduledFor) {
    return classCode + ' (' + selectedIds.length + ' students) scheduled for ' +
           Utilities.formatDate(new Date(scheduledFor), TZ, 'EEE d MMM \'at\' HH:mm') +
           '. You\'ll receive a summary email when sent.';
  }
  return result;
}


// ═══════════════════════════════════════════════════════════
// QUEUE MANAGEMENT
// ═══════════════════════════════════════════════════════════

/**
 * Add a class to the queue.
 * scheduledFor = null  → Send Now (processed at next trigger cycle)
 * scheduledFor = ms timestamp → Scheduled (processed when time arrives)
 *
 * CRASH GUARD 1: Entire operation inside LockService
 */
function addToQueue_(classCode, mode, scheduledFor, selectedIds) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    return 'System is busy — please try again in a moment.';
  }

  try {
    var queue   = getQueue_();
    var current = getCurrentProcessing_();

    // CRASH GUARD 2: state validation
    if (current && current.classCode === classCode) {
      return classCode + ' is currently being processed. You\'ll receive a summary email when complete.';
    }
    var existing = queue.filter(function(q){ return q.classCode === classCode; });
    if (existing.length > 0) {
      return classCode + ' is already in the queue. Use Queue Status to reschedule or cancel it.';
    }
    if (queue.length >= MAX_QUEUE_SIZE) {
      return 'The queue is full (' + MAX_QUEUE_SIZE + ' classes). Please try again in a few minutes.';
    }

    var entry = {
      id:           Utilities.getUuid(),
      classCode:    classCode,
      mode:         mode,
      teacherEmail: Session.getActiveUser().getEmail(),
      queuedAt:     Date.now(),
      scheduledFor: scheduledFor || null,
      selectedIds:  selectedIds || null,  // null = send to all eligible students
    };

    queue.push(entry);
    saveQueue_(queue);

    var position = queue.filter(function(q){ return !q.scheduledFor; }).length;
    var studentNote = selectedIds ? ' (' + selectedIds.length + ' student' + (selectedIds.length !== 1 ? 's' : '') + ' selected)' : '';

    if (scheduledFor) {
      return classCode + studentNote + ' added to the queue.';
    } else {
      var pos = position === 1 ? '1st' : position === 2 ? '2nd' :
                position === 3 ? '3rd' : position + 'th';
      var waitMsg = current ? ' Another class is currently being processed.' : '';
      return classCode + studentNote + ' added to the queue (' + pos + ' in line).' + waitMsg +
             '\n\nYour class will be processed automatically — you don\'t need to do anything.' +
             '\nYou\'ll receive a summary email when complete.';
    }

  } finally {
    lock.releaseLock();
  }
}

// ── Queue reschedule ──────────────────────────────────────
// Called from Queue Manager dialog via google.script.run
function rescheduleQueueEntry(entryId, dateStr, timeStr) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { throw new Error('System busy — try again.'); }

  try {
    var queue = getQueue_();
    var idx   = indexOfEntry_(queue, entryId);

    if (idx === -1) {
      throw new Error('This entry is no longer in the queue — it may have already been processed.');
    }

    var parts     = dateStr.split('-');
    var timeParts = timeStr.split(':');
    var schedDateStr = parts[0] + '-' + pad_(parseInt(parts[1])) + '-' + pad_(parseInt(parts[2])) +
                       'T' + pad_(parseInt(timeParts[0])) + ':' + pad_(parseInt(timeParts[1])) + ':00+03:00';
    var scheduledFor = new Date(schedDateStr).getTime();

    if (isNaN(scheduledFor) || scheduledFor <= Date.now()) {
      throw new Error('Please choose a date and time in the future.');
    }

    // CRASH GUARD 1+2: update atomically inside lock
    queue[idx].scheduledFor = scheduledFor;
    saveQueue_(queue);

    return queue[idx].classCode + ' rescheduled for ' +
           Utilities.formatDate(new Date(scheduledFor), TZ, 'EEE d MMM \'at\' HH:mm') + '.';

  } finally {
    lock.releaseLock();
  }
}

// ── Cancel queue entry ────────────────────────────────────
function cancelQueueEntry(entryId) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { throw new Error('System busy — try again.'); }

  try {
    var queue = getQueue_();
    var idx   = indexOfEntry_(queue, entryId);

    if (idx === -1) {
      throw new Error('This entry is no longer in the queue — it may have already been processed.');
    }

    var cc = queue[idx].classCode;
    queue.splice(idx, 1);
    saveQueue_(queue);
    return cc + ' has been removed from the queue.';

  } finally {
    lock.releaseLock();
  }
}

// ── Convert scheduled → Send Now ─────────────────────────
function sendNowFromQueue(entryId) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { throw new Error('System busy — try again.'); }

  try {
    var queue = getQueue_();
    var idx   = indexOfEntry_(queue, entryId);

    if (idx === -1) {
      throw new Error('This entry is no longer in the queue — it may have already been processed.');
    }

    // CRASH GUARD 2: check it's not already processing
    var current = getCurrentProcessing_();
    var cc      = queue[idx].classCode;
    if (current && current.classCode === cc) {
      throw new Error(cc + ' is currently being processed. You\'ll receive a summary email shortly.');
    }

    var entry = queue.splice(idx, 1)[0];
    entry.scheduledFor = null;   // convert from scheduled to Send Now

    // FIX BUG 8: insert after the last existing Send Now entry (fair queue position),
    // not at position 0 which would jump ahead of everyone already waiting
    var insertPos = 0;
    for (var j = 0; j < queue.length; j++) {
      if (!queue[j].scheduledFor) { insertPos = j + 1; }
    }
    queue.splice(insertPos, 0, entry);
    saveQueue_(queue);

    return cc + ' moved to Send Now — will be processed at the next queue check (within 5 minutes).';

  } finally {
    lock.releaseLock();
  }
}

// ── Raw queue accessors ───────────────────────────────────
function getQueue_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(QUEUE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch(e) {
    // Corrupt JSON — log and return empty to prevent queue processor crash
    Logger.log('KS3 ERROR: getQueue_ JSON.parse failed: ' + e.toString());
    return [];
  }
}

function saveQueue_(queue) {
  PropertiesService.getScriptProperties().setProperty(QUEUE_KEY, JSON.stringify(queue));
}

function getCurrentProcessing_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(PROCESSING_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch(e) {
    // Corrupt JSON — clear it so processing can resume
    Logger.log('KS3 ERROR: getCurrentProcessing_ JSON.parse failed: ' + e.toString());
    PropertiesService.getScriptProperties().deleteProperty(PROCESSING_KEY);
    return null;
  }
}

function setCurrentProcessing_(entry) {
  PropertiesService.getScriptProperties().setProperty(PROCESSING_KEY, JSON.stringify(entry));
}

function clearCurrentProcessing_() {
  PropertiesService.getScriptProperties().deleteProperty(PROCESSING_KEY);
}

function indexOfEntry_(queue, entryId) {
  for (var i = 0; i < queue.length; i++) {
    if (queue[i].id === entryId) { return i; }
  }
  return -1;
}


// ═══════════════════════════════════════════════════════════
// QUEUE STATUS & MANAGEMENT DIALOG
// ═══════════════════════════════════════════════════════════

function showQueueManager() {
  var queue   = getQueue_();
  var current = getCurrentProcessing_();
  var now     = Date.now();

  // Build current processing section
  var currentHtml = '';
  if (current) {
    var elapsed = Math.round((now - current.startedAt) / 60000);
    currentHtml =
      '<div class="section">' +
      '<div class="section-title processing-title">⚙️ Currently Processing</div>' +
      '<div class="entry processing-entry">' +
      '<span class="cc">' + escHtml_(current.classCode) + '</span>' +
      '<span class="meta">Started ' + elapsed + ' min ago · ' + escHtml_(current.mode) + '</span>' +
      '<span class="action-note">Processing — cannot be modified</span>' +
      '</div></div>';
  }

  // Build queue sections
  var scheduledHtml = '';
  var sendNowHtml   = '';

  queue.forEach(function(entry) {
    var timeLabel = entry.scheduledFor
      ? Utilities.formatDate(new Date(entry.scheduledFor), TZ, 'EEE d MMM \'at\' HH:mm')
      : 'As soon as possible';
    var isScheduled = !!entry.scheduledFor;

    var entryHtml =
      '<div class="entry" id="entry-' + entry.id + '">' +
      '<div class="entry-top">' +
      '<span class="cc">' + escHtml_(entry.classCode) + '</span>' +
      '<span class="meta">' + (isScheduled ? '🕐 ' : '⚡ ') + escHtml_(timeLabel) + ' · ' + escHtml_(entry.mode) + '</span>' +
      '</div>' +
      '<div class="entry-actions">' +
      (isScheduled
        ? '<button class="btn btn-sec" onclick="openReschedule(\'' + entry.id + '\',\'' + escAttr_(entry.classCode) + '\')">Reschedule</button>' +
          '<button class="btn btn-danger" onclick="doCancel(\'' + entry.id + '\',\'' + escAttr_(entry.classCode) + '\')">Cancel</button>' +
          '<button class="btn btn-primary" onclick="doSendNow(\'' + entry.id + '\',\'' + escAttr_(entry.classCode) + '\')">Send Now</button>'
        : '<button class="btn btn-danger" onclick="doCancel(\'' + entry.id + '\',\'' + escAttr_(entry.classCode) + '\')">Cancel</button>'
      ) +
      '</div></div>';

    if (isScheduled) {
      scheduledHtml += entryHtml;
    } else {
      sendNowHtml += entryHtml;
    }
  });

  var emptyHtml = (!current && queue.length === 0)
    ? '<div style="padding:20px;text-align:center;color:#9e9e9e;font-size:13px;">Queue is empty — nothing scheduled or pending.</div>'
    : '';

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    'body{font-family:Arial,sans-serif;font-size:13px;padding:0;margin:0;color:#212121;background:#f5f5f5;}' +
    '.header{background:#1a237e;color:#fff;padding:16px 20px;}' +
    '.header h3{margin:0;font-size:15px;}' +
    '.body{padding:16px;}' +
    '.section{margin-bottom:16px;}' +
    '.section-title{font-size:11px;font-weight:bold;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;color:#5c6bc0;}' +
    '.processing-title{color:#e65100;}' +
    '.entry{background:#fff;border:1px solid #e0e0e0;border-radius:6px;padding:12px;margin-bottom:8px;}' +
    '.processing-entry{border-color:#ff8f00;background:#fff8e1;}' +
    '.entry-top{margin-bottom:8px;}' +
    '.cc{font-weight:bold;font-size:14px;color:#1a237e;margin-right:10px;}' +
    '.meta{font-size:12px;color:#666;}' +
    '.entry-actions{display:flex;gap:8px;flex-wrap:wrap;}' +
    '.action-note{font-size:12px;color:#e65100;font-style:italic;}' +
    '.btn{padding:6px 12px;border:none;border-radius:4px;cursor:pointer;font-size:12px;font-weight:bold;}' +
    '.btn-primary{background:#1a237e;color:#fff;}' +
    '.btn-sec{background:#e8eaf6;color:#3949ab;}' +
    '.btn-danger{background:#ffebee;color:#c62828;}' +
    '.btn:hover{opacity:0.85;}' +
    '.reschedule-panel{display:none;margin-top:10px;padding:10px;background:#f9fbe7;border-radius:6px;border:1px solid #dce775;}' +
    '.reschedule-panel input{padding:6px;border:1px solid #ccc;border-radius:4px;font-size:12px;margin-right:6px;}' +
    '#feedback{padding:10px;border-radius:6px;margin-top:12px;display:none;font-size:13px;}' +
    '.fb-ok{background:#e8f5e9;color:#2e7d32;}' +
    '.fb-err{background:#ffebee;color:#c62828;}' +
    '</style></head><body>' +
    '<div class="header"><h3>📋 Queue Status &amp; Management</h3></div>' +
    '<div class="body">' +
    emptyHtml +
    currentHtml +
    (scheduledHtml ? '<div class="section"><div class="section-title">🕐 Scheduled</div>' + scheduledHtml + '</div>' : '') +
    (sendNowHtml   ? '<div class="section"><div class="section-title">⚡ Send Now Queue</div>' + sendNowHtml + '</div>'   : '') +
    '<div id="feedback"></div>' +

    // Reschedule panel (shared, shown below target entry)
    '<div id="reschedulePanel" class="reschedule-panel">' +
    '<strong id="rescheduleLabel" style="display:block;margin-bottom:8px;font-size:12px;"></strong>' +
    '<input type="date" id="rDate"> <input type="time" id="rTime">' +
    '<div style="margin-top:8px;display:flex;gap:8px;">' +
    '<button class="btn btn-primary" onclick="confirmReschedule()">Confirm</button>' +
    '<button class="btn btn-sec" onclick="closeReschedule()">Cancel</button>' +
    '</div></div>' +

    '</div>' +
    '<script>' +
    'var pendingId = null;' +
    'var pendingCc = null;' +
    'function fb(msg,ok){var el=document.getElementById("feedback");el.textContent=msg;el.className="fb-"+(ok?"ok":"err");el.style.display="block";}' +
    'function openReschedule(id,cc){' +
    '  pendingId=id;pendingCc=cc;' +
    '  document.getElementById("rescheduleLabel").textContent="Reschedule " + cc;' +
    '  var panel=document.getElementById("reschedulePanel");' +
    '  var entry=document.getElementById("entry-"+id);' +
    '  entry.appendChild(panel);panel.style.display="block";' +
    '}' +
    'function closeReschedule(){document.getElementById("reschedulePanel").style.display="none";}' +
    'function confirmReschedule(){' +
    '  var d=document.getElementById("rDate").value;' +
    '  var t=document.getElementById("rTime").value;' +
    '  if(!d||!t){fb("Please select a date and time.",false);return;}' +
    // FIX BUG 4: was calling non-existent getQueueManagerData server function
    // FIX BUG 5: was calling location.reload() which is blocked in modal iframe
    // Solution: show success, close panel, tell user to reopen Queue Status to see update
    '  google.script.run' +
    '  .withSuccessHandler(function(r){fb(r + " Close this dialog and reopen Queue Status to see the update.",true);closeReschedule();})'  +
    '  .withFailureHandler(function(e){fb(e.message,false);})'  +
    '  .rescheduleQueueEntry(pendingId,d,t);' +
    '}' +
    'function doCancel(id,cc){' +
    '  if(!confirm("Remove " + cc + " from the queue?")) return;' +
    '  google.script.run' +
    '  .withSuccessHandler(function(r){fb(r,true);removeEntry(id);})'  +
    '  .withFailureHandler(function(e){fb(e.message,false);})'  +
    '  .cancelQueueEntry(id);' +
    '}' +
    'function doSendNow(id,cc){' +
    '  if(!confirm("Move " + cc + " to Send Now? It will be processed within 5 minutes.")) return;' +
    // FIX BUG 5: location.reload() blocked in modal — close dialog instead
    '  google.script.run' +
    '  .withSuccessHandler(function(r){fb(r + " You can close this dialog.",true);})'  +
    '  .withFailureHandler(function(e){fb(e.message,false);})'  +
    '  .sendNowFromQueue(id);' +
    '}' +
    'function removeEntry(id){var el=document.getElementById("entry-"+id);if(el)el.remove();}' +
    '</script></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(500).setHeight(480),
    'Queue Status & Management'
  );
}


// ═══════════════════════════════════════════════════════════
// QUEUE PROCESSOR — runs every 5 minutes via time trigger
// ═══════════════════════════════════════════════════════════

function processP1Queue() {
  var lock = LockService.getScriptLock();

  // If another trigger cycle is already inside here, skip silently
  if (!lock.tryLock(3000)) { return; }

  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var now = Date.now();

    // CRASH GUARD 4: recover stuck runs before doing anything
    recoverStuckRuns_(ss);

    // CRASH GUARD 3: if something is still processing, don't start another
    if (getCurrentProcessing_()) {
      lock.releaseLock();
      return;
    }

    // Find the next entry that is ready to process
    // "Ready" = scheduledFor is null (Send Now) OR scheduledFor <= now
    var queue = getQueue_();
    var nextIdx = -1;
    for (var i = 0; i < queue.length; i++) {
      if (!queue[i].scheduledFor || queue[i].scheduledFor <= now) {
        nextIdx = i; break;
      }
    }

    if (nextIdx === -1) {
      lock.releaseLock();
      return;  // nothing ready yet
    }

    var next = queue.splice(nextIdx, 1)[0];
    saveQueue_(queue);

    // CRASH GUARD 3: mark as processing BEFORE releasing lock
    next.startedAt = Date.now();
    setCurrentProcessing_(next);

    lock.releaseLock();

    // Do the heavy work outside the lock
    try {
      processP1Class_(ss, next);
    } catch(e) {
      logError_(ss, next, e.toString());
      clearCurrentProcessing_();
    }

  } catch(e) {
    // FIX BUG 9: log outer catch errors so they're not silently swallowed
    try {
      Logger.log('KS3 processP1Queue outer error: ' + e.toString());
      var ss2 = SpreadsheetApp.getActiveSpreadsheet();
      logError_(ss2, { classCode: 'UNKNOWN', mode: 'TRIGGER', teacherEmail: 'system' }, e.toString());
    } catch(e3) {}
    try { lock.releaseLock(); } catch(e2) {}
  }
}


// ═══════════════════════════════════════════════════════════
// STUCK RUN RECOVERY
// ═══════════════════════════════════════════════════════════

function recoverStuckRuns_(ss) {
  var current = getCurrentProcessing_();
  if (!current) { return; }

  var elapsedMin = (Date.now() - current.startedAt) / 60000;
  if (elapsedMin <= STUCK_MINUTES) { return; }

  clearCurrentProcessing_();

  try {
    var config = getConfig_(ss);
    MailApp.sendEmail({
      to:      config.hoksEmail,
      subject: '[KS3 Science] Stuck run recovered — ' + current.classCode,
      body:    'A Paper 1 processing run for ' + current.classCode +
               ' was stuck for ' + Math.round(elapsedMin) + ' minutes and has been reset.\n\n' +
               'The class has NOT been re-queued automatically.\n' +
               'Please ask the teacher to re-submit from the KS3 Science menu.\n\n' +
               'Run details:\n' + JSON.stringify(current, null, 2),
    });
  } catch(e) {}
}

function resetStuckRuns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  recoverStuckRuns_(ss);
  SpreadsheetApp.getUi().alert('Stuck run check complete. Check your email if any runs were recovered.');
}


// ═══════════════════════════════════════════════════════════
// CORE CLASS PROCESSOR
// ═══════════════════════════════════════════════════════════

function processP1Class_(ss, queueEntry) {
  var classCode = queueEntry.classCode;
  var mode      = queueEntry.mode;

  var config      = getConfig_(ss);
  var classInfo   = getClassInfo_(ss, classCode, config.academicYear);
  var answerKey   = getAnswerKey_(ss, config.p1MaxQ);
  var tags        = getTagData_(ss, config.p1MaxQ);
  var tasks       = getFollowUpTasks_(ss);
  var allStudents = getStudentData_(ss, classCode);

  // Filter to selected students only if teacher made a specific selection
  var selectedIds = queueEntry.selectedIds || null;
  var students = selectedIds
    ? allStudents.filter(function(s) {
        return selectedIds.some(function(id) {
          return safeId_(id) === safeId_(s.studentId);
        });
      })
    : allStudents;

  // Validate before creating the log row (these throw before logRow exists — caught by outer catch)
  if (!classInfo) {
    throw new Error('Class ' + classCode + ' not found in Classes tab.');
  }
  if (students.length === 0) {
    throw new Error('No active students found for class ' + classCode + ' in Student Register.');
  }

  var p1Sheet   = ss.getSheetByName('Paper 1');
  var p1LastRow = p1Sheet.getLastRow();
  if (p1LastRow < P1_ROW.dataStart) {
    throw new Error('Paper 1 has no student data. Please enter answers before sending emails.');
  }

  // Create log row AFTER validation — guarantees if we create it we will also finalize it
  var logRow  = initProcessingLog_(ss, queueEntry, config);
  var processed = 0;
  var failed    = 0;
  var errors    = [];
  var results   = [];

  // FIX BUG 6: use try/finally to guarantee log finalization and clearCurrentProcessing
  try {
    var p1Data = p1Sheet.getRange(
      P1_ROW.dataStart, 1,
      p1LastRow - P1_ROW.dataStart + 1,
      P1_COL.status
    ).getValues();

    // Map studentId → row (with name fallback for ID format mismatches)
    var p1Map = {};
    p1Data.forEach(function(row) {
      var sid  = safeId_(row[P1_COL.studentId - 1]);
      var name = String(row[P1_COL.fullName - 1]).trim().toLowerCase();
      if (sid)  { p1Map[sid]  = row; }
      if (name) { p1Map['name:' + name] = row; }
    });

  // FIX ISSUE 6: build status map ONCE before processing students
  // instead of reading the full ID column for every student
  var p1StatusMap    = buildP1StatusMap_(ss);
  var statusUpdates  = [];
  var studentLogRows = [];   // per-student audit trail

  // CRASH GUARD 5: per-student try/catch
  students.forEach(function(student) {
    try {
      // Try ID match first, fall back to name match
      var p1Row = p1Map[safeId_(student.studentId)] ||
                  p1Map['name:' + student.fullName.toLowerCase()];

      if (!p1Row) {
        results.push({ student: student, status: 'skipped', reason: 'No P1 data entered' });
        studentLogRows.push({ student: student, status: 'SKIPPED', reason: 'No P1 data entered', score: '', wrongQs: [] });
        return;
      }

      // CRASH GUARD 7: idempotent — never double-send in LIVE mode
      // TEST mode bypasses this so tests can be re-run freely on same data
      var currentStatus = String(p1Row[P1_COL.status - 1]).trim();
      if (mode !== 'TEST' && (currentStatus === 'Emails Sent' || currentStatus === 'Complete')) {
        results.push({ student: student, status: 'skipped', reason: 'Already processed' });
        studentLogRows.push({ student: student, status: 'SKIPPED', reason: 'Already processed', score: '', wrongQs: [] });
        return;
      }

      var score   = p1Row[P1_COL.score - 1];
      var pct     = p1Row[P1_COL.pct - 1];
      var wrongQs = identifyWrongQuestions_(p1Row, answerKey, tags, config.p1MaxQ);

      var emailHtml = buildP1StudentEmail_(student, wrongQs, tasks, score, pct, classInfo, config, mode);
      sendStudentEmail_(student, emailHtml, classInfo, config, mode, 'Paper 1 Follow-Up');

      // Queue status update — only in LIVE mode so TEST can be re-run freely
      if (mode !== 'TEST') {
        statusUpdates.push({ studentId: student.studentId, newStatus: 'Emails Sent' });
      }

      // Record per-student log entry
      var emailSentTo = mode === 'TEST' ? config.hoksEmail + ' (TEST — redirected from ' + student.email + ')' : student.email;
      studentLogRows.push({
        student:    student,
        status:     'SENT',
        reason:     '',
        score:      score + '/' + config.p1MaxQ + ' (' + Math.round(parseFloat(pct)||0) + '%)',
        emailTo:    emailSentTo,
        wrongQs:    wrongQs.map(function(q){ return 'Q' + q.number; }),
      });

      processed++;
      results.push({ student: student, status: 'sent', score: score, pct: pct, wrongCount: wrongQs.length });

    } catch(e) {
      failed++;
      errors.push(student.fullName + ': ' + e.toString());
      results.push({ student: student, status: 'failed', reason: e.toString() });
      studentLogRows.push({ student: student, status: 'FAILED', reason: e.toString(), score: '', wrongQs: [] });
    }
  });

  // Write all status updates and student log rows
  flushP1StatusUpdates_(ss, statusUpdates, p1StatusMap);
  writeStudentLogRows_(ss, logRow, studentLogRows, config);

  // Send teacher summary
  sendTeacherSummary_(classInfo, results, config, mode, 'Paper 1');

  } catch(e) {
    // Unexpected error during processing — record it
    failed++;
    errors.push('Unexpected error: ' + e.toString());
  } finally {
    // FIX BUG 6: ALWAYS finalize log and clear processing, no matter what happened
    finaliseProcessingLog_(ss, logRow, processed, failed, errors);
    clearCurrentProcessing_();
  }
}


// ═══════════════════════════════════════════════════════════
// DATA READERS
// ═══════════════════════════════════════════════════════════

function getConfig_(ss) {
  var sheet = ss.getSheetByName('Year Controller');
  var data  = sheet.getRange(1, 2, 40, 1).getValues();

  function val(row) {
    var v = data[row - 1][0];
    return (v !== null && v !== undefined) ? String(v).trim() : '';
  }

  return {
    academicYear:   val(YC_ROW.academicYear),
    currentTerm:    val(YC_ROW.currentTerm),
    ownerEmail:     val(YC_ROW.ownerEmail),
    hoksName:       val(YC_ROW.hoksName),
    hoksEmail:      val(YC_ROW.hoksEmail),
    signOffName:    val(YC_ROW.signOffName),
    p1MaxQ:         parseInt(val(YC_ROW.p1MaxQ)) || 40,
    bannerP1:       extractFileId_(val(YC_ROW.bannerP1)),
    bannerModelAns: extractFileId_(val(YC_ROW.bannerModelAns)),
    bannerChaseUp:  extractFileId_(val(YC_ROW.bannerChaseUp)),
    bannerTeacher:  extractFileId_(val(YC_ROW.bannerTeacher)),
    masterFolderId: val(YC_ROW.masterFolderId),
  };
}

function getClassInfo_(ss, classCode, academicYear) {
  var sheet = ss.getSheetByName('Classes');
  var last  = sheet.getLastRow();
  if (last < 3) { return null; }
  var data  = sheet.getRange(3, 1, last - 2, 17).getValues();

  for (var i = 0; i < data.length; i++) {
    var rowAY = String(data[i][CL_COL.academicYear - 1]).trim();
    var rowCC = String(data[i][CL_COL.classCode - 1]).trim();
    // FIX ISSUE 5: match by academic year (or accept blank AY for backward compat)
    if (rowCC === classCode && (!rowAY || !academicYear || rowAY === academicYear)) {
      return {
        classCode:   rowCC,
        leadName:    String(data[i][CL_COL.leadName - 1]).trim(),
        leadEmail:   String(data[i][CL_COL.leadEmail - 1]).trim(),
        leadFirst:   String(data[i][CL_COL.leadFirst - 1]).trim(),
        otherName:   String(data[i][CL_COL.otherName - 1]).trim(),
        otherEmail:  String(data[i][CL_COL.otherEmail - 1]).trim(),
        otherFirst:  String(data[i][CL_COL.otherFirst - 1]).trim(),
        lessonSplit: String(data[i][CL_COL.lessonSplit - 1]).trim(),
      };
    }
  }
  return null;
}

function getStudentData_(ss, classCode) {
  var sheet = ss.getSheetByName('Student Register');
  var last  = sheet.getLastRow();
  if (last < 3) { return []; }
  var data  = sheet.getRange(3, 1, last - 2, 14).getValues();
  var out   = [];

  data.forEach(function(row) {
    var sid    = safeId_(row[SR_COL.studentId - 1]);
    var cc     = String(row[SR_COL.classCode - 1]).trim();
    var status = String(row[SR_COL.status - 1]).trim();
    if (!sid || cc !== classCode || status !== 'Active') { return; }

    out.push({
      studentId:   sid,
      firstName:   String(row[SR_COL.firstName - 1]).trim(),
      lastName:    String(row[SR_COL.lastName - 1]).trim(),
      fullName:    (String(row[SR_COL.fullName - 1]).trim() ||
                   String(row[SR_COL.firstName - 1]).trim() + ' ' + String(row[SR_COL.lastName - 1]).trim()),
      email:       String(row[SR_COL.email - 1]).trim(),
      fatherEmail: String(row[SR_COL.fatherEmail - 1]).trim(),
      motherEmail: String(row[SR_COL.motherEmail - 1]).trim(),
    });
  });
  return out;
}

// FIX BUG 7: accept maxQ parameter instead of hardcoding 40
// FIX BUG 10: single getRange batch call for all tag rows
function getAnswerKey_(ss, maxQ) {
  maxQ = maxQ || 40;
  var sheet  = ss.getSheetByName('Paper 1');
  var keyRow = sheet.getRange(P1_ROW.answerKey, P1_COL.q1, 1, maxQ).getValues()[0];
  var key    = {};
  for (var i = 0; i < keyRow.length; i++) {
    var ans = String(keyRow[i]).trim().toUpperCase();
    if (ans) { key[i + 1] = ans; }
  }
  return key;
}

// FIX BUG 7: accept maxQ; FIX BUG 10: single batch API call covers all 5 tag rows
function getTagData_(ss, maxQ) {
  maxQ = maxQ || 40;
  var sheet  = ss.getSheetByName('Paper 1');
  // Read all 5 tag rows in ONE API call (rows 2–6, starting at tagRowFirst)
  var allTags = sheet.getRange(P1_ROW.aoTag, P1_COL.q1, 5, maxQ).getValues();
  return {
    ao:         allTags[0],   // row 2 — AO
    discipline: allTags[1],   // row 3 — Discipline
    topic:      allTags[2],   // row 4 — Topic
    blooms:     allTags[3],   // row 5 — Bloom's Level
    qType:      allTags[4],   // row 6 — Question Type
  };
}

function getFollowUpTasks_(ss) {
  var sheet = ss.getSheetByName('Follow-Up Tasks');
  var last  = sheet.getLastRow();
  if (last < 3) { return {}; }
  var data  = sheet.getRange(3, 1, last - 2, 17).getValues();
  var tasks = {};

  data.forEach(function(row) {
    var type = String(row[FU_COL.taskType - 1]).trim();
    var qNum = String(row[FU_COL.questionNum - 1]).trim();
    if (type !== 'P1-Variant' || !qNum) { return; }
    if (!tasks[qNum]) { tasks[qNum] = []; }
    tasks[qNum].push({
      questionText: String(row[FU_COL.questionText - 1]).trim(),
      optionA:      String(row[FU_COL.optionA - 1]).trim(),
      optionB:      String(row[FU_COL.optionB - 1]).trim(),
      optionC:      String(row[FU_COL.optionC - 1]).trim(),
      optionD:      String(row[FU_COL.optionD - 1]).trim(),
      correctAns:   String(row[FU_COL.correctAns - 1]).trim(),
      explanation:  String(row[FU_COL.explanation - 1]).trim(),
      topic:        String(row[FU_COL.topic - 1]).trim(),
      ao:           String(row[FU_COL.ao - 1]).trim(),
      discipline:   String(row[FU_COL.discipline - 1]).trim(),
    });
  });
  return tasks;
}

function getPermittedClasses_(ss, userEmail, config) {
  var isAdmin = isAdminAccount_(userEmail, config);
  var sheet   = ss.getSheetByName('Classes');
  var last    = sheet.getLastRow();
  if (last < 3) { return []; }
  var data    = sheet.getRange(3, 1, last - 2, 10).getValues();
  var out     = [];

  data.forEach(function(row) {
    var cc    = String(row[CL_COL.classCode - 1]).trim();
    var rowAY = String(row[CL_COL.academicYear - 1]).trim();
    var lead  = String(row[CL_COL.leadEmail - 1]).trim().toLowerCase();
    var other = String(row[CL_COL.otherEmail - 1]).trim().toLowerCase();
    if (!cc) { return; }
    // FIX ISSUE 5: only show classes for current academic year (or blank AY rows)
    if (rowAY && config.academicYear && rowAY !== config.academicYear) { return; }
    if (isAdmin || lead === userEmail.toLowerCase() || other === userEmail.toLowerCase()) {
      out.push({
        classCode:   cc,
        leadName:    String(row[CL_COL.leadName - 1]).trim(),
        lessonSplit: String(row[CL_COL.lessonSplit - 1]).trim(),
      });
    }
  });
  return out;
}

function isAdminAccount_(email, config) {
  return email.toLowerCase() === (config.ownerEmail || '').toLowerCase();
}


// ═══════════════════════════════════════════════════════════
// WRONG QUESTION IDENTIFIER
// ═══════════════════════════════════════════════════════════

function identifyWrongQuestions_(p1Row, answerKey, tags, maxQ) {
  var wrong = [];
  for (var q = 1; q <= maxQ; q++) {
    var correct = answerKey[q];
    if (!correct) { continue; }
    var student = String(p1Row[P1_COL.q1 + q - 2]).trim().toUpperCase();
    if (!student || student === '-') { continue; }
    if (student !== correct) {
      wrong.push({
        number:        q,
        studentAnswer: student,
        correctAnswer: correct,
        ao:            String(tags.ao[q - 1]).trim(),
        discipline:    String(tags.discipline[q - 1]).trim(),
        topic:         String(tags.topic[q - 1]).trim(),
        blooms:        String(tags.blooms[q - 1]).trim(),
      });
    }
  }
  return wrong;
}


// ═══════════════════════════════════════════════════════════
// EMAIL BUILDER — STUDENT (Redesigned v5)
// ═══════════════════════════════════════════════════════════

function buildP1StudentEmail_(student, wrongQs, tasks, score, pct, classInfo, config, mode) {
  var firstName    = student.firstName || student.fullName.split(' ')[0];
  var teacherName  = classInfo.leadFirst || classInfo.leadName;
  var teacherEmail = classInfo.leadEmail;
  var bannerUrl    = config.bannerP1
    ? 'https://drive.google.com/uc?export=view&id=' + config.bannerP1 : '';

  var pctNum       = parseFloat(pct) || 0;
  var pctDisplay   = (pct !== '' && pct !== null) ? Math.round(pctNum) + '%' : 'N/A';
  var scoreDisplay = (score !== '' && score !== null) ? score + ' / ' + config.p1MaxQ : 'N/A';
  var scoreColour  = pctNum >= 80 ? '#2e7d32' : pctNum >= 60 ? '#e65100' : '#b71c1c';

  var testBanner = mode === 'TEST'
    ? '<div style="background:#ff8f00;color:#fff;padding:10px;text-align:center;font-weight:bold;font-size:12px;">' +
      'TEST MODE — This email was redirected from ' + escHtml_(student.email) + '</div>'
    : '';

  // ── Score summary box ─────────────────────────────────────
  var scoreBox =
    '<table style="width:100%;border-collapse:collapse;background:#e8eaf6;border-radius:8px;">' +
    '<tr>' +
    '<td style="width:50%;padding:18px 20px;text-align:center;border-right:1px solid #c5cae9;">' +
    '<div style="font-size:10px;color:#5c6bc0;text-transform:uppercase;font-weight:bold;letter-spacing:0.8px;margin-bottom:6px;">Your Score</div>' +
    '<div style="font-size:32px;font-weight:bold;color:' + scoreColour + ';line-height:1;">' + escHtml_(pctDisplay) + '</div>' +
    '<div style="font-size:12px;color:#666;margin-top:4px;">' + escHtml_(scoreDisplay) + '</div>' +
    '</td>' +
    '<td style="width:50%;padding:18px 20px;text-align:center;">' +
    '<div style="font-size:10px;color:#5c6bc0;text-transform:uppercase;font-weight:bold;letter-spacing:0.8px;margin-bottom:6px;">Questions to Review</div>' +
    '<div style="font-size:32px;font-weight:bold;color:#1a237e;line-height:1;">' + wrongQs.length + '</div>' +
    '<div style="font-size:12px;color:#666;margin-top:4px;">out of ' + config.p1MaxQ + '</div>' +
    '</td>' +
    '</tr></table>';

  // ── Instructions section ──────────────────────────────────
  var instructions =
    '<div style="margin:24px 0 0;padding:18px 20px;background:#f8f9fe;border-radius:8px;border-left:4px solid #1a237e;">' +
    '<h3 style="margin:0 0 10px;font-size:14px;color:#1a237e;">Instructions</h3>' +
    '<p style="margin:0 0 8px;font-size:13px;color:#424242;line-height:1.6;">' +
    'Complete <strong>ALL</strong> tasks below in your <strong>exercise book</strong>. ' +
    'Bring your completed follow-up tasks to the due date lesson so your teacher can check your work. ' +
    'Model answers will be reviewed together at the start of that lesson.' +
    '</p>' +
    '</div>';

  // ── Why these tasks matter ────────────────────────────────
  var whyMatters =
    '<div style="margin:16px 0 0;padding:18px 20px;background:#f9fbe7;border-radius:8px;border-left:4px solid #c6e03a;">' +
    '<h3 style="margin:0 0 10px;font-size:14px;color:#33691e;">Why Do These Tasks Matter?</h3>' +
    '<p style="margin:0;font-size:13px;color:#424242;line-height:1.6;">' +
    'These follow-up tasks help you revisit and strengthen the ideas you found tricky — not just repeat them. ' +
    'This process, called <strong>retrieval practice</strong>, trains your brain to pull knowledge back out, making it stick much longer. ' +
    'It also builds <strong>metacognition</strong> — learning to think about how you learn. ' +
    'Doing this now boosts your confidence, deepens your understanding, and will help you improve in future assessments.' +
    '</p>' +
    '</div>';

  // ── AO explanation ────────────────────────────────────────
  var aoExplain =
    '<div style="margin:16px 0 0;padding:18px 20px;background:#f5f5f5;border-radius:8px;">' +
    '<h3 style="margin:0 0 10px;font-size:14px;color:#37474f;">Assessment Objectives (AO)</h3>' +
    '<p style="margin:0 0 6px;font-size:12px;color:#424242;"><strong style="color:#1a237e;">AO1</strong> — Do you know it? Facts, definitions and key scientific ideas.</p>' +
    '<p style="margin:0 0 6px;font-size:12px;color:#424242;"><strong style="color:#1a237e;">AO2</strong> — Can you use it? Apply your knowledge to unfamiliar questions.</p>' +
    '<p style="margin:0;font-size:12px;color:#424242;"><strong style="color:#1a237e;">AO3</strong> — Can you think like a scientist? Interpret data, analyse trends and evaluate methods.</p>' +
    '</div>';

  // ── Question summary ──────────────────────────────────────
  var questionSummary = '';
  if (wrongQs.length > 0) {
    // List of question numbers
    var qNums = wrongQs.map(function(q){ return 'Q' + q.number; }).join(', ');

    // Topic analysis — group wrong questions by topic
    var topicMap = {};
    wrongQs.forEach(function(q) {
      var t = q.topic || 'Untagged';
      if (!topicMap[t]) { topicMap[t] = 0; }
      topicMap[t]++;
    });
    var topicRows = Object.keys(topicMap).sort(function(a,b){ return topicMap[b] - topicMap[a]; }).map(function(t) {
      return '<tr><td style="padding:4px 0;font-size:12px;color:#424242;">' + escHtml_(t) + '</td>' +
             '<td style="padding:4px 0;font-size:12px;color:#1a237e;font-weight:bold;text-align:right;">' + topicMap[t] + ' question' + (topicMap[t] > 1 ? 's' : '') + ' wrong</td></tr>';
    }).join('');

    // Estimated time (approx 5 min per question)
    var estMins = wrongQs.length * 5;
    var estTime = estMins >= 60
      ? Math.floor(estMins/60) + ' hour' + (Math.floor(estMins/60) > 1 ? 's' : '') + (estMins % 60 > 0 ? ' ' + (estMins%60) + ' minutes' : '')
      : estMins + ' minutes';

    questionSummary =
      '<div style="margin:16px 0 0;padding:18px 20px;background:#fff;border:1px solid #e0e0e0;border-radius:8px;">' +
      '<h3 style="margin:0 0 12px;font-size:14px;color:#37474f;">Your Task Summary</h3>' +
      '<p style="margin:0 0 8px;font-size:12px;color:#424242;">The following questions are what you answered incorrectly in your test:</p>' +
      '<p style="margin:0 0 14px;font-size:13px;color:#1a237e;font-weight:bold;line-height:1.6;">' + escHtml_(qNums) + '</p>' +
      '<p style="margin:0 0 8px;font-size:12px;color:#37474f;font-weight:bold;">Topic Breakdown:</p>' +
      '<table style="width:100%;border-collapse:collapse;">' + topicRows + '</table>' +
      '<p style="margin:14px 0 0;font-size:12px;color:#666;font-style:italic;">Estimated time to complete all tasks: ' + escHtml_(estTime) + '</p>' +
      '</div>';
  }

  // ── Question task blocks ──────────────────────────────────
  var taskSection = '';
  if (wrongQs.length === 0) {
    taskSection = '<div style="margin:16px 0;padding:18px 20px;background:#e8f5e9;border-radius:8px;text-align:center;">' +
      '<p style="color:#2e7d32;font-weight:bold;font-size:15px;margin:0;">🎉 Excellent work — no incorrect answers to review!</p>' +
      '</div>';
  } else {
    taskSection =
      '<h2 style="font-size:16px;color:#1a237e;border-bottom:2px solid #e8eaf6;padding-bottom:8px;margin:24px 0 16px;">Your Follow-Up Tasks</h2>' +
      wrongQs.map(function(q) { return buildQuestionBlock_(q, pickTask_(tasks, q.number)); }).join('');
  }

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
    '<body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif;">' +
    testBanner +
    '<div style="max-width:640px;margin:0 auto;background:#fff;">' +

    // Banner image
    (bannerUrl ? '<img src="' + bannerUrl + '" width="640" style="display:block;width:100%;max-width:640px;" alt="KS3 Science">' : '') +

    // Subtle year/term strip replacing the ugly blue box
    '<div style="background:#f5f5f5;padding:8px 30px;border-bottom:1px solid #e0e0e0;">' +
    '<p style="margin:0;font-size:11px;color:#9e9e9e;letter-spacing:0.3px;">KS3 Science · Doha College · ' + escHtml_(config.academicYear) + ' · ' + escHtml_(config.currentTerm) + '</p>' +
    '</div>' +

    '<div style="padding:24px 30px;">' +
    '<p style="font-size:15px;color:#212121;margin:0 0 4px;">Dear ' + escHtml_(firstName) + ',</p>' +
    '<p style="font-size:13px;color:#666;margin:0 0 20px;">Your Paper 1 results are below. Please read all sections carefully.</p>' +

    // Score box (fixed centering using table not flex)
    scoreBox +
    instructions +
    whyMatters +
    aoExplain +
    questionSummary +
    taskSection +

    // Sign off
    '<div style="margin-top:30px;padding-top:20px;border-top:1px solid #e8eaf6;">' +
    '<p style="font-size:13px;color:#424242;">If you have any questions about these tasks, please speak to me in class or reply to this email.</p>' +
    '<p style="font-size:13px;color:#424242;margin-bottom:4px;">Best wishes,</p>' +
    '<p style="font-size:14px;color:#1a237e;font-weight:bold;margin:0;">' + escHtml_(teacherName) + '</p>' +
    '<p style="font-size:11px;color:#9e9e9e;margin:2px 0 0;">KS3 Science · Doha College · <a href="mailto:' + escHtml_(teacherEmail) + '" style="color:#9e9e9e;">' + escHtml_(teacherEmail) + '</a></p>' +
    '</div></div>' +

    // Footer
    '<div style="background:#e8eaf6;padding:10px 30px;text-align:center;">' +
    '<p style="font-size:10px;color:#9e9e9e;margin:0;">Automated message from KS3 Science · Doha College. To reply, use your teacher\'s email address above.</p>' +
    '</div>' +
    '</div></body></html>';
}

function buildQuestionBlock_(q, task) {
  var discColour = q.discipline === 'Biology'   ? '#2e7d32' :
                   q.discipline === 'Chemistry'  ? '#1565c0' :
                   q.discipline === 'Physics'    ? '#6a1b9a' : '#37474f';

  var bloomsColour = '#78909c';  // neutral grey for Bloom's

  // Task block — NO answers shown to maintain test integrity
  var taskHtml = '';
  if (task && task.questionText) {
    taskHtml =
      '<div style="background:#f9fbe7;border-left:4px solid #c6e03a;padding:14px 16px;margin-top:10px;border-radius:0 6px 6px 0;">' +
      '<div style="font-size:10px;color:#558b2f;font-weight:bold;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;">Practice Question</div>' +
      '<p style="font-size:13px;color:#212121;margin:0 0 10px;line-height:1.5;">' + escHtml_(task.questionText) + '</p>' +
      '<table style="width:100%;border-collapse:collapse;">' +
      '<tr><td style="padding:4px 8px;font-size:13px;color:#333;width:50%;"><strong>A.</strong> ' + escHtml_(task.optionA) + '</td>' +
          '<td style="padding:4px 8px;font-size:13px;color:#333;width:50%;"><strong>B.</strong> ' + escHtml_(task.optionB) + '</td></tr>' +
      '<tr><td style="padding:4px 8px;font-size:13px;color:#333;"><strong>C.</strong> ' + escHtml_(task.optionC) + '</td>' +
          '<td style="padding:4px 8px;font-size:13px;color:#333;"><strong>D.</strong> ' + escHtml_(task.optionD) + '</td></tr>' +
      '</table></div>';
  } else {
    taskHtml =
      '<div style="background:#fff3e0;border-left:4px solid #ff8f00;padding:12px 16px;margin-top:10px;border-radius:0 6px 6px 0;">' +
      '<p style="font-size:13px;color:#e65100;margin:0;">Your teacher will provide follow-up materials for this topic in class.</p>' +
      '</div>';
  }

  return '<div style="border:1px solid #e0e0e0;border-radius:8px;padding:16px;margin-bottom:14px;">' +
    // Header row: Q number circle + topic
    '<table style="width:100%;border-collapse:collapse;margin-bottom:10px;"><tr>' +
    '<td style="width:36px;vertical-align:middle;">' +
    '<div style="background:#1a237e;color:#fff;border-radius:50%;width:32px;height:32px;text-align:center;line-height:32px;font-size:12px;font-weight:bold;">Q' + q.number + '</div>' +
    '</td>' +
    '<td style="vertical-align:middle;padding-left:10px;">' +
    '<div style="font-size:13px;font-weight:bold;color:#212121;">' + escHtml_(q.topic || 'Topic not yet tagged') + '</div>' +
    '<div style="margin-top:4px;">' +
    (q.discipline ? '<span style="background:' + discColour + ';color:#fff;border-radius:3px;padding:2px 7px;font-size:10px;margin-right:4px;">' + escHtml_(q.discipline) + '</span>' : '') +
    (q.ao ? '<span style="background:#e8eaf6;color:#3949ab;border-radius:3px;padding:2px 7px;font-size:10px;margin-right:4px;">' + escHtml_(q.ao) + '</span>' : '') +
    (q.blooms ? '<span style="background:#eceff1;color:' + bloomsColour + ';border-radius:3px;padding:2px 7px;font-size:10px;">' + escHtml_(q.blooms) + '</span>' : '') +
    '</div>' +
    '</td></tr></table>' +
    taskHtml +
    '</div>';
}

function pickTask_(tasks, questionNumber) {
  var list = tasks[String(questionNumber)];
  if (!list || list.length === 0) { return null; }
  return list[Math.floor(Math.random() * list.length)];
}


// ═══════════════════════════════════════════════════════════
// EMAIL SENDER
// ═══════════════════════════════════════════════════════════

function sendStudentEmail_(student, html, classInfo, config, mode, subject) {
  var to  = mode === 'TEST' ? config.hoksEmail : student.email;
  var cc  = [];
  if (mode === 'LIVE') {
    if (student.fatherEmail) { cc.push(student.fatherEmail); }
    if (student.motherEmail) { cc.push(student.motherEmail); }
  }

  // FIX ISSUE 4: guard against blank student email in live mode (P2 already had this)
  if (mode === 'LIVE' && !to) {
    throw new Error('Student email is blank for ' + student.fullName +
                    ' (ID: ' + student.studentId + '). Please update the Student Register.');
  }

  var opts = {
    to:       to,
    subject:  subject + ' — ' + student.fullName,
    htmlBody: html,
    name:     buildFromName_(classInfo),
    replyTo:  classInfo.leadEmail,
  };
  if (cc.length > 0) { opts.cc = cc.join(','); }

  MailApp.sendEmail(opts);
}


// ═══════════════════════════════════════════════════════════
// TEACHER SUMMARY EMAIL
// ═══════════════════════════════════════════════════════════

function sendTeacherSummary_(classInfo, results, config, mode, paperName) {
  var sent    = results.filter(function(r){ return r.status === 'sent'; });
  var skipped = results.filter(function(r){ return r.status === 'skipped'; });
  var failed  = results.filter(function(r){ return r.status === 'failed'; });
  var to      = mode === 'TEST' ? config.hoksEmail : classInfo.leadEmail;

  var rows = '';
  sent.forEach(function(r) {
    var pn = parseFloat(r.pct) || 0;
    var c  = pn >= 80 ? '#2e7d32' : pn >= 60 ? '#e65100' : '#b71c1c';
    rows += '<tr style="border-bottom:1px solid #e0e0e0;">' +
      '<td style="padding:7px 10px;font-size:13px;">' + escHtml_(r.student.fullName) + '</td>' +
      '<td style="padding:7px 10px;font-size:13px;text-align:center;color:' + c + ';font-weight:bold;">' + Math.round(pn) + '%</td>' +
      '<td style="padding:7px 10px;font-size:13px;text-align:center;">' + (r.wrongCount || 0) + '</td>' +
      '<td style="padding:7px 10px;font-size:13px;color:#2e7d32;">&#10003; Sent</td></tr>';
  });
  skipped.forEach(function(r) {
    rows += '<tr style="border-bottom:1px solid #e0e0e0;opacity:0.6;">' +
      '<td style="padding:7px 10px;font-size:13px;">' + escHtml_(r.student.fullName) + '</td>' +
      '<td style="padding:7px 10px;text-align:center;">—</td>' +
      '<td style="padding:7px 10px;text-align:center;">—</td>' +
      '<td style="padding:7px 10px;font-size:13px;color:#9e9e9e;">Skipped (' + escHtml_(r.reason) + ')</td></tr>';
  });
  failed.forEach(function(r) {
    rows += '<tr style="border-bottom:1px solid #e0e0e0;">' +
      '<td style="padding:7px 10px;font-size:13px;">' + escHtml_(r.student.fullName) + '</td>' +
      '<td colspan="2" style="padding:7px 10px;font-size:13px;color:#b71c1c;">Error</td>' +
      '<td style="padding:7px 10px;font-size:13px;color:#b71c1c;">&#10007; Failed</td></tr>';
  });

  var testBanner = mode === 'TEST'
    ? '<div style="background:#ff8f00;color:#fff;padding:10px;text-align:center;font-weight:bold;font-size:12px;">TEST MODE SUMMARY</div>'
    : '';

  var html = '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;margin:0;padding:0;background:#f5f5f5;">' +
    testBanner +
    '<div style="max-width:620px;margin:0 auto;background:#fff;">' +
    '<div style="background:#1a237e;padding:20px 30px;">' +
    '<h2 style="color:#fff;margin:0;font-size:18px;">' + escHtml_(paperName) + ' Emails — Class Summary</h2>' +
    '<p style="color:#c5cae9;margin:4px 0 0;font-size:13px;">' + escHtml_(classInfo.classCode) + ' · ' + escHtml_(config.academicYear) + ' · ' + escHtml_(config.currentTerm) + '</p>' +
    '</div>' +
    '<div style="padding:24px 30px;">' +
    '<table style="width:100%;border-collapse:collapse;margin-bottom:20px;"><tr>' +
    '<td style="width:33%;text-align:center;background:#e8f5e9;border-radius:8px;padding:14px;">' +
    '<div style="font-size:28px;font-weight:bold;color:#2e7d32;">' + sent.length + '</div>' +
    '<div style="font-size:12px;color:#388e3c;">Emails Sent</div></td>' +
    '<td style="width:4%;"></td>' +
    '<td style="width:33%;text-align:center;background:#fff3e0;border-radius:8px;padding:14px;">' +
    '<div style="font-size:28px;font-weight:bold;color:#e65100;">' + skipped.length + '</div>' +
    '<div style="font-size:12px;color:#ef6c00;">Skipped</div></td>' +
    '<td style="width:4%;"></td>' +
    '<td style="width:33%;text-align:center;background:' + (failed.length > 0 ? '#ffebee' : '#f5f5f5') + ';border-radius:8px;padding:14px;">' +
    '<div style="font-size:28px;font-weight:bold;color:' + (failed.length > 0 ? '#c62828' : '#9e9e9e') + ';">' + failed.length + '</div>' +
    '<div style="font-size:12px;color:' + (failed.length > 0 ? '#c62828' : '#9e9e9e') + ';">Failed</div></td>' +
    '</tr></table>' +
    '<table style="width:100%;border-collapse:collapse;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden;">' +
    '<thead><tr style="background:#e8eaf6;">' +
    '<th style="padding:9px 10px;text-align:left;font-size:12px;color:#3949ab;">Student</th>' +
    '<th style="padding:9px 10px;text-align:center;font-size:12px;color:#3949ab;">Score %</th>' +
    '<th style="padding:9px 10px;text-align:center;font-size:12px;color:#3949ab;">Wrong Qs</th>' +
    '<th style="padding:9px 10px;text-align:left;font-size:12px;color:#3949ab;">Status</th>' +
    '</tr></thead><tbody>' + rows + '</tbody></table>' +
    '<p style="font-size:13px;color:#666;margin-top:20px;">Best wishes,<br><strong>' + escHtml_(config.signOffName) + '</strong><br>KS3 Science · Doha College</p>' +
    '</div></div></body></html>';

  MailApp.sendEmail({
    to:       to,
    subject:  (mode === 'TEST' ? '[TEST] ' : '') + '✅ ' + classInfo.classCode + ' ' + paperName + ' — ' + sent.length + ' emails sent',
    htmlBody: html,
    name:     'KS3 Science System',
    replyTo:  config.hoksEmail,
  });

  // CC Andy on live runs (unless he's the teacher)
  if (mode === 'LIVE' && classInfo.leadEmail.toLowerCase() !== config.hoksEmail.toLowerCase()) {
    MailApp.sendEmail({
      to:       config.hoksEmail,
      subject:  '📋 ' + classInfo.classCode + ' P1 complete — ' + sent.length + '/' + results.length + ' sent',
      htmlBody: html,
      name:     'KS3 Science System',
      replyTo:  config.hoksEmail,
    });
  }
}


// ═══════════════════════════════════════════════════════════
// PROCESSING LOG
// ═══════════════════════════════════════════════════════════

function initProcessingLog_(ss, queueEntry, config) {
  var sheet   = ss.getSheetByName('Processing Log');
  var lastRow = Math.max(sheet.getLastRow(), 2) + 1;
  var logId   = 'P1-' + queueEntry.classCode + '-' +
                Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmm');

  var vals = [logId, new Date(), 'Paper1_FollowUp',
    'Y9-P1-' + config.currentTerm, config.academicYear, config.currentTerm,
    queueEntry.classCode, queueEntry.mode, 'PROCESSING',
    '', '', '', queueEntry.teacherEmail, ''];
  sheet.getRange(lastRow, 1, 1, vals.length).setValues([vals]);
  SpreadsheetApp.flush();
  return lastRow;
}

function finaliseProcessingLog_(ss, logRow, processed, failed, errors) {
  var sheet = ss.getSheetByName('Processing Log');
  sheet.getRange(logRow, PL_COL.status).setValue(failed === 0 ? 'COMPLETE' : 'FAILED');
  sheet.getRange(logRow, PL_COL.processed).setValue(processed);
  sheet.getRange(logRow, PL_COL.failed).setValue(failed);
  if (errors.length > 0) {
    sheet.getRange(logRow, PL_COL.errors).setValue(errors.join('\n'));
  }
  SpreadsheetApp.flush();
}

function logError_(ss, queueEntry, errorMsg) {
  try {
    var sheet   = ss.getSheetByName('Processing Log');
    var lastRow = Math.max(sheet.getLastRow(), 2) + 1;
    var vals    = ['ERR-' + Date.now(), new Date(), 'Paper1_FollowUp',
      '', '', '', queueEntry.classCode, queueEntry.mode, 'FAILED',
      '', '', errorMsg, queueEntry.teacherEmail, ''];
    sheet.getRange(lastRow, 1, 1, vals.length).setValues([vals]);
  } catch(e) {}
}

/**
 * Write one log row per student beneath the class summary row.
 * Cols 1-14 repeat the run context. Cols 15-20 are student-specific.
 * This creates a full audit trail — if a student claims they didn't
 * receive an email, we have a timestamped record of exactly when it
 * was sent and to which address.
 */
function writeStudentLogRows_(ss, classLogRow, studentLogRows, config) {
  if (!studentLogRows || studentLogRows.length === 0) { return; }
  try {
    var sheet   = ss.getSheetByName('Processing Log');
    var startRow = classLogRow + 1;
    var rows    = [];

    // Read the class summary row to repeat context in each student row
    var classRowVals = sheet.getRange(classLogRow, 1, 1, 14).getValues()[0];
    var logId        = classRowVals[PL_COL.logId - 1];
    var paperId      = classRowVals[PL_COL.paperId - 1];
    var academicYear = classRowVals[PL_COL.academicYear - 1];
    var term         = classRowVals[PL_COL.term - 1];
    var classCode    = classRowVals[PL_COL.classCode - 1];
    var triggerType  = classRowVals[PL_COL.triggerType - 1];
    var runBy        = classRowVals[PL_COL.runBy - 1];

    studentLogRows.forEach(function(entry, idx) {
      var rowId = logId + '-S' + (idx + 1);
      rows.push([
        rowId,                                          // col 1  Log ID
        new Date(),                                     // col 2  Timestamp
        'Paper1_Student',                               // col 3  Script Name
        paperId,                                        // col 4  Paper ID
        academicYear,                                   // col 5  Academic Year
        term,                                           // col 6  Term
        classCode,                                      // col 7  Class Code
        triggerType,                                    // col 8  Trigger Type
        entry.status,                                   // col 9  Status
        entry.status === 'SENT' ? 1 : 0,               // col 10 Records Processed
        entry.status === 'FAILED' ? 1 : 0,             // col 11 Records Failed
        entry.reason || '',                             // col 12 Error Messages
        runBy,                                          // col 13 Run By
        '',                                             // col 14 Notes
        entry.student.fullName,                         // col 15 Student Name
        safeId_(entry.student.studentId),               // col 16 Student ID
        entry.emailTo || '',                            // col 17 Email Sent To
        entry.score   || '',                            // col 18 Score
        entry.wrongQs ? entry.wrongQs.length : '',      // col 19 Wrong Q Count
        entry.wrongQs ? entry.wrongQs.join(', ') : '',  // col 20 Wrong Q List
      ]);
    });

    if (rows.length > 0) {
      // Insert rows to avoid overwriting anything below
      sheet.insertRowsAfter(classLogRow, rows.length);
      sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
      // Style student rows subtly — indented appearance
      sheet.getRange(startRow, 1, rows.length, 20)
           .setFontSize(9)
           .setFontColor('#555555')
           .setBackground('#fafafa');
      SpreadsheetApp.flush();
    }
  } catch(e) {
    Logger.log('KS3: writeStudentLogRows_ failed: ' + e.toString());
  }
}


// ═══════════════════════════════════════════════════════════
// STATUS UPDATER — batched for performance (FIX ISSUE 6)
// Reads ID column ONCE, builds a map, writes all statuses in one batch
// ═══════════════════════════════════════════════════════════

/**
 * Build a map of studentId → row index (0-based) for the Paper 1 data range.
 * Called once before processing students, then used for O(1) lookups.
 */
function buildP1StatusMap_(ss) {
  var sheet = ss.getSheetByName('Paper 1');
  var last  = sheet.getLastRow();
  if (last < P1_ROW.dataStart) { return {}; }
  var ids   = sheet.getRange(P1_ROW.dataStart, P1_COL.studentId,
                last - P1_ROW.dataStart + 1, 1).getValues();
  var map   = {};
  for (var i = 0; i < ids.length; i++) {
    var sid = safeId_(ids[i][0]);
    if (sid) { map[sid] = i; }
  }
  return map;
}

/**
 * Write all pending status updates in a single batch.
 * statusUpdates = array of { studentId, newStatus }
 */
function flushP1StatusUpdates_(ss, statusUpdates, p1StatusMap) {
  if (!statusUpdates || statusUpdates.length === 0) { return; }
  var sheet = ss.getSheetByName('Paper 1');
  statusUpdates.forEach(function(update) {
    var idx = p1StatusMap[safeId_(update.studentId)];
    if (idx !== undefined) {
      sheet.getRange(P1_ROW.dataStart + idx, P1_COL.status).setValue(update.newStatus);
    }
  });
}


// ═══════════════════════════════════════════════════════════
// PREVIEW SIDEBAR
// ═══════════════════════════════════════════════════════════

function showPreviewSidebar_(classCode) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var config    = getConfig_(ss);
  var classInfo = getClassInfo_(ss, classCode, config.academicYear);
  var answerKey = getAnswerKey_(ss, config.p1MaxQ);
  var tags      = getTagData_(ss, config.p1MaxQ);
  var tasks     = getFollowUpTasks_(ss);
  var students  = getStudentData_(ss, classCode);

  if (!classInfo) {
    SpreadsheetApp.getUi().alert('Class ' + classCode + ' not found in Classes tab.');
    return;
  }
  if (students.length === 0) {
    SpreadsheetApp.getUi().alert('No active students found for ' + classCode + '.');
    return;
  }

  var student = students[0];
  var p1Sheet = ss.getSheetByName('Paper 1');
  var p1Last  = p1Sheet.getLastRow();
  var p1Data  = p1Last >= P1_ROW.dataStart
    ? p1Sheet.getRange(P1_ROW.dataStart, 1, p1Last - P1_ROW.dataStart + 1, P1_COL.status).getValues()
    : [];

  var p1Row = null;
  for (var i = 0; i < p1Data.length; i++) {
    if (safeId_(p1Data[i][0]) === safeId_(student.studentId)) { p1Row = p1Data[i]; break; }
  }

  var wrongQs = p1Row ? identifyWrongQuestions_(p1Row, answerKey, tags, config.p1MaxQ) : [];
  var score   = p1Row ? p1Row[P1_COL.score - 1] : 0;
  var pct     = p1Row ? p1Row[P1_COL.pct - 1]   : 0;
  var html    = buildP1StudentEmail_(student, wrongQs, tasks, score, pct, classInfo, config, 'PREVIEW');

  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;font-size:11px;color:#666;padding:8px 12px;background:#fff3e0;border-bottom:1px solid #ffe082;">' +
      '📧 PREVIEW — ' + escHtml_(classCode) + ' — ' + escHtml_(student.fullName) +
      ' &nbsp;|&nbsp; ' + wrongQs.length + ' wrong question' + (wrongQs.length !== 1 ? 's' : '') +
      '</div>' + html
    ).setTitle('Email Preview — ' + classCode).setWidth(660)
  );
}


// ═══════════════════════════════════════════════════════════
// TRIGGER MANAGEMENT
// ═══════════════════════════════════════════════════════════

function installQueueTrigger() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var config = getConfig_(ss);

  if (!isAdminAccount_(Session.getActiveUser().getEmail(), config)) {
    SpreadsheetApp.getUi().alert('Only the script owner (' + config.ownerEmail + ') can install triggers.');
    return;
  }

  // Remove any existing P1 queue triggers (prevent duplicates)
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processP1Queue') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('processP1Queue').timeBased().everyMinutes(1).create();

  SpreadsheetApp.getUi().alert(
    '✅ Queue trigger installed.\n\n' +
    'The system will check for queued and scheduled classes every 1 minute.\n\n' +
    'You only need to do this once — the trigger persists until you remove it.'
  );
}


// ═══════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════

function buildFromName_(classInfo) {
  var first   = (classInfo.leadFirst || '').trim();
  var surname = (classInfo.leadName  || '').trim().split(' ').pop();
  var display = (first && surname) ? first + ' ' + surname : (classInfo.leadName || 'KS3 Science').trim();
  // FIX ISSUE 7: ensure space before em-dash separator
  return display + ' \u2014 KS3 Science';
}

function extractFileId_(urlOrId) {
  if (!urlOrId) { return ''; }
  var match = String(urlOrId).match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : String(urlOrId).trim();
}

function escHtml_(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escAttr_(str) {
  return String(str || '').replace(/'/g, '&#39;').replace(/"/g, '&quot;');
}

function pad_(n) {
  return n < 10 ? '0' + n : String(n);
}

/**
 * Safely convert a cell value to a Student ID string.
 * Handles:
 *  - Leading zeros (stored as text) → preserved
 *  - Large numbers (13+ digits) → Google converts to float,
 *    JS then produces scientific notation e.g. 1.3175e+12
 *    This converts back to the full integer string.
 *  - Already-correct strings → returned as-is
 */
function safeId_(val) {
  if (val === null || val === undefined || val === '') { return ''; }
  var s = String(val).trim();
  // Detect scientific notation (e.g. "1.3175e+12")
  if (s.indexOf('e') !== -1 || s.indexOf('E') !== -1) {
    var n = Number(s);
    if (!isNaN(n)) { s = n.toFixed(0); }
  }
  return s;
}
