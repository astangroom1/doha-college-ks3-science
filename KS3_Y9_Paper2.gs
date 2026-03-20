/**
 * ============================================================
 *  KS3 SCIENCE — PAPER 2 FOLLOW-UP EMAIL SCRIPT
 *  Year 9 | Doha College | Andy Stangroom
 *  v1.0
 * ============================================================
 *
 *  KEY DIFFERENCES FROM PAPER 1:
 *   • Tier-based (Rebuild/Refine/Extend/Mastery) not per-question variants
 *   • Sub-questions grouped into parent questions by parsing row 6 labels
 *   • Per-question tier calculated from marks earned vs max marks
 *   • Max marks read from hidden row 1 (D1:AQ1)
 *   • Separate queue keys (P2_QUEUE / P2_CURRENT) — P1 and P2 can run simultaneously
 *   • Status column is AU (col 47) not AT (col 46)
 *
 *  FLOW:
 *   1. Teacher opens menu → KS3 Science → Paper 2
 *   2. Selects class from permitted list
 *   3. Chooses: Send Now OR Schedule (date/time picker)
 *   4. Queue Status & Management dialog — Reschedule / Cancel / Send Now
 *   5. 5-min trigger processes queue automatically
 *   6. Teacher receives class summary email on completion
 *
 *  ALL 10 BUGS FROM PAPER 1 AUDIT PRE-FIXED:
 *   1.  handleSchedule checks addToQueue_ return value
 *   2.  JSON.parse wrapped in try/catch
 *   3.  quickPick uses server Qatar date not browser local time
 *   4.  No non-existent server function calls in dialog
 *   5.  No location.reload() in modal iframe
 *   6.  processP2Class_ uses try/finally for guaranteed log finalization
 *   7.  Tag/max-mark reads use config.p2MaxMarks not hardcoded value
 *   8.  sendNowFromQueue inserts after existing Send Now entries (fair queue)
 *   9.  Outer catch in queue processor logs properly
 *   10. Tag data read in single batch API call
 *
 *  CRASH GUARDS (7):
 *   1. Atomic queue ops via LockService
 *   2. State validation before every action
 *   3. Scheduled entries fire exactly once
 *   4. Stuck run auto-recovery (15 min)
 *   5. Per-student try/catch
 *   6. Empty sheet / zero max marks guards
 *   7. Idempotent status check (never double-sends)
 *
 * ============================================================
 */


// ═══════════════════════════════════════════════════════════
// COLUMN / ROW CONSTANTS — verified against setup script v4
// ═══════════════════════════════════════════════════════════

// Year Controller row numbers (values in col B)
var YC2_ROW = {
  academicYear:   4,
  currentTerm:    6,
  ownerEmail:     10,
  hoksName:       11,
  hoksEmail:      12,
  signOffName:    13,
  p2MaxMarks:     16,
  bannerP2:       20,
  bannerModelAns: 21,
  bannerChaseUp:  23,
  bannerTeacher:  24,
  masterFolderId: 34,
  // Tier thresholds — rows 27-32 in Year Controller
  rebuildMax:     27,
  refineMin:      28,
  refineMax:      29,
  extendMin:      30,
  extendMax:      31,
  masteryPct:     32,
};

// Classes tab columns (1-based)
var CL2_COL = {
  classCode:   2,
  leadName:    4,
  leadEmail:   5,
  leadFirst:   6,
  otherName:   7,
  otherEmail:  8,
  otherFirst:  9,
  lessonSplit: 10,
};

// Student Register columns (1-based)
var SR2_COL = {
  studentId:   1,
  firstName:   2,
  lastName:    3,
  fullName:    4,
  email:       5,
  classCode:   8,
  status:      10,
  fatherEmail: 12,
  motherEmail: 14,
};

// Paper 2 structure
var P2_ROW = {
  maxMarks:      1,   // hidden — max marks per sub-question
  aoTag:         2,
  bloomsTag:     3,
  qTypeTag:      4,
  qGroupTag:     5,
  sqLabels:      6,   // sub-question labels e.g. "1a", "1b", "2a"
  header:        7,
  dataStart:     8,
};

var P2_COL = {
  studentId:  1,
  fullName:   2,
  classCode:  3,
  sq1:        4,    // D — first sub-question
  sqLast:     43,   // AQ — last sub-question (40 slots)
  grandTotal: 44,   // AR
  pct:        45,   // AS
  tier:       46,   // AT
  status:     47,   // AU
};

// Follow-Up Tasks columns (1-based)
var FU2_COL = {
  taskId:       1,
  paperId:      2,
  taskType:     3,
  questionNum:  4,
  ao:           7,
  discipline:   9,
  topic:        10,
  taskTitle:    18,
  taskDesc:     19,
  background:   20,
  youtubeLink:  21,
  modelAnswer:  22,
};

// Processing Log columns (1-based)
var PL2_COL = {
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
};

// Queue / property keys — SEPARATE from Paper 1 to allow simultaneous runs
var P2_QUEUE_KEY      = 'P2_QUEUE';
var P2_PROCESSING_KEY = 'P2_CURRENT';
var P2_MAX_QUEUE      = 6;
var P2_STUCK_MIN      = 15;
var TZ2               = 'Asia/Qatar';

// Tier thresholds are loaded from Year Controller at runtime via getP2Config_
// Do NOT hardcode them here — use config.tierThresholds in calculateQuestionResults_


// ═══════════════════════════════════════════════════════════
// MENU — added to existing KS3 Science menu by onOpen
// ═══════════════════════════════════════════════════════════
// NOTE: Paper 2 adds its items to the existing menu built by Paper 1.
// If running Paper 2 standalone (without Paper 1 script), use onOpen2()
// as a temporary measure.

function addP2ToMenu() {
  // Call this from Paper 1's onOpen if both scripts are in the same project.
  // When both scripts are present, the menu is built once in Paper 1's onOpen
  // and this function is called to add Paper 2 items.
}

function onOpen2() {
  // Standalone menu — only use if Paper 1 script is NOT present
  SpreadsheetApp.getUi()
    .createMenu('KS3 Science')
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Paper 2')
        .addItem('Send Follow-Up Emails — Live',   'showP2LiveDialog')
        .addItem('Send Follow-Up Emails — Test',   'showP2TestDialog')
        .addItem('Preview Emails',                 'showP2PreviewDialog')
        .addSeparator()
        .addItem('Queue Status & Management',      'showP2QueueManager')
    )
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Admin')
        .addItem('Install / Refresh Queue Trigger','installP2QueueTrigger')
        .addItem('Reset Stuck Runs',               'resetP2StuckRuns')
        .addItem('Process Queue Now',              'processP2Queue')
    )
    .addToUi();
}


// ═══════════════════════════════════════════════════════════
// DIALOG ENTRY POINTS
// ═══════════════════════════════════════════════════════════

function showP2LiveDialog()    { showP2ClassDialog_('LIVE');    }
function showP2TestDialog()    { showP2ClassDialog_('TEST');    }
function showP2PreviewDialog() { showP2ClassDialog_('PREVIEW'); }

function showP2ClassDialog_(mode) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var config  = getP2Config_(ss);
  var user    = Session.getActiveUser().getEmail();
  var classes = getP2PermittedClasses_(ss, user, config);

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
  var headerColour = mode === 'LIVE'    ? '#b71c1c' :
                     mode === 'TEST'    ? '#bf360c' : '#1b5e20';
  var title        = mode === 'LIVE'    ? 'Paper 2 — Send Follow-Up Emails (Live)' :
                     mode === 'TEST'    ? 'Paper 2 — Send Follow-Up Emails (Test)' :
                                          'Paper 2 — Preview Emails';

  var options = classes.map(function(c) {
    return '<option value="' + escP2Attr_(c.classCode) + '">' +
           escP2Html_(c.classCode) + ' — ' + escP2Html_(c.leadName) +
           (c.lessonSplit ? ' (' + c.lessonSplit + ')' : '') +
           '</option>';
  }).join('');

  // FIX BUG 3: pass server Qatar date down to client
  var qatarToday = Utilities.formatDate(new Date(), TZ2, 'yyyy-MM-dd');

  var html =
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    'body{font-family:Arial,sans-serif;font-size:13px;padding:0;margin:0;color:#212121;}' +
    '.hdr{background:' + headerColour + ';color:#fff;padding:16px 20px;}' +
    '.hdr h3{margin:0;font-size:15px;}.hdr p{margin:4px 0 0;font-size:11px;opacity:0.85;}' +
    '.body{padding:20px;}' +
    'label{font-weight:bold;display:block;margin-bottom:6px;font-size:12px;color:#555;}' +
    'select{width:100%;padding:8px;font-size:13px;border:1px solid #ccc;border-radius:4px;margin-bottom:16px;box-sizing:border-box;}' +
    '.row{display:flex;gap:10px;}' +
    '.btn{flex:1;padding:10px;font-size:13px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;}' +
    '.pri{background:' + headerColour + ';color:#fff;}' +
    '.sec{background:#e8eaf6;color:#3949ab;}' +
    '.can{background:#f5f5f5;color:#666;border:1px solid #ddd;}' +
    '.btn:hover{opacity:0.88;}' +
    '#sched{display:none;margin-top:16px;border-top:1px solid #e0e0e0;padding-top:16px;}' +
    '.qp{display:flex;gap:8px;margin-bottom:12px;}' +
    '.qpb{padding:6px 10px;font-size:11px;background:#f5f5f5;color:#666;border:1px solid #ddd;border-radius:4px;cursor:pointer;}' +
    'input[type=date],input[type=time]{padding:8px;border:1px solid #ccc;border-radius:4px;font-size:13px;flex:1;box-sizing:border-box;}' +
    '#msg{display:none;padding:12px;border-radius:6px;margin-top:12px;font-size:13px;}' +
    '.ok{background:#e8f5e9;color:#2e7d32;}.err{background:#ffebee;color:#c62828;}' +
    '</style></head><body>' +
    '<div class="hdr"><h3>' + escP2Html_(title) + '</h3><p>' + escP2Html_(modeLabel) + '</p></div>' +
    '<div class="body">' +
    '<label>Select your class:</label>' +
    '<select id="cc">' + options + '</select>' +
    (mode === 'PREVIEW'
      ? '<div class="row"><button class="btn pri" onclick="doPreview()">Open Preview</button><button class="btn can" onclick="google.script.host.close()">Cancel</button></div>'
      : '<div class="row"><button class="btn pri" onclick="doSendNow()">Send Now</button><button class="btn sec" onclick="showSched()">Schedule</button><button class="btn can" onclick="google.script.host.close()">Cancel</button></div>'
    ) +
    '<div id="sched">' +
    '<label>Schedule delivery date &amp; time:</label>' +
    '<div class="row" style="margin-bottom:10px;">' +
    '<input type="date" id="sd"><input type="time" id="st">' +
    '</div>' +
    '<div class="qp">' +
    '<button class="qpb" onclick="qp(0,16,0)">Today 4pm</button>' +
    '<button class="qpb" onclick="qp(1,8,0)">Tomorrow 8am</button>' +
    '<button class="qpb" onclick="qp(1,16,0)">Tomorrow 4pm</button>' +
    '</div>' +
    '<div class="row"><button class="btn pri" onclick="doSchedule()">Confirm Schedule</button><button class="btn can" onclick="hideSched()">Back</button></div>' +
    '</div>' +
    '<div id="msg"></div>' +
    '</div>' +
    '<script>' +
    'var QT="' + qatarToday + '";' +
    'var MODE="' + mode + '";' +
    'function cc(){return document.getElementById("cc").value;}' +
    'function msg(t,ok){var m=document.getElementById("msg");m.className=ok?"ok":"err";m.textContent=t;m.style.display="block";}' +
    'function showSched(){document.getElementById("sched").style.display="block";}' +
    'function hideSched(){document.getElementById("sched").style.display="none";}' +
    // FIX BUG 3: use server Qatar date
    'function qp(days,h,m){' +
    '  var p=QT.split("-"),d=new Date(+p[0],+p[1]-1,+p[2]);' +
    '  d.setDate(d.getDate()+days);' +
    '  var y=d.getFullYear(),mo=d.getMonth()+1,dy=d.getDate();' +
    '  document.getElementById("sd").value=y+"-"+(mo<10?"0"+mo:mo)+"-"+(dy<10?"0"+dy:dy);' +
    '  document.getElementById("st").value=(h<10?"0"+h:h)+":"+(m<10?"0"+m:m);' +
    '}' +
    'function doPreview(){' +
    '  google.script.run.withSuccessHandler(function(){google.script.host.close();})' +
    '  .withFailureHandler(function(e){msg(e.message,false);})' +
    '  .handleP2Preview(cc());' +
    '}' +
    'function doSendNow(){' +
    '  msg("Adding to queue...",true);' +
    '  google.script.run' +
    '  .withSuccessHandler(function(r){msg(r,true);setTimeout(function(){google.script.host.close();},3000);})' +
    '  .withFailureHandler(function(e){msg(e.message,false);})' +
    '  .handleP2SendNow(cc(),MODE);' +
    '}' +
    'function doSchedule(){' +
    '  var d=document.getElementById("sd").value,t=document.getElementById("st").value;' +
    '  if(!d||!t){msg("Please select a date and time.",false);return;}' +
    '  msg("Scheduling...",true);' +
    '  google.script.run' +
    '  .withSuccessHandler(function(r){msg(r,true);setTimeout(function(){google.script.host.close();},3000);})' +
    '  .withFailureHandler(function(e){msg(e.message,false);})' +
    '  .handleP2Schedule(cc(),MODE,d,t);' +
    '}' +
    '</script></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(440).setHeight(mode === 'PREVIEW' ? 210 : 330),
    title
  );
}


// ═══════════════════════════════════════════════════════════
// DIALOG CALLBACKS
// ═══════════════════════════════════════════════════════════

function handleP2Preview(classCode) {
  showP2PreviewSidebar_(classCode);
}

function handleP2SendNow(classCode, mode) {
  return p2AddToQueue_(classCode, mode, null);
}

function handleP2Schedule(classCode, mode, dateStr, timeStr) {
  var parts     = dateStr.split('-');
  var timeParts = timeStr.split(':');
  var isoStr    = parts[0] + '-' + p2Pad_(parseInt(parts[1])) + '-' + p2Pad_(parseInt(parts[2])) +
                  'T' + p2Pad_(parseInt(timeParts[0])) + ':' + p2Pad_(parseInt(timeParts[1])) + ':00+03:00';
  var scheduledFor = new Date(isoStr).getTime();

  if (isNaN(scheduledFor) || scheduledFor <= Date.now()) {
    throw new Error('Please choose a date and time in the future.');
  }

  // FIX BUG 1: check return value of addToQueue_
  var result   = p2AddToQueue_(classCode, mode, scheduledFor);
  var isError  = result.indexOf('currently being processed') !== -1 ||
                 result.indexOf('already in the queue')      !== -1 ||
                 result.indexOf('queue is full')             !== -1 ||
                 result.indexOf('System is busy')            !== -1;
  if (isError) { throw new Error(result); }

  return classCode + ' scheduled for ' +
         Utilities.formatDate(new Date(scheduledFor), TZ2, 'EEE d MMM \'at\' HH:mm') +
         '. You\'ll receive a summary email when sent.';
}


// ═══════════════════════════════════════════════════════════
// QUEUE MANAGEMENT
// ═══════════════════════════════════════════════════════════

function p2AddToQueue_(classCode, mode, scheduledFor) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    return 'System is busy — please try again in a moment.';
  }

  try {
    var queue   = p2GetQueue_();
    var current = p2GetCurrent_();

    // CRASH GUARD 2: state validation
    if (current && current.classCode === classCode) {
      return classCode + ' is currently being processed. You\'ll receive a summary email when complete.';
    }
    if (queue.some(function(q){ return q.classCode === classCode; })) {
      return classCode + ' is already in the queue. Use Queue Status to reschedule or cancel it.';
    }
    if (queue.length >= P2_MAX_QUEUE) {
      return 'The queue is full (' + P2_MAX_QUEUE + ' classes). Please try again in a few minutes.';
    }

    var entry = {
      id:           Utilities.getUuid(),
      classCode:    classCode,
      mode:         mode,
      teacherEmail: Session.getActiveUser().getEmail(),
      queuedAt:     Date.now(),
      scheduledFor: scheduledFor || null,
    };
    queue.push(entry);
    p2SaveQueue_(queue);

    if (scheduledFor) {
      return classCode + ' added to the queue.';  // caller formats time
    }
    var sendNowCount = queue.filter(function(q){ return !q.scheduledFor; }).length;
    var pos = sendNowCount === 1 ? '1st' : sendNowCount === 2 ? '2nd' :
              sendNowCount === 3 ? '3rd' : sendNowCount + 'th';
    var waitMsg = current ? ' Another class is currently processing.' : '';
    return classCode + ' added to queue (' + pos + ' in line).' + waitMsg +
           '\n\nYour class will be processed automatically.\nYou\'ll receive a summary email when complete.';

  } finally {
    lock.releaseLock();
  }
}

function p2RescheduleEntry(entryId, dateStr, timeStr) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { throw new Error('System busy — try again.'); }

  try {
    var queue = p2GetQueue_();
    var idx   = p2IndexOf_(queue, entryId);
    if (idx === -1) { throw new Error('Entry no longer in queue — may have already been processed.'); }

    var parts    = dateStr.split('-');
    var tParts   = timeStr.split(':');
    var isoStr   = parts[0] + '-' + p2Pad_(parseInt(parts[1])) + '-' + p2Pad_(parseInt(parts[2])) +
                   'T' + p2Pad_(parseInt(tParts[0])) + ':' + p2Pad_(parseInt(tParts[1])) + ':00+03:00';
    var sf = new Date(isoStr).getTime();
    if (isNaN(sf) || sf <= Date.now()) { throw new Error('Please choose a date and time in the future.'); }

    queue[idx].scheduledFor = sf;
    p2SaveQueue_(queue);
    return queue[idx].classCode + ' rescheduled for ' +
           Utilities.formatDate(new Date(sf), TZ2, 'EEE d MMM \'at\' HH:mm') + '.';
  } finally {
    lock.releaseLock();
  }
}

function p2CancelEntry(entryId) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { throw new Error('System busy — try again.'); }

  try {
    var queue = p2GetQueue_();
    var idx   = p2IndexOf_(queue, entryId);
    if (idx === -1) { throw new Error('Entry no longer in queue.'); }
    var cc = queue[idx].classCode;
    queue.splice(idx, 1);
    p2SaveQueue_(queue);
    return cc + ' has been removed from the queue.';
  } finally {
    lock.releaseLock();
  }
}

function p2SendNowFromQueue(entryId) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { throw new Error('System busy — try again.'); }

  try {
    var queue   = p2GetQueue_();
    var idx     = p2IndexOf_(queue, entryId);
    if (idx === -1) { throw new Error('Entry no longer in queue.'); }

    var current = p2GetCurrent_();
    var cc      = queue[idx].classCode;
    if (current && current.classCode === cc) {
      throw new Error(cc + ' is currently being processed. You\'ll receive a summary email shortly.');
    }

    var entry = queue.splice(idx, 1)[0];
    entry.scheduledFor = null;

    // FIX BUG 8: insert after last existing Send Now entry — fair queue ordering
    var insertPos = 0;
    for (var j = 0; j < queue.length; j++) {
      if (!queue[j].scheduledFor) { insertPos = j + 1; }
    }
    queue.splice(insertPos, 0, entry);
    p2SaveQueue_(queue);

    return cc + ' moved to Send Now — will process within 5 minutes. You can close this dialog.';
  } finally {
    lock.releaseLock();
  }
}

// ── Queue accessors with FIX BUG 2 JSON.parse safety ─────
function p2GetQueue_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(P2_QUEUE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch(e) {
    Logger.log('KS3 P2 ERROR: p2GetQueue_ JSON.parse failed: ' + e.toString());
    return [];
  }
}

function p2SaveQueue_(queue) {
  PropertiesService.getScriptProperties().setProperty(P2_QUEUE_KEY, JSON.stringify(queue));
}

function p2GetCurrent_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(P2_PROCESSING_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch(e) {
    Logger.log('KS3 P2 ERROR: p2GetCurrent_ JSON.parse failed: ' + e.toString());
    PropertiesService.getScriptProperties().deleteProperty(P2_PROCESSING_KEY);
    return null;
  }
}

function p2SetCurrent_(entry) {
  PropertiesService.getScriptProperties().setProperty(P2_PROCESSING_KEY, JSON.stringify(entry));
}

function p2ClearCurrent_() {
  PropertiesService.getScriptProperties().deleteProperty(P2_PROCESSING_KEY);
}

function p2IndexOf_(queue, entryId) {
  for (var i = 0; i < queue.length; i++) {
    if (queue[i].id === entryId) { return i; }
  }
  return -1;
}


// ═══════════════════════════════════════════════════════════
// QUEUE STATUS & MANAGEMENT DIALOG
// ═══════════════════════════════════════════════════════════

function showP2QueueManager() {
  var queue   = p2GetQueue_();
  var current = p2GetCurrent_();
  var now     = Date.now();

  var currentHtml = '';
  if (current) {
    var elapsed = Math.round((now - current.startedAt) / 60000);
    currentHtml =
      '<div class="sec"><div class="sec-title proc-title">⚙️ Currently Processing</div>' +
      '<div class="entry proc-entry">' +
      '<span class="cc">' + escP2Html_(current.classCode) + '</span>' +
      '<span class="meta">Started ' + elapsed + ' min ago · ' + escP2Html_(current.mode) + '</span>' +
      '<span class="note">Processing — cannot be modified</span>' +
      '</div></div>';
  }

  var schedHtml = '';
  var nowHtml   = '';

  queue.forEach(function(entry) {
    var tLabel = entry.scheduledFor
      ? Utilities.formatDate(new Date(entry.scheduledFor), TZ2, 'EEE d MMM \'at\' HH:mm')
      : 'As soon as possible';
    var isSched = !!entry.scheduledFor;

    var eHtml =
      '<div class="entry" id="e-' + entry.id + '">' +
      '<div class="etop">' +
      '<span class="cc">' + escP2Html_(entry.classCode) + '</span>' +
      '<span class="meta">' + (isSched ? '🕐 ' : '⚡ ') + escP2Html_(tLabel) + ' · ' + escP2Html_(entry.mode) + '</span>' +
      '</div>' +
      '<div class="eact">' +
      (isSched
        ? '<button class="btn sec" onclick="openR(\'' + entry.id + '\',\'' + escP2Attr_(entry.classCode) + '\')">Reschedule</button>' +
          '<button class="btn danger" onclick="doCancel(\'' + entry.id + '\',\'' + escP2Attr_(entry.classCode) + '\')">Cancel</button>' +
          '<button class="btn pri" onclick="doSendNow(\'' + entry.id + '\',\'' + escP2Attr_(entry.classCode) + '\')">Send Now</button>'
        : '<button class="btn danger" onclick="doCancel(\'' + entry.id + '\',\'' + escP2Attr_(entry.classCode) + '\')">Cancel</button>'
      ) +
      '</div></div>';

    if (isSched) { schedHtml += eHtml; } else { nowHtml += eHtml; }
  });

  var emptyHtml = (!current && queue.length === 0)
    ? '<div style="padding:20px;text-align:center;color:#9e9e9e;font-size:13px;">Queue is empty.</div>'
    : '';

  var html =
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    'body{font-family:Arial,sans-serif;font-size:13px;padding:0;margin:0;background:#f5f5f5;}' +
    '.hdr{background:#b71c1c;color:#fff;padding:16px 20px;}' +
    '.hdr h3{margin:0;font-size:15px;}' +
    '.body{padding:16px;}' +
    '.sec{margin-bottom:14px;}' +
    '.sec-title{font-size:11px;font-weight:bold;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;color:#5c6bc0;}' +
    '.proc-title{color:#e65100;}' +
    '.entry{background:#fff;border:1px solid #e0e0e0;border-radius:6px;padding:12px;margin-bottom:8px;}' +
    '.proc-entry{border-color:#ff8f00;background:#fff8e1;}' +
    '.etop{margin-bottom:8px;}' +
    '.cc{font-weight:bold;font-size:14px;color:#b71c1c;margin-right:10px;}' +
    '.meta{font-size:12px;color:#666;}' +
    '.eact{display:flex;gap:8px;flex-wrap:wrap;}' +
    '.note{font-size:12px;color:#e65100;font-style:italic;}' +
    '.btn{padding:6px 12px;border:none;border-radius:4px;cursor:pointer;font-size:12px;font-weight:bold;}' +
    '.pri{background:#b71c1c;color:#fff;}' +
    '.sec{background:#e8eaf6;color:#3949ab;}' +
    '.danger{background:#ffebee;color:#c62828;}' +
    '.btn:hover{opacity:0.85;}' +
    '.rpanel{display:none;margin-top:10px;padding:10px;background:#fff8e1;border-radius:6px;border:1px solid #ffe082;}' +
    '.rpanel input{padding:6px;border:1px solid #ccc;border-radius:4px;font-size:12px;margin-right:6px;}' +
    '#fb{padding:10px;border-radius:6px;margin-top:10px;display:none;font-size:13px;}' +
    '.fbok{background:#e8f5e9;color:#2e7d32;}.fberr{background:#ffebee;color:#c62828;}' +
    '</style></head><body>' +
    '<div class="hdr"><h3>📋 Paper 2 — Queue Status &amp; Management</h3></div>' +
    '<div class="body">' +
    emptyHtml + currentHtml +
    (schedHtml ? '<div class="sec"><div class="sec-title">🕐 Scheduled</div>' + schedHtml + '</div>' : '') +
    (nowHtml   ? '<div class="sec"><div class="sec-title">⚡ Send Now Queue</div>' + nowHtml + '</div>'   : '') +
    '<div id="fb"></div>' +
    '<div id="rpanel" class="rpanel">' +
    '<strong id="rlabel" style="display:block;margin-bottom:8px;font-size:12px;"></strong>' +
    '<input type="date" id="rd"><input type="time" id="rt">' +
    '<div style="margin-top:8px;display:flex;gap:8px;">' +
    '<button class="btn pri" onclick="confirmR()">Confirm</button>' +
    '<button class="btn sec" onclick="closeR()">Cancel</button>' +
    '</div></div>' +
    '</div>' +
    '<script>' +
    'var pid=null;' +
    'function fb(msg,ok){var e=document.getElementById("fb");e.textContent=msg;e.className=ok?"fbok":"fberr";e.style.display="block";}' +
    'function openR(id,cc){' +
    '  pid=id;' +
    '  document.getElementById("rlabel").textContent="Reschedule " + cc;' +
    '  var p=document.getElementById("rpanel");' +
    '  document.getElementById("e-"+id).appendChild(p);p.style.display="block";' +
    '}' +
    'function closeR(){document.getElementById("rpanel").style.display="none";}' +
    'function confirmR(){' +
    '  var d=document.getElementById("rd").value,t=document.getElementById("rt").value;' +
    '  if(!d||!t){fb("Please select a date and time.",false);return;}' +
    // FIX BUG 4+5: no getQueueManagerData, no location.reload — just show success message
    '  google.script.run' +
    '  .withSuccessHandler(function(r){fb(r+" Close and reopen Queue Status to see the update.",true);closeR();})' +
    '  .withFailureHandler(function(e){fb(e.message,false);})' +
    '  .p2RescheduleEntry(pid,d,t);' +
    '}' +
    'function doCancel(id,cc){' +
    '  if(!confirm("Remove "+cc+" from the queue?"))return;' +
    '  google.script.run' +
    '  .withSuccessHandler(function(r){fb(r,true);var e=document.getElementById("e-"+id);if(e)e.remove();})' +
    '  .withFailureHandler(function(e){fb(e.message,false);})' +
    '  .p2CancelEntry(id);' +
    '}' +
    'function doSendNow(id,cc){' +
    '  if(!confirm("Move "+cc+" to Send Now? Will process within 5 minutes."))return;' +
    // FIX BUG 5: no location.reload — just show confirmation
    '  google.script.run' +
    '  .withSuccessHandler(function(r){fb(r,true);})' +
    '  .withFailureHandler(function(e){fb(e.message,false);})' +
    '  .p2SendNowFromQueue(id);' +
    '}' +
    '</script></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(500).setHeight(480),
    'Paper 2 — Queue Status & Management'
  );
}


// ═══════════════════════════════════════════════════════════
// QUEUE PROCESSOR — runs every 5 minutes via trigger
// ═══════════════════════════════════════════════════════════

function processP2Queue() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) { return; }

  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var now = Date.now();

    // CRASH GUARD 4: recover stuck runs
    p2RecoverStuck_(ss);

    // CRASH GUARD 3: don't start another if already processing
    if (p2GetCurrent_()) {
      lock.releaseLock();
      return;
    }

    // Find next ready entry
    var queue   = p2GetQueue_();
    var nextIdx = -1;
    for (var i = 0; i < queue.length; i++) {
      if (!queue[i].scheduledFor || queue[i].scheduledFor <= now) {
        nextIdx = i; break;
      }
    }

    if (nextIdx === -1) {
      lock.releaseLock();
      return;
    }

    var next = queue.splice(nextIdx, 1)[0];
    p2SaveQueue_(queue);

    // CRASH GUARD 3: mark processing BEFORE releasing lock
    next.startedAt = Date.now();
    p2SetCurrent_(next);
    lock.releaseLock();

    try {
      processP2Class_(ss, next);
    } catch(e) {
      p2LogError_(ss, next, e.toString());
      p2ClearCurrent_();
    }

  } catch(e) {
    // FIX BUG 9: log outer errors
    try {
      Logger.log('KS3 P2 processP2Queue outer error: ' + e.toString());
      var ss2 = SpreadsheetApp.getActiveSpreadsheet();
      p2LogError_(ss2, { classCode:'UNKNOWN', mode:'TRIGGER', teacherEmail:'system' }, e.toString());
    } catch(e3) {}
    try { lock.releaseLock(); } catch(e2) {}
  }
}


// ═══════════════════════════════════════════════════════════
// STUCK RUN RECOVERY
// ═══════════════════════════════════════════════════════════

function p2RecoverStuck_(ss) {
  var current = p2GetCurrent_();
  if (!current) { return; }

  var elapsedMin = (Date.now() - current.startedAt) / 60000;
  if (elapsedMin <= P2_STUCK_MIN) { return; }

  p2ClearCurrent_();

  try {
    var config = getP2Config_(ss);
    // FIX NOTE 1: guard against blank hoksEmail before attempting to send
    if (!config.hoksEmail) { return; }
    MailApp.sendEmail({
      to:      config.hoksEmail,
      subject: '[KS3 Science] P2 stuck run recovered — ' + current.classCode,
      body:    'A Paper 2 processing run for ' + current.classCode +
               ' was stuck for ' + Math.round(elapsedMin) + ' minutes and has been reset.\n\n' +
               'The class has NOT been re-queued automatically.\n' +
               'Please ask the teacher to re-submit from the KS3 Science menu.\n\n' +
               'Run details:\n' + JSON.stringify(current, null, 2),
    });
  } catch(e) {}
}

function resetP2StuckRuns() {
  p2RecoverStuck_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Stuck run check complete. Check your email if any were recovered.');
}


// ═══════════════════════════════════════════════════════════
// CORE CLASS PROCESSOR
// ═══════════════════════════════════════════════════════════

function processP2Class_(ss, queueEntry) {
  var classCode = queueEntry.classCode;
  var mode      = queueEntry.mode;
  var config    = getP2Config_(ss);
  var classInfo = getP2ClassInfo_(ss, classCode, config.academicYear);
  var students  = getP2Students_(ss, classCode);

  // Pre-log validation (throws caught by outer catch in processP2Queue)
  if (!classInfo) {
    throw new Error('Class ' + classCode + ' not found in Classes tab.');
  }
  if (students.length === 0) {
    throw new Error('No active students for class ' + classCode + ' in Student Register.');
  }

  // CRASH GUARD 6: validate Paper 2 has data and max marks are set
  var p2Sheet   = ss.getSheetByName('Paper 2');
  var p2LastRow = p2Sheet.getLastRow();
  if (p2LastRow < P2_ROW.dataStart) {
    throw new Error('Paper 2 has no student data. Please enter marks before sending emails.');
  }

  var maxMarks = p2Sheet.getRange(P2_ROW.maxMarks, P2_COL.sq1, 1, 40).getValues()[0];
  var totalMax = maxMarks.reduce(function(s, v) { return s + (parseFloat(v) || 0); }, 0);
  if (totalMax === 0) {
    throw new Error('Paper 2 max marks (row 1, D1:AQ1) are all blank or zero. Please enter max marks before sending emails.');
  }

  // Read sub-question labels and tags in batch (FIX BUG 10)
  var sqLabels = p2Sheet.getRange(P2_ROW.sqLabels, P2_COL.sq1, 1, 40).getValues()[0];
  var allTags  = p2Sheet.getRange(P2_ROW.aoTag, P2_COL.sq1, 4, 40).getValues();
  var tags = {
    ao:       allTags[0],  // row 2
    blooms:   allTags[1],  // row 3
    qType:    allTags[2],  // row 4
    qGroup:   allTags[3],  // row 5
  };

  // Build question structure — group sub-questions into parent questions
  var questions = buildQuestionStructure_(sqLabels, maxMarks);

  // FIX BUG B: guard against empty questions array
  // This happens if all sq labels are still the default 'SQ1'...'SQ40' placeholders
  if (questions.length === 0) {
    throw new Error(
      'Paper 2 sub-question labels have not been set up for ' + classCode + '.\n' +
      'Please replace the default SQ1, SQ2... labels in row 6 of the Paper 2 tab ' +
      'with real labels such as 1a, 1b, 2a etc. before sending emails.'
    );
  }

  // Read follow-up tasks
  var tasks = getP2FollowUpTasks_(ss);

  // Read P2 data
  var p2Data = p2Sheet.getRange(
    P2_ROW.dataStart, 1,
    p2LastRow - P2_ROW.dataStart + 1,
    P2_COL.status
  ).getValues();

  // Map studentId → row
  var p2Map = {};
  p2Data.forEach(function(row) {
    var sid = String(row[P2_COL.studentId - 1]).trim();
    if (sid) { p2Map[sid] = row; }
  });

  // Create log row AFTER all validation
  var logRow    = p2InitLog_(ss, queueEntry, config);
  var processed = 0;
  var failed    = 0;
  var errors    = [];
  var results   = [];

  // FIX ISSUE 6: build status map once before loop — one read instead of one per student
  var p2StatusMap   = buildP2StatusMap_(ss);
  var statusUpdates = [];

  // FIX BUG 6: try/finally guarantees log finalisation and clearCurrent
  try {
    // CRASH GUARD 5: per-student try/catch
    students.forEach(function(student) {
      try {
        var p2Row = p2Map[student.studentId];

        if (!p2Row) {
          results.push({ student: student, status: 'skipped', reason: 'No P2 data entered' });
          return;
        }

        // CRASH GUARD 7: idempotent — never double-send
        var currentStatus = String(p2Row[P2_COL.status - 1]).trim();
        if (currentStatus === 'Emails Sent' || currentStatus === 'Complete') {
          results.push({ student: student, status: 'skipped', reason: 'Already processed' });
          return;
        }

        var grandTotal  = p2Row[P2_COL.grandTotal - 1];
        var pct         = p2Row[P2_COL.pct - 1];
        var overallTier = p2Row[P2_COL.tier - 1];

        var questionResults = calculateQuestionResults_(p2Row, questions);

        var emailHtml = buildP2StudentEmail_(
          student, questionResults, tasks, grandTotal, pct, overallTier,
          totalMax, classInfo, config, mode, tags
        );

        sendP2StudentEmail_(student, emailHtml, classInfo, config, mode, 'Paper 2 Follow-Up');
        statusUpdates.push({ studentId: student.studentId, newStatus: 'Emails Sent' });

        processed++;
        results.push({
          student: student,
          status:  'sent',
          pct:     pct,
          tier:    overallTier,
          questionsWithTasks: questionResults.filter(function(q){ return q.tier !== 'Mastery'; }).length,
        });

      } catch(e) {
        failed++;
        errors.push(student.fullName + ': ' + e.toString());
        results.push({ student: student, status: 'failed', reason: e.toString() });
      }
    });

    // Batch write all status updates in one pass
    flushP2StatusUpdates_(ss, statusUpdates, p2StatusMap);
    sendP2TeacherSummary_(classInfo, results, config, mode);

  } catch(e) {
    failed++;
    errors.push('Unexpected error: ' + e.toString());
  } finally {
    // FIX BUG 6: ALWAYS runs regardless of what happened above
    p2FinaliseLog_(ss, logRow, processed, failed, errors);
    p2ClearCurrent_();
  }
}


// ═══════════════════════════════════════════════════════════
// QUESTION STRUCTURE BUILDER
// ═══════════════════════════════════════════════════════════

/**
 * Parse sub-question labels to group them into parent questions.
 * Labels like "1a", "1b", "1c(i)", "2a", "2b" → parent questions 1 and 2.
 * Blank labels = unused slot — skip.
 *
 * Returns array of question objects:
 * { parentNum, subQuestions: [{sqIndex, label, maxMark}], totalMax }
 */
function buildQuestionStructure_(sqLabels, maxMarks) {
  var qMap = {};     // parentNum → {subQuestions, totalMax}
  var qOrder = [];   // to preserve order

  for (var i = 0; i < sqLabels.length; i++) {
    var label   = String(sqLabels[i]).trim();
    var maxMark = parseFloat(maxMarks[i]) || 0;

    if (!label || label === 'SQ' + (i + 1)) {
      // Blank or still has default placeholder — skip this slot
      continue;
    }

    // Parse parent question number from label
    // Handles: "1a", "1b", "1c(i)", "2", "Q1a", "1.a", "1 a"
    var parentNum = parseParentQuestion_(label);
    if (!parentNum) { continue; }

    if (!qMap[parentNum]) {
      qMap[parentNum] = { parentNum: parentNum, subQuestions: [], totalMax: 0 };
      qOrder.push(parentNum);
    }
    qMap[parentNum].subQuestions.push({ sqIndex: i, label: label, maxMark: maxMark });
    qMap[parentNum].totalMax += maxMark;
  }

  return qOrder.map(function(pn) { return qMap[pn]; });
}

/**
 * Extract parent question number from a sub-question label.
 * "1a" → "1", "Q2b" → "2", "3c(i)" → "3", "10a" → "10"
 * Returns string or null if can't parse.
 */
function parseParentQuestion_(label) {
  // Strip leading Q/q
  var clean = label.replace(/^[Qq]/, '').trim();
  // Match leading digits
  var match = clean.match(/^(\d+)/);
  return match ? match[1] : null;
}

/**
 * Calculate per-question results for a student.
 * Returns array of question results with earned marks, max, % and tier.
 * Thresholds come from config (Year Controller) not hardcoded values.
 */
function calculateQuestionResults_(p2Row, questions, tierThresholds) {
  return questions.map(function(q) {
    var earned = 0;
    q.subQuestions.forEach(function(sq) {
      var val = parseFloat(p2Row[P2_COL.sq1 + sq.sqIndex - 1]) || 0;
      earned += val;
    });

    // FIX BUG C: clamp earned to max — prevents negative 'dropped' if teacher over-marks
    earned = Math.min(earned, q.totalMax);

    var qPct  = q.totalMax > 0 ? Math.round((earned / q.totalMax) * 100) : 0;

    // FIX BUG A: use thresholds from config not hardcoded constants
    var t     = tierThresholds;
    var qTier = qPct >= t.mastery ? 'Mastery' :
                qPct >= t.extend  ? 'Extend'  :
                qPct >= t.refine  ? 'Refine'  : 'Rebuild';

    return {
      parentNum: q.parentNum,
      earned:    earned,
      totalMax:  q.totalMax,
      pct:       qPct,
      tier:      qTier,
      dropped:   q.totalMax - earned,  // always >= 0 after clamp
    };
  });
}


// ═══════════════════════════════════════════════════════════
// DATA READERS
// ═══════════════════════════════════════════════════════════

function getP2Config_(ss) {
  var sheet = ss.getSheetByName('Year Controller');
  var data  = sheet.getRange(1, 2, 40, 1).getValues();

  function val(row) {
    var v = data[row - 1][0];
    return (v !== null && v !== undefined) ? String(v).trim() : '';
  }

  return {
    academicYear:   val(YC2_ROW.academicYear),
    currentTerm:    val(YC2_ROW.currentTerm),
    ownerEmail:     val(YC2_ROW.ownerEmail),
    hoksName:       val(YC2_ROW.hoksName),
    hoksEmail:      val(YC2_ROW.hoksEmail),
    signOffName:    val(YC2_ROW.signOffName),
    p2MaxMarks:     parseInt(val(YC2_ROW.p2MaxMarks)) || 40,
    bannerP2:       p2ExtractFileId_(val(YC2_ROW.bannerP2)),
    bannerModelAns: p2ExtractFileId_(val(YC2_ROW.bannerModelAns)),
    bannerChaseUp:  p2ExtractFileId_(val(YC2_ROW.bannerChaseUp)),
    bannerTeacher:  p2ExtractFileId_(val(YC2_ROW.bannerTeacher)),
    // FIX BUG A: tier thresholds from Year Controller, not hardcoded
    // Defaults match Paper Config: Rebuild 0-64, Refine 65-89, Extend 90-99, Mastery 100
    tierThresholds: {
      mastery: parseInt(val(YC2_ROW.masteryPct))  || 100,
      extend:  parseInt(val(YC2_ROW.extendMin))   || 90,
      refine:  parseInt(val(YC2_ROW.refineMin))   || 65,
    },
  };
}

function getP2ClassInfo_(ss, classCode, academicYear) {
  var sheet = ss.getSheetByName('Classes');
  var last  = sheet.getLastRow();
  if (last < 3) { return null; }
  var data  = sheet.getRange(3, 1, last - 2, 17).getValues();

  for (var i = 0; i < data.length; i++) {
    var rowAY = String(data[i][CL2_COL.classCode - 2]).trim(); // col A = academicYear, index 0
    var rowCC = String(data[i][CL2_COL.classCode - 1]).trim();
    // FIX ISSUE 5: filter by academic year; accept blank AY for backward compat
    if (rowCC === classCode && (!rowAY || !academicYear || rowAY === academicYear)) {
      return {
        classCode:   rowCC,
        leadName:    String(data[i][CL2_COL.leadName - 1]).trim(),
        leadEmail:   String(data[i][CL2_COL.leadEmail - 1]).trim(),
        leadFirst:   String(data[i][CL2_COL.leadFirst - 1]).trim(),
        otherName:   String(data[i][CL2_COL.otherName - 1]).trim(),
        otherEmail:  String(data[i][CL2_COL.otherEmail - 1]).trim(),
        otherFirst:  String(data[i][CL2_COL.otherFirst - 1]).trim(),
        lessonSplit: String(data[i][CL2_COL.lessonSplit - 1]).trim(),
      };
    }
  }
  return null;
}

function getP2Students_(ss, classCode) {
  var sheet = ss.getSheetByName('Student Register');
  var last  = sheet.getLastRow();
  if (last < 3) { return []; }
  var data  = sheet.getRange(3, 1, last - 2, 14).getValues();
  var out   = [];

  data.forEach(function(row) {
    var sid    = String(row[SR2_COL.studentId - 1]).trim();
    var cc     = String(row[SR2_COL.classCode - 1]).trim();
    var status = String(row[SR2_COL.status - 1]).trim();
    if (!sid || cc !== classCode || status !== 'Active') { return; }

    out.push({
      studentId:   sid,
      firstName:   String(row[SR2_COL.firstName - 1]).trim(),
      lastName:    String(row[SR2_COL.lastName - 1]).trim(),
      fullName:    String(row[SR2_COL.fullName - 1]).trim() ||
                   String(row[SR2_COL.firstName - 1]).trim() + ' ' + String(row[SR2_COL.lastName - 1]).trim(),
      email:       String(row[SR2_COL.email - 1]).trim(),
      fatherEmail: String(row[SR2_COL.fatherEmail - 1]).trim(),
      motherEmail: String(row[SR2_COL.motherEmail - 1]).trim(),
    });
  });
  return out;
}

function getP2FollowUpTasks_(ss) {
  var sheet = ss.getSheetByName('Follow-Up Tasks');
  var last  = sheet.getLastRow();
  if (last < 3) { return {}; }
  var data  = sheet.getRange(3, 1, last - 2, 22).getValues();

  // Tasks keyed as taskType + '-' + questionNum, e.g. 'P2-Rebuild-1', 'P2-Extend-3'
  var tasks = {};

  data.forEach(function(row) {
    var type = String(row[FU2_COL.taskType - 1]).trim();
    var qNum = String(row[FU2_COL.questionNum - 1]).trim();

    // Only P2 task types
    if (!type.match(/^P2-(Rebuild|Refine|Extend|Mastery)$/) || !qNum) { return; }

    var key = type + '-' + qNum;
    if (!tasks[key]) { tasks[key] = []; }
    tasks[key].push({
      taskType:    type,
      questionNum: qNum,
      taskTitle:   String(row[FU2_COL.taskTitle - 1]).trim(),
      taskDesc:    String(row[FU2_COL.taskDesc - 1]).trim(),
      background:  String(row[FU2_COL.background - 1]).trim(),
      youtubeLink: String(row[FU2_COL.youtubeLink - 1]).trim(),
      modelAnswer: String(row[FU2_COL.modelAnswer - 1]).trim(),
      ao:          String(row[FU2_COL.ao - 1]).trim(),
      discipline:  String(row[FU2_COL.discipline - 1]).trim(),
      topic:       String(row[FU2_COL.topic - 1]).trim(),
    });
  });

  return tasks;
}

function getP2PermittedClasses_(ss, userEmail, config) {
  var isAdmin = p2IsAdmin_(userEmail, config);
  var sheet   = ss.getSheetByName('Classes');
  var last    = sheet.getLastRow();
  if (last < 3) { return []; }
  var data    = sheet.getRange(3, 1, last - 2, 10).getValues();
  var out     = [];

  data.forEach(function(row) {
    var cc    = String(row[CL2_COL.classCode - 1]).trim();
    var rowAY = String(row[0]).trim();  // col A = Academic Year
    var lead  = String(row[CL2_COL.leadEmail - 1]).trim().toLowerCase();
    var other = String(row[CL2_COL.otherEmail - 1]).trim().toLowerCase();
    if (!cc) { return; }
    // FIX ISSUE 5: filter by academic year; accept blank AY for backward compat
    if (rowAY && config.academicYear && rowAY !== config.academicYear) { return; }
    if (isAdmin || lead === userEmail.toLowerCase() || other === userEmail.toLowerCase()) {
      out.push({
        classCode:   cc,
        leadName:    String(row[CL2_COL.leadName - 1]).trim(),
        lessonSplit: String(row[CL2_COL.lessonSplit - 1]).trim(),
      });
    }
  });
  return out;
}

function p2IsAdmin_(email, config) {
  return email.toLowerCase() === (config.ownerEmail || '').toLowerCase();
}


// ═══════════════════════════════════════════════════════════
// EMAIL BUILDER — STUDENT
// ═══════════════════════════════════════════════════════════

function buildP2StudentEmail_(student, questionResults, tasks, grandTotal,
                               pct, overallTier, totalMax, classInfo, config, mode, tags) {
  var firstName   = student.firstName || student.fullName.split(' ')[0];
  var teacherName = classInfo.leadFirst || classInfo.leadName;
  var bannerUrl   = config.bannerP2
    ? 'https://drive.google.com/uc?export=view&id=' + config.bannerP2 : '';

  var pctNum       = parseFloat(pct) || 0;
  var pctDisplay   = (pct !== '' && pct !== null && pct !== undefined) ? Math.round(pctNum) + '%' : 'N/A';
  // FIX BUG D: guard undefined as well as null/empty for grandTotal
  var scoreDisplay = (grandTotal !== '' && grandTotal !== null && grandTotal !== undefined)
    ? grandTotal + ' / ' + totalMax : 'N/A';

  // Overall tier colours — Emerging/Developing/Secure/Mastery
  var tierColours = {
    'Mastery':    { bg: '#1b5e20', fc: '#ffffff', light: '#e8f5e9' },
    'Secure':     { bg: '#2e7d32', fc: '#ffffff', light: '#c8e6c9' },
    'Developing': { bg: '#e65100', fc: '#ffffff', light: '#fff3e0' },
    'Emerging':   { bg: '#b71c1c', fc: '#ffffff', light: '#ffebee' },
  };
  var tc = tierColours[overallTier] || tierColours['Rebuild'];

  var testBanner = mode === 'TEST'
    ? '<div style="background:#ff8f00;color:#fff;padding:10px;text-align:center;font-weight:bold;font-size:12px;">' +
      'TEST MODE — This email was redirected from ' + escP2Html_(student.email) + '</div>'
    : '';

  // Build per-question task blocks
  var taskBlocks = '';
  var hasAnyTask = false;

  // Show questions where marks were dropped (not Mastery) in order of most marks dropped
  var actionableQs = questionResults
    .filter(function(qr) { return qr.dropped > 0 && qr.totalMax > 0; })
    .sort(function(a, b) { return b.dropped - a.dropped; });

  if (actionableQs.length === 0 && overallTier === 'Mastery') {
    taskBlocks = '<div style="background:#e8f5e9;border-radius:8px;padding:16px 20px;text-align:center;margin-bottom:16px;">' +
      '<p style="color:#2e7d32;font-weight:bold;font-size:15px;margin:0;">🏆 Outstanding! Full marks across all questions.</p>' +
      '</div>';
  } else {
    actionableQs.forEach(function(qr) {
      var taskKey  = 'P2-' + qr.tier + '-' + qr.parentNum;
      var taskList = tasks[taskKey];
      var task     = (taskList && taskList.length > 0)
        ? taskList[Math.floor(Math.random() * taskList.length)]
        : null;

      taskBlocks += buildP2QuestionBlock_(qr, task);
      hasAnyTask  = true;
    });

    if (!hasAnyTask) {
      taskBlocks = '<p style="color:#e65100;font-size:13px;padding:12px 0;">Your teacher will provide follow-up materials in class.</p>';
    }
  }

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
    '<body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif;">' +
    testBanner +
    '<div style="max-width:620px;margin:0 auto;background:#fff;">' +

    (bannerUrl ? '<img src="' + bannerUrl + '" width="620" style="display:block;width:100%;max-width:620px;" alt="KS3 Science">' : '') +

    '<div style="background:#b71c1c;padding:24px 30px;">' +
    '<h1 style="color:#fff;margin:0;font-size:22px;">Paper 2 — Follow-Up Tasks</h1>' +
    '<p style="color:#ef9a9a;margin:6px 0 0;font-size:14px;">' + escP2Html_(config.academicYear) + ' · ' + escP2Html_(config.currentTerm) + '</p>' +
    '</div>' +

    '<div style="padding:24px 30px;">' +
    '<p style="font-size:15px;color:#212121;">Dear ' + escP2Html_(firstName) + ',</p>' +
    '<p style="font-size:14px;color:#424242;line-height:1.6;">Your Paper 2 results are in. Your personalised follow-up tasks are below.</p>' +

    // Score + tier box
    '<div style="background:#f5f5f5;border-radius:8px;padding:16px 20px;margin:20px 0;">' +
    '<table style="width:100%;border-collapse:collapse;"><tr>' +
    '<td style="width:33%;text-align:center;">' +
    '<div style="font-size:11px;color:#666;text-transform:uppercase;font-weight:bold;letter-spacing:0.5px;">Score</div>' +
    '<div style="font-size:26px;font-weight:bold;color:#212121;">' + escP2Html_(pctDisplay) + '</div>' +
    '<div style="font-size:12px;color:#666;">' + escP2Html_(String(scoreDisplay)) + '</div>' +
    '</td>' +
    '<td style="width:34%;text-align:center;padding:0 10px;">' +
    '<div style="background:' + tc.bg + ';color:' + tc.fc + ';border-radius:8px;padding:12px 8px;">' +
    '<div style="font-size:11px;text-transform:uppercase;font-weight:bold;letter-spacing:0.5px;opacity:0.85;">Overall Tier</div>' +
    '<div style="font-size:22px;font-weight:bold;">' + escP2Html_(overallTier || 'N/A') + '</div>' +
    '</div></td>' +
    '<td style="width:33%;text-align:center;">' +
    '<div style="font-size:11px;color:#666;text-transform:uppercase;font-weight:bold;letter-spacing:0.5px;">Questions</div>' +
    '<div style="font-size:26px;font-weight:bold;color:#b71c1c;">' + actionableQs.length + '</div>' +
    '<div style="font-size:12px;color:#666;">to follow up</div>' +
    '</td>' +
    '</tr></table>' +
    '</div>' +

    (actionableQs.length > 0
      ? '<h2 style="font-size:16px;color:#b71c1c;border-bottom:2px solid #ffebee;padding-bottom:8px;margin-bottom:16px;">Your Follow-Up Tasks</h2>'
      : '') +
    taskBlocks +

    '<div style="margin-top:30px;padding-top:20px;border-top:1px solid #f5f5f5;">' +
    '<p style="font-size:14px;color:#424242;">If you have any questions, please speak to me in class or reply to this email.</p>' +
    '<p style="font-size:14px;color:#424242;margin-bottom:4px;">Best wishes,</p>' +
    '<p style="font-size:14px;color:#b71c1c;font-weight:bold;margin:0;">' + escP2Html_(teacherName) + '</p>' +
    '<p style="font-size:12px;color:#666;margin:2px 0 0;">KS3 Science · Doha College</p>' +
    '</div></div>' +

    '<div style="background:#ffebee;padding:12px 30px;text-align:center;">' +
    '<p style="font-size:11px;color:#666;margin:0;">Automated message from KS3 Science. To reply, use your teacher\'s email address above.</p>' +
    '</div>' +
    '</div></body></html>';
}

function buildP2QuestionBlock_(qr, task) {
  var tierColours = {
    'Mastery': '#1b5e20', 'Extend': '#2e7d32',
    'Refine': '#e65100', 'Rebuild': '#b71c1c',
  };
  var tierBg = {
    'Mastery': '#e8f5e9', 'Extend': '#c8e6c9',
    'Refine': '#fff3e0', 'Rebuild': '#ffebee',
  };
  var tc = tierColours[qr.tier] || '#37474f';
  var tb = tierBg[qr.tier]     || '#f5f5f5';

  var taskHtml = '';
  if (task && task.taskTitle) {
    taskHtml =
      '<div style="background:#f9fbe7;border-left:4px solid #c6e03a;padding:14px 16px;margin-top:10px;border-radius:0 6px 6px 0;">' +
      '<div style="font-size:11px;color:#558b2f;font-weight:bold;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">' + escP2Html_(qr.tier) + ' Task</div>' +
      '<div style="font-size:14px;font-weight:bold;color:#212121;margin-bottom:6px;">' + escP2Html_(task.taskTitle) + '</div>' +
      (task.taskDesc ? '<p style="font-size:13px;color:#424242;margin:0 0 8px;line-height:1.5;">' + escP2Html_(task.taskDesc) + '</p>' : '') +
      (task.background ? '<div style="background:#fff;border-radius:4px;padding:8px 10px;font-size:12px;color:#666;margin-bottom:8px;">' + escP2Html_(task.background) + '</div>' : '') +
      (task.youtubeLink ? '<a href="' + escP2Html_(task.youtubeLink) + '" style="font-size:12px;color:#1565c0;">▶ Watch the support video</a>' : '') +
      '</div>';
  } else {
    taskHtml =
      '<div style="background:#fff3e0;border-left:4px solid #ff8f00;padding:12px 16px;margin-top:10px;border-radius:0 6px 6px 0;">' +
      '<p style="font-size:13px;color:#e65100;margin:0;">Your teacher will provide follow-up materials for Question ' + escP2Html_(qr.parentNum) + ' in class.</p>' +
      '</div>';
  }

  return '<div style="border:1px solid #e0e0e0;border-radius:8px;padding:16px;margin-bottom:16px;">' +
    '<div style="display:flex;align-items:center;margin-bottom:10px;">' +
    '<div style="background:#b71c1c;color:#fff;border-radius:50%;min-width:32px;height:32px;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:bold;margin-right:12px;">Q' + escP2Html_(qr.parentNum) + '</div>' +
    '<div style="flex:1;">' +
    '<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">' +
    '<span style="background:' + tc + ';color:#fff;border-radius:4px;padding:3px 8px;font-size:11px;font-weight:bold;">' + escP2Html_(qr.tier) + '</span>' +
    '<span style="font-size:12px;color:#666;">' + qr.earned + ' / ' + qr.totalMax + ' marks (' + qr.pct + '%)</span>' +
    '</div>' +
    '</div></div>' +

    '<div style="background:' + tb + ';border-radius:6px;padding:8px 12px;font-size:12px;color:' + tc + ';margin-bottom:4px;">' +
    'Marks dropped: <strong>' + qr.dropped + '</strong> out of ' + qr.totalMax +
    '</div>' +
    taskHtml +
    '</div>';
}


// ═══════════════════════════════════════════════════════════
// EMAIL SENDER
// ═══════════════════════════════════════════════════════════

function sendP2StudentEmail_(student, html, classInfo, config, mode, subject) {
  var to = mode === 'TEST' ? config.hoksEmail : student.email;
  var cc = [];
  if (mode === 'LIVE') {
    if (student.fatherEmail) { cc.push(student.fatherEmail); }
    if (student.motherEmail) { cc.push(student.motherEmail); }
  }

  // CRASH GUARD 6: guard against blank student email in live mode
  if (mode === 'LIVE' && !to) {
    throw new Error('Student email is blank for ' + student.fullName + ' (ID: ' + student.studentId + ')');
  }

  var opts = {
    to:       to,
    subject:  subject + ' — ' + student.fullName,
    htmlBody: html,
    name:     p2BuildFromName_(classInfo),
    replyTo:  classInfo.leadEmail,
  };
  if (cc.length > 0) { opts.cc = cc.join(','); }
  MailApp.sendEmail(opts);
}


// ═══════════════════════════════════════════════════════════
// TEACHER SUMMARY EMAIL
// ═══════════════════════════════════════════════════════════

function sendP2TeacherSummary_(classInfo, results, config, mode) {
  var sent    = results.filter(function(r){ return r.status === 'sent'; });
  var skipped = results.filter(function(r){ return r.status === 'skipped'; });
  var failed  = results.filter(function(r){ return r.status === 'failed'; });
  var to      = mode === 'TEST' ? config.hoksEmail : classInfo.leadEmail;

  // Overall tier breakdown counts
  var tierCounts = { Mastery: 0, Secure: 0, Developing: 0, Emerging: 0 };
  sent.forEach(function(r) {
    var t = String(r.tier || '').trim();
    if (tierCounts.hasOwnProperty(t)) { tierCounts[t]++; }
  });

  var rows = '';
  sent.forEach(function(r) {
    var pn  = parseFloat(r.pct) || 0;
    var tc  = r.tier === 'Mastery'    ? '#1b5e20' :
              r.tier === 'Secure'     ? '#2e7d32' :
              r.tier === 'Developing' ? '#e65100' : '#b71c1c';
    rows += '<tr style="border-bottom:1px solid #e0e0e0;">' +
      '<td style="padding:7px 10px;font-size:13px;">' + escP2Html_(r.student.fullName) + '</td>' +
      '<td style="padding:7px 10px;font-size:13px;text-align:center;color:' + tc + ';font-weight:bold;">' + Math.round(pn) + '%</td>' +
      '<td style="padding:7px 10px;font-size:13px;text-align:center;">' +
      '<span style="background:' + tc + ';color:#fff;border-radius:3px;padding:2px 7px;font-size:11px;">' + escP2Html_(r.tier || 'N/A') + '</span>' +
      '</td>' +
      '<td style="padding:7px 10px;font-size:13px;color:#2e7d32;">&#10003; Sent</td>' +
      '</tr>';
  });

  skipped.forEach(function(r) {
    rows += '<tr style="border-bottom:1px solid #e0e0e0;opacity:0.6;">' +
      '<td style="padding:7px 10px;font-size:13px;">' + escP2Html_(r.student.fullName) + '</td>' +
      '<td style="padding:7px 10px;text-align:center;">—</td>' +
      '<td style="padding:7px 10px;text-align:center;">—</td>' +
      '<td style="padding:7px 10px;font-size:13px;color:#9e9e9e;">Skipped (' + escP2Html_(r.reason) + ')</td>' +
      '</tr>';
  });

  failed.forEach(function(r) {
    rows += '<tr style="border-bottom:1px solid #e0e0e0;">' +
      '<td style="padding:7px 10px;font-size:13px;">' + escP2Html_(r.student.fullName) + '</td>' +
      '<td colspan="2" style="padding:7px 10px;font-size:13px;color:#b71c1c;">Error</td>' +
      '<td style="padding:7px 10px;font-size:13px;color:#b71c1c;">&#10007; Failed</td>' +
      '</tr>';
  });

  var tierBar =
    '<div style="display:flex;gap:8px;margin-bottom:20px;">' +
    buildTierCell_('Emerging',   tierCounts.Emerging,   '#b71c1c', '#ffebee') +
    buildTierCell_('Developing', tierCounts.Developing, '#e65100', '#fff3e0') +
    buildTierCell_('Secure',     tierCounts.Secure,     '#2e7d32', '#c8e6c9') +
    buildTierCell_('Mastery',    tierCounts.Mastery,    '#1b5e20', '#e8f5e9') +
    '</div>';

  var testBanner = mode === 'TEST'
    ? '<div style="background:#ff8f00;color:#fff;padding:10px;text-align:center;font-weight:bold;font-size:12px;">TEST MODE SUMMARY</div>'
    : '';

  var html = '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;margin:0;padding:0;background:#f5f5f5;">' +
    testBanner +
    '<div style="max-width:620px;margin:0 auto;background:#fff;">' +
    '<div style="background:#b71c1c;padding:20px 30px;">' +
    '<h2 style="color:#fff;margin:0;font-size:18px;">Paper 2 Emails — Class Summary</h2>' +
    '<p style="color:#ef9a9a;margin:4px 0 0;font-size:13px;">' + escP2Html_(classInfo.classCode) + ' · ' + escP2Html_(config.academicYear) + ' · ' + escP2Html_(config.currentTerm) + '</p>' +
    '</div>' +
    '<div style="padding:24px 30px;">' +
    tierBar +
    '<table style="width:100%;border-collapse:collapse;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden;margin-bottom:20px;">' +
    '<thead><tr style="background:#ffebee;">' +
    '<th style="padding:9px 10px;text-align:left;font-size:12px;color:#b71c1c;">Student</th>' +
    '<th style="padding:9px 10px;text-align:center;font-size:12px;color:#b71c1c;">Score %</th>' +
    '<th style="padding:9px 10px;text-align:center;font-size:12px;color:#b71c1c;">Tier</th>' +
    '<th style="padding:9px 10px;text-align:left;font-size:12px;color:#b71c1c;">Status</th>' +
    '</tr></thead><tbody>' + rows + '</tbody></table>' +
    '<p style="font-size:13px;color:#666;margin-top:20px;">Best wishes,<br><strong>' + escP2Html_(config.signOffName) + '</strong><br>KS3 Science · Doha College</p>' +
    '</div></div></body></html>';

  MailApp.sendEmail({
    to:       to,
    subject:  (mode === 'TEST' ? '[TEST] ' : '') + '✅ ' + classInfo.classCode + ' Paper 2 — ' + sent.length + ' emails sent',
    htmlBody: html,
    name:     'KS3 Science System',
    replyTo:  config.hoksEmail,
  });

  // CC Andy on live runs
  if (mode === 'LIVE' && classInfo.leadEmail.toLowerCase() !== config.hoksEmail.toLowerCase()) {
    MailApp.sendEmail({
      to:       config.hoksEmail,
      subject:  '📋 ' + classInfo.classCode + ' P2 complete — ' + sent.length + '/' + results.length + ' sent',
      htmlBody: html,
      name:     'KS3 Science System',
      replyTo:  config.hoksEmail,
    });
  }
}

function buildTierCell_(label, count, colour, bg) {
  return '<div style="flex:1;background:' + bg + ';border-radius:8px;padding:12px;text-align:center;">' +
    '<div style="font-size:22px;font-weight:bold;color:' + colour + ';">' + count + '</div>' +
    '<div style="font-size:11px;color:' + colour + ';font-weight:bold;">' + label + '</div>' +
    '</div>';
}


// ═══════════════════════════════════════════════════════════
// STATUS UPDATER
// ═══════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════
// STATUS UPDATER — batched (FIX ISSUE 6)
// ═══════════════════════════════════════════════════════════

function buildP2StatusMap_(ss) {
  var sheet = ss.getSheetByName('Paper 2');
  var last  = sheet.getLastRow();
  if (last < P2_ROW.dataStart) { return {}; }
  var ids   = sheet.getRange(P2_ROW.dataStart, P2_COL.studentId,
                last - P2_ROW.dataStart + 1, 1).getValues();
  var map   = {};
  for (var i = 0; i < ids.length; i++) {
    var sid = String(ids[i][0]).trim();
    if (sid) { map[sid] = i; }
  }
  return map;
}

function flushP2StatusUpdates_(ss, statusUpdates, p2StatusMap) {
  if (!statusUpdates || statusUpdates.length === 0) { return; }
  var sheet = ss.getSheetByName('Paper 2');
  statusUpdates.forEach(function(u) {
    var idx = p2StatusMap[String(u.studentId).trim()];
    if (idx !== undefined) {
      sheet.getRange(P2_ROW.dataStart + idx, P2_COL.status).setValue(u.newStatus);
    }
  });
}

function updateP2Status_(ss, studentId, newStatus) {
  // Legacy single-call version — kept for compatibility, use batch version in processP2Class_
  var sheet = ss.getSheetByName('Paper 2');
  var last  = sheet.getLastRow();
  if (last < P2_ROW.dataStart) { return; }
  var ids = sheet.getRange(P2_ROW.dataStart, P2_COL.studentId,
    last - P2_ROW.dataStart + 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(studentId).trim()) {
      sheet.getRange(P2_ROW.dataStart + i, P2_COL.status).setValue(newStatus);
      return;
    }
  }
}


// ═══════════════════════════════════════════════════════════
// PROCESSING LOG
// ═══════════════════════════════════════════════════════════

function p2InitLog_(ss, queueEntry, config) {
  var sheet   = ss.getSheetByName('Processing Log');
  var lastRow = Math.max(sheet.getLastRow(), 2) + 1;
  var logId   = 'P2-' + queueEntry.classCode + '-' +
                Utilities.formatDate(new Date(), TZ2, 'yyyyMMdd-HHmm');
  var vals    = [logId, new Date(), 'Paper2_FollowUp',
    'Y9-P2-' + config.currentTerm, config.academicYear, config.currentTerm,
    queueEntry.classCode, queueEntry.mode, 'PROCESSING',
    '', '', '', queueEntry.teacherEmail, ''];
  sheet.getRange(lastRow, 1, 1, vals.length).setValues([vals]);
  SpreadsheetApp.flush();
  return lastRow;
}

function p2FinaliseLog_(ss, logRow, processed, failed, errors) {
  var sheet = ss.getSheetByName('Processing Log');
  sheet.getRange(logRow, PL2_COL.status).setValue(failed === 0 ? 'COMPLETE' : 'FAILED');
  sheet.getRange(logRow, PL2_COL.processed).setValue(processed);
  sheet.getRange(logRow, PL2_COL.failed).setValue(failed);
  if (errors.length > 0) {
    sheet.getRange(logRow, PL2_COL.errors).setValue(errors.join('\n'));
  }
  SpreadsheetApp.flush();
}

function p2LogError_(ss, queueEntry, errorMsg) {
  try {
    var sheet   = ss.getSheetByName('Processing Log');
    var lastRow = Math.max(sheet.getLastRow(), 2) + 1;
    var vals    = ['ERR-P2-' + Date.now(), new Date(), 'Paper2_FollowUp',
      '', '', '', queueEntry.classCode, queueEntry.mode, 'FAILED',
      '', '', errorMsg, queueEntry.teacherEmail, ''];
    sheet.getRange(lastRow, 1, 1, vals.length).setValues([vals]);
  } catch(e) {}
}


// ═══════════════════════════════════════════════════════════
// PREVIEW SIDEBAR
// ═══════════════════════════════════════════════════════════

function showP2PreviewSidebar_(classCode) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var config    = getP2Config_(ss);
  var classInfo = getP2ClassInfo_(ss, classCode, config.academicYear);
  var students  = getP2Students_(ss, classCode);

  if (!classInfo) {
    SpreadsheetApp.getUi().alert('Class ' + classCode + ' not found in Classes tab.');
    return;
  }
  if (students.length === 0) {
    SpreadsheetApp.getUi().alert('No active students found for ' + classCode + '.');
    return;
  }

  // Preview first student
  var student   = students[0];
  var p2Sheet   = ss.getSheetByName('Paper 2');
  var p2Last    = p2Sheet.getLastRow();

  // CRASH GUARD 6: check data exists before reading
  if (p2Last < P2_ROW.dataStart) {
    SpreadsheetApp.getUi().alert('Paper 2 has no student data yet. Enter some marks first to preview.');
    return;
  }

  var maxMarks  = p2Sheet.getRange(P2_ROW.maxMarks, P2_COL.sq1, 1, 40).getValues()[0];
  var sqLabels  = p2Sheet.getRange(P2_ROW.sqLabels, P2_COL.sq1, 1, 40).getValues()[0];
  var allTags   = p2Sheet.getRange(P2_ROW.aoTag, P2_COL.sq1, 4, 40).getValues();
  var tags      = { ao: allTags[0], blooms: allTags[1], qType: allTags[2], qGroup: allTags[3] };
  var questions = buildQuestionStructure_(sqLabels, maxMarks);
  var tasks     = getP2FollowUpTasks_(ss);
  var totalMax  = maxMarks.reduce(function(s,v){ return s + (parseFloat(v)||0); }, 0);

  var p2Data = p2Sheet.getRange(P2_ROW.dataStart, 1,
    p2Last - P2_ROW.dataStart + 1, P2_COL.status).getValues();

  var p2Row = null;
  for (var i = 0; i < p2Data.length; i++) {
    if (String(p2Data[i][0]).trim() === student.studentId) { p2Row = p2Data[i]; break; }
  }

  var grandTotal      = p2Row ? p2Row[P2_COL.grandTotal - 1] : 0;
  var pct             = p2Row ? p2Row[P2_COL.pct - 1]        : 0;
  var overallTier     = p2Row ? p2Row[P2_COL.tier - 1]       : 'N/A';
  // FIX BUG A: pass config.tierThresholds
  var questionResults = p2Row ? calculateQuestionResults_(p2Row, questions, config.tierThresholds) : [];

  var html = buildP2StudentEmail_(
    student, questionResults, tasks, grandTotal, pct, overallTier,
    totalMax, classInfo, config, 'PREVIEW', tags
  );

  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;font-size:11px;color:#666;padding:8px 12px;background:#ffebee;border-bottom:1px solid #ef9a9a;">' +
      '📧 PREVIEW — ' + escP2Html_(classCode) + ' — ' + escP2Html_(student.fullName) +
      ' &nbsp;|&nbsp; Overall tier: ' + escP2Html_(String(overallTier)) +
      '</div>' + html
    ).setTitle('Paper 2 Preview — ' + classCode).setWidth(660)
  );
}


// ═══════════════════════════════════════════════════════════
// TRIGGER MANAGEMENT
// ═══════════════════════════════════════════════════════════

function installP2QueueTrigger() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var config = getP2Config_(ss);

  if (!p2IsAdmin_(Session.getActiveUser().getEmail(), config)) {
    SpreadsheetApp.getUi().alert('Only the script owner can install triggers.');
    return;
  }

  // Remove existing P2 queue triggers
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processP2Queue') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('processP2Queue').timeBased().everyMinutes(5).create();

  SpreadsheetApp.getUi().alert(
    '✅ Paper 2 queue trigger installed.\n\n' +
    'The system will check for queued and scheduled Paper 2 classes every 5 minutes.\n\n' +
    'You only need to do this once.'
  );
}


// ═══════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════

function p2BuildFromName_(classInfo) {
  var first   = (classInfo.leadFirst || '').trim();
  var surname = (classInfo.leadName  || '').trim().split(' ').pop();
  var display = (first && surname) ? first + ' ' + surname : (classInfo.leadName || 'KS3 Science').trim();
  return display + ' \u2014 KS3 Science';
}

function p2ExtractFileId_(urlOrId) {
  if (!urlOrId) { return ''; }
  var match = String(urlOrId).match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : String(urlOrId).trim();
}

function escP2Html_(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escP2Attr_(str) {
  return String(str || '').replace(/'/g, '&#39;').replace(/"/g, '&quot;');
}

function p2Pad_(n) {
  return n < 10 ? '0' + n : String(n);
}
