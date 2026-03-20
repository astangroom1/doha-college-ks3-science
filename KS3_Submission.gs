/**
 * ============================================================
 *  KS3 SCIENCE — STUDENT SUBMISSION WEB APP
 *  Doha College | Andy Stangroom
 *  v1.1
 * ============================================================
 *
 *  DEPLOY ONCE:
 *   Apps Script → Deploy → New Deployment → Web App
 *   Execute as: Me (astangroom@dohacollege.com)
 *   Who has access: Anyone
 *   Copy deployment URL → paste into Year Controller row 36
 *
 *  FLOW:
 *   1. Student clicks link in email → doGet() serves submission page
 *   2. Student photographs work, uploads image
 *   3. Photo saved to Drive, submission recorded in Submissions sheet
 *   4. Teacher receives digest (07:30 + 13:00) with thumbnails
 *   5. Teacher clicks Approve/Reject → approval web app page
 *   6. Approve → model answers sent instantly
 *   7. Reject → reason + free text → new upload link sent to student
 *   8. Max 3 attempts → auto-escalate to Andy
 *
 *  v1.1 CHANGES (surgical only):
 *   • getSubConfig_            — adds bannerChaseUp (YC row 23); wraps banner IDs in subExtractFileId_
 *   • sendModelAnswers_        — complete rewrite: personalised per-student wrong questions
 *   • buildP1ModelAnswerSection_ — NEW helper for P1 model answers
 *   • buildP2ModelAnswerSection_ — NEW helper for P2 model answers
 *   • sendRejectionEmail_      — adds bannerChaseUp image banner
 *   • handleChaseUp_           — adds bannerChaseUp image banner to student emails
 *   • subExtractFileId_        — NEW utility (extracts Drive ID from full URL or bare ID)
 * ============================================================
 */

var SUB_TZ           = 'Asia/Qatar';
var SUB_MAX_ATTEMPTS = 3;
var SUB_TOKEN_DAYS   = 7;   // token expires after 7 days

// Submission sheet columns (1-based)
var SC = {
  token:          1,
  studentId:      2,
  studentName:    3,
  studentEmail:   4,
  classCode:      5,
  paperId:        6,
  academicYear:   7,
  term:           8,
  teacherEmail:   9,
  createdAt:      10,
  expiresAt:      11,
  submittedAt:    12,
  driveFileId:    13,
  driveFileUrl:   14,
  status:         15,   // Pending / Submitted / Approved / Rejected / Escalated / Expired
  attemptNumber:  16,
  reviewedAt:     17,
  rejectionReason:18,
  rejectionNote:  19,
  escalatedAt:    20,
};

var REJECTION_REASONS = [
  'Work not clearly visible in photo',
  'Incomplete — not all questions attempted',
  'Photo too blurry or dark to assess',
  'Wrong page / incorrect work submitted',
  'Work not in exercise book',
  'Answers copied — does not demonstrate understanding',
  'Other (see note below)',
];


// ═══════════════════════════════════════════════════════════
// WEB APP ENTRY POINT
// ═══════════════════════════════════════════════════════════

function doGet(e) {
  var params    = e ? e.parameter : {};
  var token     = params.token  || '';
  var action    = params.action || 'submit';
  var classCode = params['class'] || '';
  var isAdmin   = params.admin === '1';

  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var config = getSubConfig_(ss);

  // Chase-up: teacher sends emails to non-submitters for a class
  if (action === 'chaseup' && classCode) {
    return HtmlService.createHtmlOutput(handleChaseUp_(classCode, config))
      .setTitle('KS3 Science — Chase-Up').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Suppress / unsuppress: admin only
  if ((action === 'suppress' || action === 'unsuppress') && classCode && isAdmin) {
    return HtmlService.createHtmlOutput(handleSuppressAction_(classCode, action))
      .setTitle('KS3 Science — Admin').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Token-based actions require a valid token
  if (!token) {
    return HtmlService.createHtmlOutput(errorPage_('No submission token provided. Please use the link from your email.'))
      .setTitle('KS3 Science — Submission').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (action === 'approve') { return handleApprovalPage_(token, 'approve'); }
  if (action === 'reject')  { return handleApprovalPage_(token, 'reject');  }

  return serveSubmissionPage_(token);
}


// ═══════════════════════════════════════════════════════════
// STUDENT SUBMISSION PAGE
// ═══════════════════════════════════════════════════════════

function serveSubmissionPage_(token) {
  var sub = getSubmissionByToken_(token);

  if (!sub) {
    return HtmlService.createHtmlOutput(
      errorPage_('This submission link is not recognised. Please use the link from your email.'))
      .setTitle('KS3 Science — Submission');
  }

  if (sub.status === 'Approved') {
    return HtmlService.createHtmlOutput(
      infoPage_('Already Approved',
        'Your work for ' + sub.paperId + ' has already been approved. ' +
        'Check your email for your model answers.', '#2e7d32'))
      .setTitle('KS3 Science — Already Submitted');
  }

  var now       = new Date();
  var expiresAt = new Date(sub.expiresAt);
  if (now > expiresAt) {
    return HtmlService.createHtmlOutput(
      errorPage_('This submission link has expired (links are valid for ' + SUB_TOKEN_DAYS + ' days). ' +
        'Please speak to your teacher for a new link.'))
      .setTitle('KS3 Science — Link Expired');
  }

  if (parseInt(sub.attemptNumber) >= SUB_MAX_ATTEMPTS && sub.status === 'Escalated') {
    return HtmlService.createHtmlOutput(
      errorPage_('You have reached the maximum number of submission attempts. ' +
        'Your teacher has been notified and will speak to you directly.'))
      .setTitle('KS3 Science — Maximum Attempts Reached');
  }

  var attemptNum   = parseInt(sub.attemptNumber) || 0;
  var attemptsLeft = SUB_MAX_ATTEMPTS - attemptNum;
  var rejNote      = (sub.status === 'Rejected' && sub.rejectionReason)
    ? '<div style="background:#ffebee;border-radius:8px;padding:14px 18px;margin-bottom:20px;">' +
      '<p style="font-weight:bold;color:#b71c1c;margin:0 0 6px;font-size:14px;">Your previous submission was not accepted:</p>' +
      '<p style="color:#c62828;margin:0 0 6px;font-size:13px;">' + escSub_(sub.rejectionReason) + '</p>' +
      (sub.rejectionNote ? '<p style="color:#c62828;margin:0;font-size:13px;font-style:italic;">' + escSub_(sub.rejectionNote) + '</p>' : '') +
      '</div>'
    : '';

  var MAX_PHOTOS = 4;

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>KS3 Science \u2014 Submit Work</title>' +
    '<style>' +
    'body{font-family:Arial,sans-serif;margin:0;padding:0;background:#f5f5f5;color:#212121;}' +
    '.card{max-width:520px;margin:20px auto;background:#fff;border-radius:12px;box-shadow:0 2px 12px rgba(0,0,0,0.1);overflow:hidden;}' +
    '.hdr{background:#1a237e;padding:20px 24px;color:#fff;}' +
    '.hdr h1{margin:0;font-size:19px;}' +
    '.hdr p{margin:5px 0 0;font-size:12px;opacity:0.85;}' +
    '.body{padding:20px 24px;}' +
    '.info{background:#e8eaf6;border-radius:8px;padding:10px 14px;margin-bottom:16px;}' +
    '.info p{margin:2px 0;font-size:12px;color:#3949ab;}' +
    '.info strong{color:#1a237e;}' +
    '.instr{background:#f9fbe7;border-radius:8px;padding:12px 16px;margin-bottom:16px;}' +
    '.instr h3{color:#33691e;margin:0 0 6px;font-size:13px;}' +
    '.instr ol{margin:0;padding-left:16px;}' +
    '.instr li{font-size:12px;color:#558b2f;line-height:1.7;}' +
    '.photos-label{font-size:12px;font-weight:bold;color:#37474f;margin-bottom:8px;}' +
    '.photos-grid{display:flex;gap:8px;margin-bottom:12px;}' +
    '.slot{flex:1;aspect-ratio:1;border:2px dashed #c5cae9;border-radius:8px;position:relative;overflow:hidden;background:#fafafa;box-sizing:border-box;}' +
    '.slot label{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;cursor:pointer;font-size:22px;color:#c5cae9;}' +
    '.slot label:hover{background:#f0f2fc;}' +
    '.slot input[type=file]{display:none;}' +
    '.slot img{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;border-radius:6px;}' +
    '.slot .rm{position:absolute;top:3px;right:3px;z-index:5;background:rgba(183,28,28,0.9);color:#fff;border:none;border-radius:50%;width:22px;height:22px;font-size:13px;cursor:pointer;line-height:22px;text-align:center;padding:0;}' +
    '.photo-count{font-size:11px;color:#9e9e9e;text-align:center;margin-bottom:12px;}' +
    '.btn{display:block;width:100%;padding:14px;background:#1a237e;color:#fff;border:none;border-radius:8px;font-size:15px;font-weight:bold;cursor:pointer;box-sizing:border-box;}' +
    '.btn:disabled{background:#9e9e9e;cursor:not-allowed;}' +
    '.btn:hover:not(:disabled){background:#283593;}' +
    '.attempts{font-size:11px;color:#9e9e9e;text-align:center;margin-top:8px;}' +
    '.progress{width:100%;background:#e0e0e0;border-radius:4px;height:6px;margin-top:8px;}' +
    '.progress-bar{height:6px;background:#1a237e;border-radius:4px;transition:width 0.3s;}' +
    '.spinner,.success{display:none;text-align:center;padding:16px;}' +
    '.err{background:#ffebee;color:#c62828;border-radius:6px;padding:10px 14px;font-size:13px;margin-top:10px;display:none;}' +
    '</style></head><body>' +
    '<div class="card">' +
    '<div class="hdr"><h1>Submit Your Work</h1>' +
    '<p>KS3 Science \xb7 ' + escSub_(sub.paperId) + ' \xb7 ' + escSub_(sub.term) + ' \xb7 ' + escSub_(sub.academicYear) + '</p></div>' +
    '<div class="body">' +
    rejNote +
    '<div class="info">' +
    '<p><strong>Student:</strong> ' + escSub_(sub.studentName) + '</p>' +
    '<p><strong>Class:</strong> ' + escSub_(sub.classCode) + ' &nbsp;\xb7&nbsp; <strong>Paper:</strong> ' + escSub_(sub.paperId) + '</p>' +
    '</div>' +
    '<div class="instr"><h3>How to Submit</h3><ol>' +
    '<li>Open your exercise book to your completed follow-up tasks</li>' +
    '<li>Take a <strong>clear, well-lit photo</strong> of each page of work</li>' +
    '<li>Add up to <strong>4 photos</strong> if your work spans multiple pages</li>' +
    '<li>Tap <strong>Submit My Work</strong> and wait for confirmation</li>' +
    '</ol></div>' +
    '<div class="photos-label">Your Photos (tap a slot to add)</div>' +
    '<div class="photos-grid">' +
    '<div class="slot" id="s0"><label for="f0">\ud83d\udcf7</label><input type="file" id="f0" accept="image/*" onchange="addPhoto(this,0)"></div>' +
    '<div class="slot" id="s1" style="opacity:0.35"><label for="f1" id="l1" style="pointer-events:none">+</label><input type="file" id="f1" accept="image/*" onchange="addPhoto(this,1)"></div>' +
    '<div class="slot" id="s2" style="opacity:0.35"><label for="f2" id="l2" style="pointer-events:none">+</label><input type="file" id="f2" accept="image/*" onchange="addPhoto(this,2)"></div>' +
    '<div class="slot" id="s3" style="opacity:0.35"><label for="f3" id="l3" style="pointer-events:none">+</label><input type="file" id="f3" accept="image/*" onchange="addPhoto(this,3)"></div>' +
    '</div>' +
    '<p class="photo-count" id="pc">0 of 4 photos added</p>' +
    '<button class="btn" id="sub" onclick="go()" disabled>Submit My Work</button>' +
    '<p class="attempts">Submission ' + (attemptNum + 1) + ' of ' + SUB_MAX_ATTEMPTS + ' allowed</p>' +
    '<div class="err" id="err"></div>' +
    '<div class="spinner" id="spin"><p style="color:#666;font-size:14px" id="smsg">Uploading...</p><div class="progress"><div class="progress-bar" id="pb" style="width:0%"></div></div></div>' +
    '<div class="success" id="ok"><p style="font-size:36px">\u2705</p><p style="font-weight:bold;color:#2e7d32;font-size:16px">Work Submitted!</p><p style="font-size:13px;color:#555">Your teacher will review your work and send your model answers once approved.</p><p style="font-size:12px;color:#9e9e9e;margin-top:12px">You can now close this page.</p></div>' +
    '</div></div>' +
    '<script>' +
    'var P={};var TK="' + token + '";' +

    'function addPhoto(inp,n){' +
    '  if(!inp.files||!inp.files[0])return;' +
    '  var f=inp.files[0];' +
    '  var objUrl=URL.createObjectURL(f);' +
    '  P[n]={file:f,name:f.name};' +
    '  var s=document.getElementById("s"+n);' +
    '  s.innerHTML="<img src=\'"+objUrl+"\'><button class=\'rm\' onclick=\'rm("+n+")\'>&#x2715;</button>";' +
    '  s.style.opacity="1";' +
    '  upd();' +
    '}' +

    'function rm(n){' +
    '  delete P[n];' +
    '  var s=document.getElementById("s"+n);' +
    '  var ico=n===0?"\ud83d\udcf7":"+";' +
    '  s.innerHTML="<label for=\'f"+n+"\'>" + ico + "</label><input type=\'file\' id=\'f"+n+"\' accept=\'image/*\' onchange=\'addPhoto(this,"+n+")\'>"; ' +
    '  upd();' +
    '}' +

    'function upd(){' +
    '  var n=Object.keys(P).length;' +
    '  document.getElementById("pc").textContent=n+" of 4 photos added";' +
    '  document.getElementById("sub").disabled=n===0;' +
    '  for(var i=1;i<4;i++){' +
    '    var hasPrev=P[i-1]!=null;' +
    '    var lbl=document.getElementById("l"+i);' +
    '    var sl=document.getElementById("s"+i);' +
    '    if(lbl&&!P[i]){lbl.style.pointerEvents=hasPrev?"auto":"none";sl.style.opacity=hasPrev?"1":"0.35";}' +
    '  }' +
    '}' +

    'function go(){' +
    '  var keys=Object.keys(P).sort();' +
    '  if(!keys.length){showErr("Please add at least one photo.");return;}' +
    '  document.getElementById("sub").disabled=true;' +
    '  document.getElementById("sub").style.display="none";' +
    '  document.getElementById("spin").style.display="block";' +
    '  document.getElementById("smsg").textContent="Reading photos...";' +
    '  readAndSend(keys,0,[]);' +
    '}' +

    'function readAndSend(keys,i,items){' +
    '  if(i>=keys.length){' +
    '    next(items,0,items.length);' +
    '    return;' +
    '  }' +
    '  var k=parseInt(keys[i]);' +
    '  var f=P[k].file;' +
    '  document.getElementById("smsg").textContent="Processing photo "+(i+1)+" of "+keys.length+"...";' +
    '  var r=new FileReader();' +
    '  r.onload=function(e){' +
    '    var dataUrl=e.target.result;' +
    '    var base64=dataUrl.split(",")[1];' +
    '    var img=new Image();' +
    '    img.onload=function(){' +
    '      try{' +
    '        var mw=1800,mh=2400,w=img.width,h=img.height;' +
    '        if(w>mw){h=Math.round(h*mw/w);w=mw;}' +
    '        if(h>mh){w=Math.round(w*mh/h);h=mh;}' +
    '        var c=document.createElement("canvas");c.width=w;c.height=h;' +
    '        c.getContext("2d").drawImage(img,0,0,w,h);' +
    '        var compressed=c.toDataURL("image/jpeg",0.85).split(",")[1];' +
    '        items.push({n:k+1,d:compressed,f:P[k].name});' +
    '      }catch(ex){' +
    '        items.push({n:k+1,d:base64,f:P[k].name});' +
    '      }' +
    '      readAndSend(keys,i+1,items);' +
    '    };' +
    '    img.onerror=function(){' +
    '      items.push({n:k+1,d:base64,f:P[k].name});' +
    '      readAndSend(keys,i+1,items);' +
    '    };' +
    '    img.src=dataUrl;' +
    '  };' +
    '  r.onerror=function(){' +
    '    showErr("Could not read photo "+(i+1)+". Please try again.");' +
    '    fail();' +
    '  };' +
    '  r.readAsDataURL(f);' +
    '}' +

    'function next(items,i,total){' +
    '  if(i>=total){document.getElementById("spin").style.display="none";document.getElementById("ok").style.display="block";return;}' +
    '  document.getElementById("smsg").textContent="Uploading photo "+(i+1)+" of "+total+"...";' +
    '  document.getElementById("pb").style.width=Math.round(i/total*100)+"%";' +
    '  google.script.run' +
    '  .withSuccessHandler(function(r){if(r.ok){next(items,i+1,total);}else{showErr(r.error||"Upload failed.");fail();}})' +
    '  .withFailureHandler(function(e){showErr("Error: "+e.message);fail();})' +
    '  .handleStudentSubmission(TK,items[i].d,items[i].f,items[i].n,total);' +
    '}' +

    'function showErr(msg){document.getElementById("err").textContent=msg;document.getElementById("err").style.display="block";}' +
    'function fail(){document.getElementById("spin").style.display="none";document.getElementById("sub").style.display="block";document.getElementById("sub").disabled=false;}' +
    '<\/script></body></html>';

  return HtmlService.createHtmlOutput(html).setTitle('KS3 Science \u2014 Submit Work').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ═══════════════════════════════════════════════════════════
// TEACHER APPROVAL PAGE
// ═══════════════════════════════════════════════════════════

function handleApprovalPage_(token, action) {
  var sub = getSubmissionByToken_(token);
  if (!sub) {
    return HtmlService.createHtmlOutput(errorPage_('Submission not found.')).setTitle('KS3 Science');
  }

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>KS3 Science — Review Submission</title>' +
    '<style>' +
    'body{font-family:Arial,sans-serif;margin:0;padding:20px;background:#f5f5f5;color:#212121;}' +
    '.card{max-width:560px;margin:0 auto;background:#fff;border-radius:12px;box-shadow:0 2px 12px rgba(0,0,0,0.1);overflow:hidden;}' +
    '.hdr{background:' + (action==='approve'?'#1b5e20':'#b71c1c') + ';padding:20px 24px;color:#fff;}' +
    '.hdr h1{margin:0;font-size:18px;}' +
    '.body{padding:20px 24px;}' +
    '.info{background:#f5f5f5;border-radius:8px;padding:12px 16px;margin-bottom:16px;font-size:13px;}' +
    '.photo{margin-bottom:16px;text-align:center;}' +
    '.photo img{max-width:100%;border-radius:8px;border:1px solid #e0e0e0;}' +
    'select,textarea{width:100%;padding:8px 10px;border:1px solid #ddd;border-radius:6px;font-size:13px;box-sizing:border-box;font-family:Arial,sans-serif;margin-bottom:12px;}' +
    '.btn{display:block;width:100%;padding:13px;border:none;border-radius:8px;font-size:15px;font-weight:bold;cursor:pointer;box-sizing:border-box;margin-bottom:8px;}' +
    '.btn-approve{background:#1b5e20;color:#fff;}' +
    '.btn-reject{background:#b71c1c;color:#fff;}' +
    '.btn:disabled{background:#9e9e9e;}' +
    '.result{display:none;padding:16px;border-radius:8px;text-align:center;font-weight:bold;font-size:15px;}' +
    '.ok{background:#e8f5e9;color:#2e7d32;}.err{background:#ffebee;color:#c62828;}' +
    '</style></head><body>' +
    '<div class="card">' +
    '<div class="hdr"><h1>' + (action==='approve' ? '✅ Approve Submission' : '✗ Reject Submission') + '</h1>' +
    '<p style="margin:4px 0 0;font-size:12px;opacity:0.85;">' + escSub_(sub.studentName) + ' · ' + escSub_(sub.classCode) + ' · ' + escSub_(sub.paperId) + '</p></div>' +
    '<div class="body">' +
    '<div class="info">' +
    '<strong>' + escSub_(sub.studentName) + '</strong> · ' + escSub_(sub.classCode) + '<br>' +
    'Submitted: ' + (sub.submittedAt ? new Date(sub.submittedAt).toLocaleString() : 'Unknown') + '<br>' +
    'Attempt: ' + escSub_(sub.attemptNumber) + ' of ' + SUB_MAX_ATTEMPTS +
    '</div>' +
    (sub.driveFileUrl ? '<div class="photo"><img src="' + sub.driveFileUrl + '" alt="Student submission"><p style="font-size:11px;color:#9e9e9e;margin-top:6px;"><a href="' + sub.driveFileUrl + '" target="_blank">Open full image</a></p></div>' : '') +
    (action === 'reject'
      ? '<label style="font-size:12px;font-weight:bold;color:#555;display:block;margin-bottom:4px;">Reason for rejection:</label>' +
        '<select id="reason"><option value="">Select a reason...</option>' +
        REJECTION_REASONS.map(function(r){ return '<option>' + r + '</option>'; }).join('') +
        '</select>' +
        '<label style="font-size:12px;font-weight:bold;color:#555;display:block;margin-bottom:4px;">Additional note (optional):</label>' +
        '<textarea id="note" rows="3" placeholder="Add any extra guidance for the student..."></textarea>'
      : '') +
    '<button class="btn ' + (action==='approve'?'btn-approve':'btn-reject') + '" id="actionBtn" onclick="doAction()">' +
    (action==='approve' ? 'Confirm Approval & Send Model Answers' : 'Reject & Send Feedback to Student') +
    '</button>' +
    '<div class="result" id="result"></div>' +
    '</div></div>' +
    '<script>' +
    'function doAction(){' +
    '  document.getElementById("actionBtn").disabled=true;' +
    (action === 'reject'
      ? '  var reason=document.getElementById("reason").value;' +
        '  var note=document.getElementById("note").value;' +
        '  if(!reason){document.getElementById("actionBtn").disabled=false;alert("Please select a rejection reason.");return;}' +
        '  google.script.run.withSuccessHandler(cb).withFailureHandler(err).processApproval("' + token + '","reject",reason,note);'
      : '  google.script.run.withSuccessHandler(cb).withFailureHandler(err).processApproval("' + token + '","approve","","");') +
    '}' +
    'function cb(r){var el=document.getElementById("result");el.className="result "+(r.ok?"ok":"err");el.textContent=r.message||"Done";el.style.display="block";}' +
    'function err(e){var el=document.getElementById("result");el.className="result err";el.textContent="Error: "+e.message;el.style.display="block";document.getElementById("actionBtn").disabled=false;}' +
    '</script></body></html>';

  return HtmlService.createHtmlOutput(html).setTitle('KS3 Science — Review').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ═══════════════════════════════════════════════════════════
// SERVER FUNCTIONS — called from web app pages
// ═══════════════════════════════════════════════════════════

/**
 * Called when student submits photo.
 * base64Data = base64-encoded image file.
 */
function handleStudentSubmission(token, base64Data, fileName, photoIndex, totalPhotos) {
  try {
    var sub = getSubmissionByToken_(token);
    if (!sub) { return { ok: false, error: 'Invalid submission token.' }; }

    var now = new Date();
    if (now > new Date(sub.expiresAt)) {
      return { ok: false, error: 'This link has expired. Please speak to your teacher.' };
    }
    if (sub.status === 'Approved') {
      return { ok: false, error: 'This submission has already been approved.' };
    }

    photoIndex   = photoIndex   || 1;
    totalPhotos  = totalPhotos  || 1;

    // Only increment attempt number on first photo of a submission
    var attemptNum = photoIndex === 1
      ? (parseInt(sub.attemptNumber) || 0) + 1
      : (parseInt(sub.attemptNumber) || 1);

    if (photoIndex === 1 && attemptNum > SUB_MAX_ATTEMPTS) {
      return { ok: false, error: 'Maximum submissions reached. Your teacher has been notified.' };
    }

    // Save photo to Drive
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var config   = getSubConfig_(ss);

    // Extract folder ID from full URL if pasted (e.g. https://drive.google.com/drive/folders/XXXX)
    var rawId    = config.masterFolderId || '';
    var folderId = rawId.replace(/.*\/folders\/([^\/\?]+).*/, '$1').replace(/.*\/d\/([^\/\?]+).*/, '$1').trim();

    var termFolder;
    try {
      var masterFolder   = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
      var feedbackFolder = getOrCreateFolder_(masterFolder, 'KS3 Science \u2014 Student Feedback');
      var studentFolderName = sub.studentName + ' \u2014 ' + sub.studentId;
      var studentFolder  = getOrCreateFolder_(feedbackFolder, studentFolderName);
      var ayShort        = sub.academicYear.replace(/(\d{4})-(\d{2})(\d{2})/, '$1-$3') || sub.academicYear;
      var yearGroup      = config.yearGroup || 'Year 9';
      var yearFolder     = getOrCreateFolder_(studentFolder, ayShort + ' (' + yearGroup + ')');
      termFolder         = getOrCreateFolder_(yearFolder, sub.term || config.currentTerm || 'T1');
    } catch(driveErr) {
      // If folder hierarchy fails, fall back to a flat KS3 submissions folder in root
      Logger.log('KS3 Drive folder error (using fallback): ' + driveErr.toString());
      try {
        var rootFallback = DriveApp.getRootFolder();
        termFolder = getOrCreateFolder_(rootFallback, 'KS3 Science \u2014 Submissions (Fallback)');
      } catch(e2) {
        return { ok: false, error: 'Drive access error: ' + driveErr.toString() };
      }
    }

    // File name
    var dateStamp = Utilities.formatDate(now, SUB_TZ, 'ddMMyy');
    var ext       = (fileName && fileName.indexOf('.') > -1) ? fileName.split('.').pop().toLowerCase() : 'jpg';
    var mimeType  = (ext === 'png') ? 'image/png' : 'image/jpeg';
    var safeId    = sub.paperId.replace(/[^a-zA-Z0-9_-]/g, '_');
    var blobName  = safeId + '_Submission_attempt' + attemptNum +
                    '_photo' + photoIndex + 'of' + totalPhotos +
                    '_' + dateStamp + '.jpg';  // always save as jpg (compressed)

    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', blobName);
    var file = termFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId  = file.getId();
    var fileUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;

    // Update submission row — on first photo set status to Submitted,
    // on subsequent photos just add the additional file URL
    if (photoIndex === 1) {
      updateSubmissionRow_(token, {
        submittedAt:     now.toISOString(),
        driveFileId:     fileId,
        driveFileUrl:    fileUrl,
        status:          'Submitted',
        attemptNumber:   attemptNum,
        rejectionReason: '',
        rejectionNote:   '',
      });
    }
    // Additional photos: append URL to existing (teachers can open Drive folder for all)

    // If all photos sent and this is max attempts, escalate
    if (photoIndex === totalPhotos && attemptNum >= SUB_MAX_ATTEMPTS) {
      escalateToAndy_(SpreadsheetApp.getActiveSpreadsheet(), sub, config, attemptNum);
    }

    return { ok: true };

  } catch(e) {
    Logger.log('handleStudentSubmission error: ' + e.toString());
    return { ok: false, error: e.toString() };
  }
}

/**
 * Called from approval web app page.
 * action = 'approve' or 'reject'
 */
function processApproval(token, action, reason, note) {
  try {
    var sub = getSubmissionByToken_(token);
    if (!sub) { return { ok: false, message: 'Submission not found.' }; }
    if (sub.status === 'Approved') { return { ok: false, message: 'Already approved.' }; }

    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var config = getSubConfig_(ss);
    var now    = new Date();

    if (action === 'approve') {
      updateSubmissionRow_(token, {
        status:     'Approved',
        reviewedAt: now.toISOString(),
      });
      sendModelAnswers_(ss, sub, config);
      return { ok: true, message: '✅ Approved! Model answers sent to ' + sub.studentEmail + '.' };

    } else {
      var nextAttempt = parseInt(sub.attemptNumber) || 0;
      var hasMore     = nextAttempt < SUB_MAX_ATTEMPTS;

      updateSubmissionRow_(token, {
        status:          hasMore ? 'Rejected' : 'Escalated',
        reviewedAt:      now.toISOString(),
        rejectionReason: reason,
        rejectionNote:   note || '',
      });

      if (hasMore) {
        sendRejectionEmail_(ss, sub, config, reason, note);
        return { ok: true, message: 'Rejection sent. Student can resubmit.' };
      } else {
        escalateToAndy_(ss, sub, config, nextAttempt);
        return { ok: true, message: 'Max attempts reached. Andy has been notified.' };
      }
    }
  } catch(e) {
    Logger.log('processApproval error: ' + e.toString());
    return { ok: false, message: 'Error: ' + e.toString() };
  }
}


// ═══════════════════════════════════════════════════════════
// MODEL ANSWERS & REJECTION EMAILS
// ═══════════════════════════════════════════════════════════

/**
 * Send model answers to a student after teacher approval.
 * v1.1: Reads the student's specific wrong questions from Paper 1/2 sheet
 * and delivers personalised model answers from the Follow-Up Tasks sheet.
 */
function sendModelAnswers_(ss, sub, config) {
  var bannerUrl  = config.bannerModelAns
    ? 'https://drive.google.com/uc?export=view&id=' + config.bannerModelAns : '';
  var bannerHtml = bannerUrl
    ? '<img src="' + bannerUrl + '" alt="KS3 Science" ' +
      'style="display:block;width:100%;max-width:640px;height:auto;border:0;" width="640">'
    : '';

  var firstName = sub.studentName ? sub.studentName.split(' ')[0] : 'Student';
  var isPaper1  = sub.paperId.indexOf('-P1-') !== -1;

  var answerSection = '';
  if (isPaper1) {
    answerSection = buildP1ModelAnswerSection_(ss, sub);
  } else {
    answerSection = buildP2ModelAnswerSection_(ss, sub);
  }

  // Graceful fallback if sheet data unavailable
  if (!answerSection) {
    answerSection =
      '<div style="background:#e8eaf6;border-radius:8px;padding:16px 18px;margin-bottom:18px;text-align:center;">' +
      '<p style="font-size:13px;color:#3949ab;margin:0;font-weight:bold;">Your teacher will go through the model answers with you in the next lesson.</p>' +
      '</div>';
  }

  var html =
    '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
    '<body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif;">' +
    '<div style="max-width:640px;margin:0 auto;background:#fff;">' +
    bannerHtml +
    '<div style="background:#f5f5f5;padding:7px 30px;border-bottom:1px solid #e0e0e0;">' +
    '<p style="margin:0;font-size:10px;color:#9e9e9e;">KS3 Science \xb7 Doha College</p>' +
    '</div>' +
    '<div style="padding:22px 28px;">' +
    '<p style="font-size:15px;color:#212121;margin:0 0 4px;">Dear ' + escSub_(firstName) + ',</p>' +
    '<p style="font-size:13px;color:#666;margin:0 0 18px;">Your teacher has reviewed your follow-up submission for ' +
    '<strong>' + escSub_(sub.paperId) + '</strong> and approved it. ' +
    'Your personalised model answers are below.</p>' +
    '<div style="background:#e8f5e9;border-radius:8px;padding:14px 18px;margin-bottom:18px;text-align:center;">' +
    '<div style="font-size:28px;margin-bottom:4px;">&#x2705;</div>' +
    '<p style="font-size:14px;font-weight:bold;color:#1b5e20;margin:0;">Submission Approved</p>' +
    '</div>' +
    answerSection +
    '<div style="margin-top:20px;padding-top:16px;border-top:1px solid #e8eaf6;">' +
    '<p style="font-size:13px;color:#424242;">Use these model answers to check your work and identify any remaining gaps. ' +
    'If you have any questions, speak to your teacher in class.</p>' +
    '</div></div>' +
    '<div style="background:#e8eaf6;padding:8px 28px;text-align:center;">' +
    '<p style="font-size:10px;color:#9e9e9e;margin:0;">Automated message from KS3 Science \xb7 Doha College.</p>' +
    '</div></div></body></html>';

  MailApp.sendEmail({
    to:       sub.studentEmail,
    subject:  'Your Model Answers \u2014 ' + sub.paperId + ' Follow-Up',
    htmlBody: html,
    name:     'KS3 Science \u2014 Doha College',
  });
}


/**
 * Build personalised P1 model answer section.
 * Looks up the student's actual wrong questions then fetches the correct
 * answer and explanation from Follow-Up Tasks (col P=correctAns, col Q=explanation).
 */
function buildP1ModelAnswerSection_(ss, sub) {
  try {
    var p1Sheet = ss.getSheetByName('Paper 1');
    if (!p1Sheet || p1Sheet.getLastRow() < 8) { return ''; }

    // Find student row in Paper 1 (data starts row 8, studentId in col A)
    var p1LastRow  = p1Sheet.getLastRow();
    var p1Data     = p1Sheet.getRange(8, 1, p1LastRow - 7, 46).getValues();
    var studentRow = null;
    for (var i = 0; i < p1Data.length; i++) {
      var rowSid = String(p1Data[i][0]).trim();
      if (rowSid && rowSid === String(sub.studentId).trim()) { studentRow = p1Data[i]; break; }
    }
    if (!studentRow) { return ''; }

    // Read answer key (row 1, cols D onwards = index 3+) and tag rows (rows 2-4)
    var keyRow  = p1Sheet.getRange(1, 4, 1, 40).getValues()[0];
    var tagRows = p1Sheet.getRange(2, 4, 3, 40).getValues();
    var aoTags    = tagRows[0];   // row 2
    var discTags  = tagRows[1];   // row 3
    var topicTags = tagRows[2];   // row 4

    // Identify wrong questions (student answers in cols D onwards = row index 3+)
    var wrongQs = [];
    for (var q = 1; q <= 40; q++) {
      var correctAns = String(keyRow[q - 1]).trim().toUpperCase();
      if (!correctAns) { continue; }
      var studentAns = String(studentRow[3 + q - 1]).trim().toUpperCase();
      if (!studentAns || studentAns === '-') { continue; }
      if (studentAns !== correctAns) {
        wrongQs.push({
          number:        q,
          correctAnswer: correctAns,
          studentAnswer: studentAns,
          topic:         String(topicTags[q - 1]).trim(),
          ao:            String(aoTags[q - 1]).trim(),
          discipline:    String(discTags[q - 1]).trim(),
        });
      }
    }
    if (wrongQs.length === 0) {
      return '<div style="background:#e8f5e9;border-radius:8px;padding:14px 18px;margin-bottom:18px;text-align:center;">' +
        '<p style="font-size:13px;font-weight:bold;color:#1b5e20;margin:0;">All questions were correct \u2014 nothing to review!</p>' +
        '</div>';
    }

    // Read Follow-Up Tasks sheet
    // Col C (index 2) = taskType, Col D (index 3) = questionNum
    // Col P (index 15) = correctAns, Col Q (index 16) = explanation
    var tasksSheet = ss.getSheetByName('Follow-Up Tasks');
    var taskLookup = {};   // questionNum (string) → {correctAns, explanation}

    if (tasksSheet && tasksSheet.getLastRow() >= 3) {
      var taskData = tasksSheet.getRange(3, 1, tasksSheet.getLastRow() - 2, 17).getValues();
      taskData.forEach(function(row) {
        var taskType = String(row[2]).trim();    // col C
        var qNum     = String(row[3]).trim();    // col D
        var correct  = String(row[15]).trim();   // col P
        var expl     = String(row[16]).trim();   // col Q
        if (taskType !== 'P1-Variant' || !qNum) { return; }
        if (!taskLookup[qNum]) {
          taskLookup[qNum] = { correctAns: correct, explanation: expl };
        }
      });
    }

    // Build answer blocks — one per wrong question
    var discColours = { 'Biology': '#2e7d32', 'Chemistry': '#1565c0', 'Physics': '#6a1b9a' };

    var blocks = wrongQs.map(function(q) {
      var task  = taskLookup[String(q.number)];
      var discC = discColours[q.discipline] || '#37474f';

      var tagHtml =
        (q.discipline ? '<span style="background:' + discC + ';color:#fff;border-radius:3px;padding:2px 6px;font-size:10px;margin-right:4px;">' + escSub_(q.discipline) + '</span>' : '') +
        (q.ao ? '<span style="background:#e8eaf6;color:#3949ab;border-radius:3px;padding:2px 6px;font-size:10px;">' + escSub_(q.ao) + '</span>' : '');

      var explanationHtml = (task && task.explanation)
        ? '<div style="background:#f9fbe7;border-radius:6px;padding:10px 12px;margin-top:8px;">' +
          '<div style="font-size:10px;color:#558b2f;font-weight:bold;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:4px;">Explanation</div>' +
          '<p style="font-size:12px;color:#33691e;margin:0;line-height:1.5;">' + escSub_(task.explanation) + '</p>' +
          '</div>'
        : '<div style="background:#f5f5f5;border-radius:6px;padding:8px 12px;margin-top:8px;">' +
          '<p style="font-size:12px;color:#9e9e9e;margin:0;">Your teacher will explain this concept in class.</p>' +
          '</div>';

      return '<div style="border:1px solid #e0e0e0;border-radius:8px;padding:12px 14px;margin-bottom:10px;">' +
        '<table style="width:100%;border-collapse:collapse;"><tr>' +
        '<td style="width:32px;vertical-align:top;padding-top:2px;">' +
        '<div style="background:#1a237e;color:#fff;border-radius:50%;width:28px;height:28px;text-align:center;line-height:28px;font-size:11px;font-weight:bold;">Q' + q.number + '</div>' +
        '</td>' +
        '<td style="vertical-align:top;padding-left:10px;">' +
        '<div style="font-size:13px;font-weight:bold;color:#212121;margin-bottom:3px;">' + escSub_(q.topic || 'Topic not yet tagged') + '</div>' +
        '<div>' + tagHtml + '</div>' +
        '</td>' +
        '<td style="vertical-align:top;text-align:right;padding-left:8px;white-space:nowrap;">' +
        '<div><span style="font-size:11px;color:#b71c1c;background:#ffebee;border-radius:3px;padding:2px 6px;display:inline-block;margin-bottom:2px;">You: ' + escSub_(q.studentAnswer) + '</span></div>' +
        '<div><span style="font-size:11px;color:#1b5e20;background:#e8f5e9;border-radius:3px;padding:2px 6px;display:inline-block;">Correct: ' + escSub_(q.correctAnswer) + '</span></div>' +
        '</td>' +
        '</tr></table>' +
        explanationHtml +
        '</div>';
    }).join('');

    return '<div style="margin-bottom:18px;">' +
      '<h3 style="font-size:14px;color:#1a237e;margin:0 0 12px;padding-bottom:6px;border-bottom:2px solid #e8eaf6;">' +
      'Model Answers \u2014 ' + wrongQs.length + ' question' + (wrongQs.length !== 1 ? 's' : '') + ' to review' +
      '</h3>' +
      blocks +
      '</div>';

  } catch(e) {
    Logger.log('buildP1ModelAnswerSection_ error: ' + e.toString());
    return '';
  }
}


/**
 * Build personalised P2 model answer section.
 * Identifies questions where marks were dropped, fetches task title,
 * description and model answer from Follow-Up Tasks
 * (col R=taskTitle, col S=taskDesc, col V=modelAnswer).
 */
function buildP2ModelAnswerSection_(ss, sub) {
  try {
    var p2Sheet = ss.getSheetByName('Paper 2');
    if (!p2Sheet || p2Sheet.getLastRow() < 8) { return ''; }

    // Find student row in Paper 2
    var p2LastRow  = p2Sheet.getLastRow();
    var p2Data     = p2Sheet.getRange(8, 1, p2LastRow - 7, 47).getValues();
    var studentRow = null;
    for (var i = 0; i < p2Data.length; i++) {
      if (String(p2Data[i][0]).trim() === String(sub.studentId).trim()) { studentRow = p2Data[i]; break; }
    }
    if (!studentRow) { return ''; }

    // Read max marks (row 1) and sub-question labels (row 6), starting col D (index 3)
    var maxMarksRow = p2Sheet.getRange(1, 4, 1, 40).getValues()[0];
    var sqLabels    = p2Sheet.getRange(6, 4, 1, 40).getValues()[0];

    // Build question structure (mirrors buildQuestionStructure_ in Paper2.gs)
    var qMap = {};
    var qOrder = [];
    for (var j = 0; j < sqLabels.length; j++) {
      var label   = String(sqLabels[j]).trim();
      var maxMark = parseFloat(maxMarksRow[j]) || 0;
      if (!label || label === 'SQ' + (j + 1)) { continue; }
      var parentMatch = label.replace(/^[Qq]/, '').match(/^(\d+)/);
      if (!parentMatch) { continue; }
      var pn = parentMatch[1];
      if (!qMap[pn]) { qMap[pn] = { parentNum: pn, subQuestions: [], totalMax: 0 }; qOrder.push(pn); }
      qMap[pn].subQuestions.push({ sqIndex: j, maxMark: maxMark });
      qMap[pn].totalMax += maxMark;
    }

    // Read tier thresholds from Year Controller (rows 28, 30, 32)
    var ycSheet    = ss.getSheetByName('Year Controller');
    var extendMin  = 90;
    var refineMin  = 65;
    var masteryPct = 100;
    if (ycSheet) {
      var ycData = ycSheet.getRange(1, 2, 35, 1).getValues();
      extendMin  = parseInt(String(ycData[29][0]).trim()) || 90;
      refineMin  = parseInt(String(ycData[27][0]).trim()) || 65;
      masteryPct = parseInt(String(ycData[31][0]).trim()) || 100;
    }

    // Calculate earned marks per question; collect those with dropped marks
    var droppedQs = [];
    qOrder.forEach(function(pn) {
      var q = qMap[pn];
      var earned = 0;
      q.subQuestions.forEach(function(sq) {
        earned += parseFloat(studentRow[3 + sq.sqIndex]) || 0;
      });
      earned = Math.min(earned, q.totalMax);  // clamp to max
      var dropped = q.totalMax - earned;
      if (dropped > 0 && q.totalMax > 0) {
        var pct  = Math.round((earned / q.totalMax) * 100);
        var tier = pct >= masteryPct ? 'Mastery' : pct >= extendMin ? 'Extend' : pct >= refineMin ? 'Refine' : 'Rebuild';
        droppedQs.push({ parentNum: pn, earned: earned, totalMax: q.totalMax, dropped: dropped, pct: pct, tier: tier });
      }
    });

    if (droppedQs.length === 0) {
      return '<div style="background:#e8f5e9;border-radius:8px;padding:14px 18px;margin-bottom:18px;text-align:center;">' +
        '<p style="font-size:13px;font-weight:bold;color:#1b5e20;margin:0;">Full marks across all questions \u2014 outstanding work!</p>' +
        '</div>';
    }

    // Read Follow-Up Tasks sheet for P2
    // Col C (index 2) = taskType, Col D (index 3) = questionNum
    // Col R (index 17) = taskTitle, Col S (index 18) = taskDesc, Col V (index 21) = modelAnswer
    var tasksSheet = ss.getSheetByName('Follow-Up Tasks');
    var taskLookup = {};   // 'P2-Tier-QNum' → {taskTitle, taskDesc, modelAnswer}

    if (tasksSheet && tasksSheet.getLastRow() >= 3) {
      var taskData = tasksSheet.getRange(3, 1, tasksSheet.getLastRow() - 2, 22).getValues();
      taskData.forEach(function(row) {
        var taskType  = String(row[2]).trim();    // col C
        var qNum      = String(row[3]).trim();    // col D
        var taskTitle = String(row[17]).trim();   // col R
        var taskDesc  = String(row[18]).trim();   // col S
        var modelAns  = String(row[21]).trim();   // col V
        if (!taskType.match(/^P2-(Rebuild|Refine|Extend|Mastery)$/) || !qNum) { return; }
        var key = taskType + '-' + qNum;
        if (!taskLookup[key]) {
          taskLookup[key] = { taskTitle: taskTitle, taskDesc: taskDesc, modelAnswer: modelAns };
        }
      });
    }

    // Build answer blocks
    var tierColours = { Mastery: '#1b5e20', Extend: '#2e7d32', Refine: '#e65100', Rebuild: '#b71c1c' };
    var tierBg      = { Mastery: '#e8f5e9', Extend: '#c8e6c9', Refine: '#fff3e0', Rebuild: '#ffebee' };

    var blocks = droppedQs.map(function(q) {
      var tc   = tierColours[q.tier] || '#37474f';
      var tb   = tierBg[q.tier]     || '#f5f5f5';
      var task = taskLookup['P2-' + q.tier + '-' + q.parentNum];

      var taskHtml = '';
      if (task && task.taskTitle) {
        taskHtml =
          '<div style="background:#f9fbe7;border-radius:6px;padding:10px 12px;margin-top:8px;">' +
          '<div style="font-size:10px;color:#558b2f;font-weight:bold;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">Model Answer \u2014 ' + escSub_(q.tier) + ' Task</div>' +
          (task.taskTitle ? '<div style="font-size:13px;font-weight:bold;color:#212121;margin-bottom:6px;">' + escSub_(task.taskTitle) + '</div>' : '') +
          (task.taskDesc  ? '<p style="font-size:12px;color:#424242;margin:0 0 6px;line-height:1.5;">' + escSub_(task.taskDesc) + '</p>' : '') +
          (task.modelAnswer
            ? '<div style="background:#fff;border-radius:4px;padding:8px 10px;border-left:3px solid #c6e03a;">' +
              '<div style="font-size:10px;color:#558b2f;font-weight:bold;margin-bottom:3px;">Answer</div>' +
              '<p style="font-size:12px;color:#33691e;margin:0;line-height:1.5;">' + escSub_(task.modelAnswer) + '</p>' +
              '</div>'
            : '') +
          '</div>';
      } else {
        taskHtml =
          '<div style="background:#f5f5f5;border-radius:6px;padding:8px 12px;margin-top:8px;">' +
          '<p style="font-size:12px;color:#9e9e9e;margin:0;">Your teacher will provide feedback for Question ' + escSub_(q.parentNum) + ' in class.</p>' +
          '</div>';
      }

      return '<div style="border:1px solid #e0e0e0;border-radius:8px;padding:12px 14px;margin-bottom:10px;">' +
        '<table style="width:100%;border-collapse:collapse;"><tr>' +
        '<td style="width:32px;vertical-align:top;padding-top:2px;">' +
        '<div style="background:#b71c1c;color:#fff;border-radius:50%;width:28px;height:28px;text-align:center;line-height:28px;font-size:11px;font-weight:bold;">Q' + escSub_(q.parentNum) + '</div>' +
        '</td>' +
        '<td style="vertical-align:top;padding-left:10px;">' +
        '<span style="background:' + tc + ';color:#fff;border-radius:3px;padding:2px 8px;font-size:11px;font-weight:bold;">' + escSub_(q.tier) + '</span>' +
        '</td>' +
        '<td style="vertical-align:top;text-align:right;white-space:nowrap;">' +
        '<span style="font-size:12px;color:#424242;">' + q.earned + '/' + q.totalMax + ' (' + q.pct + '%)</span>' +
        '<br><span style="font-size:11px;color:' + tc + ';">' + q.dropped + ' mark' + (q.dropped !== 1 ? 's' : '') + ' dropped</span>' +
        '</td>' +
        '</tr></table>' +
        taskHtml +
        '</div>';
    }).join('');

    return '<div style="margin-bottom:18px;">' +
      '<h3 style="font-size:14px;color:#b71c1c;margin:0 0 12px;padding-bottom:6px;border-bottom:2px solid #ffebee;">' +
      'Model Answers \u2014 ' + droppedQs.length + ' question' + (droppedQs.length !== 1 ? 's' : '') + ' to review' +
      '</h3>' +
      blocks +
      '</div>';

  } catch(e) {
    Logger.log('buildP2ModelAnswerSection_ error: ' + e.toString());
    return '';
  }
}


/**
 * Send rejection feedback email to student.
 * v1.1: adds bannerChaseUp image banner at top.
 */
function sendRejectionEmail_(ss, sub, config, reason, note) {
  var webAppUrl    = config.webAppUrl || '';
  var attemptsDone = parseInt(sub.attemptNumber) || 0;
  var attemptsLeft = SUB_MAX_ATTEMPTS - attemptsDone;
  var resubmitLink = (webAppUrl && attemptsLeft > 0)
    ? '<a href="' + webAppUrl + '?token=' + sub.token + '" ' +
      'style="display:inline-block;background:#1a237e;color:#fff;padding:12px 24px;border-radius:8px;' +
      'font-size:14px;font-weight:bold;text-decoration:none;margin-top:10px;">Resubmit My Work</a>'
    : '';

  var firstName = sub.studentName ? sub.studentName.split(' ')[0] : 'Student';

  // Banner — bannerChaseUp (row 23 in Year Controller)
  var bannerUrl  = config.bannerChaseUp
    ? 'https://drive.google.com/uc?export=view&id=' + config.bannerChaseUp : '';
  var bannerHtml = bannerUrl
    ? '<img src="' + bannerUrl + '" alt="KS3 Science" ' +
      'style="display:block;width:100%;max-width:640px;height:auto;border:0;" width="640">'
    : '';

  var html =
    '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
    '<body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif;">' +
    '<div style="max-width:640px;margin:0 auto;background:#fff;">' +
    bannerHtml +
    '<div style="background:#f5f5f5;padding:7px 30px;border-bottom:1px solid #e0e0e0;">' +
    '<p style="margin:0;font-size:10px;color:#9e9e9e;">KS3 Science &middot; Doha College</p>' +
    '</div>' +
    '<div style="padding:22px 28px;">' +
    '<p style="font-size:15px;color:#212121;margin:0 0 4px;">Dear ' + escSub_(firstName) + ',</p>' +
    '<p style="font-size:13px;color:#666;margin:0 0 18px;">Your submission for <strong>' + escSub_(sub.paperId) +
    '</strong> has been reviewed but could not be accepted this time.</p>' +
    '<div style="background:#ffebee;border-radius:8px;padding:14px 18px;margin-bottom:18px;">' +
    '<h3 style="margin:0 0 6px;font-size:13px;color:#b71c1c;font-weight:bold;">Reason</h3>' +
    '<p style="font-size:13px;color:#c62828;margin:0;">' + escSub_(reason) + '</p>' +
    (note ? '<p style="font-size:13px;color:#c62828;margin:8px 0 0;font-style:italic;">' + escSub_(note) + '</p>' : '') +
    '</div>' +
    (attemptsLeft > 0
      ? '<div style="background:#e8eaf6;border-radius:8px;padding:14px 18px;margin-bottom:18px;text-align:center;">' +
        '<p style="font-size:13px;color:#3949ab;margin:0 0 10px;">You have <strong>' + attemptsLeft +
        ' attempt' + (attemptsLeft > 1 ? 's' : '') + '</strong> remaining. ' +
        'Please correct the issue above and resubmit using the button below.</p>' +
        resubmitLink +
        '</div>'
      : '<div style="background:#fff3e0;border-radius:8px;padding:14px 18px;margin-bottom:18px;">' +
        '<p style="font-size:13px;color:#e65100;margin:0;">You have used all your submission attempts. ' +
        'Your teacher has been notified and will speak to you directly.</p>' +
        '</div>') +
    '<div style="margin-top:20px;padding-top:16px;border-top:1px solid #e8eaf6;">' +
    '<p style="font-size:13px;color:#424242;">If you have any questions, speak to your teacher in class.</p>' +
    '</div></div>' +
    '<div style="background:#e8eaf6;padding:8px 28px;text-align:center;">' +
    '<p style="font-size:10px;color:#9e9e9e;margin:0;">Automated message from KS3 Science &middot; Doha College.</p>' +
    '</div></div></body></html>';

  MailApp.sendEmail({
    to:       sub.studentEmail,
    subject:  'Action Required \u2014 Resubmit Your Work (' + sub.paperId + ')',
    htmlBody: html,
    name:     'KS3 Science \u2014 Doha College',
  });
}


/**
 * Generate a unique submission token for a student.
 * ALSO expires any existing Pending/Rejected tokens for this student+paper
 * so the digest doesn't show duplicate rows from repeated test runs.
 * Writes a new row to the Submissions sheet and returns the token.
 */
function generateSubmissionToken_(ss, student, classCode, paperId, academicYear, term, teacherEmail) {
  var token     = 'SUB-' + Utilities.getUuid().replace(/-/g, '').substring(0, 16).toUpperCase();
  var now       = new Date();
  var expiresAt = new Date(now.getTime() + SUB_TOKEN_DAYS * 24 * 60 * 60 * 1000);

  var sheet = ss.getSheetByName('Submissions');
  if (!sheet) {
    Logger.log('KS3: Submissions sheet not found — submission links disabled');
    return null;
  }

  // Expire any existing Pending or Rejected tokens for this student+paper
  // Prevents duplicate rows in the teacher digest from repeated sends
  if (sheet.getLastRow() >= 3) {
    var existing = sheet.getRange(3, 1, sheet.getLastRow() - 2, 20).getValues();
    existing.forEach(function(row, i) {
      var rowStudentId = String(row[SC.studentId - 1]).trim();
      var rowPaperId   = String(row[SC.paperId - 1]).trim();
      var rowStatus    = String(row[SC.status - 1]).trim();
      if (rowStudentId === String(student.studentId).trim() &&
          rowPaperId === paperId &&
          (rowStatus === 'Pending' || rowStatus === 'Rejected')) {
        sheet.getRange(i + 3, SC.status).setValue('Expired');
      }
    });
    SpreadsheetApp.flush();
  }

  // Write the new token row
  var row = new Array(20).fill('');
  row[SC.token - 1]          = token;
  row[SC.studentId - 1]      = student.studentId;
  row[SC.studentName - 1]    = student.fullName;
  row[SC.studentEmail - 1]   = student.email;
  row[SC.classCode - 1]      = classCode;
  row[SC.paperId - 1]        = paperId;
  row[SC.academicYear - 1]   = academicYear;
  row[SC.term - 1]           = term;
  row[SC.teacherEmail - 1]   = teacherEmail;
  row[SC.createdAt - 1]      = now.toISOString();
  row[SC.expiresAt - 1]      = expiresAt.toISOString();
  row[SC.status - 1]         = 'Pending';
  row[SC.attemptNumber - 1]  = 0;

  var lastRow = Math.max(sheet.getLastRow(), 2) + 1;
  sheet.getRange(lastRow, 1, 1, row.length).setValues([row]);

  return token;
}

/**
 * Look up a submission row by token. Returns a plain object or null.
 */
function getSubmissionByToken_(token) {
  if (!token) { return null; }
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Submissions');
    if (!sheet || sheet.getLastRow() < 3) { return null; }

    var data  = sheet.getRange(3, 1, sheet.getLastRow() - 2, 20).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][SC.token - 1]).trim() === token) {
        return {
          rowIndex:        i + 3,
          token:           String(data[i][SC.token - 1]).trim(),
          studentId:       String(data[i][SC.studentId - 1]).trim(),
          studentName:     String(data[i][SC.studentName - 1]).trim(),
          studentEmail:    String(data[i][SC.studentEmail - 1]).trim(),
          classCode:       String(data[i][SC.classCode - 1]).trim(),
          paperId:         String(data[i][SC.paperId - 1]).trim(),
          academicYear:    String(data[i][SC.academicYear - 1]).trim(),
          term:            String(data[i][SC.term - 1]).trim(),
          teacherEmail:    String(data[i][SC.teacherEmail - 1]).trim(),
          createdAt:       data[i][SC.createdAt - 1],
          expiresAt:       data[i][SC.expiresAt - 1],
          submittedAt:     data[i][SC.submittedAt - 1],
          driveFileId:     String(data[i][SC.driveFileId - 1]).trim(),
          driveFileUrl:    String(data[i][SC.driveFileUrl - 1]).trim(),
          status:          String(data[i][SC.status - 1]).trim(),
          attemptNumber:   data[i][SC.attemptNumber - 1],
          reviewedAt:      data[i][SC.reviewedAt - 1],
          rejectionReason: String(data[i][SC.rejectionReason - 1]).trim(),
          rejectionNote:   String(data[i][SC.rejectionNote - 1]).trim(),
        };
      }
    }
  } catch(e) {
    Logger.log('getSubmissionByToken_ error: ' + e.toString());
  }
  return null;
}

/**
 * Update specific fields of a submission row.
 */
function updateSubmissionRow_(token, updates) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Submissions');
  if (!sheet) { return; }

  var sub = getSubmissionByToken_(token);
  if (!sub) { return; }

  var colMap = {
    submittedAt:     SC.submittedAt,
    driveFileId:     SC.driveFileId,
    driveFileUrl:    SC.driveFileUrl,
    status:          SC.status,
    attemptNumber:   SC.attemptNumber,
    reviewedAt:      SC.reviewedAt,
    rejectionReason: SC.rejectionReason,
    rejectionNote:   SC.rejectionNote,
  };

  Object.keys(updates).forEach(function(key) {
    if (colMap[key]) {
      sheet.getRange(sub.rowIndex, colMap[key]).setValue(updates[key]);
    }
  });
  SpreadsheetApp.flush();
}


// ═══════════════════════════════════════════════════════════
// ON-DEMAND DIGEST — called from Paper1.gs menu
// ═══════════════════════════════════════════════════════════

/**
 * Targeted version of sendSubmissionDigest for a specific user.
 * Called from Paper1.gs sendMySubmissionDigestNow().
 * userEmail = who to send to (and whose classes to show if not admin).
 * isAdmin = if true, show all classes regardless of teacher.
 *
 * Returns: 'sent' | 'no_submissions' | 'all_complete'
 */
function sendSubmissionDigestForUser_(userEmail, isAdmin) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName('Submissions');
  if (!sheet || sheet.getLastRow() < 3) { return 'no_submissions'; }

  var config = getSubConfig_(ss);
  var webUrl = config.webAppUrl;
  if (!webUrl) { return 'no_submissions'; }

  var props = PropertiesService.getScriptProperties();
  var now   = new Date();
  var data  = sheet.getRange(3, 1, sheet.getLastRow() - 2, 20).getValues();

  // Build class map — same logic as sendSubmissionDigest but filtered by user
  var classMap = {};
  data.forEach(function(row) {
    var token        = String(row[SC.token - 1]).trim();
    var studentName  = String(row[SC.studentName - 1]).trim();
    var studentEmail = String(row[SC.studentEmail - 1]).trim();
    var classCode    = String(row[SC.classCode - 1]).trim();
    var paperId      = String(row[SC.paperId - 1]).trim();
    var teacherEmail = String(row[SC.teacherEmail - 1]).trim();
    var status       = String(row[SC.status - 1]).trim();
    var submittedAt  = row[SC.submittedAt - 1];
    var reviewedAt   = row[SC.reviewedAt - 1];
    var driveUrl     = String(row[SC.driveFileUrl - 1]).trim();
    var expiresAt    = row[SC.expiresAt - 1];
    var attemptNum   = row[SC.attemptNumber - 1];

    if (!classCode) { return; }
    if (props.getProperty('P1_SUPPRESS_' + classCode) === 'true') { return; }

    // Filter: admin sees all, teacher sees own classes only
    if (!isAdmin && teacherEmail.toLowerCase() !== userEmail.toLowerCase()) { return; }

    if (!classMap[classCode]) {
      classMap[classCode] = {
        paperId: paperId,
        dueDate: props.getProperty('P1_DUE_' + classCode) || '',
        rows: []
      };
    }
    classMap[classCode].rows.push({
      token: token, studentName: studentName, studentEmail: studentEmail,
      status: status, submittedAt: submittedAt, reviewedAt: reviewedAt,
      driveUrl: driveUrl, expiresAt: expiresAt, attemptNum: attemptNum
    });
  });

  if (Object.keys(classMap).length === 0) { return 'no_submissions'; }

  // Check if anything active
  var activeKeys = Object.keys(classMap).filter(function(cc) {
    return classMap[cc].rows.some(function(r) {
      return r.status === 'Submitted' || r.status === 'Pending';
    });
  });

  if (activeKeys.length === 0) { return 'all_complete'; }

  // Send to the requesting user (not the original teacher address)
  sendTeacherDigestEmail_(userEmail, classMap, activeKeys, webUrl, config, now);
  return 'sent';
}


function sendSubmissionDigest() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName('Submissions');
  if (!sheet || sheet.getLastRow() < 3) { return; }
  var config = getSubConfig_(ss);
  var webUrl = config.webAppUrl;
  if (!webUrl) { Logger.log('KS3: Web App URL not set. Digest skipped.'); return; }
  var props = PropertiesService.getScriptProperties();
  var now   = new Date();
  var data  = sheet.getRange(3, 1, sheet.getLastRow() - 2, 20).getValues();

  var byTeacher = {};
  data.forEach(function(row) {
    var token        = String(row[SC.token - 1]).trim();
    var studentName  = String(row[SC.studentName - 1]).trim();
    var studentEmail = String(row[SC.studentEmail - 1]).trim();
    var classCode    = String(row[SC.classCode - 1]).trim();
    var paperId      = String(row[SC.paperId - 1]).trim();
    var teacherEmail = String(row[SC.teacherEmail - 1]).trim();
    var status       = String(row[SC.status - 1]).trim();
    var submittedAt  = row[SC.submittedAt - 1];
    var reviewedAt   = row[SC.reviewedAt - 1];
    var driveUrl     = String(row[SC.driveFileUrl - 1]).trim();
    var expiresAt    = row[SC.expiresAt - 1];
    var attemptNum   = row[SC.attemptNumber - 1];
    if (!teacherEmail || !classCode) { return; }
    if (props.getProperty('P1_SUPPRESS_' + classCode) === 'true') { return; }
    if (!byTeacher[teacherEmail]) { byTeacher[teacherEmail] = {}; }
    if (!byTeacher[teacherEmail][classCode]) {
      byTeacher[teacherEmail][classCode] = {
        paperId: paperId,
        dueDate: props.getProperty('P1_DUE_' + classCode) || '',
        rows: []
      };
    }
    byTeacher[teacherEmail][classCode].rows.push({
      token: token, studentName: studentName, studentEmail: studentEmail,
      status: status, submittedAt: submittedAt, reviewedAt: reviewedAt,
      driveUrl: driveUrl, expiresAt: expiresAt, attemptNum: attemptNum
    });
  });

  Object.keys(byTeacher).forEach(function(teacherEmail) {
    var classMap = byTeacher[teacherEmail];
    var activeKeys = Object.keys(classMap).filter(function(cc) {
      return classMap[cc].rows.some(function(r) {
        return r.status === 'Submitted' || r.status === 'Pending';
      });
    });
    if (activeKeys.length === 0) { return; }
    sendTeacherDigestEmail_(teacherEmail, classMap, activeKeys, webUrl, config, now);
  });
}

function sendTeacherDigestEmail_(teacherEmail, classMap, activeKeys, webUrl, config, now) {
  var dateStr = Utilities.formatDate(now, SUB_TZ, 'EEEE d MMMM yyyy');

  var classSections = activeKeys.map(function(cc) {
    var cls     = classMap[cc];
    var rows    = cls.rows;
    var dueDate = cls.dueDate;
    var paperId = cls.paperId;

    var awaiting    = rows.filter(function(r){ return r.status === 'Submitted'; });
    var completed   = rows.filter(function(r){ return r.status === 'Approved' || r.status === 'Escalated'; });
    var outstanding = rows.filter(function(r){
      return r.status === 'Pending' && r.expiresAt && new Date(r.expiresAt) > now;
    });

    var dueLine = '';
    if (dueDate) {
      var due      = new Date(dueDate + 'T23:59:59+03:00');
      var diffDays = Math.ceil((due - now) / (1000 * 60 * 60 * 24));
      if (diffDays > 0) {
        dueLine = '<span style="font-size:11px;color:#00695c;font-weight:bold;">Due in ' + diffDays + ' day' + (diffDays > 1 ? 's' : '') + ' (' + dueDate + ')</span>';
      } else if (diffDays === 0) {
        dueLine = '<span style="font-size:11px;color:#e65100;font-weight:bold;">Due TODAY</span>';
      } else {
        dueLine = '<span style="font-size:11px;color:#b71c1c;font-weight:bold;">OVERDUE \u2014 was due ' + dueDate + '</span>';
      }
    }

    // Section 1: Awaiting review
    var awaitingHtml = '';
    if (awaiting.length > 0) {
      awaitingHtml = '<div style="margin-bottom:14px;">' +
        '<div style="font-size:11px;font-weight:bold;color:#e64a19;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;">Awaiting Your Review (' + awaiting.length + ')</div>' +
        awaiting.map(function(r) {
          var approveUrl = webUrl + '?token=' + r.token + '&action=approve';
          var rejectUrl  = webUrl + '?token=' + r.token + '&action=reject';
          var subTime    = r.submittedAt ? Utilities.formatDate(new Date(r.submittedAt), SUB_TZ, 'd MMM HH:mm') : '';
          var timingBadge = '';
          if (dueDate && r.submittedAt) {
            var due2      = new Date(dueDate + 'T23:59:59+03:00');
            var submitted = new Date(r.submittedAt);
            var daysDiff  = Math.floor((due2 - submitted) / (1000 * 60 * 60 * 24));
            timingBadge = daysDiff >= 3
              ? '<span style="background:#e8f5e9;color:#1b5e20;border-radius:3px;padding:1px 6px;font-size:10px;margin-left:6px;">Early</span>'
              : daysDiff >= 0
              ? '<span style="background:#fff9c4;color:#f57f17;border-radius:3px;padding:1px 6px;font-size:10px;margin-left:6px;">On time</span>'
              : '<span style="background:#ffebee;color:#b71c1c;border-radius:3px;padding:1px 6px;font-size:10px;margin-left:6px;">Late</span>';
          }
          return '<table style="width:100%;border-collapse:collapse;margin-bottom:8px;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden;">' +
            '<tr style="background:#fafafa;"><td style="padding:10px 12px;vertical-align:middle;">' +
            '<div style="font-weight:bold;font-size:13px;color:#1a237e;">' + escSub_(r.studentName) + timingBadge + '</div>' +
            '<div style="font-size:11px;color:#9e9e9e;margin-top:2px;">Submitted ' + subTime + ' &middot; Attempt ' + (r.attemptNum || 1) + '</div></td>' +
            (r.driveUrl ? '<td style="padding:8px 10px;width:60px;vertical-align:middle;">' +
              '<a href="' + r.driveUrl + '" target="_blank"><img src="' + r.driveUrl + '" width="50" height="50" style="border-radius:4px;border:1px solid #e0e0e0;object-fit:cover;" alt="Photo"></a></td>' : '') +
            '<td style="padding:8px 10px;vertical-align:middle;text-align:right;white-space:nowrap;">' +
            '<a href="' + approveUrl + '" style="display:inline-block;background:#1b5e20;color:#fff;padding:6px 12px;border-radius:4px;font-size:11px;font-weight:bold;text-decoration:none;margin-bottom:4px;">Approve</a><br>' +
            '<a href="' + rejectUrl + '" style="display:inline-block;background:#ffebee;color:#b71c1c;padding:6px 12px;border-radius:4px;font-size:11px;font-weight:bold;text-decoration:none;">Reject</a>' +
            '</td></tr></table>';
        }).join('') + '</div>';
    }

    // Section 2: Completed
    var completedHtml = '';
    if (completed.length > 0) {
      completedHtml = '<div style="margin-bottom:14px;">' +
        '<div style="font-size:11px;font-weight:bold;color:#2e7d32;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">Completed (' + completed.length + ')</div>' +
        '<table style="width:100%;border-collapse:collapse;">' +
        completed.map(function(r) {
          var reviewTime = r.reviewedAt ? Utilities.formatDate(new Date(r.reviewedAt), SUB_TZ, 'd MMM') : '';
          return '<tr style="border-bottom:1px solid #f5f5f5;">' +
            '<td style="padding:5px 8px;font-size:12px;color:#424242;">' + escSub_(r.studentName) + '</td>' +
            '<td style="padding:5px 8px;font-size:11px;color:#2e7d32;text-align:right;">Approved' + (reviewTime ? ' &middot; ' + reviewTime : '') + '</td>' +
            '</tr>';
        }).join('') + '</table></div>';
    }

    // Section 3: Outstanding
    var outstandingHtml = '';
    if (outstanding.length > 0) {
      var chaseUrl = webUrl + '?action=chaseup&class=' + encodeURIComponent(cc);
      outstandingHtml = '<div style="margin-bottom:8px;">' +
        '<div style="font-size:11px;font-weight:bold;color:#9e9e9e;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">Not Yet Submitted (' + outstanding.length + ')</div>' +
        '<table style="width:100%;border-collapse:collapse;">' +
        outstanding.map(function(r) {
          var expiry = r.expiresAt ? Utilities.formatDate(new Date(r.expiresAt), SUB_TZ, 'd MMM') : '';
          return '<tr style="border-bottom:1px solid #f5f5f5;">' +
            '<td style="padding:5px 8px;font-size:12px;color:#424242;">' + escSub_(r.studentName) + '</td>' +
            '<td style="padding:5px 8px;font-size:11px;color:#9e9e9e;text-align:right;">Link expires ' + expiry + '</td></tr>';
        }).join('') + '</table>' +
        '<div style="margin-top:8px;text-align:center;">' +
        '<a href="' + chaseUrl + '" style="display:inline-block;background:#e8eaf6;color:#1a237e;padding:8px 16px;border-radius:6px;font-size:12px;font-weight:bold;text-decoration:none;">Send Chase-Up Email to Outstanding Students</a>' +
        '</div></div>';
    }

    return '<div style="margin-bottom:20px;border:1px solid #e8eaf6;border-radius:8px;overflow:hidden;">' +
      '<div style="background:#e8eaf6;padding:12px 16px;">' +
      '<table style="width:100%;border-collapse:collapse;"><tr>' +
      '<td><span style="font-size:14px;font-weight:bold;color:#1a237e;">' + escSub_(cc) + '</span>' +
      '<span style="font-size:12px;color:#666;margin-left:8px;">' + escSub_(paperId) + '</span></td>' +
      '<td style="text-align:right;">' + dueLine + '</td></tr></table></div>' +
      '<div style="padding:12px 16px;">' + awaitingHtml + completedHtml + outstandingHtml + '</div></div>';
  }).join('');

  MailApp.sendEmail({
    to:       teacherEmail,
    subject:  '[KS3 Science] Daily Submission Update \u2014 ' + dateStr,
    htmlBody: '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
      '<body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif;">' +
      '<div style="max-width:640px;margin:0 auto;background:#fff;">' +
      '<div style="background:#1a237e;padding:20px 28px;">' +
      '<h1 style="color:#fff;margin:0;font-size:17px;">Your Daily Submission Update</h1>' +
      '<p style="color:#c5cae9;margin:4px 0 0;font-size:11px;">' + escSub_(dateStr) + ' &middot; KS3 Science &middot; Doha College</p>' +
      '</div><div style="padding:18px 24px;">' + classSections + '</div>' +
      '<div style="background:#e8eaf6;padding:8px 24px;text-align:center;">' +
      '<p style="font-size:10px;color:#9e9e9e;margin:0;">KS3 Science automated digest &middot; Doha College</p>' +
      '</div></div></body></html>',
    name: 'KS3 Science \u2014 Doha College',
  });
}


// ═══════════════════════════════════════════════════════════
// CHASE-UP EMAIL
// ═══════════════════════════════════════════════════════════

/**
 * v1.1: adds bannerChaseUp image banner to each student chase-up email.
 */
function handleChaseUp_(classCode, config) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Submissions');
  if (!sheet || sheet.getLastRow() < 3) {
    return infoPage_('No Records Found', 'No submission records found for ' + classCode + '.', '#9e9e9e');
  }
  var now     = new Date();
  var data    = sheet.getRange(3, 1, sheet.getLastRow() - 2, 20).getValues();
  var sent    = 0;
  var dueDate = PropertiesService.getScriptProperties().getProperty('P1_DUE_' + classCode) || '';

  // Build banner once for all chase-up emails — bannerChaseUp (YC row 23)
  var chaseBannerUrl  = config.bannerChaseUp
    ? 'https://drive.google.com/uc?export=view&id=' + config.bannerChaseUp : '';
  var chaseBannerHtml = chaseBannerUrl
    ? '<img src="' + chaseBannerUrl + '" alt="KS3 Science" ' +
      'style="display:block;width:100%;max-width:640px;height:auto;border:0;" width="640">'
    : '';

  data.forEach(function(row) {
    var cc      = String(row[SC.classCode - 1]).trim();
    var status  = String(row[SC.status - 1]).trim();
    var expiry  = row[SC.expiresAt - 1];
    var token   = String(row[SC.token - 1]).trim();
    var sEmail  = String(row[SC.studentEmail - 1]).trim();
    var sName   = String(row[SC.studentName - 1]).trim();
    var paperId = String(row[SC.paperId - 1]).trim();
    if (cc !== classCode || status !== 'Pending') { return; }
    if (!expiry || new Date(expiry) <= now || !sEmail) { return; }
    var dueLine   = dueDate
      ? 'Please submit your work <strong>by ' + dueDate + '</strong>.'
      : 'Please submit your work as soon as possible.';
    var submitUrl = config.webAppUrl + '?token=' + token;
    var expiryStr = expiry ? Utilities.formatDate(new Date(expiry), SUB_TZ, 'd MMM') : 'expiry';
    try {
      MailApp.sendEmail({
        to: sEmail,
        subject: 'Reminder \u2014 Please Submit Your Follow-Up Work (' + paperId + ')',
        htmlBody:
          '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
          '<body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif;">' +
          '<div style="max-width:640px;margin:0 auto;background:#fff;">' +
          chaseBannerHtml +
          '<div style="background:#f5f5f5;padding:7px 30px;border-bottom:1px solid #e0e0e0;">' +
          '<p style="margin:0;font-size:10px;color:#9e9e9e;">KS3 Science &middot; Doha College</p></div>' +
          '<div style="padding:22px 28px;">' +
          '<p style="font-size:15px;color:#212121;margin:0 0 4px;">Dear ' + escSub_(sName.split(' ')[0]) + ',</p>' +
          '<p style="font-size:13px;color:#666;margin:0 0 18px;">This is a reminder that your follow-up work for ' +
          '<strong>' + escSub_(paperId) + '</strong> has not yet been submitted.</p>' +
          '<div style="background:#fff9c4;border-radius:8px;padding:14px 18px;margin-bottom:18px;">' +
          '<p style="font-size:13px;color:#f57f17;margin:0;">' + dueLine +
          ' Use the button below to take a photo of your completed exercise book work and submit it.</p></div>' +
          '<div style="text-align:center;margin-bottom:18px;">' +
          '<a href="' + submitUrl + '" style="display:inline-block;background:#1a237e;color:#fff;' +
          'padding:13px 30px;border-radius:8px;font-size:14px;font-weight:bold;text-decoration:none;">Submit My Work Now</a></div>' +
          '<p style="font-size:12px;color:#9e9e9e;text-align:center;">This link is personal to you, valid until ' + expiryStr + '</p>' +
          '</div><div style="background:#e8eaf6;padding:8px 28px;text-align:center;">' +
          '<p style="font-size:10px;color:#9e9e9e;margin:0;">Automated reminder from KS3 Science &middot; Doha College.</p>' +
          '</div></div></body></html>',
        name: 'KS3 Science \u2014 Doha College',
      });
      sent++;
    } catch(e) { Logger.log('Chase-up failed for ' + sName + ': ' + e.toString()); }
  });

  return infoPage_('Chase-Up Emails Sent',
    sent + ' chase-up email' + (sent !== 1 ? 's' : '') + ' sent to outstanding students in ' + classCode + '.', '#1a237e');
}


// ═══════════════════════════════════════════════════════════
// PRE-DEADLINE REMINDER TO TEACHER
// ═══════════════════════════════════════════════════════════

function sendPreDeadlineReminders() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Submissions');
  if (!sheet || sheet.getLastRow() < 3) { return; }
  var config   = getSubConfig_(ss);
  var props    = PropertiesService.getScriptProperties();
  var allProps = props.getProperties();
  var now      = new Date();
  var tomorrow    = new Date(now.getTime() + 24 * 60 * 60 * 1000);
  var tomorrowStr = Utilities.formatDate(tomorrow, SUB_TZ, 'yyyy-MM-dd');

  Object.keys(allProps).forEach(function(key) {
    if (key.indexOf('P1_DUE_') !== 0) { return; }
    if (allProps[key] !== tomorrowStr) { return; }
    var classCode = key.replace('P1_DUE_', '');
    if (props.getProperty('P1_SUPPRESS_' + classCode) === 'true') { return; }

    var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 20).getValues();
    var outstanding = [], teacherEmail = '', paperId = '';

    data.forEach(function(row) {
      var cc     = String(row[SC.classCode - 1]).trim();
      var status = String(row[SC.status - 1]).trim();
      var expiry = row[SC.expiresAt - 1];
      var sName  = String(row[SC.studentName - 1]).trim();
      var te     = String(row[SC.teacherEmail - 1]).trim();
      var pid    = String(row[SC.paperId - 1]).trim();
      if (cc !== classCode) { return; }
      if (te)  { teacherEmail = te; }
      if (pid) { paperId = pid; }
      if (status === 'Pending' && expiry && new Date(expiry) > now) { outstanding.push(sName); }
    });

    if (outstanding.length === 0 || !teacherEmail) { return; }

    var chaseUrl = (config.webAppUrl || '') + '?action=chaseup&class=' + encodeURIComponent(classCode);

    try {
      MailApp.sendEmail({
        to: teacherEmail,
        subject: '[KS3 Science] Deadline Tomorrow \u2014 ' + outstanding.length + ' student' + (outstanding.length > 1 ? 's' : '') + ' not yet submitted (' + classCode + ')',
        htmlBody: '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
          '<body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif;">' +
          '<div style="max-width:640px;margin:0 auto;background:#fff;">' +
          '<div style="background:#e64a19;padding:18px 24px;">' +
          '<h2 style="color:#fff;margin:0;font-size:16px;">Deadline Reminder &mdash; ' + escSub_(classCode) + '</h2>' +
          '<p style="color:#ffccbc;margin:4px 0 0;font-size:11px;">' + escSub_(paperId) + ' &middot; Due tomorrow (' + tomorrowStr + ')</p>' +
          '</div><div style="padding:18px 24px;">' +
          '<p style="font-size:13px;color:#424242;margin:0 0 14px;">The following <strong>' + outstanding.length + ' student' + (outstanding.length > 1 ? 's have' : ' has') + '</strong> not yet submitted their follow-up work. The deadline is tomorrow.</p>' +
          '<div style="background:#ffebee;border-radius:8px;padding:12px 16px;margin-bottom:16px;">' +
          outstanding.map(function(n){ return '<div style="font-size:13px;color:#b71c1c;padding:2px 0;">&bull; ' + escSub_(n) + '</div>'; }).join('') +
          '</div>' +
          '<div style="text-align:center;margin-bottom:14px;">' +
          '<a href="' + chaseUrl + '" style="display:inline-block;background:#e64a19;color:#fff;padding:11px 24px;border-radius:8px;font-size:13px;font-weight:bold;text-decoration:none;">\ud83d\udce7 Send Chase-Up Emails to All Outstanding Students</a>' +
          '<p style="font-size:11px;color:#9e9e9e;margin:6px 0 0;">One click sends a reminder email to each student listed above with their personal submission link.</p>' +
          '</div>' +
          '<p style="font-size:12px;color:#9e9e9e;">Alternatively, use your daily digest email or KS3 Science &rarr; Paper 1 &rarr; Resend Submission Link to manage individual students.</p>' +
          '</div><div style="background:#e8eaf6;padding:8px 24px;text-align:center;">' +
          '<p style="font-size:10px;color:#9e9e9e;margin:0;">KS3 Science automated reminder &middot; Doha College</p>' +
          '</div></div></body></html>',
        name: 'KS3 Science \u2014 Doha College',
      });
    } catch(e) { Logger.log('Pre-deadline reminder failed for ' + classCode + ': ' + e.toString()); }
  });
}


// ═══════════════════════════════════════════════════════════
// SUPPRESS / UNSUPPRESS — admin only via daily digest link
// ═══════════════════════════════════════════════════════════

function handleSuppressAction_(classCode, action) {
  var props = PropertiesService.getScriptProperties();
  if (action === 'suppress') {
    props.setProperty('P1_SUPPRESS_' + classCode, 'true');
    return infoPage_('Reminders Suppressed',
      'Teacher reminders for class ' + classCode + ' have been suppressed indefinitely. ' +
      'They will clear automatically when you run Archive and Clear at end of term. ' +
      'To re-enable sooner, click Re-enable in your next daily digest.',
      '#9e9e9e');
  } else {
    props.deleteProperty('P1_SUPPRESS_' + classCode);
    return infoPage_('Reminders Re-enabled',
      'Teacher digest emails for class ' + classCode + ' have been re-enabled.', '#2e7d32');
  }
}

/**
 * Returns HTML table of all classes with suppress links for Andy's 6am digest.
 * Called from Paper1.gs sendDailyDigest().
 */
function getSubmissionSummaryForAdminDigest_(webUrl) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Submissions');
  if (!sheet || sheet.getLastRow() < 3) { return ''; }
  var props = PropertiesService.getScriptProperties();
  var now   = new Date();
  var data  = sheet.getRange(3, 1, sheet.getLastRow() - 2, 20).getValues();

  var classMap = {};
  data.forEach(function(row) {
    var cc           = String(row[SC.classCode - 1]).trim();
    var status       = String(row[SC.status - 1]).trim();
    var teacherEmail = String(row[SC.teacherEmail - 1]).trim();
    var paperId      = String(row[SC.paperId - 1]).trim();
    if (!cc) { return; }
    if (!classMap[cc]) {
      classMap[cc] = {
        paperId: paperId, teacherEmail: teacherEmail,
        suppressed: props.getProperty('P1_SUPPRESS_' + cc) === 'true',
        approved: 0, submitted: 0, pending: 0,
      };
    }
    if (status === 'Approved')  { classMap[cc].approved++;  }
    if (status === 'Submitted') { classMap[cc].submitted++; }
    if (status === 'Pending')   { classMap[cc].pending++;   }
  });

  if (Object.keys(classMap).length === 0) { return ''; }

  var rows = Object.keys(classMap).map(function(cc) {
    var c = classMap[cc];
    var suppressUrl   = webUrl + '?action=suppress&class='   + encodeURIComponent(cc) + '&admin=1';
    var unsuppressUrl = webUrl + '?action=unsuppress&class=' + encodeURIComponent(cc) + '&admin=1';
    return '<tr style="border-bottom:1px solid #f0f0f0;">' +
      '<td style="padding:8px 10px;font-size:12px;font-weight:bold;color:#1a237e;">' + escSub_(cc) + '</td>' +
      '<td style="padding:8px 10px;font-size:11px;color:#424242;">' + escSub_((c.teacherEmail || '').split('@')[0]) + '</td>' +
      '<td style="padding:8px 10px;font-size:11px;text-align:center;color:#2e7d32;">' + c.approved + '</td>' +
      '<td style="padding:8px 10px;font-size:11px;text-align:center;color:#f9a825;">' + c.submitted + '</td>' +
      '<td style="padding:8px 10px;font-size:11px;text-align:center;color:#9e9e9e;">' + c.pending + '</td>' +
      '<td style="padding:8px 10px;text-align:right;">' +
      (c.suppressed
        ? '<a href="' + unsuppressUrl + '" style="font-size:10px;color:#2e7d32;text-decoration:none;padding:3px 8px;border:1px solid #2e7d32;border-radius:3px;">Re-enable</a>'
        : '<a href="' + suppressUrl   + '" style="font-size:10px;color:#9e9e9e;text-decoration:none;padding:3px 8px;border:1px solid #ddd;border-radius:3px;">Suppress</a>') +
      '</td></tr>';
  }).join('');

  return '<div style="margin-top:16px;">' +
    '<h3 style="font-size:13px;color:#37474f;margin:0 0 8px;">Submission Window Status</h3>' +
    '<table style="width:100%;border-collapse:collapse;border:1px solid #e8eaf6;border-radius:8px;overflow:hidden;">' +
    '<thead><tr style="background:#e8eaf6;">' +
    '<th style="padding:7px 10px;text-align:left;font-size:11px;color:#3949ab;">Class</th>' +
    '<th style="padding:7px 10px;text-align:left;font-size:11px;color:#3949ab;">Teacher</th>' +
    '<th style="padding:7px 10px;text-align:center;font-size:11px;color:#2e7d32;">Approved</th>' +
    '<th style="padding:7px 10px;text-align:center;font-size:11px;color:#f9a825;">Pending Review</th>' +
    '<th style="padding:7px 10px;text-align:center;font-size:11px;color:#9e9e9e;">Not Submitted</th>' +
    '<th style="padding:7px 10px;font-size:11px;"></th>' +
    '</tr></thead><tbody>' + rows + '</tbody></table></div>';
}


// ═══════════════════════════════════════════════════════════
// ESCALATION TO ANDY
// ═══════════════════════════════════════════════════════════

function escalateToAndy_(ss, sub, config, attemptNum) {
  try {
    updateSubmissionRow_(sub.token, { status: 'Escalated', escalatedAt: new Date().toISOString() });
    if (!config.hoksEmail) { return; }
    MailApp.sendEmail({
      to:      config.hoksEmail,
      subject: '[KS3 Science] Max attempts reached \u2014 ' + sub.studentName + ' (' + sub.classCode + ')',
      body:    'Student: ' + sub.studentName + '\n' +
               'Class: ' + sub.classCode + '\n' +
               'Paper: ' + sub.paperId + '\n' +
               'Attempts: ' + attemptNum + ' of ' + SUB_MAX_ATTEMPTS + '\n' +
               'Teacher: ' + sub.teacherEmail + '\n\n' +
               'This student has reached the maximum number of submission attempts. ' +
               'Please follow up with them directly.\n\n' +
               (sub.driveFileUrl ? 'Last submission: ' + sub.driveFileUrl : 'No submission photo on file.'),
      name: 'KS3 Science System',
    });
  } catch(e) {
    Logger.log('escalateToAndy_ error: ' + e.toString());
  }
}


// ═══════════════════════════════════════════════════════════
// TRIGGER INSTALLATION
// ═══════════════════════════════════════════════════════════

function installDigestTriggers() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var config = getSubConfig_(ss);
  var user   = Session.getActiveUser().getEmail();

  if (user.toLowerCase() !== (config.ownerEmail || '').toLowerCase()) {
    SpreadsheetApp.getUi().alert('Only the script owner can install triggers.');
    return;
  }

  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'sendSubmissionDigest' || fn === 'sendPreDeadlineReminders') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Teacher digest — 6:30am Qatar time = 3:30am UTC
  ScriptApp.newTrigger('sendSubmissionDigest').timeBased()
    .atHour(3).nearMinute(30).everyDays(1).inTimezone('UTC').create();

  // Pre-deadline reminder — 6:00am Qatar time = 3:00am UTC
  ScriptApp.newTrigger('sendPreDeadlineReminders').timeBased()
    .atHour(3).nearMinute(0).everyDays(1).inTimezone('UTC').create();

  SpreadsheetApp.getUi().alert(
    'Submission triggers installed.\n\n' +
    'Teacher digest: 6:30am Qatar time (daily, only when work is outstanding)\n' +
    'Pre-deadline reminder: 6:00am Qatar time (day before due date only)\n\n' +
    'Note: 13:00 digest removed. Teachers now receive one email per day maximum.'
  );
}


// ═══════════════════════════════════════════════════════════
// CONFIG & UTILITIES
// ═══════════════════════════════════════════════════════════

/**
 * v1.1: adds bannerChaseUp (YC row 23); wraps banner IDs in subExtractFileId_
 * so both bare IDs and full Drive URLs work when pasted into Year Controller.
 */
function getSubConfig_(ss) {
  var sheet = ss.getSheetByName('Year Controller');
  if (!sheet) { return {}; }
  var data  = sheet.getRange(1, 2, 40, 1).getValues();
  function v(row) { var val = data[row-1][0]; return (val !== null && val !== undefined) ? String(val).trim() : ''; }
  return {
    academicYear:   v(4),
    yearGroup:      v(5),
    currentTerm:    v(6),
    ownerEmail:     v(10),
    hoksName:       v(11),
    hoksEmail:      v(12),
    bannerModelAns: subExtractFileId_(v(21)),   // extracts ID from full URL or bare ID
    bannerChaseUp:  subExtractFileId_(v(23)),   // NEW — row 23
    masterFolderId: v(34),
    webAppUrl:      v(36),
  };
}

function escSub_(str) {
  return String(str || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * Extracts a Google Drive file ID from either a bare ID or a full Drive URL.
 * E.g. "https://drive.google.com/file/d/FILEID/view" → "FILEID"
 * E.g. "FILEID" → "FILEID"
 */
function subExtractFileId_(urlOrId) {
  if (!urlOrId) { return ''; }
  var s     = String(urlOrId).trim();
  var match = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : s;
}

function getOrCreateFolder_(parentFolder, name) {
  var existing = parentFolder.getFoldersByName(name);
  if (existing.hasNext()) { return existing.next(); }
  return parentFolder.createFolder(name);
}

function errorPage_(msg) {
  return '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{font-family:Arial,sans-serif;padding:30px;background:#f5f5f5;} .card{max-width:460px;margin:0 auto;background:#fff;border-radius:12px;padding:28px;box-shadow:0 2px 8px rgba(0,0,0,0.1);} h2{color:#b71c1c;} p{color:#555;font-size:14px;line-height:1.6;}</style></head>' +
    '<body><div class="card"><h2>Link Error</h2><p>' + escSub_(msg) + '</p></div></body></html>';
}

function infoPage_(title, msg, colour) {
  return '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{font-family:Arial,sans-serif;padding:30px;background:#f5f5f5;} .card{max-width:460px;margin:0 auto;background:#fff;border-radius:12px;padding:28px;box-shadow:0 2px 8px rgba(0,0,0,0.1);} h2{color:' + (colour||'#1a237e') + ';} p{color:#555;font-size:14px;line-height:1.6;}</style></head>' +
    '<body><div class="card"><h2>' + escSub_(title) + '</h2><p>' + escSub_(msg) + '</p></div></body></html>';
}
