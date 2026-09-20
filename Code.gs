/**********************************************************************************
 * Franchise-Setup-Operations-Tracker
* ------------------------------------------------------------------------------
 * Features:
 *   1. Team_Master Management & Dropdown Allocation
 *   2. Manual Email Notification Trigger for Staff Allocation (with Checkbox)
 *   3. Direct Sheet Edit Tracking (Owner_Audit_Log)
 *   4. Project Closure & Archiving
 **********************************************************************************/

var CONFIG = {
  OWNER_EMAIL:       'mksiva0699@gmail.com',   
  DIGEST_FREQUENCY:  'DAILY',                  
  DIGEST_HOUR:       8,                        
  COMPANY_NAME:      'Golisoda Franchise Operations'
};

var SHEETS = {
  submissions:        'Submissions',           
  queries:            'Queries',               
  projects:           'Projects',              
  master:             'Master',                
  teamMaster:         'Team_Master',           // New Staff Master Sheet
  dashboard:          'Dashboard',             
  franchise:          'Franchise View',        
  ownerAuditLog:      'Owner_Audit_Log',       // New Owner Audit Trail Log
  archivedProjects:   'Archived_Projects',     
  archivedSubmissions:'Archived_Submissions'    
};

/**
 * Custom Governance Menu
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('👑 Project Governance (PMP)')
    .addItem('1. Initialize Architecture & Master Sheets', 'setupAll')
    .addItem('2. Send Staff Allocation Emails (Selected Checkboxes)', 'sendStaffAllocationEmails')
    .addSeparator()
    .addItem('3. Execute Project Closure (Archive Project)', 'archiveCompletedProject')
    .addToUi();
}

/**
 * 1. INITIALIZE ARCHITECTURE & TEAM MASTER
 */
function setupAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create Team_Master if not exists
  var tmSheet = ss.getSheetByName(SHEETS.teamMaster) || ss.insertSheet(SHEETS.teamMaster);
  if (tmSheet.getLastRow() === 0) {
    tmSheet.appendRow(['Staff ID', 'Staff Name', 'Email ID', 'Phone Number', 'Role / Skill', 'Active?']);
    tmSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#d9ead3');
    // Sample Data
    tmSheet.appendRow(['ST01', 'Suresh K.', 'suresh@example.com', '9876543210', 'Senior Technician', 'Yes']);
    tmSheet.appendRow(['ST02', 'Anil P.', 'anil@example.com', '9876543211', 'Assembly Specialist', 'Yes']);
  }

  // Create Owner_Audit_Log if not exists
  var logSheet = ss.getSheetByName(SHEETS.ownerAuditLog) || ss.insertSheet(SHEETS.ownerAuditLog);
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(['Timestamp', 'User / Role', 'Sheet Name', 'Cell Location', 'Old Value', 'New Value', 'Action Summary']);
    logSheet.getRange('A1:G1').setFontWeight('bold').setBackground('#fce5cd');
  }

  // Update Data Validation Dropdowns in Projects Sheet
  updateStaffDropdowns_(ss);

  var upd = buildUpdateForm_(ss);
  var con = buildConcernForm_(ss);
  installTriggers_(ss, upd.getId(), con.getId());
  saveFormLinks_(ss, upd, con);

  SpreadsheetApp.getActiveSpreadsheet().toast('System Initialized with Team Master & Governance Logs.', 'Setup Complete', 8);
}

/**
 * 2. SEND STAFF ALLOCATION EMAILS (Triggered by Button / Menu)
 */
function sendStaffAllocationEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pSheet = ss.getSheetByName(SHEETS.projects);
  var tmSheet = ss.getSheetByName(SHEETS.teamMaster);
  
  if (!pSheet || !tmSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Projects or Team_Master sheet missing.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Load Team Master Emails (Staff Name -> Email)
  var tmData = tmSheet.getRange(2, 1, Math.max(tmSheet.getLastRow() - 1, 1), 6).getValues();
  var staffEmailMap = {};
  for (var i = 0; i < tmData.length; i++) {
    var name = String(tmData[i][1]).trim();
    var email = String(tmData[i][2]).trim();
    var active = String(tmData[i][5]).trim().toLowerCase();
    if (name && email && active !== 'no') {
      staffEmailMap[name] = email;
    }
  }

  var pData = pSheet.getRange(5, 1, pSheet.getLastRow() - 4, 15).getValues(); 
  // Col A: ProjID, Col B: Franchise, Col C: Location, Col L: Tech1, Col M: Tech2, Col O: Send Checkbox
  var count = 0;

  for (var r = 0; r < pData.length; r++) {
    var row = pData[r];
    var projectId = row[0];
    var franchiseName = row[1];
    var location = row[2];
    var tech1 = String(row[11] || '').trim();
    var tech2 = String(row[12] || '').trim();
    var isChecked = row[14]; // Checkbox Column O

    if (isChecked === true && franchiseName !== '- available -') {
      var assignedTechs = [tech1, tech2].filter(function(t) { return t !== ''; });
      
      assignedTechs.forEach(function(techName) {
        var email = staffEmailMap[techName];
        if (email) {
          var subject = '[' + CONFIG.COMPANY_NAME + '] New Project Assignment: ' + franchiseName + ' (' + projectId + ')';
          var body = 'Hello ' + techName + ',\n\n'
                   + 'You have been assigned to a new franchise installation project:\n\n'
                   + '• Project ID: ' + projectId + '\n'
                   + '• Franchise: ' + franchiseName + '\n'
                   + '• Location: ' + location + '\n\n'
                   + 'Please coordinate with the Operations Head for execution details.\n\n'
                   + 'Regards,\n' + CONFIG.COMPANY_NAME;
          
          MailApp.sendEmail(email, subject, body);
          count++;
        }
      });

      // Uncheck the checkbox after sending email
      pSheet.getRange(r + 5, 15).setValue(false);
    }
  }

  SpreadsheetApp.getUi().alert('Success', count + ' Allocation Email(s) sent successfully!', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 3. AUTOMATIC OWNER EDIT TRACKER (Logs edits in Owner_Audit_Log)
 */
function onEdit(e) {
  if (!e) return;
  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();

  // Log changes made in Projects, Tracker, or Master sheets
  if (sheetName === SHEETS.projects || sheetName === SHEETS.master || sheetName === 'Tracker') {
    var ss = e.source;
    var logSheet = ss.getSheetByName(SHEETS.ownerAuditLog);
    if (!logSheet) return;

    var user = Session.getActiveUser().getEmail() || 'Owner / PM';
    var cellLoc = range.getA1Notation();
    var oldValue = e.oldValue || '(Empty)';
    var newValue = e.value || '(Cleared)';

    logSheet.appendRow([
      new Date(),
      user,
      sheetName,
      cellLoc,
      oldValue,
      newValue,
      'Direct Cell Edit on ' + sheetName + ' at ' + cellLoc
    ]);
  }
}

/**
 * 4. PROJECT CLOSURE (ARCHIVING)
 */
function archiveCompletedProject() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var response = ui.prompt('Project Closure', 'Enter Project ID to Close & Archive (e.g., P01):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var projectId = response.getResponseText().trim();

  if (!projectId) return;

  var projSheet = ss.getSheetByName(SHEETS.projects);
  var subSheet  = ss.getSheetByName(SHEETS.submissions);

  var archProjSheet = ss.getSheetByName(SHEETS.archivedProjects) || ss.insertSheet(SHEETS.archivedProjects);
  var archSubSheet  = ss.getSheetByName(SHEETS.archivedSubmissions) || ss.insertSheet(SHEETS.archivedSubmissions);

  if (archProjSheet.getLastRow() === 0) {
    archProjSheet.appendRow(['Project ID', 'Franchise Name', 'Location', 'Owner', 'Contact', 'Start Date', 'Target Go-Live', 'Lifecycle Phase', '% Progress', 'Archived?', 'Closure Date']);
  }
  if (archSubSheet.getLastRow() === 0) {
    archSubSheet.appendRow(['Timestamp', 'Project ID', 'Task ID', 'Resource/Tech', 'Status', 'Actual/Qty', 'Remarks', 'Artifact URL', 'Seq']);
  }

  var pData = projSheet.getRange(5, 1, projSheet.getLastRow() - 4, 11).getValues();
  for (var i = 0; i < pData.length; i++) {
    if (String(pData[i][0]).trim().toLowerCase() === projectId.toLowerCase()) {
      var rowToArchive = pData[i];
      rowToArchive[9] = 'Yes';
      rowToArchive[10] = new Date(); 
      archProjSheet.appendRow(rowToArchive);
      
      projSheet.getRange(i + 5, 2).setValue('- available -');
      projSheet.getRange(i + 5, 10).setValue('Yes');
      break;
    }
  }

  var sData = subSheet.getRange(5, 1, Math.max(subSheet.getLastRow() - 4, 1), 9).getValues();
  for (var j = sData.length - 1; j >= 0; j--) {
    if (String(sData[j][1]).trim().toLowerCase() === projectId.toLowerCase()) {
      archSubSheet.appendRow(sData[j]);
      subSheet.deleteRow(j + 5);
    }
  }

  ui.alert('Project Closed!', 'Project ' + projectId + ' records moved to Archive successfully.', ui.ButtonSet.OK);
}

/* Helper Functions */

function updateStaffDropdowns_(ss) {
  var tmSheet = ss.getSheetByName(SHEETS.teamMaster);
  var pSheet = ss.getSheetByName(SHEETS.projects);
  if (!tmSheet || !pSheet) return;

  var lastRow = tmSheet.getLastRow();
  if (lastRow < 2) return;

  var staffNames = tmSheet.getRange(2, 2, lastRow - 1, 1).getValues().map(function(r) { return r[0]; }).filter(String);
  if (staffNames.length === 0) return;

  var rule = SpreadsheetApp.newDataValidation().requireValueInList(staffNames, true).build();
  // Apply validation to Tech 1 (Col L) & Tech 2 (Col M) in Projects sheet
  pSheet.getRange('L5:M14').setDataValidation(rule);
  
  // Apply Checkbox validation to Col O (Send Allocation Email?)
  var cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  pSheet.getRange('O5:O14').setDataValidation(cbRule);
  pSheet.getRange('O4').setValue('Send Email?').setFontWeight('bold');
}

function buildUpdateForm_(ss) {
  var projects = getActiveProjectIds_(ss);
  var tasks    = getActiveTaskIds_(ss);
  var statuses = getColumn_(ss, SHEETS.master, 9, 5);

  var form = FormApp.create(CONFIG.COMPANY_NAME + ' — Work Performance Entry')
    .setDescription('Submit task performance metrics for project lifecycle tracking.')
    .setCollectEmail(false).setAllowResponseEdits(false).setLimitOneResponsePerUser(false);

  form.addListItem().setTitle('Project ID').setChoiceValues(projects).setRequired(true);
  form.addListItem().setTitle('Task ID').setChoiceValues(tasks).setRequired(true);
  form.addTextItem().setTitle('Resource Name').setRequired(true);
  form.addListItem().setTitle('Status').setChoiceValues(statuses).setRequired(true);
  form.addTextItem().setTitle('Actual / Deliverable Metric');
  form.addParagraphTextItem().setTitle('Performance Remarks');
  form.addTextItem().setTitle('Photo link (optional)');
  return form;
}

function buildConcernForm_(ss) {
  var projects = getActiveProjectIds_(ss);
  var form = FormApp.create(CONFIG.COMPANY_NAME + ' — Issue Log & Risk Intake')
    .setDescription('Log an operational issue or project risk for Project Manager review.')
    .setCollectEmail(false).setAllowResponseEdits(false);

  form.addListItem().setTitle('Project ID').setChoiceValues(projects).setRequired(true);
  form.addTextItem().setTitle('Reported By').setRequired(true);
  form.addListItem().setTitle('Type').setChoiceValues(['Risk', 'Issue', 'Change Request']).setRequired(true);
  form.addListItem().setTitle('Category').setChoiceValues(['Procurement', 'Logistics', 'Technical', 'Quality', 'Financial']).setRequired(true);
  form.addParagraphTextItem().setTitle('Description').setRequired(true);
  form.addListItem().setTitle('Severity Level').setChoiceValues(['High', 'Medium', 'Low']).setRequired(true);
  return form;
}

function installTriggers_(ss, updateFormId, concernFormId) {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onUpdateSubmit').forForm(FormApp.openById(updateFormId)).onFormSubmit().create();
  ScriptApp.newTrigger('onConcernSubmit').forForm(FormApp.openById(concernFormId)).onFormSubmit().create();

  var b = ScriptApp.newTrigger('sendDigest').timeBased().atHour(CONFIG.DIGEST_HOUR);
  if (CONFIG.DIGEST_FREQUENCY === 'WEEKLY') { b.onWeekDay(ScriptApp.WeekDay.MONDAY); }
  else { b.everyDays(1); }
  b.create();

  var props = PropertiesService.getDocumentProperties();
  props.setProperty('UPDATE_FORM_ID', updateFormId);
  props.setProperty('CONCERN_FORM_ID', concernFormId);
}

function onUpdateSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.submissions);
  var a  = mapResponse_(e);

  var row = [
    new Date(),
    a['Project ID'] || '',
    a['Task ID'] || '',
    a['Resource Name'] || '',
    a['Status'] || '',
    a['Actual / Deliverable Metric'] || '',
    a['Performance Remarks'] || '',
    a['Photo link (optional)'] || ''
  ];
  sh.appendRow(row);
  var r = sh.getLastRow();
  sh.getRange(r, 9).setFormula('=ROW()');

  if ((a['Status'] || '').toString().toLowerCase() === 'issue') {
    sendIssueAlert_(a);
  }
}

function onConcernSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.queries);
  var a  = mapResponse_(e);

  var qid = nextQueryId_(sh);
  sh.appendRow([
    qid, new Date(), a['Project ID'] || '', a['Reported By'] || '',
    a['Type'] || '', a['Category'] || '', a['Description'] || '',
    a['Severity Level'] || '', 'Open', '', ''
  ]);

  MailApp.sendEmail(CONFIG.OWNER_EMAIL,
    '[' + CONFIG.COMPANY_NAME + '] Risk/Issue Log Alert — ' + (a['Project ID'] || ''),
    'A new ' + (a['Type'] || 'Issue') + ' has been logged.');
}

function sendIssueAlert_(a) {
  MailApp.sendEmail(CONFIG.OWNER_EMAIL,
    '⚠ CRITICAL ISSUE — ' + (a['Project ID'] || '') + ' / ' + (a['Task ID'] || ''),
    'Work Package flagged with Critical Issue:\n\nProject: ' + (a['Project ID'] || '') + '\nTask: ' + (a['Task ID'] || ''));
}

function sendDigest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName(SHEETS.dashboard);
  var data = dash.getRange(15, 1, 10, 13).getValues();
  var lines = [];

  data.forEach(function (r) {
    var pid = r[0], name = r[1];
    if (!name || name === '- available -') return;
    var pct = Math.round((r[5] || 0) * 100);
    var stage = r[4];
    lines.push('• ' + pid + '  ' + name + '  —  Progress: ' + pct + '%  (Phase: ' + stage + ')');
  });

  MailApp.sendEmail(CONFIG.OWNER_EMAIL,
    '[' + CONFIG.COMPANY_NAME + '] Project Portfolio Digest', lines.join('\n'));
}

function mapResponse_(e) {
  var out = {};
  e.response.getItemResponses().forEach(function (ir) {
    out[ir.getItem().getTitle()] = ir.getResponse();
  });
  return out;
}

function getColumn_(ss, sheetName, colIndex, startRow) {
  var sh = ss.getSheetByName(sheetName);
  var last = sh.getLastRow();
  if (last < startRow) return [];
  var vals = sh.getRange(startRow, colIndex, last - startRow + 1, 1).getValues();
  return vals.map(function (r) { return r[0]; }).filter(function (v) { return v !== '' && v != null; });
}

function getActiveProjectIds_(ss) {
  var sh = ss.getSheetByName(SHEETS.projects);
  var last = sh.getLastRow(); if (last < 5) return [];
  var vals = sh.getRange(5, 1, last - 4, 2).getValues();
  var out = [];
  vals.forEach(function (r) {
    if (r[0] && r[1] && r[1] !== '- available -') out.push(String(r[0]));
  });
  return out.length ? out : ['P01'];
}

function getActiveTaskIds_(ss) {
  var sh = ss.getSheetByName(SHEETS.master);
  var last = sh.getLastRow(); if (last < 5) return [];
  var vals = sh.getRange(5, 1, last - 4, 7).getValues();
  var out = [];
  vals.forEach(function (r) {
    if (r[0] && String(r[6]).toLowerCase() !== 'no') out.push(String(r[0]));
  });
  return out;
}

function nextQueryId_(sh) {
  var last = sh.getLastRow();
  var n = 0;
  if (last >= 5) {
    var ids = sh.getRange(5, 1, last - 4, 1).getValues();
    ids.forEach(function (r) {
      var m = String(r[0]).match(/Q(\d+)/);
      if (m) n = Math.max(n, parseInt(m[1], 10));
    });
  }
  return 'Q' + ('0' + (n + 1)).slice(-2);
}

function saveFormLinks_(ss, upd, con) {
  var name = 'Form Links';
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  sh.getRange('A1').setValue('FORM LINKS — Stakeholder & Team Interfaces').setFontWeight('bold').setFontSize(13);
  sh.getRange('A3').setValue('Work Performance Form (Team)');
  sh.getRange('B3').setValue(upd.getPublishedUrl());
  sh.getRange('A4').setValue('Issue Log Form (Stakeholders)');
  sh.getRange('B4').setValue(con.getPublishedUrl());
  sh.getRange('A6').setValue('Work Performance Form — EDIT (PM Only)');
  sh.getRange('B6').setValue(upd.getEditUrl());
  sh.getRange('A7').setValue('Issue Log Form — EDIT (PM Only)');
  sh.getRange('B7').setValue(con.getEditUrl());
  sh.setColumnWidth(1, 240); sh.setColumnWidth(2, 520);
}
