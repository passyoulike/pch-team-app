const SPREADSHEET_ID = '1X6mw_MzNMgIGbS1XMFxdXEGjONEGPRIrwGPaBrYeiqs';
const SHEET_NAME = 'Applicants';
const RESUME_FOLDER_ID = '1WftKBM2t-75Kev02byM74qpBzNrRbZlJ';

function doGet(e) {
  var appUrl = ScriptApp.getService().getUrl();
  if (e && e.parameter && e.parameter.page === 'apply') {
  return HtmlService.createHtmlOutput(getApplyHtml(appUrl))
  .setTitle('PCH Online Application Form')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  var template = HtmlService.createTemplateFromFile('Index');
  template.appUrl = appUrl;
  template.policyOptions = getPolicyOptions();
  template.departmentOptions = getDepartmentOptions();
  return template.evaluate()
  .setTitle('app.pch.hospital')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

                              function submitApplication(data) {
                                var folder = DriveApp.getFolderById(RESUME_FOLDER_ID);
                                  var resumeUrl = '';

                                    if (data.resumeBase64 && data.resumeFileName) {
                                        var decoded = Utilities.base64Decode(data.resumeBase64);
                                            var blob = Utilities.newBlob(decoded, data.resumeMimeType, data.resumeFileName);
                                                var file = folder.createFile(blob);
                                                    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
                                                        resumeUrl = file.getUrl();
                                                          }

                                                            var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
                                                              var sheet = ss.getSheetByName(SHEET_NAME);

                                                                sheet.appendRow([
                                                                    new Date(),
                                                                        data.lastName,
                                                                            data.firstName,
                                                                                data.middleName,
                                                                                    data.email,
                                                                                        data.contactNumber,
                                                                                            data.address,
                                                                                                data.position,
                                                                                                    '',
                                                                                                        '',
                                                                                                            resumeUrl
                                                                                                              ]);

                                                                                                                var emailSheet = ss.getSheetByName('EMAIL');
                                                                                                                if (emailSheet) {
                                                                                                                var lastRow = emailSheet.getLastRow();
                                                                                                                var recipients = [];
                                                                                                                if (lastRow > 1) {
                                                                                                                var emailData = emailSheet.getRange(2, 1, lastRow - 1, 1).getValues();
                                                                                                                for (var i = 0; i < emailData.length; i++) {
                                                                                                                var addr = emailData[i][0];
                                                                                                                if (addr && addr.toString().indexOf('@') > -1) {
                                                                                                                recipients.push(addr.toString());
                                                                                                                }
                                                                                                                }
                                                                                                                }
                                                                                                                if (recipients.length > 0) {
                                                                                                                var subject = 'New Application: ' + data.position + ' - ' + data.lastName + ', ' + data.firstName;
                                                                                                                var body = '<h2>New Job Application Received</h2>' +
                                                                                                                '<p><strong>Position:</strong> ' + data.position + '</p>' +
                                                                                                                '<p><strong>Name:</strong> ' + data.lastName + ', ' + data.firstName + ' ' + data.middleName + '</p>' +
                                                                                                                '<p><strong>Email:</strong> ' + data.email + '</p>' +
                                                                                                                '<p><strong>Contact Number:</strong> ' + data.contactNumber + '</p>' +
                                                                                                                '<p><strong>Address:</strong> ' + data.address + '</p>' +
                                                                                                                '<p><strong>Resume:</strong> <a href="' + resumeUrl + '">' + resumeUrl + '</a></p>';
                                                                                                                MailApp.sendEmail({
                                                                                                                to: recipients.join(','),
                                                                                                                subject: subject,
                                                                                                                htmlBody: body
                                                                                                                });
                                                                                                                }
                                                                                                                }
                                                                                                                
                                                                                                                return 'success';
                                                                                                                }

                                                                                                                function getApplyHtml(appUrl) {
                                                                                                                var html = '<!DOCTYPE html><html><head>';
                                                                                                                html += '<base target="_top">';
                                                                                                                html += '<meta name="viewport" content="width=device-width, initial-scale=1">';
                                                                                                                html += '<style>';
                                                                                                                html += '*{box-sizing:border-box;margin:0;padding:0;font-family:\'Segoe UI\', Arial, sans-serif;}';
                                                                                                                html += 'body{min-height:100vh;background:radial-gradient(circle at 15% 15%, rgba(0,200,150,0.25), transparent 40%),radial-gradient(circle at 85% 85%, rgba(0,100,180,0.25), transparent 40%),linear-gradient(135deg, #071a14 0%, #0c2a1f 45%, #08202e 100%);background-attachment:fixed;color:#eafff5;padding:20px;}';
                                                                                                                html += '.container{max-width:520px;margin:auto;}';
                                                                                                                html += '.back{display:inline-block;margin-bottom:16px;color:#7fe3c0;text-decoration:none;font-size:13px;font-weight:700;}';
                                                                                                                html += '.header{background:linear-gradient(135deg, #00a86b, #007a52 60%, #004d33);color:white;padding:30px 24px;border-radius:24px;text-align:center;box-shadow:0 10px 30px rgba(0,0,0,0.35), inset 0 1px 0 rgba(255,255,255,0.15);margin-bottom:24px;}';
                                                                                                                html += '.header h1{font-size:24px;font-weight:800;margin-bottom:6px;}';
                                                                                                                html += '.header p{font-size:14px;opacity:0.95;}';
                                                                                                                html += '.card{background:rgba(255,255,255,0.06);backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);border:1px solid rgba(255,255,255,0.12);border-radius:20px;padding:24px;box-shadow:0 4px 18px rgba(0,0,0,0.25);}';
                                                                                                                html += 'label{display:block;font-size:12px;font-weight:700;color:#7fe3c0;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;margin-top:16px;}';
                                                                                                                html += 'label:first-child{margin-top:0;}';
                                                                                                                html += 'select, input[type=text], input[type=email], input[type=tel], textarea{width:100%;padding:12px 14px;border-radius:12px;border:1px solid rgba(255,255,255,0.18);background:rgba(255,255,255,0.05);color:#eafff5;font-size:15px;}';
                                                                                                                html += 'select option{color:#000;}';
                                                                                                                html += 'textarea{resize:vertical;min-height:70px;font-family:inherit;}';
                                                                                                                html += 'input[type=file]{width:100%;padding:10px;border-radius:12px;border:1px dashed rgba(255,255,255,0.3);background:rgba(255,255,255,0.05);color:#c9f2e2;font-size:13px;}';
                                                                                                                html += 'button{width:100%;margin-top:24px;padding:14px;border:none;border-radius:14px;background:linear-gradient(135deg, #00c896, #007a52);color:white;font-size:16px;font-weight:700;cursor:pointer;box-shadow:0 4px 14px rgba(0,200,150,0.35);}';
                                                                                                                html += 'button:disabled{opacity:0.6;cursor:not-allowed;}';
                                                                                                                html += '#status{margin-top:16px;text-align:center;font-size:14px;}';
                                                                                                                html += '#status.success{color:#7fe3c0;}';
                                                                                                                html += '#status.error{color:#ff9999;}';
                                                                                                                html += '</style></head><body>';
                                                                                                                html += '<div class="container">';
                                                                                                                html += '<a class="back" href="' + appUrl + '">&larr; Back to PCH Team App</a>';
                                                                                                                html += '<div class="header"><h1>PCH Online Application Form</h1><p>Puerto Community Hospital Careers</p></div>';
                                                                                                                html += '<div class="card"><form id="appForm">';
                                                                                                                html += '<label>Position Applied For</label>';
                                                                                                                html += '<select id="position" required>';
                                                                                                                html += '<option value="" disabled selected>Select Position</option>';
                                                                                                                html += '<option value="Midwife">Midwife</option>';
                                                                                                                html += '<option value="Nurse">Nurse</option>';
                                                                                                                html += '<option value="Radtech">Radtech</option>';
                                                                                                                html += '<option value="Medtech">Medtech</option>';
                                                                                                                html += '</select>';
                                                                                                                html += '<label>Last Name</label><input type="text" id="lastName" required>';
                                                                                                                html += '<label>First Name</label><input type="text" id="firstName" required>';
                                                                                                                html += '<label>Middle Name</label><input type="text" id="middleName">';
                                                                                                                html += '<label>Email Address</label><input type="email" id="email" required>';
                                                                                                                html += '<label>Contact Number</label><input type="tel" id="contactNumber" required>';
                                                                                                                html += '<label>Address</label><textarea id="address" required></textarea>';
                                                                                                                html += '<label>Upload Resume</label><input type="file" id="resume" accept=".pdf,.doc,.docx" required>';
                                                                                                                html += '<button type="submit" id="submitBtn">Submit Application</button>';
                                                                                                                html += '<div id="status"></div>';
                                                                                                                html += '</form></div></div>';

                                                                                                                html += '<script>';
                                                                                                                html += 'document.getElementById("appForm").addEventListener("submit", function(e) {';
                                                                                                                html += 'e.preventDefault();';
                                                                                                                html += 'var btn = document.getElementById("submitBtn");';
                                                                                                                html += 'var status = document.getElementById("status");';
                                                                                                                html += 'var fileInput = document.getElementById("resume");';
                                                                                                                html += 'var file = fileInput.files[0];';
                                                                                                                html += 'if (!file) { status.textContent = "Please attach your resume."; status.className = "error"; return; }';
                                                                                                                html += 'btn.disabled = true; btn.textContent = "Submitting..."; status.textContent = ""; status.className = "";';
                                                                                                                html += 'var reader = new FileReader();';
                                                                                                                html += 'reader.onload = function() {';
                                                                                                                html += 'var base64 = reader.result.split(",")[1];';
                                                                                                                html += 'var data = {';
                                                                                                                html += 'position: document.getElementById("position").value,';
                                                                                                                html += 'lastName: document.getElementById("lastName").value,';
                                                                                                                html += 'firstName: document.getElementById("firstName").value,';
                                                                                                                html += 'middleName: document.getElementById("middleName").value,';
                                                                                                                html += 'email: document.getElementById("email").value,';
                                                                                                                html += 'contactNumber: document.getElementById("contactNumber").value,';
                                                                                                                html += 'address: document.getElementById("address").value,';
                                                                                                                html += 'resumeBase64: base64,';
                                                                                                                html += 'resumeFileName: file.name,';
                                                                                                                html += 'resumeMimeType: file.type';
                                                                                                                html += '};';
                                                                                                                html += 'google.script.run.withSuccessHandler(function() {';
                                                                                                                html += 'status.textContent = "Application submitted successfully!"; status.className = "success";';
                                                                                                                html += 'btn.textContent = "Submitted";';
                                                                                                                html += 'document.getElementById("appForm").reset();';
                                                                                                                html += '}).withFailureHandler(function(err) {';
                                                                                                                html += 'status.textContent = "Error: " + err.message; status.className = "error";';
                                                                                                                html += 'btn.disabled = false; btn.textContent = "Submit Application";';
                                                                                                                html += '}).submitApplication(data);';
                                                                                                                html += '};';
                                                                                                                html += 'reader.readAsDataURL(file);';
                                                                                                                html += '});';
                                                                                                                html += '</script>';
                                                                                                                html += '</body></html>';
                                                                                                                return html;
                                                                                                                }



function submitSupplyRequest(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Supply Request 2');
  var lastRow = sheet.getLastRow();
  var targetRow = lastRow < 23 ? 24 : lastRow + 1;

  sheet.getRange(targetRow, 1, 1, 9).setValues([[
    new Date(),
    0,
    data.employeeName,
    data.phoneNumber,
    data.employeeEmail,
    data.department,
    data.supplyRequested,
    '',
    ''
  ]]);

  return 'success';
}

function sendSupplyRequestNotification(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var emailSheet = ss.getSheetByName('EMAIL');
  if (emailSheet) {
    var emailLastRow = emailSheet.getLastRow();
    var recipients = [];
    if (emailLastRow > 1) {
      var emailData = emailSheet.getRange(2, 3, emailLastRow - 1, 1).getValues();
      for (var i = 0; i < emailData.length; i++) {
        var addr = emailData[i][0];
        if (addr && addr.toString().indexOf('@') > -1) { recipients.push(addr.toString()); }
      }
    }
    if (recipients.length > 0) {
      var subject = 'New Supply Request: ' + data.employeeName + ' - ' + data.department;
      var body = '<h2>New Supply Request</h2>' +
        '<p><strong>Employee Name:</strong> ' + data.employeeName + '</p>' +
        '<p><strong>Phone Number:</strong> ' + data.phoneNumber + '</p>' +
        '<p><strong>Employee Email:</strong> ' + data.employeeEmail + '</p>' +
        '<p><strong>Department:</strong> ' + data.department + '</p>' +
        '<p><strong>Supply Requested:</strong> ' + data.supplyRequested + '</p>';
      MailApp.sendEmail({ to: recipients.join(','), subject: subject, htmlBody: body });
    }
  }
}

function submitEndorsement(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Endorsement');
  if (!sheet) {
    sheet = ss.insertSheet('Endorsement');
    sheet.appendRow(['Timestamp', 'Date', 'Shift', 'Names', 'Department', 'BP Apparatus', 'Thermometers', 'Pulse Oximeter', 'Stethoscope', 'Suction Machine', 'Nebulizer', 'Others']);
  }

  sheet.appendRow([
    new Date(),
    data.date,
    data.shift,
    data.names,
    data.department,
    data.bpApparatus,
    data.thermometers,
    data.pulseOximeter,
    data.stethoscope,
    data.suctionMachine,
    data.nebulizer,
    data.others
  ]);

  return 'success';
}

function getEndorsementReports() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Endorsement');
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var numRows = Math.min(lastRow - 1, 20);
  var startRow = lastRow - numRows + 1;
  var values = sheet.getRange(startRow, 1, numRows, 12).getValues();

  var reports = [];
  for (var i = values.length - 1; i >= 0; i--) {
    var row = values[i];
    reports.push({
      date: row[1],
      shift: row[2],
      names: row[3],
      department: row[4],
      bpApparatus: row[5],
      thermometers: row[6],
      pulseOximeter: row[7],
      stethoscope: row[8],
      suctionMachine: row[9],
      nebulizer: row[10],
      others: row[11]
    });
  }
  return reports;
}

function formatDateKey_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value.toString().trim();
}

function getDailyEndorsementSummary() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Endorsement');
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, 12).getValues();

  var requiredShifts = ['7am-3pm', '3pm-11pm', '11pm-7am'];
  var equipmentFields = [
    { col: 5, label: 'BP Apparatus' },
    { col: 6, label: 'Thermometers' },
    { col: 7, label: 'Pulse Oximeter' },
    { col: 8, label: 'Stethoscope' },
    { col: 9, label: 'Suction Machine' },
    { col: 10, label: 'Nebulizer' }
  ];

  var byDate = {};
  var order = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var dateKey = formatDateKey_(row[1]);
    if (!dateKey) continue;
    var shift = row[2];

    if (!byDate[dateKey]) {
      byDate[dateKey] = {};
      order.push(dateKey);
    }
    byDate[dateKey][shift] = row;
  }

  var days = [];
  for (var d = 0; d < order.length; d++) {
    var dateKey = order[d];
    var shiftsData = byDate[dateKey];
    var shiftsSubmitted = requiredShifts.filter(function(s) { return !!shiftsData[s]; });
    var shiftsMissing = requiredShifts.filter(function(s) { return !shiftsData[s]; });
    var complete = shiftsMissing.length === 0;

    var missingDevices = [];
    if (complete) {
      requiredShifts.forEach(function(shift) {
        var row = shiftsData[shift];
        equipmentFields.forEach(function(f) {
          var val = row[f.col];
          if (val !== null && val !== undefined && val.toString().trim() === '0') {
            missingDevices.push({ shift: shift, department: row[4], device: f.label });
          }
        });
      });
    }

    days.push({
      date: dateKey,
      complete: complete,
      shiftsSubmitted: shiftsSubmitted,
      shiftsMissing: shiftsMissing,
      missingDevices: missingDevices
    });
  }

  days.sort(function(a, b) { return a.date < b.date ? 1 : -1; });
  return days.slice(0, 30);
}

function sendEndorsementNotification(data) {
  var dateKey = formatDateKey_(data.date);
  var days = getDailyEndorsementSummary();
  var day = null;
  for (var i = 0; i < days.length; i++) {
    if (days[i].date === dateKey) { day = days[i]; break; }
  }
  if (!day) return;

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var emailSheet = ss.getSheetByName('EMAIL');
  if (!emailSheet) return;

  var emailLastRow = emailSheet.getLastRow();
  var recipients = [];
  if (emailLastRow > 1) {
    var emailData = emailSheet.getRange(2, 4, emailLastRow - 1, 1).getValues();
    for (var j = 0; j < emailData.length; j++) {
      var addr = emailData[j][0];
      if (addr && addr.toString().indexOf('@') > -1) { recipients.push(addr.toString()); }
    }
  }
  if (recipients.length === 0) return;

  var subject = 'Endorsement Report - ' + day.date + (day.complete ? (day.missingDevices.length ? ' (Missing Devices)' : ' (Complete)') : ' (Incomplete)');

  var body = '<h2>Endorsement Report - ' + day.date + '</h2>';
  body += '<p><strong>Shifts Submitted:</strong> ' + (day.shiftsSubmitted.length ? day.shiftsSubmitted.join(', ') : 'None') + '</p>';

  if (!day.complete) {
    body += '<p><strong>Waiting For:</strong> ' + day.shiftsMissing.join(', ') + '</p>';
  } else if (day.missingDevices.length > 0) {
    body += '<p><strong>Missing Devices:</strong></p><ul>';
    day.missingDevices.forEach(function(m) {
      body += '<li>' + m.device + ' &mdash; ' + m.department + ' (' + m.shift + ' shift)</li>';
    });
    body += '</ul>';
  } else {
    body += '<p>All equipment accounted for across all 3 shifts.</p>';
  }

  MailApp.sendEmail({ to: recipients.join(','), subject: subject, htmlBody: body });
}

function getDepartmentOptions() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('SELECTION');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var options = [];
  for (var i = 0; i < values.length; i++) {
    var v = values[i][0];
    if (v) options.push(v.toString());
  }
  return options;
}

function getPolicyOptions() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('SELECTION');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var options = [];
  for (var i = 0; i < values.length; i++) {
    var v = values[i][0];
    if (v) options.push(v.toString());
  }
  return options;
}

function submitIncidentReport(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Incident Report');
  if (!sheet) {
    sheet = ss.insertSheet('Incident Report');
    sheet.appendRow(['Timestamp', 'Date', 'Reported By', 'Reported Individual', 'Subject of the Incident', 'Details of the Incident', 'Employee ID', 'Policy']);
  }

  sheet.appendRow([
    new Date(),
    data.date,
    data.reportedBy,
    data.reportedIndividual,
    data.subject,
    data.details,
    data.employeeId,
    data.policy
  ]);

  return 'success';
}

function sendIncidentReportNotification(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var emailSheet = ss.getSheetByName('EMAIL');
  if (emailSheet) {
    var emailLastRow = emailSheet.getLastRow();
    var recipients = [];
    if (emailLastRow > 1) {
      var emailData = emailSheet.getRange(2, 3, emailLastRow - 1, 1).getValues();
      for (var i = 0; i < emailData.length; i++) {
        var addr = emailData[i][0];
        if (addr && addr.toString().indexOf('@') > -1) { recipients.push(addr.toString()); }
      }
    }
    if (recipients.length > 0) {
      var subject = 'New Incident Report: ' + data.subject + ' - ' + data.reportedIndividual;
      var body = '<h2>New Incident Report</h2>' +
        '<p><strong>Date:</strong> ' + data.date + '</p>' +
        '<p><strong>Reported By:</strong> ' + data.reportedBy + '</p>' +
        '<p><strong>Reported Individual:</strong> ' + data.reportedIndividual + '</p>' +
        '<p><strong>Subject of the Incident:</strong> ' + data.subject + '</p>' +
        '<p><strong>Details of the Incident:</strong> ' + data.details + '</p>' +
        '<p><strong>Employee ID:</strong> ' + data.employeeId + '</p>' +
        '<p><strong>Policy:</strong> ' + data.policy + '</p>';
      MailApp.sendEmail({ to: recipients.join(','), subject: subject, htmlBody: body });
    }
  }
}
