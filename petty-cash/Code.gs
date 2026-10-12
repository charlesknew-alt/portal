// ═══════════════════════════════════════════════════════════════
// ABSOLUTE LIVING — PETTY CASH v2
// Google Apps Script (Web App)
// ═══════════════════════════════════════════════════════════════
//
// SETUP:
// 1. Run setupSheet() once to create the Google Sheet and Receipts folder
// 2. Go to Project Settings (gear icon) → Script Properties → Add:
//    Key: ANTHROPIC_API_KEY   Value: your-api-key-here
// 3. Deploy as Web App
// ═══════════════════════════════════════════════════════════════

var SHEET_NAME = 'Claims';
var STAFF_SHEET = 'Staff';
var SETTINGS_SHEET = 'Mileage Profiles';
var SUMMARY_SHEET = 'Quarterly Summary';
var SUBMISSIONS_SHEET = 'Submissions';
var OWNER_EMAIL = 'charles@charlesknew.com';

// ─── SETUP ───

function setupSheet() {
  var ss = SpreadsheetApp.create('Absolute Living — Petty Cash');
  var id = ss.getId();

  var sheet = ss.getSheetByName('Sheet1');
  sheet.setName(SHEET_NAME);
  sheet.getRange(1, 1, 1, 10).setValues([['Date','Description','Staff','Amount','Category','Fuel £/L','Quarter','Settled','Ref','Receipt']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 100); sheet.setColumnWidth(2, 200); sheet.setColumnWidth(3, 100); sheet.setColumnWidth(10, 200);
  sheet.getRange('A:A').setNumberFormat('dd/MM/yyyy');
  sheet.getRange('D:D').setNumberFormat('#,##0.00');
  sheet.getRange('F:F').setNumberFormat('0.000');

  var st = ss.insertSheet(STAFF_SHEET);
  st.getRange(1, 1, 1, 4).setValues([['ID','Name','Email','Active']]).setFontWeight('bold');
  st.setColumnWidth(3, 220);
  st.getRange(2, 1, 3, 4).setValues([['charlie','Charlie','','YES'],['harvey','Harvey','','YES'],['shakil','Shakil','','YES']]);

  var mp = ss.insertSheet(SETTINGS_SHEET);
  mp.getRange(1, 1, 1, 5).setValues([['Staff ID','Miles Return','MPG','Trips/Week','Annual Miles']]).setFontWeight('bold');
  mp.getRange(2, 1, 2, 5).setValues([['harvey',50,28,3,'=B2*D2*52'],['shakil',11,30,6,'=B3*D3*52']]);

  var qs = ss.insertSheet(SUMMARY_SHEET);
  qs.getRange(1, 1, 1, 8).setValues([['Quarter','Staff','Gross Claims','Avg Fuel £/L','Personal Mileage Cost','Net Payable','Settled Date','Entries']]).setFontWeight('bold');
  qs.setFrozenRows(1);

  var sb = ss.insertSheet(SUBMISSIONS_SHEET);
  sb.getRange(1, 1, 1, 6).setValues([['Quarter','Staff ID','Staff Name','Date Submitted','Claim Count','Total Amount']]).setFontWeight('bold');
  sb.setFrozenRows(1);

  var folder = DriveApp.createFolder('AL Petty Cash — Receipts');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
  PropertiesService.getScriptProperties().setProperty('RECEIPTS_FOLDER_ID', folder.getId());
  Logger.log('Sheet: https://docs.google.com/spreadsheets/d/' + id);
  Logger.log('Receipts folder: ' + folder.getUrl());
}

// ─── WEB APP ───

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('PettyCash')
    .setTitle('Petty Cash — Absolute Living')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// ─── RECEIPT VALIDATION ───

function validateReceipt(base64Data, mimeType) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) {
    Logger.log('No ANTHROPIC_API_KEY set — skipping validation');
    return { valid: true, reason: 'Validation skipped — no API key configured' };
  }

  try {
    var mediaType = mimeType || 'image/jpeg';
    var payload = {
      model: 'claude-sonnet-4-20250514',
      max_tokens: 200,
      messages: [{
        role: 'user',
        content: [
          {
            type: 'image',
            source: { type: 'base64', media_type: mediaType, data: base64Data }
          },
          {
            type: 'text',
            text: 'You are a receipt validator for a petty cash system. Look at this image and determine if it is a genuine receipt, invoice, or proof of purchase. Reply with ONLY a JSON object, no other text:\n{"valid": true, "reason": "brief description"} if it IS a receipt/invoice/proof of purchase\n{"valid": false, "reason": "brief explanation why"} if it is NOT (e.g. blank page, random photo, screenshot of something unrelated, selfie, photo of a non-receipt item)\nBe reasonable — handwritten receipts, till receipts, fuel pump receipts, card machine printouts, and invoices are all acceptable. Slightly blurry but readable is fine.'
          }
        ]
      }]
    };

    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var result = JSON.parse(response.getContentText());
    if (result.content && result.content[0] && result.content[0].text) {
      var text = result.content[0].text.trim();
      // Extract JSON from response
      var jsonMatch = text.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        var parsed = JSON.parse(jsonMatch[0]);
        return { valid: !!parsed.valid, reason: parsed.reason || '' };
      }
    }
    // If we can't parse, allow it through
    return { valid: true, reason: 'Could not parse validation response' };
  } catch (e) {
    Logger.log('Receipt validation error: ' + e.toString());
    // On error, allow through rather than blocking
    return { valid: true, reason: 'Validation error — allowed through' };
  }
}

// ─── RECEIPT UPLOAD ───

function uploadReceipt(base64Data, filename, mimeType) {
  var folderId = PropertiesService.getScriptProperties().getProperty('RECEIPTS_FOLDER_ID');
  if (!folderId) {
    var folder = DriveApp.createFolder('AL Petty Cash — Receipts');
    folderId = folder.getId();
    PropertiesService.getScriptProperties().setProperty('RECEIPTS_FOLDER_ID', folderId);
  }
  var folder = DriveApp.getFolderById(folderId);
  var decoded = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(decoded, mimeType, filename);
  var file = folder.createFile(blob);
  return { id: file.getId(), url: file.getUrl() };
}

// ─── STAFF ───

function getStaffList() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(STAFF_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var staff = [];
  for (var i = 0; i < values.length; i++) {
    staff.push({ id: values[i][0], name: values[i][1], email: values[i][2] || '', active: values[i][3] === 'YES' });
  }
  return staff;
}

function addStaff(data) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(STAFF_SHEET);
  var id = data.name.toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
  var existing = getStaffList();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].id === id) throw new Error('"' + data.name + '" already exists');
  }
  sheet.appendRow([id, data.name, data.email || '', 'YES']);
  return { status: 'ok', id: id };
}

function updateStaffEmail(data) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(STAFF_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No staff');
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === data.staffId) { sheet.getRange(i + 2, 3).setValue(data.email); return { status: 'ok' }; }
  }
  throw new Error('Staff not found');
}

function toggleStaff(staffId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(STAFF_SHEET);
  var lastRow = sheet.getLastRow();
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === staffId) {
      var current = sheet.getRange(i + 2, 4).getValue();
      sheet.getRange(i + 2, 4).setValue(current === 'YES' ? 'NO' : 'YES');
      return { status: 'ok', active: current !== 'YES' };
    }
  }
  throw new Error('Staff not found');
}

function removeStaff(staffId) {
  var ss = getSpreadsheet();
  var claims = ss.getSheetByName(SHEET_NAME);
  var claimsLast = claims.getLastRow();
  if (claimsLast >= 2) {
    var staffCol = claims.getRange(2, 3, claimsLast - 1, 1).getValues();
    for (var j = 0; j < staffCol.length; j++) {
      if (staffCol[j][0] === staffId) throw new Error('Has existing claims — deactivate instead.');
    }
  }
  var sheet = ss.getSheetByName(STAFF_SHEET);
  var lastRow = sheet.getLastRow();
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === staffId) { sheet.deleteRow(i + 2); return { status: 'ok' }; }
  }
  throw new Error('Staff not found');
}

// ─── MILEAGE ───

function getProfiles() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  var values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var profiles = {};
  for (var i = 0; i < values.length; i++) {
    if (values[i][0]) profiles[values[i][0]] = { milesReturn: values[i][1], mpg: values[i][2], tripsPerWeek: values[i][3] };
  }
  return profiles;
}

function saveProfile(data) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] === data.staffId) {
        sheet.getRange(i + 2, 2, 1, 3).setValues([[data.milesReturn, data.mpg, data.tripsPerWeek]]);
        sheet.getRange(i + 2, 5).setFormula('=B' + (i+2) + '*D' + (i+2) + '*52');
        return { status: 'ok' };
      }
    }
  }
  var nr = lastRow + 1;
  sheet.appendRow([data.staffId, data.milesReturn, data.mpg, data.tripsPerWeek]);
  sheet.getRange(nr, 5).setFormula('=B' + nr + '*D' + nr + '*52');
  return { status: 'ok' };
}

function removeProfile(staffId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No profiles');
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === staffId) { sheet.deleteRow(i + 2); return { status: 'ok' }; }
  }
  throw new Error('Profile not found');
}

// ─── CLAIMS ───

function getNextRef() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return '001';
  var refs = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < refs.length; i++) {
    var val = parseInt(refs[i][0], 10);
    if (!isNaN(val) && val > max) max = val;
  }
  return (max + 1).toString().padStart(3, '0');
}

function submitClaim(data) {
  // Validate receipt first
  if (!data.receiptData) throw new Error('Receipt photo is required');

  var validation = validateReceipt(data.receiptData, data.receiptType || 'image/jpeg');
  if (!validation.valid) {
    throw new Error('Receipt rejected: ' + validation.reason);
  }

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var date = new Date(data.date);
  var quarter = getQuarterLabel(date);
  var ref = getNextRef();

  var receiptResult = uploadReceipt(data.receiptData, ref + '_receipt.' + (data.receiptType === 'image/png' ? 'png' : 'jpg'), data.receiptType || 'image/jpeg');
  var receiptUrl = receiptResult.url;

  sheet.appendRow([date, data.description || data.category, data.staff, data.amount, data.category, data.fuelPrice || '', quarter, '', ref, receiptUrl]);
  return { status: 'ok', quarter: quarter, ref: ref };
}

function getEntries() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { entries: [], summary: {} };
  var values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var entries = [], summary = {};
  for (var i = 0; i < values.length; i++) {
    var row = values[i], date = row[0], staffId = row[2], amount = row[3], settled = row[7] === 'YES';
    if (staffId && date && amount) {
      entries.push({
        rowIndex: i + 2, date: Utilities.formatDate(new Date(date), 'Europe/London', 'yyyy-MM-dd'),
        description: row[1], staff: staffId, amount: amount, category: row[4],
        fuelPrice: row[5] || null, quarter: row[6], settled: settled, ref: row[8] || '',
        receipt: row[9] || ''
      });
      var q = row[6];
      if (q && !settled) {
        if (!summary[q]) summary[q] = {};
        if (!summary[q][staffId]) summary[q][staffId] = { total: 0, fuelPrices: [], count: 0 };
        summary[q][staffId].total += amount;
        summary[q][staffId].count++;
        if (row[5] && row[5] !== '') summary[q][staffId].fuelPrices.push(row[5]);
      }
    }
  }
  return { entries: entries, summary: summary };
}

function deleteEntry(rowIndex) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (rowIndex < 2) throw new Error('Invalid row');
  if (sheet.getRange(rowIndex, 8).getValue() === 'YES') throw new Error('Cannot delete settled entry');
  sheet.deleteRow(rowIndex);
  return { status: 'ok' };
}

// ─── QUARTERLY SUBMISSION ───

function getSubmissions() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var subs = [];
  for (var i = 0; i < values.length; i++) {
    subs.push({
      quarter: values[i][0], staffId: values[i][1], staffName: values[i][2],
      date: values[i][3] ? Utilities.formatDate(new Date(values[i][3]), 'Europe/London', 'dd/MM/yyyy') : '',
      count: values[i][4], total: values[i][5]
    });
  }
  return subs;
}

function submitQuarter(data) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SUBMISSIONS_SHEET);
    sheet.getRange(1, 1, 1, 6).setValues([['Quarter','Staff ID','Staff Name','Date Submitted','Claim Count','Total Amount']]).setFontWeight('bold');
  }
  var existing = getSubmissions();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].quarter === data.quarter && existing[i].staffId === data.staffId) throw new Error('Already submitted for ' + data.quarter);
  }
  var staffList = getStaffList();
  var staffName = data.staffId, staffEmail = '';
  for (var j = 0; j < staffList.length; j++) {
    if (staffList[j].id === data.staffId) { staffName = staffList[j].name; staffEmail = staffList[j].email; break; }
  }
  var now = new Date();
  sheet.appendRow([data.quarter, data.staffId, staffName, now, data.count, data.total]);

  try {
    MailApp.sendEmail({
      to: OWNER_EMAIL,
      subject: 'Petty Cash — ' + staffName + ' submitted ' + data.quarter + ' (£' + Number(data.total).toFixed(2) + ')',
      htmlBody: '<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto">' +
        '<div style="background:#6B2D5B;padding:20px;border-radius:10px 10px 0 0"><div style="color:#fff;font-size:18px;font-weight:700">Quarter Submission</div><div style="color:#B87DA8;font-size:12px;letter-spacing:1px;text-transform:uppercase;margin-top:2px">Absolute Living — Petty Cash</div></div>' +
        '<div style="background:#fff;padding:24px;border:1px solid #E5E0E3;border-top:none;border-radius:0 0 10px 10px">' +
          '<p style="font-size:14px;color:#3A2832;margin-bottom:16px"><strong>' + staffName + '</strong> has submitted their claims for <strong>' + data.quarter + '</strong>.</p>' +
          '<table style="width:100%;border-collapse:collapse;margin-bottom:16px"><tr><td style="padding:8px 0;font-size:14px;color:#998E94">Claims</td><td style="padding:8px 0;font-size:14px;text-align:right;font-weight:600">' + data.count + '</td></tr><tr><td style="padding:8px 0;font-size:14px;color:#998E94">Gross Total</td><td style="padding:8px 0;font-size:14px;text-align:right;font-weight:600">£' + Number(data.total).toFixed(2) + '</td></tr></table>' +
          '<p style="font-size:13px;color:#998E94">Log in to settle when all staff have submitted.</p>' +
          '<p style="font-size:12px;color:#CCC5C9;margin-top:16px">Submitted ' + Utilities.formatDate(now, 'Europe/London', 'dd/MM/yyyy HH:mm') + '</p>' +
        '</div></div>'
    });
  } catch (e) { Logger.log('Email error: ' + e); }

  return { status: 'ok', date: Utilities.formatDate(now, 'Europe/London', 'dd/MM/yyyy') };
}

// ─── SETTLEMENT ───

function settleQuarter(data) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var summarySheet = ss.getSheetByName(SUMMARY_SHEET);
  var settledDate = Utilities.formatDate(new Date(), 'Europe/London', 'dd/MM/yyyy');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No entries');

  var quarters = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
  var settledCol = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
  var count = 0;
  for (var i = 0; i < quarters.length; i++) {
    if (quarters[i][0] === data.quarter && settledCol[i][0] !== 'YES') { sheet.getRange(i + 2, 8).setValue('YES'); count++; }
  }

  var staffList = getStaffList();
  var nameMap = {}, emailMap = {};
  staffList.forEach(function(s) { nameMap[s.id] = s.name; emailMap[s.id] = s.email; });

  var allEntries = getEntries().entries;
  var qtrEntries = {};
  allEntries.forEach(function(e) { if (e.quarter === data.quarter) { if (!qtrEntries[e.staff]) qtrEntries[e.staff] = []; qtrEntries[e.staff].push(e); } });

  if (data.staffSummary) {
    data.staffSummary.forEach(function(s) {
      summarySheet.appendRow([data.quarter, nameMap[s.staff] || s.staff, s.gross, s.avgFuelPrice || '', s.mileageDeduction || 0, s.net, settledDate, s.entryCount || 0]);
    });
  }

  try {
    if (data.staffSummary) {
      data.staffSummary.forEach(function(s) {
        var email = emailMap[s.staff]; var name = nameMap[s.staff] || s.staff;
        if (!email) return;
        var entries = qtrEntries[s.staff] || [];
        var claimLines = '';
        entries.forEach(function(e) {
          claimLines += '<tr><td style="padding:6px 10px;border-bottom:1px solid #eee;font-size:13px">' + e.ref + '</td><td style="padding:6px 10px;border-bottom:1px solid #eee;font-size:13px">' + e.date + '</td><td style="padding:6px 10px;border-bottom:1px solid #eee;font-size:13px">' + e.description + '</td><td style="padding:6px 10px;border-bottom:1px solid #eee;font-size:13px;text-align:right">£' + Number(e.amount).toFixed(2) + '</td></tr>';
        });
        MailApp.sendEmail({ to: email, subject: 'Petty Cash Settlement — ' + data.quarter + ' — £' + Number(s.net).toFixed(2),
          htmlBody: '<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto"><div style="background:#6B2D5B;padding:20px;border-radius:10px 10px 0 0"><div style="color:#fff;font-size:18px;font-weight:700">Petty Cash Settlement</div><div style="color:#B87DA8;font-size:12px;letter-spacing:1px;text-transform:uppercase;margin-top:2px">Absolute Living</div></div><div style="background:#fff;padding:24px;border:1px solid #E5E0E3;border-top:none;border-radius:0 0 10px 10px"><p style="font-size:14px;color:#3A2832;margin-bottom:16px">Hi ' + name + ',</p><p style="font-size:14px;color:#3A2832;margin-bottom:16px">Your petty cash for <strong>' + data.quarter + '</strong> has been settled:</p><table style="width:100%;border-collapse:collapse;margin-bottom:16px"><thead><tr style="background:#F5EEF3"><th style="padding:8px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#998E94">Ref</th><th style="padding:8px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#998E94">Date</th><th style="padding:8px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#998E94">Description</th><th style="padding:8px 10px;text-align:right;font-size:10px;text-transform:uppercase;color:#998E94">Amount</th></tr></thead><tbody>' + claimLines + '</tbody></table><table style="width:100%;border-collapse:collapse;margin-bottom:16px"><tr><td style="padding:6px 10px;font-size:14px">Gross Claims</td><td style="padding:6px 10px;font-size:14px;text-align:right;font-weight:600">£' + Number(s.gross).toFixed(2) + '</td></tr>' + (s.mileageDeduction > 0 ? '<tr><td style="padding:6px 10px;font-size:14px;color:#998E94">Mileage Deduction' + (s.avgFuelPrice ? ' (avg £' + Number(s.avgFuelPrice).toFixed(3) + '/L)' : '') + '</td><td style="padding:6px 10px;font-size:14px;text-align:right;color:#C0392B">−£' + Number(s.mileageDeduction).toFixed(2) + '</td></tr>' : '') + '<tr style="border-top:2px solid #6B2D5B"><td style="padding:10px;font-size:16px;font-weight:700;color:#6B2D5B">Net Payable</td><td style="padding:10px;font-size:16px;font-weight:700;text-align:right;color:#6B2D5B">£' + Number(s.net).toFixed(2) + '</td></tr></table><p style="font-size:12px;color:#998E94">Settled on ' + settledDate + '</p></div></div>'
        });
      });
    }

    var ownerRows = '', totalNet = 0;
    if (data.staffSummary) {
      data.staffSummary.forEach(function(s) {
        totalNet += s.net;
        ownerRows += '<tr><td style="padding:8px 10px;border-bottom:1px solid #eee;font-size:13px">' + (nameMap[s.staff] || s.staff) + '</td><td style="padding:8px 10px;border-bottom:1px solid #eee;font-size:13px;text-align:right">£' + Number(s.gross).toFixed(2) + '</td><td style="padding:8px 10px;border-bottom:1px solid #eee;font-size:13px;text-align:right;color:#998E94">' + (s.avgFuelPrice ? '£' + Number(s.avgFuelPrice).toFixed(3) : '–') + '</td><td style="padding:8px 10px;border-bottom:1px solid #eee;font-size:13px;text-align:right;color:#C0392B">' + (s.mileageDeduction > 0 ? '(£' + Number(s.mileageDeduction).toFixed(2) + ')' : '–') + '</td><td style="padding:8px 10px;border-bottom:1px solid #eee;font-size:13px;text-align:right;font-weight:700">£' + Number(s.net).toFixed(2) + '</td></tr>';
      });
    }
    MailApp.sendEmail({ to: OWNER_EMAIL, subject: 'Petty Cash Settled — ' + data.quarter + ' — Total £' + Number(totalNet).toFixed(2),
      htmlBody: '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto"><div style="background:#6B2D5B;padding:20px;border-radius:10px 10px 0 0"><div style="color:#fff;font-size:18px;font-weight:700">Settlement Summary</div><div style="color:#B87DA8;font-size:12px;letter-spacing:1px;text-transform:uppercase;margin-top:2px">Absolute Living — Petty Cash</div></div><div style="background:#fff;padding:24px;border:1px solid #E5E0E3;border-top:none;border-radius:0 0 10px 10px"><p style="font-size:14px;margin-bottom:16px"><strong>' + data.quarter + '</strong> settled on ' + settledDate + '. ' + count + ' entries.</p><table style="width:100%;border-collapse:collapse;margin-bottom:16px"><thead><tr style="background:#F5EEF3"><th style="padding:8px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#998E94">Staff</th><th style="padding:8px 10px;text-align:right;font-size:10px;text-transform:uppercase;color:#998E94">Gross</th><th style="padding:8px 10px;text-align:right;font-size:10px;text-transform:uppercase;color:#998E94">Avg £/L</th><th style="padding:8px 10px;text-align:right;font-size:10px;text-transform:uppercase;color:#998E94">Mileage</th><th style="padding:8px 10px;text-align:right;font-size:10px;text-transform:uppercase;color:#998E94">Net</th></tr></thead><tbody>' + ownerRows + '</tbody><tfoot><tr style="border-top:2px solid #6B2D5B"><td style="padding:10px;font-weight:700;color:#6B2D5B">Total</td><td colspan="3"></td><td style="padding:10px;font-weight:700;text-align:right;color:#6B2D5B">£' + Number(totalNet).toFixed(2) + '</td></tr></tfoot></table></div></div>'
    });
  } catch (emailErr) { Logger.log('Email error: ' + emailErr); }

  return { status: 'ok', settled: count, date: settledDate };
}

// ─── HELPERS ───

function getSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Run setupSheet() first');
  return SpreadsheetApp.openById(id);
}

function getQuarterLabel(date) {
  var m = date.getMonth(), y = date.getFullYear();
  var labels = ['Q1 Jan-Mar','Q2 Apr-Jun','Q3 Jul-Sep','Q4 Oct-Dec'];
  return y + ' ' + labels[Math.floor(m / 3)];
}
