// ============================================================
// Greenwood 100 Inc. — Google Apps Script Backend
// ============================================================
// SETUP (one time):
// 1. Go to script.google.com → New Project → paste this file
// 2. Replace SHEET_ID below with your Google Sheet ID
//    (from URL: docs.google.com/spreadsheets/d/SHEET_ID/edit)
// 3. Click Deploy → New Deployment
//    - Type: Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Copy the Web App URL
// 5. In index.html find: const SUBMIT_URL = '/api/submit';
//    Replace with:        const SUBMIT_URL = 'YOUR_WEB_APP_URL';
// ============================================================

const SHEET_ID = '1O2E2sSgEApsZofvddGLzWzMrEe2SDUBi--oUPuLyF6U';

const TAB_MAP = {
  'signup':             'Signups',
  'contact':            'Contact',
  'volunteer':          'Volunteers',
  'investor-interest':  'Investor Interest',
  'investor-register':  'Investor Register',
  'land-trust':         'Land Trust',
  'message':            'Messages',
  'task':               'Tasks',
  'receipt':            'Receipts',
  'otp':                'OTP',
};

const HEADERS = {
  'signup':             ['Timestamp','Row ID','First Name','Last Name','Phone','Email','Account Type','Business Name','Category','Address','Description','Website','Referral Program','Investor Wanted','Verified','Status'],
  'contact':            ['Timestamp','Row ID','Name','Email','Subject','Message','Status'],
  'volunteer':          ['Timestamp','Row ID','Name','Email','Phone','Opportunity','Status'],
  'investor-interest':  ['Timestamp','Row ID','Name','Email','Investment Range','Message','Business Interest','Status'],
  'investor-register':  ['Timestamp','Row ID','Name','Email','Investment Range','Areas of Interest','Status'],
  'land-trust':         ['Timestamp','Row ID','Name','Email','Phone','Interest Type','Message','Status'],
  'message':            ['Timestamp','Message ID','Thread ID','From','To','To Name','Subject','Message','Read','Flagged','Deleted'],
  'task':               ['Timestamp','Task ID','Title','Description','Due Date','Assigned To','Assigned By','Status','Notes','Last Updated'],
  'receipt':            ['Timestamp','Receipt ID','User Email','Business Name','Amount','Receipt Date','Image Hash','Image Data','Status','Points Awarded','Admin Notes'],
  'otp':                ['Timestamp','Email','OTP','Expires','Used'],
};

// ---- READ (admin dashboard fetches live data) ----
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'ping';
  let result;

  if (action === 'messages') {
    const userFilter = e.parameter.user || '';
    const threadFilter = e.parameter.thread || '';
    const showDeleted = e.parameter.showDeleted === 'true';
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName('Messages');
      if (!sheet || sheet.getLastRow() < 2) {
        result = { ok: true, rows: [] };
      } else {
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        let rows = data.slice(1).map(row => {
          const obj = {};
          headers.forEach((h, i) => { obj[h] = row[i]; });
          return obj;
        });
        if (!showDeleted) rows = rows.filter(r => r['Deleted'] !== 'Yes');
        if (threadFilter) {
          rows = rows.filter(r => r['Thread ID'] === threadFilter);
        } else if (userFilter) {
          rows = rows.filter(r => r['From'] === userFilter || r['To'] === userFilter);
        }
        result = { ok: true, rows };
      }
    } catch (err) {
      result = { ok: false, error: err.toString() };
    }
  } else if (action === 'list') {
    const sheetName = e.parameter.sheet || 'Signups';
    const filterStatus = e.parameter.status || ''; // e.g. 'Pending' or ''
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) {
        result = { ok: true, rows: [] };
      } else {
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const rows = data.slice(1).map(row => {
          const obj = {};
          headers.forEach((h, i) => { obj[h] = row[i]; });
          return obj;
        }).filter(row => {
          if (!filterStatus) return true;
          const s = row['Status'] || 'Pending';
          return s === filterStatus;
        });
        result = { ok: true, rows };
      }
    } catch (err) {
      result = { ok: false, error: err.toString() };
    }
  } else if (action === 'tasks') {
    const assignee = e.parameter.assignee || '';
    const all = e.parameter.all === 'true';
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName('Tasks');
      if (!sheet || sheet.getLastRow() < 2) {
        result = { ok: true, rows: [] };
      } else {
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        let rows = data.slice(1).map(row => {
          const obj = {};
          headers.forEach((h, i) => { obj[h] = row[i]; });
          return obj;
        });
        if (!all && assignee) {
          rows = rows.filter(r => r['Assigned To'] === assignee);
        }
        result = { ok: true, rows };
      }
    } catch(err) {
      result = { ok: false, error: err.toString() };
    }
  } else if (action === 'receipts') {
    const userFilter = e.parameter.user || '';
    const allFlag = e.parameter.all === 'true';
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName('Receipts');
      if (!sheet || sheet.getLastRow() < 2) { result = { ok: true, rows: [] }; }
      else {
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        let rows = data.slice(1).map(row => {
          const obj = {};
          headers.forEach((h, i) => { obj[h] = row[i]; });
          return obj;
        });
        if (!allFlag && userFilter) rows = rows.filter(r => r['User Email'] === userFilter);
        result = { ok: true, rows };
      }
    } catch(err) { result = { ok: false, error: err.toString() }; }
  } else if (action === 'send-otp') {
    const email = e.parameter.email || '';
    try {
      // Generate 6-digit OTP
      const otp = String(Math.floor(100000 + Math.random() * 900000));
      const expires = new Date(Date.now() + 10 * 60 * 1000).toISOString(); // 10 min
      const ts = new Date().toISOString();

      // Store in OTP sheet
      const ss = SpreadsheetApp.openById(SHEET_ID);
      let sheet = ss.getSheetByName('OTP');
      if (!sheet) {
        sheet = ss.insertSheet('OTP');
        const hr = sheet.getRange(1,1,1,5);
        hr.setValues([['Timestamp','Email','OTP','Expires','Used']]);
        hr.setFontWeight('bold').setBackground('#2d6a2d').setFontColor('#ffffff');
      }
      sheet.appendRow([ts, email, otp, expires, 'No']);

      // Send email
      MailApp.sendEmail({
        to: email,
        subject: 'Greenwood 100 — Your Login Code',
        body: `Your Greenwood 100 verification code is:\n\n${otp}\n\nThis code expires in 10 minutes. If you did not request this, please ignore this email.\n\n— Greenwood 100 Inc.`,
        htmlBody: `<div style="font-family:sans-serif;max-width:480px;margin:0 auto;padding:32px;background:#f9fafb;border-radius:12px;">
          <div style="text-align:center;margin-bottom:24px;">
            <h1 style="color:#2d6a2d;font-size:24px;margin:0;">Greenwood 100</h1>
          </div>
          <div style="background:white;border-radius:8px;padding:24px;text-align:center;">
            <p style="color:#374151;margin-bottom:16px;">Your verification code is:</p>
            <div style="font-size:40px;font-weight:bold;letter-spacing:8px;color:#2d6a2d;padding:16px;background:#f0fdf4;border-radius:8px;margin:0 auto 16px;">${otp}</div>
            <p style="color:#6b7280;font-size:14px;">Expires in 10 minutes</p>
          </div>
          <p style="color:#9ca3af;font-size:12px;text-align:center;margin-top:16px;">If you didn't request this code, you can safely ignore this email.</p>
        </div>`
      });
      result = { ok: true };
    } catch(err) { result = { ok: false, error: err.toString() }; }

  } else if (action === 'verify-otp') {
    const email = e.parameter.email || '';
    const otp = e.parameter.otp || '';
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName('OTP');
      if (!sheet) { result = { ok: false, error: 'No OTP found' }; }
      else {
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const emailCol = headers.indexOf('Email');
        const otpCol = headers.indexOf('OTP');
        const expiresCol = headers.indexOf('Expires');
        const usedCol = headers.indexOf('Used');
        let found = false;
        for (let i = data.length - 1; i >= 1; i--) {
          if (String(data[i][emailCol]).toLowerCase() === email.toLowerCase() &&
              String(data[i][otpCol]) === String(otp) &&
              data[i][usedCol] !== 'Yes') {
            const expires = new Date(data[i][expiresCol]);
            if (new Date() > expires) {
              result = { ok: false, error: 'OTP expired. Please request a new one.' };
            } else {
              // Mark as used
              sheet.getRange(i+1, usedCol+1).setValue('Yes');
              // Look up user in Signups sheet
              const signups = ss.getSheetByName('Signups');
              let userType = 'member';
              let firstName = '';
              let lastName = '';
              let businessName = '';
              if (signups) {
                const sdata = signups.getDataRange().getValues();
                const sheaders = sdata[0];
                const emailIdx = sheaders.indexOf('Email');
                const typeIdx = sheaders.indexOf('Account Type');
                const firstIdx = sheaders.indexOf('First Name');
                const lastIdx = sheaders.indexOf('Last Name');
                const bizIdx = sheaders.indexOf('Business Name');
                const statusIdx = sheaders.indexOf('Status');
                for (let j = sdata.length - 1; j >= 1; j--) {
                  if (String(sdata[j][emailIdx]).toLowerCase() === email.toLowerCase() && sdata[j][statusIdx] === 'Approved') {
                    userType = (sdata[j][typeIdx] || 'member').toLowerCase();
                    firstName = sdata[j][firstIdx] || '';
                    lastName = sdata[j][lastIdx] || '';
                    businessName = sdata[j][bizIdx] || '';
                    break;
                  }
                }
              }
              result = { ok: true, email, userType, firstName, lastName, businessName };
            }
            found = true;
            break;
          }
        }
        if (!found) result = { ok: false, error: 'Invalid OTP. Please try again.' };
      }
    } catch(err) { result = { ok: false, error: err.toString() }; }

  } else {
    result = { ok: true, message: 'Greenwood 100 API is running' };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- WRITE (form submissions + admin approve/reject) ----
function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    const formType = (body.formType || '').toLowerCase();

    // Admin status update (approve / reject)
    if (formType === 'update-status') {
      result = updateRowStatus(body.sheetName, body.rowId, body.status);

    } else if (formType === 'flag-message') {
      result = updateMessageField('Messages', body.rowId, 'Flagged', 'Yes');

    } else if (formType === 'delete-message') {
      result = updateMessageField('Messages', body.rowId, 'Deleted', 'Yes');

    } else if (formType === 'mark-read') {
      result = updateMessageField('Messages', body.rowId, 'Read', 'Yes');

    } else if (formType === 'create-task') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      let sheet = ss.getSheetByName('Tasks');
      if (!sheet) {
        sheet = ss.insertSheet('Tasks');
        const hr = sheet.getRange(1,1,1,HEADERS['task'].length);
        hr.setValues([HEADERS['task']]);
        hr.setFontWeight('bold').setBackground('#2d6a2d').setFontColor('#ffffff');
      }
      const ts = new Date().toISOString();
      const taskId = Date.now().toString();
      sheet.appendRow([ts, taskId, body.title||'', body.description||'', body.dueDate||'', body.assignedTo||'', body.assignedBy||'', 'Pending', '', ts]);
      result = { ok: true, taskId };

    } else if (formType === 'update-task-status') {
      // volunteer accepts or denies
      result = updateTaskField(body.taskId, 'Status', body.status);

    } else if (formType === 'update-task-notes') {
      // association adds follow-up notes
      result = updateTaskField(body.taskId, 'Notes', body.notes);

    } else if (formType === 'complete-task') {
      result = updateTaskField(body.taskId, 'Status', 'Completed');

    } else if (formType === 'upload-receipt') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      let sheet = ss.getSheetByName('Receipts');
      if (!sheet) {
        sheet = ss.insertSheet('Receipts');
        const hr = sheet.getRange(1,1,1,HEADERS['receipt'].length);
        hr.setValues([HEADERS['receipt']]);
        hr.setFontWeight('bold').setBackground('#2d6a2d').setFontColor('#ffffff');
      }
      // Check for duplicate hash
      const hash = body.imageHash || '';
      if (hash) {
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const hashCol = headers.indexOf('Image Hash');
        if (hashCol >= 0) {
          for (let i = 1; i < data.length; i++) {
            if (data[i][hashCol] === hash) {
              result = { ok: false, duplicate: true, error: 'Duplicate receipt detected.' };
              return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
            }
          }
        }
      }
      const ts = new Date().toISOString();
      const receiptId = Date.now().toString();
      sheet.appendRow([ts, receiptId, body.userEmail||'', body.businessName||'', body.amount||'', body.receiptDate||'', hash, body.imageData||'', 'Pending', '', '']);
      result = { ok: true, receiptId };

    } else if (formType === 'update-receipt-status') {
      // admin approves/rejects receipt
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName('Receipts');
      if (!sheet) { result = { ok: false, error: 'Receipts sheet not found' }; }
      else {
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const idCol = headers.indexOf('Receipt ID');
        const statusCol = headers.indexOf('Status');
        const pointsCol = headers.indexOf('Points Awarded');
        const notesCol = headers.indexOf('Admin Notes');
        let found = false;
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][idCol]) === String(body.receiptId)) {
            sheet.getRange(i+1, statusCol+1).setValue(body.status);
            if (body.points && pointsCol >= 0) sheet.getRange(i+1, pointsCol+1).setValue(body.points);
            if (body.notes && notesCol >= 0) sheet.getRange(i+1, notesCol+1).setValue(body.notes);
            found = true;
            break;
          }
        }
        result = found ? { ok: true } : { ok: false, error: 'Receipt not found' };
      }

    } else {
      // Form submission
      const tab = TAB_MAP[formType];
      if (!tab) {
        result = { ok: false, error: 'Unknown formType: ' + formType };
      } else {
        const ss = SpreadsheetApp.openById(SHEET_ID);
        let sheet = ss.getSheetByName(tab);
        if (!sheet) {
          sheet = ss.insertSheet(tab);
          const headerRow = sheet.getRange(1, 1, 1, HEADERS[formType].length);
          headerRow.setValues([HEADERS[formType]]);
          headerRow.setFontWeight('bold')
                   .setBackground('#2d6a2d')
                   .setFontColor('#ffffff');
        }

        const ts = new Date().toISOString();
        const rowId = Date.now().toString();

        let row;
        switch (formType) {
          case 'signup':
            row = [ts, rowId, body.firstName||'', body.lastName||'', body.phone||'',
                   body.email||'', body.accountType||'', body.businessName||'',
                   body.businessCategory||'', body.address||'', body.description||'',
                   body.website||'', body.referralProgram||'',
                   body.investorWanted||'No', body.verified ? 'Yes' : 'No', 'Pending'];
            break;
          case 'contact':
            row = [ts, rowId, body.name||'', body.email||'', body.subject||'', body.message||'', 'New'];
            break;
          case 'volunteer':
            row = [ts, rowId, body.name||'', body.email||'', body.phone||'', body.opportunity||'', 'New'];
            break;
          case 'investor-interest':
            row = [ts, rowId, body.name||'', body.email||'', body.investmentRange||'', body.message||'', body.business||'', 'New'];
            break;
          case 'investor-register':
            row = [ts, rowId, body.name||'', body.email||'', body.investmentRange||'', body.areasOfInterest||'', 'New'];
            break;
          case 'land-trust':
            row = [ts, rowId, body.name||'', body.email||'', body.phone||'', body.interest||'', body.message||'', 'New'];
            break;
          case 'message':
            row = [ts, rowId,
              body.threadId || '',
              body.from || '',
              body.to || '',
              body.toName || '',
              body.subject || '',
              body.message || '',
              'No', 'No', 'No'];
            break;
          default:
            row = [ts, rowId, JSON.stringify(body), '', '', '', '', 'New'];
        }

        sheet.appendRow(row);
        result = { ok: true, rowId };
      }
    }
  } catch (err) {
    result = { ok: false, error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function updateTaskField(taskId, fieldName, value) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Tasks');
  if (!sheet) return { ok: false, error: 'Tasks sheet not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('Task ID');
  const fieldCol = headers.indexOf(fieldName);
  const updatedCol = headers.indexOf('Last Updated');
  if (idCol < 0 || fieldCol < 0) return { ok: false, error: 'Column not found' };
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(taskId)) {
      sheet.getRange(i+1, fieldCol+1).setValue(value);
      if (updatedCol >= 0) sheet.getRange(i+1, updatedCol+1).setValue(new Date().toISOString());
      return { ok: true };
    }
  }
  return { ok: false, error: 'Task not found: ' + taskId };
}

function updateMessageField(sheetName, rowId, fieldName, value) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { ok: false, error: 'Sheet not found: ' + sheetName };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('Message ID');
  const fieldCol = headers.indexOf(fieldName);
  if (idCol < 0 || fieldCol < 0) return { ok: false, error: 'Column not found' };
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(rowId)) {
      sheet.getRange(i + 1, fieldCol + 1).setValue(value);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Row not found: ' + rowId };
}

function updateRowStatus(sheetName, rowId, newStatus) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { ok: false, error: 'Sheet not found: ' + sheetName };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rowIdCol = headers.indexOf('Row ID');
  const statusCol = headers.indexOf('Status');
  if (rowIdCol < 0 || statusCol < 0) return { ok: false, error: 'Missing Row ID or Status column' };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][rowIdCol]) === String(rowId)) {
      sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Row not found: ' + rowId };
}
