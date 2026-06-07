// ============================================================
// Greenwood 100 Inc. — Google Apps Script Backend
// ============================================================
// SETUP INSTRUCTIONS:
// 1. Go to script.google.com → New Project
// 2. Paste this entire file, replacing the default code
// 3. Click Extensions → Apps Script → Save
// 4. Click Deploy → New Deployment
//    - Type: Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Click Deploy → Copy the Web App URL
// 6. In index.html, set SUBMIT_URL at the top of the <script>
//    to the URL you copied
// ============================================================

const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE'; // Replace with your Sheet ID from the URL

const TAB_MAP = {
  'signup':             'Signups',
  'contact':            'Contact',
  'volunteer':          'Volunteers',
  'investor-interest':  'Investor Interest',
  'investor-register':  'Investor Register',
  'land-trust':         'Land Trust',
};

const HEADERS = {
  'signup':             ['Timestamp', 'First Name', 'Last Name', 'Phone', 'Email', 'Account Type', 'Business Name', 'Category', 'Referral Program', 'Investor Wanted', 'Verified'],
  'contact':            ['Timestamp', 'Name', 'Email', 'Subject', 'Message'],
  'volunteer':          ['Timestamp', 'Name', 'Email', 'Phone', 'Opportunity'],
  'investor-interest':  ['Timestamp', 'Name', 'Email', 'Investment Range', 'Message', 'Business Interest'],
  'investor-register':  ['Timestamp', 'Name', 'Email', 'Investment Range', 'Areas of Interest'],
  'land-trust':         ['Timestamp', 'Name', 'Email', 'Phone', 'Interest Type', 'Message'],
};

function doPost(e) {
  const cors = ContentService.createTextOutput();
  cors.setMimeType(ContentService.MimeType.JSON);

  try {
    const body = JSON.parse(e.postData.contents);
    const formType = (body.formType || '').toLowerCase();
    const tab = TAB_MAP[formType];

    if (!tab) {
      cors.setContent(JSON.stringify({ ok: false, error: 'Unknown formType: ' + formType }));
      return cors;
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(tab);

    // Create the tab with headers if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(tab);
      sheet.appendRow(HEADERS[formType] || ['Timestamp', 'Data']);
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold').setBackground('#2d6a2d').setFontColor('#ffffff');
    }

    const ts = new Date().toISOString();
    let row;

    switch (formType) {
      case 'signup':
        row = [ts, body.firstName||'', body.lastName||'', body.phone||'', body.email||'',
               body.accountType||'', body.businessName||'', body.businessCategory||'',
               body.referralProgram||'', body.investorWanted||'No', body.verified ? 'Yes' : 'No'];
        break;
      case 'contact':
        row = [ts, body.name||'', body.email||'', body.subject||'', body.message||''];
        break;
      case 'volunteer':
        row = [ts, body.name||'', body.email||'', body.phone||'', body.opportunity||''];
        break;
      case 'investor-interest':
        row = [ts, body.name||'', body.email||'', body.investmentRange||'', body.message||'', body.business||''];
        break;
      case 'investor-register':
        row = [ts, body.name||'', body.email||'', body.investmentRange||'', body.areasOfInterest||''];
        break;
      case 'land-trust':
        row = [ts, body.name||'', body.email||'', body.phone||'', body.interest||'', body.message||''];
        break;
      default:
        row = [ts, JSON.stringify(body)];
    }

    sheet.appendRow(row);
    cors.setContent(JSON.stringify({ ok: true }));

  } catch (err) {
    cors.setContent(JSON.stringify({ ok: false, error: err.toString() }));
  }

  return cors;
}

// This handles browser preflight OPTIONS requests (CORS)
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, message: 'Greenwood 100 API running' }))
    .setMimeType(ContentService.MimeType.JSON);
}
