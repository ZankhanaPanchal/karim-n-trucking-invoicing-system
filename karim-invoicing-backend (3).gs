/**
 * KARIM N TRUCKING LTD — INVOICING SYSTEM BACKEND (v2 — full app database)
 * --------------------------------------------------------------------------
 * Same family as your POD system's script, but this one is the ENTIRE
 * database for the invoicing app — drivers, PINs, settings, and every
 * invoice — not just an email trigger. The HTML app has no storage of
 * its own; it reads and writes everything here.
 *
 * SETUP (about 10 minutes, one time):
 * 1. Go to sheets.google.com and create a new blank spreadsheet.
 *    Name it "Karim N Trucking — Invoices".
 * 2. Extensions > Apps Script.
 * 3. Delete the placeholder code and paste this entire file in.
 * 4. Near the top, set OPS_EMAIL to the address that should receive
 *    every invoice email.
 * 5. Save (Ctrl+S / Cmd+S).
 * 6. Deploy > New deployment > gear icon > pick type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Click Deploy, then click through the authorization prompt
 *    (Advanced > "Go to [project name] (unsafe)" — normal for a
 *    script you wrote yourself).
 * 7. Copy the Web app URL it gives you.
 * 8. Open karim-invoicing-app.html in a text editor, find
 *    CONFIG.SCRIPT_URL near the top, and paste that URL in. Save.
 * 9. Run authorizeGmailAccess once from the Apps Script editor (see
 *    the bottom of this file) — same as your POD system's setup —
 *    so Gmail sending is allowed.
 * 10. Host karim-invoicing-app.html the same way you hosted the POD
 *     page (e.g. drag-and-drop onto Netlify) and open that link.
 *
 * The default operations PIN is 1234 — change it from inside the
 * app (Operations Portal > Settings) once you're set up.
 */

const OPS_EMAIL = 'invoice2karimntrucking@gmail.com';

// Leave this blank if this script is opened via Extensions > Apps Script
// from inside the Sheet itself (it'll use that Sheet automatically).
// If you created this as a standalone project at script.google.com
// instead, paste the Sheet's ID here — it's the long string in the
// Sheet's URL: https://docs.google.com/spreadsheets/d/THIS_PART/edit
const SPREADSHEET_ID = '1AHSkSohPJC67lo14j_kaORXLiqWSPesvR_uDlNr_SLQ';

function getSpreadsheet_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

const DRIVERS_SHEET = 'Drivers';
const SUMMARY_SHEET = 'Invoice Summary';
const LOADS_SHEET = 'Invoice Loads';

const DRIVERS_HEADERS = ['ID', 'Name', 'Email', 'PIN', 'Active', 'Commission %', 'Company', 'Next Invoice #'];
const SUMMARY_HEADERS = [
  'Invoice #', 'Invoice ID', 'Driver ID', 'Driver Name', 'Company', 'Period Start', 'Period End', 'Loads',
  'Actual Total', 'Commission Total', 'Net Total', 'Tax Total', 'Driver Pay Total',
  'Tax Basis', 'Submitted At', 'Emailed', 'PDF Link', 'Edited At'
];
const LOADS_HEADERS = [
  'Invoice ID', 'Driver ID', 'Driver Name', 'Load Date', 'Ref Type', 'Ref Number', 'Pickup', 'Drop-off(s)',
  'Province', 'Actual Amount', 'Commission %', 'Commission Amount', 'Net After Commission',
  'Tax Rate', 'Tax Amount', 'Driver Pay'
];

// ===================== ROUTES =====================

function doGet(e) {
  try {
    return jsonOut({ success: true, drivers: readDrivers(), invoices: readInvoices(), settings: readSettings() });
  } catch (err) {
    return jsonOut({ success: false, message: String(err) });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    switch (payload.action) {
      case 'submitInvoice': return submitInvoice(payload);
      case 'updateInvoice': return updateInvoice(payload);
      case 'addDriver': return addDriver(payload);
      case 'updateDriver': return updateDriver(payload);
      case 'removeDriver': return removeDriver(payload);
      case 'updateSettings': return updateSettings(payload);
      default: return jsonOut({ success: false, message: 'Unknown action: ' + payload.action });
    }
  } catch (err) {
    return jsonOut({ success: false, message: String(err) });
  }
}

// ===================== SHEET HELPERS =====================

function getOrCreateSheet(name, headers) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    const headerRange = sh.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#33326C');
    headerRange.setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, headers.length);
  }
  return sh;
}

function formatDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v;
}

function findRowById_(sh, id) {
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

// ===================== DRIVERS =====================

function readDrivers() {
  const sh = getOrCreateSheet(DRIVERS_SHEET, DRIVERS_HEADERS);
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    out.push({
      id: String(r[0]), name: r[1] || '', email: r[2] || '', pin: String(r[3] || ''), active: String(r[4]) !== 'No',
      commission: (r[5] === '' || r[5] === null || r[5] === undefined) ? 20 : Number(r[5]),
      company: r[6] || ''
    });
  }
  return out;
}

function addDriver(payload) {
  const sh = getOrCreateSheet(DRIVERS_SHEET, DRIVERS_HEADERS);
  const id = 'drv_' + Utilities.getUuid().slice(0, 8);
  const commission = (payload.commission === undefined || payload.commission === '') ? 20 : Number(payload.commission);
  sh.appendRow([id, payload.name || '', payload.email || '', payload.pin || '', 'Yes', commission, payload.company || '', 1]);
  return jsonOut({ success: true, id: id });
}

function updateDriver(payload) {
  const sh = getOrCreateSheet(DRIVERS_SHEET, DRIVERS_HEADERS);
  const row = findRowById_(sh, payload.id);
  if (row === -1) return jsonOut({ success: false, message: 'Driver not found' });
  if (payload.name !== undefined) sh.getRange(row, 2).setValue(payload.name);
  if (payload.email !== undefined) sh.getRange(row, 3).setValue(payload.email);
  if (payload.pin !== undefined) sh.getRange(row, 4).setValue(payload.pin);
  if (payload.active !== undefined) sh.getRange(row, 5).setValue(payload.active ? 'Yes' : 'No');
  if (payload.commission !== undefined && payload.commission !== '') sh.getRange(row, 6).setValue(Number(payload.commission));
  if (payload.company !== undefined) sh.getRange(row, 7).setValue(payload.company);
  return jsonOut({ success: true });
}

function removeDriver(payload) {
  const sh = getOrCreateSheet(DRIVERS_SHEET, DRIVERS_HEADERS);
  const row = findRowById_(sh, payload.id);
  if (row === -1) return jsonOut({ success: false, message: 'Driver not found' });
  sh.deleteRow(row);
  return jsonOut({ success: true });
}

// Returns this driver's next invoice number (starting at 1) and increments
// the counter on their row, so every driver has their own sequence —
// INV-0001, INV-0002, ... — independent of every other driver's.
function nextInvoiceNumberFor_(driverId) {
  const sh = getOrCreateSheet(DRIVERS_SHEET, DRIVERS_HEADERS);
  const row = findRowById_(sh, driverId);
  if (row === -1) return 1;
  const cell = sh.getRange(row, 8);
  const current = cell.getValue();
  const num = (current === '' || current === null || isNaN(current)) ? 1 : Number(current);
  cell.setValue(num + 1);
  return num;
}

// ===================== SETTINGS =====================

function readSettings() {
  const props = PropertiesService.getScriptProperties();
  return {
    opsPin: props.getProperty('opsPin') || '1234',
    taxBasis: props.getProperty('taxBasis') || 'net'
  };
}

function updateSettings(payload) {
  const props = PropertiesService.getScriptProperties();
  if (payload.opsPin !== undefined && payload.opsPin !== '') props.setProperty('opsPin', String(payload.opsPin));
  if (payload.taxBasis !== undefined) props.setProperty('taxBasis', String(payload.taxBasis));
  return jsonOut({ success: true });
}

// ===================== INVOICES =====================

function readInvoices() {
  const ss = getSpreadsheet_();
  const summarySh = ss.getSheetByName(SUMMARY_SHEET);
  const loadsSh = ss.getSheetByName(LOADS_SHEET);

  const loadsByInvoice = {};
  if (loadsSh) {
    const rows = loadsSh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      const invId = String(r[0]);
      if (!loadsByInvoice[invId]) loadsByInvoice[invId] = [];
      loadsByInvoice[invId].push({
        date: formatDate_(r[3]), refType: r[4], refNumber: String(r[5]), pickup: r[6], dropoffsText: r[7],
        province: r[8], price: r[9], commissionPct: r[10], commission: r[11], net: r[12], rate: r[13], tax: r[14], pay: r[15]
      });
    }
  }

  const invoices = [];
  if (summarySh) {
    const rows = summarySh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[1]) continue;
      const id = String(r[1]);
      invoices.push({
        id: id, invoiceNumber: r[0], driverId: String(r[2]), driverName: r[3], company: r[4] || '',
        periodStart: formatDate_(r[5]), periodEnd: formatDate_(r[6]),
        loads: loadsByInvoice[id] || [],
        totals: { actual: r[8], commission: r[9], net: r[10], tax: r[11], pay: r[12] },
        taxBasis: r[13], submittedAt: r[14], emailed: String(r[15]) === 'Yes', pdfUrl: r[16] || '', editedAt: r[17] || ''
      });
    }
  }
  invoices.sort(function (a, b) { return new Date(b.submittedAt) - new Date(a.submittedAt); });
  return invoices;
}

function submitInvoice(payload) {
  const invoiceNumber = nextInvoiceNumberFor_(payload.driverId);
  payload.invoiceNumber = invoiceNumber;
  payload.invoiceNumberDisplay = 'INV-' + String(invoiceNumber).padStart(4, '0');

  const loadsSheet = getOrCreateSheet(LOADS_SHEET, LOADS_HEADERS);
  (payload.loads || []).forEach(function (l) {
    loadsSheet.appendRow([
      payload.ticketNo, payload.driverId, payload.driverName, l.date, l.refType, l.refNumber, l.pickup, l.dropoffs,
      l.province, l.price, l.commissionPct, l.commission, l.net, l.taxRate, l.tax, l.pay
    ]);
  });

  // Build a PDF copy of the invoice and save it to Drive. This never blocks
  // the submission — if it fails for any reason, the loads above are
  // already saved, so we just log it and carry on with an empty pdfUrl.
  let pdfUrl = '';
  let pdfBlob = null;
  try {
    pdfBlob = buildInvoicePdfBlob(payload);
    const folder = getPdfFolder();
    const file = folder.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    pdfUrl = file.getUrl();
  } catch (pdfErr) {
    Logger.log('PDF generation failed: ' + pdfErr);
  }

  // Email operations, with the driver cc'd if they have one on file, and
  // the PDF attached if it generated successfully. A mail hiccup never
  // makes the whole submission look like it failed — emailSent is just
  // reported honestly either way.
  let emailSent = false;
  let emailError = '';
  try {
    if (OPS_EMAIL.indexOf('PASTE_OPS_EMAIL_HERE') > -1) {
      throw new Error('OPS_EMAIL is still the placeholder — edit it near the top of the script.');
    }
    sendInvoiceEmail(payload, pdfBlob);
    emailSent = true;
  } catch (mailErr) {
    emailError = String(mailErr);
    Logger.log('Email send failed: ' + emailError);
  }

  const summarySheet = getOrCreateSheet(SUMMARY_SHEET, SUMMARY_HEADERS);
  const t = payload.totals || {};
  summarySheet.appendRow([
    payload.invoiceNumberDisplay, payload.ticketNo, payload.driverId, payload.driverName, payload.driverCompany || '',
    payload.periodStart, payload.periodEnd, (payload.loads || []).length, t.actual, t.commission, t.net, t.tax, t.pay,
    payload.taxBasis, payload.submittedAt, emailSent ? 'Yes' : 'No', pdfUrl, ''
  ]);

  return jsonOut({ success: true, ticketNo: payload.ticketNo, invoiceNumber: payload.invoiceNumberDisplay, emailSent: emailSent, emailError: emailError, pdfUrl: pdfUrl });
}

function findSummaryRowByTicket_(sh, ticketNo) {
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(ticketNo)) return i + 1; // column B = Invoice ID
  }
  return -1;
}

// Edits an already-submitted invoice in place: replaces its load rows,
// recomputes totals, regenerates the PDF, resends the email (marked as
// an update), and overwrites the existing summary row rather than
// appending a new one — so the original Invoice # and Submitted At are
// preserved, and an Edited At timestamp is recorded.
function updateInvoice(payload) {
  const summarySheet = getOrCreateSheet(SUMMARY_SHEET, SUMMARY_HEADERS);
  const sRow = findSummaryRowByTicket_(summarySheet, payload.ticketNo);
  if (sRow === -1) {
    return jsonOut({ success: false, message: 'Original invoice not found — it may have been removed, so this can\'t be updated.' });
  }
  const existing = summarySheet.getRange(sRow, 1, 1, SUMMARY_HEADERS.length).getValues()[0];
  payload.invoiceNumberDisplay = existing[0]; // keep the original invoice number

  const loadsSheet = getOrCreateSheet(LOADS_SHEET, LOADS_HEADERS);
  const loadRows = loadsSheet.getDataRange().getValues();
  for (let i = loadRows.length - 1; i >= 1; i--) {
    if (String(loadRows[i][0]) === String(payload.ticketNo)) loadsSheet.deleteRow(i + 1);
  }
  (payload.loads || []).forEach(function (l) {
    loadsSheet.appendRow([
      payload.ticketNo, payload.driverId, payload.driverName, l.date, l.refType, l.refNumber, l.pickup, l.dropoffs,
      l.province, l.price, l.commissionPct, l.commission, l.net, l.taxRate, l.tax, l.pay
    ]);
  });

  let pdfUrl = existing[16] || '';
  let pdfBlob = null;
  try {
    pdfBlob = buildInvoicePdfBlob(payload);
    const folder = getPdfFolder();
    const file = folder.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    pdfUrl = file.getUrl();
  } catch (pdfErr) {
    Logger.log('Invoice PDF regeneration failed: ' + pdfErr);
  }

  let emailSent = false;
  let emailError = '';
  try {
    if (OPS_EMAIL.indexOf('PASTE_OPS_EMAIL_HERE') > -1) {
      throw new Error('OPS_EMAIL is still the placeholder — edit it near the top of the script.');
    }
    sendInvoiceEmail(Object.assign({}, payload, { isEdit: true }), pdfBlob);
    emailSent = true;
  } catch (mailErr) {
    emailError = String(mailErr);
    Logger.log('Invoice update email failed: ' + emailError);
  }

  const t = payload.totals || {};
  summarySheet.getRange(sRow, 1, 1, SUMMARY_HEADERS.length).setValues([[
    existing[0], payload.ticketNo, payload.driverId, payload.driverName, payload.driverCompany || '',
    payload.periodStart, payload.periodEnd, (payload.loads || []).length, t.actual, t.commission, t.net, t.tax, t.pay,
    payload.taxBasis, existing[14], emailSent ? 'Yes' : 'No', pdfUrl, new Date().toISOString()
  ]]);

  return jsonOut({ success: true, ticketNo: payload.ticketNo, invoiceNumber: payload.invoiceNumberDisplay, emailSent: emailSent, emailError: emailError, pdfUrl: pdfUrl });
}

// Builds the shared invoice content (used for both the email body and the
// PDF) so the two always stay in sync.
function buildInvoiceContentHtml(payload) {
  const t = payload.totals || {};
  const rows = (payload.loads || []).map(function (l) {
    return row_(
      l.date + ' — ' + l.refType + '#' + l.refNumber + ' (' + l.commissionPct + '% commission)',
      l.pickup + ' → ' + l.dropoffs + ' (' + l.province + ')  $' + Number(l.pay).toFixed(2)
    );
  }).join('');
  const companyLine = payload.driverCompany ? '<p style="color:#33326C;font-weight:700;margin:0 0 2px;">' + escapeHtml_(payload.driverCompany) + '</p>' : '';
  return '<p style="color:#666;margin:0 0 2px;font-size:12px;text-transform:uppercase;letter-spacing:.04em;">' + escapeHtml_(payload.invoiceNumberDisplay || '') + '</p>' +
    companyLine +
    '<h2 style="margin:0 0 4px;color:#33326C;">Invoice — ' + escapeHtml_(payload.driverName) + '</h2>' +
    '<p style="color:#666;margin:0 0 16px;">' + payload.periodStart + ' to ' + payload.periodEnd + '</p>' +
    '<table style="border-collapse:collapse;width:100%;margin-bottom:18px;">' + rows + '</table>' +
    '<table style="border-collapse:collapse;width:100%;">' +
    row_('Actual amount', '$' + Number(t.actual).toFixed(2)) +
    row_('Commission', '-$' + Number(t.commission).toFixed(2)) +
    row_('Net after commission', '$' + Number(t.net).toFixed(2)) +
    row_('Tax', '+$' + Number(t.tax).toFixed(2)) +
    row_('Driver pay', '$' + Number(t.pay).toFixed(2)) +
    '</table>';
}

function buildInvoicePdfBlob(payload) {
  const html = '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,sans-serif;color:#222;padding:28px;}' +
    'table{font-size:13px;} td{padding:6px 10px;}' +
    '</style></head><body>' + buildInvoiceContentHtml(payload) + '</body></html>';
  const blob = Utilities.newBlob(html, 'text/html', payload.ticketNo + '.html').getAs('application/pdf');
  blob.setName((payload.invoiceNumberDisplay || payload.ticketNo) + ' - ' + payload.driverName + '.pdf');
  return blob;
}

function getPdfFolder() {
  const name = 'Karim N Trucking Invoices — PDFs';
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function sendInvoiceEmail(payload, pdfBlob) {
  const t = payload.totals || {};
  const subject = (payload.isEdit ? '[Updated] ' : '') + (payload.invoiceNumberDisplay || 'Invoice') + ' — ' + payload.driverName + ' — ' + payload.periodStart + ' to ' + payload.periodEnd;

  const lines = (payload.loads || []).map(function (l) {
    return l.date + ' | ' + l.refType + '#' + l.refNumber + ' | ' + l.pickup + ' -> ' + l.dropoffs +
      ' | ' + l.province + ' | Actual $' + Number(l.price).toFixed(2) + ' | Commission (' + l.commissionPct + '%) $' + Number(l.commission).toFixed(2) +
      ' | Tax $' + Number(l.tax).toFixed(2) + ' | Pay $' + Number(l.pay).toFixed(2);
  });

  const plainBody = [
    'New driver invoice submitted', '',
    'Invoice #: ' + (payload.invoiceNumberDisplay || ''),
    'Driver: ' + payload.driverName + (payload.driverCompany ? ' (' + payload.driverCompany + ')' : ''),
    'Period: ' + payload.periodStart + ' to ' + payload.periodEnd,
    '', 'Loads:'
  ].concat(lines).concat([
    '',
    'Actual amount: $' + Number(t.actual).toFixed(2),
    'Commission: $' + Number(t.commission).toFixed(2),
    'Net after commission: $' + Number(t.net).toFixed(2),
    'Tax: $' + Number(t.tax).toFixed(2),
    'Driver pay: $' + Number(t.pay).toFixed(2),
    '', 'A PDF copy is attached.'
  ]).join('\n');

  const htmlBody = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:560px;">' + buildInvoiceContentHtml(payload) + '</div>';

  const options = { htmlBody: htmlBody, name: 'Karim N Trucking — Invoicing', from: OPS_EMAIL };
  if (payload.driverEmail) options.cc = payload.driverEmail;
  if (pdfBlob) options.attachments = [pdfBlob];

  GmailApp.sendEmail(OPS_EMAIL, subject, plainBody, options);
}

function row_(label, value) {
  return '<tr>' +
    '<td style="padding:6px 10px;border-bottom:1px solid #eee;color:#666;white-space:nowrap;">' + escapeHtml_(label) + '</td>' +
    '<td style="padding:6px 10px;border-bottom:1px solid #eee;font-weight:600;">' + escapeHtml_(String(value)) + '</td>' +
    '</tr>';
}

function escapeHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Run this once, manually, from the Apps Script editor — same reason
 * as your POD system's version of this function: a web app call
 * can't trigger Google's permission popup, so the first Gmail send
 * needs to be authorized by hand.
 *
 * Function dropdown at the top of the editor > authorizeGmailAccess
 * > Run > click through the permission screen. Check your inbox for
 * a test email once it's done.
 */
function authorizeGmailAccess() {
  if (OPS_EMAIL.indexOf('PASTE_OPS_EMAIL_HERE') > -1) {
    throw new Error('Set OPS_EMAIL near the top of this file before running this.');
  }
  GmailApp.sendEmail(OPS_EMAIL, 'Invoicing system — test email', 'If you got this, Gmail sending is authorized and the invoicing app is ready to email submissions.');
}

/**
 * Run this ONCE from the Apps Script editor (function dropdown at the
 * top > seedDrivers > Run) to bulk-add the drivers below. It skips
 * anyone whose name is already in the Drivers sheet, so it's safe to
 * run more than once if you add more rows later.
 *
 * PINs are auto-generated — after running, open the app's
 * Operations Portal > Drivers to see (and share) each one.
 *
 * ⚠️ Two entries still have no commission %, since none was given
 * for these two brand-new drivers — I defaulted them to 20%:
 *   - Abdul Hossiaini
 *   - Dmitri Khaliavka
 * Company name is also blank for both, since none was given. Fix
 * either directly below before running, or afterward in Operations
 * Portal > Drivers — nothing here is permanent.
 *
 * "Heidar Hamidi"'s company is kept as "Header Heath Can
 * Contracting" — the spelling of that one is still a guess, worth
 * a quick double-check.
 */
function seedDrivers() {
  const list = [
    { name: 'Kamal Paramanantham', company: '17401060 Ontario Inc', commission: 21, email: 'kamallr@hotmail.com' },
    { name: 'Mohit', company: '15752175 Canada Inc', commission: 23, email: 'mohit1992165@gmail.com' },
    { name: 'Manoj Bora', company: 'Aiba Trucking Inc', commission: 21, email: 'scorpiancan@gmail.com' },
    { name: 'Gopal Mor', company: '17463197 Canada Inc', commission: 25, email: 'gopalmor3332@gmail.com' },
    { name: 'Lojithan Vanniyasingam', company: '2630858 Ontario Inc', commission: 20, email: 'rglogy85@gmail.com' },
    { name: 'Abdul Hossiaini', company: '', commission: 20, email: 'abdulhossaini239@gmail.com' },
    { name: 'Dmitri Khaliavka', company: '', commission: 20, email: 'dkhaliavka@gmail.com' },
    { name: 'Senthuran Thevarasa', company: '10316156 Canada Inc', commission: 21, email: 'thevarasasenthuran@gmail.com' },
    { name: 'Heidar Hamidi', company: 'Header Heath Can Contracting', commission: 21, email: 'hhamidi2014@hotmail.com' },
    { name: 'Jeyanthan Manaharan', company: '1000804112 Ontario Inc', commission: 21, email: 'jeyanthan_m@hotmail.com' },
    { name: 'Niroopan Sathy', company: 'SG Express Inc', commission: 21, email: 'niroopansathy@yahoo.ca' },
    { name: 'Nimalan Sivanayagam', company: '799405 Canada Inc', commission: 21, email: 'nkatransport9014@gmail.com' },
    { name: 'Roshan', company: '1916036 Ontario Inc', commission: 21, email: 'Roshantoronto@yahoo.ca' },
    { name: 'Nithysh Mano', company: '8987556 Canada Inc', commission: 21, email: 'nithysh@hotmail.com' },
    { name: 'Faisal Rehman', company: '2531061 Ontario Inc', commission: 21, email: 'faisal86rehman@gmail.com' },
    { name: 'Franc Shestani', company: 'Trans LA', commission: 6, email: 'francshestani@hotmail.com' },
    { name: 'Jeevanjot Singh', company: 'Rakh Jot Transport', commission: 23, email: 'jotjeevan1111@gmail.com' },
    { name: 'Bashir', company: '1001231795 Ontario Inc', commission: 23, email: 'bashir.28@hotmail.com' },
    { name: 'Emmanual', company: 'Willy Transport LTD', commission: 6, email: 'thirstyduck22@hotmail.com' }
  ];

  const sh = getOrCreateSheet(DRIVERS_SHEET, DRIVERS_HEADERS);
  const existing = readDrivers().map(function (d) { return d.name.toLowerCase(); });
  let added = 0, skipped = 0;

  list.forEach(function (d) {
    if (existing.indexOf(d.name.toLowerCase()) > -1) { skipped++; return; }
    const id = 'drv_' + Utilities.getUuid().slice(0, 8);
    const pin = String(Math.floor(1000 + Math.random() * 9000));
    sh.appendRow([id, d.name, d.email, pin, 'Yes', d.commission, d.company, 1]);
    added++;
  });

  Logger.log('Seed complete: ' + added + ' added, ' + skipped + ' already existed and were skipped.');
}

