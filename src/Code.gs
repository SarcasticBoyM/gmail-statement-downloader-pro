/**
 * Gmail Statement Downloader Pro
 * Google Apps Script - No Python required
 * Features: Credit Card + Bank Statements + Bank Manager via Google Sheet + Logs + Dashboard + CA Package ZIP
 */

const APP_NAME = 'Gmail Statement Downloader Pro';
const ROOT_FOLDER_NAME = 'Gmail Statement Downloader Pro';
const CONFIG_FILE_NAME = 'Statement Downloader Config';
const SHEET_BANKS = 'Banks';
const SHEET_SETTINGS = 'Settings';
const SHEET_LOGS = 'Logs';
const SHEET_HISTORY = 'History';
const SHEET_DASHBOARD = 'Dashboard';

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Statement Downloader')
      .addItem('1) Setup / Open Config', 'setupBankConfig')
      .addSeparator()
      .addItem('Download Credit Cards', 'downloadCreditCards')
      .addItem('Download Bank Statements', 'downloadBankStatements')
      .addItem('Download All', 'downloadAllStatements')
      .addSeparator()
      .addItem('Create CA Package ZIP', 'createCAPackage')
      .addItem('Open Download Folder', 'openDownloadFolder')
      .addItem('Refresh Dashboard', 'refreshDashboard')
      .addToUi();
  } catch (err) {
    Logger.log('onOpen menu skipped: ' + err.message);
  }
}

function setupBankConfig() {
  const ss = getOrCreateConfigSpreadsheet_();
  setupSheets_(ss);
  const url = ss.getUrl();
  Logger.log('Bank config ready: ' + url);
  try {
    SpreadsheetApp.getUi().alert('Bank config ready. Open this sheet:\n\n' + url);
  } catch (err) {
    Logger.log('Open this config sheet: ' + url);
  }
  return url;
}

function downloadCreditCards() {
  return downloadStatementsByType_('CREDIT_CARD');
}

function downloadBankStatements() {
  return downloadStatementsByType_('BANK_STATEMENT');
}

function downloadAllStatements() {
  return downloadStatementsByType_('ALL');
}

function createCAPackage() {
  const ss = getOrCreateConfigSpreadsheet_();
  setupSheets_(ss);

  // First download all missing enabled statements, then package everything downloaded in selected period.
  downloadAllStatements();

  const settings = getSettings_(ss);
  const periodLabel = settings.CA_PACKAGE_PERIOD || getCurrentFinancialYearLabel_();
  const root = getOrCreateFolder_(ROOT_FOLDER_NAME);
  const caFolder = getOrCreateChildFolder_(root, 'CA Packages');
  const files = [];
  const summaryRows = [['Bank','Type','File Name','Drive Folder','Downloaded At','Gmail Message ID']];

  const historySheet = ss.getSheetByName(SHEET_HISTORY);
  const history = historySheet.getDataRange().getValues();
  const headers = history[0] || [];
  const ix = indexMap_(headers);

  for (let r = 1; r < history.length; r++) {
    const row = history[r];
    const fileId = row[ix.FileId];
    if (!fileId) continue;
    try {
      const file = DriveApp.getFileById(fileId);
      const name = file.getName();
      // Include PDFs downloaded by this tool. Date filtering can be handled by Settings if needed.
      if (String(name).toLowerCase().endsWith('.pdf')) {
        const blob = file.getBlob().setName(name);
        files.push(blob);
        summaryRows.push([
          row[ix.Bank] || '',
          row[ix.Type] || '',
          name,
          row[ix.FolderPath] || '',
          row[ix.DownloadedAt] || '',
          row[ix.MessageId] || ''
        ]);
      }
    } catch (err) {
      log_(ss, 'CA_PACKAGE', '', '', 'WARN', 'Could not read fileId ' + fileId + ': ' + err.message);
    }
  }

  if (files.length === 0) {
    log_(ss, 'CA_PACKAGE', '', '', 'FAILED', 'No PDF files found in history to create package.');
    throw new Error('No PDF files found. Run Download All first.');
  }

  const csv = rowsToCsv_(summaryRows);
  files.push(Utilities.newBlob(csv, 'text/csv', 'CA_Statement_Summary.csv'));

  const zipName = 'CA_Statements_' + periodLabel + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.zip';
  const zipBlob = Utilities.zip(files, zipName);
  const zipFile = caFolder.createFile(zipBlob);

  log_(ss, 'CA_PACKAGE', '', zipName, 'SUCCESS', 'Created CA ZIP with ' + (files.length - 1) + ' PDFs: ' + zipFile.getUrl());
  refreshDashboard();

  Logger.log('CA Package ready: ' + zipFile.getUrl());
  try {
    SpreadsheetApp.getUi().alert('CA Package ZIP ready:\n\n' + zipFile.getUrl());
  } catch (err) {}
  return zipFile.getUrl();
}

function openDownloadFolder() {
  const folder = getOrCreateFolder_(ROOT_FOLDER_NAME);
  Logger.log('Download folder: ' + folder.getUrl());
  try { SpreadsheetApp.getUi().alert('Download folder:\n\n' + folder.getUrl()); } catch (err) {}
  return folder.getUrl();
}

function downloadStatementsByType_(typeFilter) {
  const ss = getOrCreateConfigSpreadsheet_();
  setupSheets_(ss);
  const banks = getEnabledBanks_(ss, typeFilter);
  const settings = getSettings_(ss);
  const root = getOrCreateFolder_(settings.ROOT_FOLDER_NAME || ROOT_FOLDER_NAME);

  let scanned = 0, downloaded = 0, duplicates = 0, failed = 0;

  banks.forEach(bank => {
    const query = buildGmailQuery_(bank, settings);
    let threads = [];
    try {
      threads = GmailApp.search(query, 0, Number(settings.MAX_THREADS_PER_BANK || 100));
      scanned += threads.length;
    } catch (err) {
      failed++;
      log_(ss, bank.Bank, bank.Type, '', 'FAILED', 'Gmail search failed: ' + err.message + ' | Query: ' + query);
      return;
    }

    const typeFolderName = bank.Type === 'CREDIT_CARD' ? 'Credit Card Statements' : 'Bank Statements';
    const typeFolder = getOrCreateChildFolder_(root, typeFolderName);
    const bankFolder = getOrCreateChildFolder_(typeFolder, safeName_(bank.Bank));

    threads.forEach(thread => {
      thread.getMessages().forEach(message => {
        const attachments = message.getAttachments({includeInlineImages:false, includeAttachments:true});
        attachments.forEach(att => {
          const originalName = att.getName() || 'statement.pdf';
          if (!isPdf_(att, originalName)) return;
          const messageId = message.getId();
          const historyKey = messageId + '|' + originalName + '|' + bank.Bank + '|' + bank.Type;
          if (historyExists_(ss, historyKey)) {
            duplicates++;
            return;
          }
          const fileName = makeStatementFileName_(bank, message, originalName);
          try {
            let finalName = fileName;
            if (fileExists_(bankFolder, finalName)) finalName = uniqueName_(bankFolder, finalName);
            const file = bankFolder.createFile(att.copyBlob().setName(finalName));
            addHistory_(ss, historyKey, bank, finalName, file.getId(), getFolderPath_(typeFolderName, bank.Bank), messageId);
            log_(ss, bank.Bank, bank.Type, finalName, 'SUCCESS', 'Downloaded');
            downloaded++;
          } catch (err) {
            failed++;
            log_(ss, bank.Bank, bank.Type, originalName, 'FAILED', err.message);
          }
        });
      });
    });
  });

  log_(ss, 'RUN_SUMMARY', typeFilter, '', 'SUCCESS', 'Scanned=' + scanned + ', Downloaded=' + downloaded + ', Duplicates=' + duplicates + ', Failed=' + failed);
  refreshDashboard();
  Logger.log('Done. Scanned=' + scanned + ', Downloaded=' + downloaded + ', Duplicates=' + duplicates + ', Failed=' + failed);
  return {scanned, downloaded, duplicates, failed};
}

function buildGmailQuery_(bank, settings) {
  let query = bank.Query || '';
  if (query.indexOf('filename:pdf') === -1) query += ' filename:pdf';
  if (settings.SEARCH_AFTER_DATE) query += ' after:' + settings.SEARCH_AFTER_DATE;
  if (settings.SEARCH_BEFORE_DATE) query += ' before:' + settings.SEARCH_BEFORE_DATE;
  return query.trim();
}

function setupSheets_(ss) {
  setupBanksSheet_(ss);
  setupSettingsSheet_(ss);
  setupLogsSheet_(ss);
  setupHistorySheet_(ss);
  setupDashboardSheet_(ss);
}

function setupBanksSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_BANKS) || ss.insertSheet(SHEET_BANKS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Bank','Type','Query','Enabled','Notes']);
    const rows = [
      ['Axis Bank','CREDIT_CARD','(from:axisbank.com OR from:axisbank.co.in OR axis credit card statement) filename:pdf','YES','Edit query if needed'],
      ['SBI Card','CREDIT_CARD','(from:sbicard.com OR from:sbicard.co.in OR "SBI Card statement") filename:pdf','YES',''],
      ['HDFC Bank','CREDIT_CARD','(from:hdfcbank.net OR from:hdfcbank.com OR "HDFC credit card statement") filename:pdf','YES',''],
      ['YES Bank','CREDIT_CARD','(from:yesbank.in OR "YES Bank credit card statement") filename:pdf','YES',''],
      ['Federal Bank','CREDIT_CARD','(from:federalbank.co.in OR "Federal Bank credit card statement") filename:pdf','YES',''],
      ['Axis Bank','BANK_STATEMENT','(from:axisbank.com OR from:axisbank.co.in OR "account statement" OR "bank statement") filename:pdf','YES',''],
      ['HDFC Bank','BANK_STATEMENT','(from:hdfcbank.net OR from:hdfcbank.com OR "account statement" OR "bank statement") filename:pdf','YES',''],
      ['YES Bank','BANK_STATEMENT','(from:yesbank.in OR "account statement" OR "bank statement") filename:pdf','YES',''],
      ['Federal Bank','BANK_STATEMENT','(from:federalbank.co.in OR "account statement" OR "bank statement") filename:pdf','YES','']
    ];
    rows.forEach(r => sh.appendRow(r));
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 5);
  }
}

function setupSettingsSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_SETTINGS) || ss.insertSheet(SHEET_SETTINGS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Key','Value','Help']);
    sh.appendRow(['ROOT_FOLDER_NAME', ROOT_FOLDER_NAME, 'Google Drive root folder name']);
    sh.appendRow(['MAX_THREADS_PER_BANK', '100', 'Increase if older statements are not found']);
    sh.appendRow(['SEARCH_AFTER_DATE', '', 'Optional Gmail date: yyyy/mm/dd']);
    sh.appendRow(['SEARCH_BEFORE_DATE', '', 'Optional Gmail date: yyyy/mm/dd']);
    sh.appendRow(['CA_PACKAGE_PERIOD', getCurrentFinancialYearLabel_(), 'Used in ZIP filename']);
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 3);
  }
}

function setupLogsSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_LOGS) || ss.insertSheet(SHEET_LOGS);
  if (sh.getLastRow() === 0) sh.appendRow(['Time','Bank','Type','File','Status','Message']);
}

function setupHistorySheet_(ss) {
  let sh = ss.getSheetByName(SHEET_HISTORY) || ss.insertSheet(SHEET_HISTORY);
  if (sh.getLastRow() === 0) sh.appendRow(['Key','DownloadedAt','Bank','Type','FileName','FileId','FolderPath','MessageId']);
}

function setupDashboardSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_DASHBOARD) || ss.insertSheet(SHEET_DASHBOARD);
  if (sh.getLastRow() === 0) refreshDashboard();
}

function refreshDashboard() {
  const ss = getOrCreateConfigSpreadsheet_();
  let dash = ss.getSheetByName(SHEET_DASHBOARD) || ss.insertSheet(SHEET_DASHBOARD);
  const logs = ss.getSheetByName(SHEET_LOGS);
  const hist = ss.getSheetByName(SHEET_HISTORY);
  const banks = ss.getSheetByName(SHEET_BANKS);
  const logRows = logs ? logs.getDataRange().getValues() : [];
  const histRows = hist ? Math.max(0, hist.getLastRow() - 1) : 0;
  const bankRows = banks ? Math.max(0, banks.getLastRow() - 1) : 0;
  let success = 0, failed = 0;
  for (let i=1;i<logRows.length;i++) {
    if (logRows[i][4] === 'SUCCESS') success++;
    if (logRows[i][4] === 'FAILED') failed++;
  }
  dash.clear();
  dash.appendRow(['Metric','Value']);
  dash.appendRow(['Total banks configured', bankRows]);
  dash.appendRow(['Total PDFs downloaded', histRows]);
  dash.appendRow(['Successful log entries', success]);
  dash.appendRow(['Failed log entries', failed]);
  dash.appendRow(['Last refresh', new Date()]);
  dash.autoResizeColumns(1, 2);
}

function getOrCreateConfigSpreadsheet_() {
  const files = DriveApp.getFilesByName(CONFIG_FILE_NAME);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  return SpreadsheetApp.create(CONFIG_FILE_NAME);
}

function getEnabledBanks_(ss, typeFilter) {
  const values = ss.getSheetByName(SHEET_BANKS).getDataRange().getValues();
  const headers = values[0];
  const ix = indexMap_(headers);
  const result = [];
  for (let r=1; r<values.length; r++) {
    const row = values[r];
    const enabled = String(row[ix.Enabled] || '').toUpperCase() === 'YES';
    const type = String(row[ix.Type] || '').toUpperCase();
    if (!enabled) continue;
    if (typeFilter !== 'ALL' && type !== typeFilter) continue;
    result.push({Bank: row[ix.Bank], Type: type, Query: row[ix.Query], Enabled: row[ix.Enabled], Notes: row[ix.Notes]});
  }
  return result;
}

function getSettings_(ss) {
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  const values = sh.getDataRange().getValues();
  const out = {};
  for (let i=1;i<values.length;i++) out[String(values[i][0])] = values[i][1];
  return out;
}

function log_(ss, bank, type, file, status, message) {
  ss.getSheetByName(SHEET_LOGS).appendRow([new Date(), bank, type, file, status, message]);
}

function addHistory_(ss, key, bank, fileName, fileId, folderPath, messageId) {
  ss.getSheetByName(SHEET_HISTORY).appendRow([key, new Date(), bank.Bank, bank.Type, fileName, fileId, folderPath, messageId]);
}

function historyExists_(ss, key) {
  const sh = ss.getSheetByName(SHEET_HISTORY);
  const finder = sh.getRange(1,1,Math.max(1,sh.getLastRow()),1).createTextFinder(key).matchEntireCell(true).findNext();
  return !!finder;
}

function getOrCreateFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function getOrCreateChildFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function fileExists_(folder, name) { return folder.getFilesByName(name).hasNext(); }

function uniqueName_(folder, name) {
  const dot = name.lastIndexOf('.');
  const base = dot > -1 ? name.substring(0, dot) : name;
  const ext = dot > -1 ? name.substring(dot) : '';
  let i = 2;
  while (fileExists_(folder, base + '_' + i + ext)) i++;
  return base + '_' + i + ext;
}

function isPdf_(blob, name) {
  return String(name).toLowerCase().endsWith('.pdf') || String(blob.getContentType()).toLowerCase().indexOf('pdf') >= 0;
}

function makeStatementFileName_(bank, message, originalName) {
  const date = Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), 'yyyy-MM');
  return safeName_(bank.Bank) + '_' + bank.Type + '_' + date + '_' + cleanFileName_(originalName);
}

function safeName_(text) { return cleanFileName_(String(text || 'Unknown')).replace(/\s+/g, '_'); }
function cleanFileName_(text) { return String(text || 'file').replace(/[\\/:*?"<>|#%{}~&]/g, '-').trim(); }
function getFolderPath_(typeFolderName, bankName) { return ROOT_FOLDER_NAME + '/' + typeFolderName + '/' + safeName_(bankName); }
function indexMap_(headers) { const m={}; headers.forEach((h,i)=>m[String(h).trim()]=i); return m; }

function rowsToCsv_(rows) {
  return rows.map(row => row.map(cell => '"' + String(cell == null ? '' : cell).replace(/"/g, '""') + '"').join(',')).join('\n');
}

function getCurrentFinancialYearLabel_() {
  const d = new Date();
  const y = d.getFullYear();
  const m = d.getMonth() + 1;
  const start = m >= 4 ? y : y - 1;
  const end = start + 1;
  return 'FY_' + start + '_' + String(end).slice(-2);
}
