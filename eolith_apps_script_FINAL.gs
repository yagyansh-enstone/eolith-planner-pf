// ============================================================
// EOLITH Production Planner — Apps Script FINAL VERSION
// ============================================================
// INSTALL STEPS:
// 1. Google Sheet kholein → Extensions → Apps Script
// 2. Poora code select (Ctrl+A) → delete → yeh paste karein
// 3. Save (Ctrl+S)
// 4. Deploy → New Deployment → Web App
//    Execute as: Me | Who has access: Anyone
// 5. Deploy → Allow → URL copy karein → HTML mein hardcode hai
// ============================================================

// ── SHEET ID — APNI SHEET KA ID YAHAN DAALEN ──
var SHEET_ID = '1bCtsS921wczGk__qbAsjcyR4yrK7xJ4KT687TX2guPM';

// ── CORS HEADERS (GitHub Pages ke liye zaroori) ──
function setCORSHeaders(output) {
  return output
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// ── GET REQUEST — Test + Dispatch Monthly Planning Status fetch ──
function doGet(e) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var action = e.parameter.action || 'test';

    if (action === 'getDispatchStatus') {
      // Dispatch Monthly Planning Status sheet se data fetch karo
      var sheet = ss.getSheetByName('Dispatch Monthly Planning Status');
      if (!sheet) {
        return setCORSHeaders(ContentService.createTextOutput(
          JSON.stringify({ status: 'error', msg: 'Dispatch Monthly Planning Status sheet nahi mili' })
        ).setMimeType(ContentService.MimeType.JSON));
      }
      var data = sheet.getDataRange().getValues();
      return setCORSHeaders(ContentService.createTextOutput(
        JSON.stringify({ status: 'ok', data: data })
      ).setMimeType(ContentService.MimeType.JSON));
    }

    // Default: API test
    var sheets = ss.getSheets().map(function(s) { return s.getName(); });
    return setCORSHeaders(ContentService.createTextOutput(
      JSON.stringify({ status: 'ok', msg: 'Eolith API Ready!', sheets: sheets })
    ).setMimeType(ContentService.MimeType.JSON));

  } catch (err) {
    return setCORSHeaders(ContentService.createTextOutput(
      JSON.stringify({ status: 'error', msg: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON));
  }
}

// ── POST REQUEST — Data save + Dispatch update ──
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var action = data.action || 'save';

    // ── ACTION: Save Production Log ──
    if (action === 'saveProduction') {
      var sheet = getOrCreateSheet(ss, 'ProductionLog', [
        'Date', 'Time', 'Plan ID', 'Size (mm)', 'Block Cutter',
        'Target (Sq Ft)', 'Min Output/24hr', 'Max Output/24hr',
        'Est Hours', 'Day Shift', 'Night Shift', 'Status', 'MIS Executive'
      ], '#1a237e');

      var rows = data.rows || [];
      rows.forEach(function(row) { sheet.appendRow(row); });
      styleNewRows(sheet, rows.length, '#e8eaf6');

      return okResponse({ saved: rows.length, tab: 'ProductionLog' });
    }

    // ── ACTION: Save Dispatch Plan ──
    if (action === 'saveDispatch') {
      var dsheet = getOrCreateSheet(ss, 'Dispatch Planning', [
        'ID', 'Client Name', 'Order Details', 'Dispatch Date',
        'Production Start', 'Est Hours', 'Status', 'Notes', 'Saved On'
      ], '#1b5e20');

      var drows = data.rows || [];
      drows.forEach(function(row) { dsheet.appendRow(row); });
      styleNewRows(dsheet, drows.length, '#e8f5e9');

      return okResponse({ saved: drows.length, tab: 'Dispatch Planning' });
    }

    // ── ACTION: Update Dispatch Monthly Planning Status ──
    if (action === 'updateDispatchStatus') {
      var dmSheet = ss.getSheetByName('Dispatch Monthly Planning Status');
      if (!dmSheet) {
        // Create with standard headers matching existing sheet
        dmSheet = ss.insertSheet('Dispatch Monthly Planning Status');
        dmSheet.appendRow([
          'Plane Date', 'Dispatch Date', 'Party Name', 'Material Size(mm)',
          'Description', 'Order Quantity', 'Balance Quantity',
          'Production Status', 'Packing Status', 'Dispatch Status'
        ]);
        dmSheet.getRange(1, 1, 1, 10).setBackground('#212121').setFontColor('#fff').setFontWeight('bold');
        dmSheet.setFrozenRows(1);
      }

      var updateAction = data.updateType;

      if (updateAction === 'addRow') {
        // New entry add karo
        var newRow = data.rowData;
        dmSheet.appendRow(newRow);
        var lr = dmSheet.getLastRow();
        if (lr % 2 === 0) {
          dmSheet.getRange(lr, 1, 1, 10).setBackground('#f5f5f5');
        }
        return okResponse({ msg: 'Row added', row: lr });
      }

      if (updateAction === 'updateStatus') {
        // Specific row ka status update karo (rowIndex 1-based, including header)
        var rowIdx = data.rowIndex;
        var colUpdates = data.colUpdates; // { col: value } pairs, 1-based
        Object.keys(colUpdates).forEach(function(col) {
          dmSheet.getRange(rowIdx, parseInt(col)).setValue(colUpdates[col]);
        });
        return okResponse({ msg: 'Status updated', row: rowIdx });
      }

      if (updateAction === 'deleteRow') {
        var delRow = data.rowIndex;
        dmSheet.deleteRow(delRow);
        return okResponse({ msg: 'Row deleted' });
      }

      return okResponse({ msg: 'No update action specified' });
    }

    return errResponse('Unknown action: ' + action);

  } catch (err) {
    return errResponse(err.toString());
  }
}

// ── HELPERS ──────────────────────────────────────────────────
function getOrCreateSheet(ss, name, headers, headerColor) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground(headerColor || '#212121')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(10);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function styleNewRows(sheet, count, color) {
  var lr = sheet.getLastRow();
  if (lr > 1 && count > 0) {
    for (var i = lr - count + 1; i <= lr; i++) {
      if (i % 2 === 0) {
        sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground(color);
      }
    }
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  }
}

function okResponse(obj) {
  obj.status = 'ok';
  return setCORSHeaders(ContentService.createTextOutput(
    JSON.stringify(obj)
  ).setMimeType(ContentService.MimeType.JSON));
}

function errResponse(msg) {
  return setCORSHeaders(ContentService.createTextOutput(
    JSON.stringify({ status: 'error', msg: msg })
  ).setMimeType(ContentService.MimeType.JSON));
}
