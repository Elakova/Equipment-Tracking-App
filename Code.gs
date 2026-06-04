// ================================================================
//  HOLOS Equipment Check — Google Apps Script
//  Paste this entire file into your Google Sheet's Script Editor:
//  Extensions → Apps Script → replace Code.gs → Save → Deploy
// ================================================================

const SPREADSHEET_ID = '1P9oBEBjHmfvXUtpon3a9XYvB1PRAhMDU6KCTc4NPyBE';
const LOG_TAB        = 'Log';
const EQUIP_TAB      = 'Equipment List';
const USERS_TAB      = 'Users';

// ── Ensure required tabs exist ──────────────────────────────────
function ensureTabs_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Log tab
  if (!ss.getSheetByName(LOG_TAB)) {
    const s = ss.insertSheet(LOG_TAB);
    s.appendRow(['#', 'Equipment', 'Status', 'Operator', 'Date & Time', 'Notes']);
    s.getRange('1:1').setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    s.setFrozenRows(1);
  }

  // Equipment List tab
  if (!ss.getSheetByName(EQUIP_TAB)) {
    const s = ss.insertSheet(EQUIP_TAB);
    s.appendRow(['Equipment Name', 'Category']);
    s.getRange('1:1').setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    s.setFrozenRows(1);
  }

  // Users tab
  if (!ss.getSheetByName(USERS_TAB)) {
    const s = ss.insertSheet(USERS_TAB);
    s.appendRow(['Name', 'Contact', 'Registered']);
    s.getRange('1:1').setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    s.setFrozenRows(1);
  }
}

// ── CORS helper ─────────────────────────────────────────────────
function corsResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET handler ─────────────────────────────────────────────────
// ?action=getEquipment  → returns { equipment: [{name, category}] }
// ?action=getLog        → returns { log: [{id,equipment,status,operator,timestamp,notes}] }
// ?action=ping          → returns { ok: true }
function doGet(e) {
  try {
    ensureTabs_();
    const action = (e.parameter && e.parameter.action) || 'ping';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    if (action === 'ping') {
      return corsResponse_({ ok: true });
    }

    if (action === 'getEquipment') {
      const sheet  = ss.getSheetByName(EQUIP_TAB);
      const values = sheet.getDataRange().getValues();
      const equipment = values.slice(1)               // skip header
        .filter(r => r[0])                            // skip empty rows
        .map(r => ({ name: String(r[0]).trim(), category: String(r[1] || '').trim() }));
      return corsResponse_({ equipment });
    }

    if (action === 'getLog') {
      const sheet  = ss.getSheetByName(LOG_TAB);
      const values = sheet.getDataRange().getValues();
      const log = values.slice(1)
        .filter(r => r[1])                            // skip blank rows
        .map((r, i) => ({
          id:        i + 1,
          equipment: String(r[1]),
          status:    String(r[2]),
          operator:  String(r[3]),
          timestamp: String(r[4]),
          notes:     String(r[5] || '')
        }))
        .reverse();                                   // newest first
      return corsResponse_({ log });
    }

    if (action === 'getUsers') {
      const sheet  = ss.getSheetByName(USERS_TAB);
      const values = sheet.getDataRange().getValues();
      const users = values.slice(1)
        .filter(r => r[0])
        .map(r => ({ name: String(r[0]), contact: String(r[1]), registered: String(r[2]) }));
      return corsResponse_({ users });
    }

    return corsResponse_({ error: 'Unknown action' });

  } catch (err) {
    return corsResponse_({ error: err.message });
  }
}

// ── POST handler ─────────────────────────────────────────────────
// Body: { entries: [{equipment, status, operator, timestamp, notes}] }
// Body: { action: 'addEquipment', name, category }
function doPost(e) {
  try {
    ensureTabs_();
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const body    = JSON.parse(e.postData.contents);

    // ── Register new user ────────────────────────
    if (body.action === 'registerUser') {
      const usersSheet = ss.getSheetByName(USERS_TAB);
      const existing = usersSheet.getDataRange().getValues()
        .slice(1).map(r => String(r[0]).toLowerCase().trim());
      if (!existing.includes(String(body.name).toLowerCase().trim())) {
        usersSheet.appendRow([
          body.name.trim(),
          (body.contact || '').trim(),
          new Date().toLocaleDateString('en-PH', { month: 'short', day: 'numeric', year: 'numeric' })
        ]);
      }
      return corsResponse_({ ok: true });
    }

    // ── Edit equipment name / category ───────────
    if (body.action === 'editEquipment') {
      const equipSheet = ss.getSheetByName(EQUIP_TAB);
      const values = equipSheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]).toLowerCase().trim() === String(body.oldName).toLowerCase().trim()) {
          equipSheet.getRange(i + 1, 1).setValue(body.newName.trim());
          equipSheet.getRange(i + 1, 2).setValue((body.newCategory || '').trim());
          break;
        }
      }
      return corsResponse_({ ok: true });
    }

    // ── Register new equipment ───────────────────
    if (body.action === 'addEquipment') {
      const equipSheet = ss.getSheetByName(EQUIP_TAB);
      // Check for duplicates (case-insensitive)
      const existing = equipSheet.getDataRange().getValues()
        .slice(1).map(r => String(r[0]).toLowerCase().trim());
      if (existing.includes(String(body.name).toLowerCase().trim())) {
        return corsResponse_({ ok: true, duplicate: true });
      }
      equipSheet.appendRow([body.name.trim(), (body.category || '').trim()]);
      return corsResponse_({ ok: true, duplicate: false });
    }

    // ── Append log entries ───────────────────────
    const logSheet = ss.getSheetByName(LOG_TAB);
    const entries  = Array.isArray(body.entries) ? body.entries : [body];

    entries.forEach(entry => {
      const rowNum = logSheet.getLastRow();   // current last row before appending
      logSheet.appendRow([
        rowNum,                               // # (auto-number)
        entry.equipment  || '',
        entry.status     || '',
        entry.operator   || '',
        entry.timestamp  || new Date().toISOString(),
        entry.notes      || ''
      ]);

      // Colour the Status cell
      const newRow = logSheet.getLastRow();
      const statusCell = logSheet.getRange(newRow, 3);
      if (entry.status === 'IN') {
        statusCell.setBackground('#d9f7be').setFontColor('#135200');
      } else {
        statusCell.setBackground('#fff1f0').setFontColor('#a8071a');
      }
    });

    return corsResponse_({ ok: true, appended: entries.length });

  } catch (err) {
    return corsResponse_({ error: err.message });
  }
}
