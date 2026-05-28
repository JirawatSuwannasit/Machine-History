// ============================================================
// CONSTANTS
// ============================================================

var SHEETS = {
  MACHINE_LIST: 'Machine_List',
  PARTS_MASTER: 'Spare_Parts_Master',
  SCHEDULE:     'Spare_Parts_Schedule',
  LOG:          'Maintenance_Log',
  DEFECT_LOG:   'Defect_Log'
};

var COL = {
  MACHINE: {
    ID: 0, SCOPE: 1, NAME: 2, MANUFACTURER: 3,
    MODEL: 4, SN: 5, RANGE: 6, OP_DATE: 7, STATUS: 8
  },
  PARTS_MASTER: { ID: 0, NAME: 1, LIFETIME: 2, DESC: 3 },
  SCHEDULE: {
    MACHINE_ID: 0, PART_ID: 1, LAST_CHANGED: 2, NEXT_DUE: 3, STATUS: 4
  },
  LOG: {
    TIMESTAMP: 0, MACHINE_ID: 1, ACTION: 2, PART_ID: 3,
    DETAILS: 4, OPERATOR: 5, MAINTENANCE_DATE: 6
  },
  DEFECT: {
    ID: 0, TIMESTAMP: 1, MACHINE_ID: 2, DATE_FOUND: 3,
    SYMPTOM: 4, SEVERITY: 5, REPORTED_BY: 6, STATUS: 7
  }
};

// ============================================================
// ENTRY POINT
// ============================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Machine History & Maintenance')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// PRIVATE HELPERS
// ============================================================

function _getSheet(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet not found: ' + name);
  return sheet;
}

function _getSheetData(name) {
  var sheet = _getSheet(name);
  if (sheet.getLastRow() <= 1) return [];
  return sheet.getDataRange().getValues().slice(1);
}

function _formatDate(date) {
  var d = (date instanceof Date) ? date : new Date(date);
  if (isNaN(d.getTime())) return '';
  return d.getFullYear() + '-' +
    String(d.getMonth() + 1).padStart(2, '0') + '-' +
    String(d.getDate()).padStart(2, '0');
}

function _formatDateTime(date) {
  var d = (date instanceof Date) ? date : new Date(date);
  if (isNaN(d.getTime())) return '';
  return _formatDate(d) + ' ' +
    String(d.getHours()).padStart(2, '0') + ':' +
    String(d.getMinutes()).padStart(2, '0') + ':' +
    String(d.getSeconds()).padStart(2, '0');
}

function _addYears(dateStr, years) {
  var d = new Date(dateStr);
  d.setDate(d.getDate() + Math.round(parseFloat(years) * 365));
  return _formatDate(d);
}

function _dateDiffDays(futureDateStr) {
  var today  = new Date(); today.setHours(0, 0, 0, 0);
  var future = new Date(futureDateStr); future.setHours(0, 0, 0, 0);
  if (isNaN(future.getTime())) return 9999;
  return Math.round((future - today) / 86400000);
}

function _scheduleStatus(days) {
  if (days < 0)   return 'OVERDUE';
  if (days <= 30) return 'DUE_SOON';
  return 'OK';
}

// Upsert Spare_Parts_Schedule by composite key [Machine_ID, Part_ID].
function _upsertSchedule(machineId, partId, lastChanged, nextDueDate, status) {
  var sheet    = _getSheet(SHEETS.SCHEDULE);
  var data     = sheet.getDataRange().getValues();
  var foundIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(machineId).trim() &&
        String(data[i][1]).trim() === String(partId).trim()) {
      foundIdx = i; break;
    }
  }
  if (foundIdx !== -1) {
    var row = foundIdx + 1;
    sheet.getRange(row, 3).setValue(lastChanged);
    sheet.getRange(row, 4).setValue(nextDueDate);
    sheet.getRange(row, 5).setValue(status);
  } else {
    sheet.appendRow([machineId, partId, lastChanged, nextDueDate, status]);
  }
}

// Generates next sequential Defect_ID (e.g. DF-0001).
// Scans existing IDs to find the highest number — safe against gaps/deletions.
function _generateDefectId() {
  var rows = _getSheetData(SHEETS.DEFECT_LOG);
  var max  = 0;
  rows.forEach(function(r) {
    var m = String(r[COL.DEFECT.ID]).match(/^DF-(\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return 'DF-' + String(max + 1).padStart(4, '0');
}

// Sets a defect's Status to 'Resolved'. Returns true if found, false otherwise.
function _resolveDefect(defectId) {
  var sheet = _getSheet(SHEETS.DEFECT_LOG);
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL.DEFECT.ID]).trim() === String(defectId).trim()) {
      sheet.getRange(i + 1, COL.DEFECT.STATUS + 1).setValue('Resolved');
      return true;
    }
  }
  return false;
}

// ============================================================
// PUBLIC FUNCTIONS — callable via google.script.run
// ============================================================

/**
 * Returns all machines for the dashboard and Machine_ID dropdown.
 */
function getMachineList() {
  try {
    return _getSheetData(SHEETS.MACHINE_LIST)
      .filter(function(r) { return String(r[0]).trim() !== ''; })
      .map(function(r) {
        return {
          machine_id:     String(r[COL.MACHINE.ID]),
          scope:          String(r[COL.MACHINE.SCOPE]),
          machine_name:   String(r[COL.MACHINE.NAME]),
          manufacturer:   String(r[COL.MACHINE.MANUFACTURER]),
          model:          String(r[COL.MACHINE.MODEL]),
          sn:             String(r[COL.MACHINE.SN]),
          range:          String(r[COL.MACHINE.RANGE]),
          operation_date: _formatDate(r[COL.MACHINE.OP_DATE]),
          status:         String(r[COL.MACHINE.STATUS])
        };
      });
  } catch (e) { return { error: e.message }; }
}

/**
 * Returns dropdown data (machines, parts, action types) for entry forms.
 */
function getDropdownData() {
  try {
    var machines = _getSheetData(SHEETS.MACHINE_LIST)
      .filter(function(r) { return String(r[0]).trim() !== ''; })
      .map(function(r) {
        return { machine_id: String(r[0]), machine_name: String(r[2]) };
      });

    var parts = _getSheetData(SHEETS.PARTS_MASTER)
      .filter(function(r) { return String(r[0]).trim() !== ''; })
      .map(function(r) {
        return {
          part_id:        String(r[0]),
          part_name:      String(r[1]),
          lifetime_years: parseFloat(r[2]) || 1,
          description:    String(r[3])
        };
      });

    return {
      machines:        machines,
      parts:           parts,
      action_types:    ['Repair', 'Part Replacement', 'PM'],
      severity_levels: ['Low', 'Medium', 'High', 'Critical']
    };
  } catch (e) { return { error: e.message }; }
}

/**
 * Returns the full profile for one machine:
 *   { machine, logs[], schedule[], defects[] }
 */
function getMachineProfile(machineId) {
  try {
    var machineRows  = _getSheetData(SHEETS.MACHINE_LIST);
    var logRows      = _getSheetData(SHEETS.LOG);
    var scheduleRows = _getSheetData(SHEETS.SCHEDULE);
    var masterRows   = _getSheetData(SHEETS.PARTS_MASTER);
    var defectRows   = _getSheetData(SHEETS.DEFECT_LOG);

    // Parts lookup map
    var partsMap = {};
    masterRows.forEach(function(r) {
      partsMap[String(r[0]).trim()] = { name: String(r[1]), lifetime: parseFloat(r[2]) || 1 };
    });

    var mId = String(machineId).trim();

    // Machine row
    var mRow = null;
    for (var i = 0; i < machineRows.length; i++) {
      if (String(machineRows[i][0]).trim() === mId) { mRow = machineRows[i]; break; }
    }
    if (!mRow) return { error: 'Machine not found: ' + machineId };

    var machine = {
      machine_id:     String(mRow[0]),
      scope:          String(mRow[1]),
      machine_name:   String(mRow[2]),
      manufacturer:   String(mRow[3]),
      model:          String(mRow[4]),
      sn:             String(mRow[5]),
      range:          String(mRow[6]),
      operation_date: _formatDate(mRow[7]),
      status:         String(mRow[8])
    };

    // Maintenance logs — newest first
    var logs = logRows
      .filter(function(r) { return String(r[COL.LOG.MACHINE_ID]).trim() === mId; })
      .map(function(r) {
        var pId = String(r[COL.LOG.PART_ID]).trim();
        return {
          timestamp:        _formatDateTime(r[COL.LOG.TIMESTAMP]),
          action_type:      String(r[COL.LOG.ACTION]),
          part_id:          pId,
          part_name:        partsMap[pId] ? partsMap[pId].name : '',
          details:          String(r[COL.LOG.DETAILS]),
          operator:         String(r[COL.LOG.OPERATOR]),
          maintenance_date: _formatDate(r[COL.LOG.MAINTENANCE_DATE])
        };
      })
      .sort(function(a, b) {
        return b.maintenance_date.localeCompare(a.maintenance_date) ||
               b.timestamp.localeCompare(a.timestamp);
      });

    // Parts schedule
    var schedule = scheduleRows
      .filter(function(r) { return String(r[COL.SCHEDULE.MACHINE_ID]).trim() === mId; })
      .map(function(r) {
        var pId  = String(r[COL.SCHEDULE.PART_ID]).trim();
        var next = _formatDate(r[COL.SCHEDULE.NEXT_DUE]);
        var days = next ? _dateDiffDays(next) : 9999;
        return {
          part_id:        pId,
          part_name:      partsMap[pId] ? partsMap[pId].name : pId,
          last_changed:   _formatDate(r[COL.SCHEDULE.LAST_CHANGED]),
          next_due_date:  next,
          status:         next ? _scheduleStatus(days) : 'OK',
          days_remaining: days
        };
      });

    // Defects — newest date_found first
    var defects = defectRows
      .filter(function(r) { return String(r[COL.DEFECT.MACHINE_ID]).trim() === mId; })
      .map(function(r) {
        return {
          defect_id:   String(r[COL.DEFECT.ID]),
          timestamp:   _formatDateTime(r[COL.DEFECT.TIMESTAMP]),
          date_found:  _formatDate(r[COL.DEFECT.DATE_FOUND]),
          symptom:     String(r[COL.DEFECT.SYMPTOM]),
          severity:    String(r[COL.DEFECT.SEVERITY]),
          reported_by: String(r[COL.DEFECT.REPORTED_BY]),
          status:      String(r[COL.DEFECT.STATUS])
        };
      })
      .sort(function(a, b) {
        return b.date_found.localeCompare(a.date_found) ||
               b.timestamp.localeCompare(a.timestamp);
      });

    return { machine: machine, logs: logs, schedule: schedule, defects: defects };
  } catch (e) { return { error: e.message }; }
}

/**
 * Saves a maintenance entry and optionally resolves a linked defect.
 *
 * Rules:
 *   - maintenance_date is always stored in the log.
 *   - part_id is REQUIRED for 'Part Replacement'; OPTIONAL otherwise.
 *   - If a part_id is provided (any action type), the parts schedule is upserted.
 *   - If action_type is 'Repair' AND defect_id is provided, that defect is
 *     marked 'Resolved' in Defect_Log.
 *
 * @param {Object} data
 *   { machine_id, maintenance_date, action_type, part_id,
 *     details, operator, defect_id }
 */
function submitMaintenanceLog(data) {
  try {
    if (!data || !data.machine_id)  return { success: false, error: 'Machine ID is required.' };
    if (!data.maintenance_date)     return { success: false, error: 'Maintenance Date is required.' };
    if (!data.action_type)          return { success: false, error: 'Action Type is required.' };
    if (!data.operator)             return { success: false, error: 'Operator name is required.' };
    if (data.action_type === 'Part Replacement' && !data.part_id) {
      return { success: false, error: 'Part ID is required for Part Replacement.' };
    }

    var now             = new Date();
    var timestamp       = _formatDateTime(now);
    var maintenanceDate = String(data.maintenance_date).trim();
    var partId          = String(data.part_id  || '').trim();
    var defectId        = String(data.defect_id || '').trim();

    // ── Append to Maintenance_Log ──
    _getSheet(SHEETS.LOG).appendRow([
      timestamp, data.machine_id, data.action_type,
      partId, data.details || '', data.operator, maintenanceDate
    ]);

    // ── Schedule upsert (any action type, if part selected) ──
    if (partId) {
      var masterRows = _getSheetData(SHEETS.PARTS_MASTER);
      var partRow    = null;
      for (var i = 0; i < masterRows.length; i++) {
        if (String(masterRows[i][0]).trim() === partId) { partRow = masterRows[i]; break; }
      }
      if (!partRow) return { success: false, error: 'Part not found: ' + partId };
      var lifetimeYears = parseFloat(partRow[COL.PARTS_MASTER.LIFETIME]) || 1;
      _upsertSchedule(data.machine_id, partId, maintenanceDate,
                      _addYears(maintenanceDate, lifetimeYears), 'Active');
    }

    // ── Resolve linked defect (Repair + defect_id selected) ──
    if (data.action_type === 'Repair' && defectId) {
      if (!_resolveDefect(defectId)) {
        return { success: false, error: 'Defect not found: ' + defectId };
      }
    }

    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

/**
 * Creates a new defect report in Defect_Log with Status = 'Pending'.
 * Auto-generates a sequential Defect_ID (e.g. DF-0001).
 *
 * @param {Object} data
 *   { machine_id, date_found, symptom, severity, reported_by }
 * @returns {{ success: boolean, defect_id?: string, error?: string }}
 */
function reportDefect(data) {
  try {
    if (!data || !data.machine_id) return { success: false, error: 'Machine ID is required.' };
    if (!data.date_found)          return { success: false, error: 'Date Found is required.' };
    if (!data.symptom)             return { success: false, error: 'Symptom description is required.' };
    if (!data.severity)            return { success: false, error: 'Severity Level is required.' };
    if (!data.reported_by)         return { success: false, error: 'Reported By is required.' };

    var defectId  = _generateDefectId();
    var timestamp = _formatDateTime(new Date());

    _getSheet(SHEETS.DEFECT_LOG).appendRow([
      defectId, timestamp, data.machine_id,
      data.date_found, data.symptom, data.severity,
      data.reported_by, 'Pending'
    ]);

    return { success: true, defect_id: defectId };
  } catch (e) { return { success: false, error: e.message }; }
}

/**
 * Returns all 'Pending' defects for a given machine.
 * Used to populate the "Link to Defect" dropdown in the Repair form.
 *
 * @param {string} machineId
 * @returns {Array<{defect_id, date_found, symptom, severity, reported_by}>}
 */
function getPendingDefects(machineId) {
  try {
    var mId = String(machineId).trim();
    return _getSheetData(SHEETS.DEFECT_LOG)
      .filter(function(r) {
        return String(r[COL.DEFECT.MACHINE_ID]).trim() === mId &&
               String(r[COL.DEFECT.STATUS]).trim()      === 'Pending';
      })
      .map(function(r) {
        return {
          defect_id:   String(r[COL.DEFECT.ID]),
          date_found:  _formatDate(r[COL.DEFECT.DATE_FOUND]),
          symptom:     String(r[COL.DEFECT.SYMPTOM]),
          severity:    String(r[COL.DEFECT.SEVERITY]),
          reported_by: String(r[COL.DEFECT.REPORTED_BY])
        };
      });
  } catch (e) { return { error: e.message }; }
}
