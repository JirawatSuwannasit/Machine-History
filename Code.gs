// ============================================================
// CONSTANTS
// ============================================================

var SHEETS = {
  MACHINE_LIST: 'Machine_List',
  PARTS_MASTER: 'Spare_Parts_Master',
  SCHEDULE:     'Spare_Parts_Schedule',
  LOG:          'Maintenance_Log',
  DEFECT_LOG:   'Defect_Log',
  AUDIT_LOG:    'Audit_Log'
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
    SYMPTOM: 4, SEVERITY: 5, REPORTED_BY: 6, STATUS: 7,
    ROOT_CAUSE: 8, CORRECTIVE_ACTION: 9, RESOLVED_BY: 10,
    RESOLVED_AT: 11, LINKED_MAINTENANCE_ID: 12
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
  var d = new Date(dateStr + 'T00:00:00');
  var y = parseFloat(years);
  if (isNaN(d.getTime()) || isNaN(y)) return '';
  var wholeYears = Math.trunc(y);
  var fractional = y - wholeYears;
  d.setFullYear(d.getFullYear() + wholeYears);
  if (fractional !== 0) d.setMonth(d.getMonth() + Math.round(fractional * 12));
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

function _required(value, label) {
  if (value == null || String(value).trim() === '') throw new Error(label + ' is required.');
  return String(value).trim();
}

function _isAllowed(value, allowedValues) {
  return allowedValues.indexOf(String(value)) !== -1;
}

function _validateDateString(dateStr) {
  var v = String(dateStr || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) return false;
  var d = new Date(v + 'T00:00:00');
  return !isNaN(d.getTime()) && _formatDate(d) === v;
}

function _machineExists(machineId) {
  var mId = String(machineId).trim();
  return _getSheetData(SHEETS.MACHINE_LIST).some(function(r){ return String(r[COL.MACHINE.ID]).trim() === mId; });
}

function _partExists(partId) {
  var pId = String(partId).trim();
  var rows = _getSheetData(SHEETS.PARTS_MASTER);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][COL.PARTS_MASTER.ID]).trim() === pId) return rows[i];
  }
  return null;
}

function _getDefectById(defectId) {
  var dId = String(defectId).trim();
  var rows = _getSheetData(SHEETS.DEFECT_LOG);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][COL.DEFECT.ID]).trim() === dId) return rows[i];
  }
  return null;
}

function _getDefectMetaColumnsOrThrow() {
  var headers = _getSheet(SHEETS.DEFECT_LOG).getRange(1, 1, 1, _getSheet(SHEETS.DEFECT_LOG).getLastColumn()).getValues()[0].map(function(h){return String(h).trim();});
  var required = ['root_cause','corrective_action','resolved_by','resolved_at','linked_maintenance_id'];
  var idx = {};
  required.forEach(function(name){
    var i = headers.indexOf(name);
    if (i === -1) throw new Error('Missing required Defect_Log column: ' + name);
    idx[name]=i+1;
  });
  return idx;
}

function _appendAuditRow(action, entityType, entityId, oldValue, newValue) {
  var sheet = _getSheet(SHEETS.AUDIT_LOG);
  var email = '';
  try { email = Session.getActiveUser().getEmail() || 'unknown'; } catch (e) { email = 'unknown'; }
  sheet.appendRow([_formatDateTime(new Date()), email, action, entityType, entityId, oldValue || '', newValue || '']);
}

function _logAudit(action, entityType, entityId, oldValue, newValue) {
  try { _appendAuditRow(action, entityType, entityId, JSON.stringify(oldValue || {}), JSON.stringify(newValue || {})); }
  catch (e) { throw new Error('Audit_Log write failed: ' + e.message); }
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
  var oldState = null;
  if (foundIdx !== -1) {
    var row = foundIdx + 1;
    oldState = { machine_id: data[foundIdx][0], part_id: data[foundIdx][1], last_changed: data[foundIdx][2], next_due: data[foundIdx][3], status: data[foundIdx][4] };
    sheet.getRange(row, 3).setValue(lastChanged);
    sheet.getRange(row, 4).setValue(nextDueDate);
    sheet.getRange(row, 5).setValue(status);
  } else {
    sheet.appendRow([machineId, partId, lastChanged, nextDueDate, status]);
  }
  _logAudit('schedule_updated', 'Spare_Parts_Schedule', machineId + '|' + partId, oldState, {
    machine_id: machineId, part_id: partId, last_changed: lastChanged, next_due: nextDueDate, status: status
  });
}

// Generates next sequential Defect_ID (e.g. DF-0001).
// Scans existing IDs to find the highest number — safe against gaps/deletions.
function _generateDefectId() {
  if (!globalThis.__WRITE_LOCK_HELD__) {
    throw new Error('_generateDefectId must be called inside a locked write operation.');
  }
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
  var metaCols = _getDefectMetaColumnsOrThrow();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL.DEFECT.ID]).trim() === String(defectId).trim()) {
      sheet.getRange(i + 1, COL.DEFECT.STATUS + 1).setValue('Resolved');
      sheet.getRange(i + 1, metaCols.resolved_at).setValue(_formatDateTime(new Date()));
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
    var defects = _getSheetData(SHEETS.DEFECT_LOG);
    var scheduleRows = _getSheetData(SHEETS.SCHEDULE);
    return _getSheetData(SHEETS.MACHINE_LIST)
      .filter(function(r) { return String(r[0]).trim() !== ''; })
      .map(function(r) {
        var machineId = String(r[COL.MACHINE.ID]).trim();
        var pendingDefectCount = 0, criticalPendingDefectCount = 0;
        defects.forEach(function(d) {
          if (String(d[COL.DEFECT.MACHINE_ID]).trim() === machineId && String(d[COL.DEFECT.STATUS]).trim() === 'Pending') {
            pendingDefectCount++;
            if (String(d[COL.DEFECT.SEVERITY]).trim() === 'Critical') criticalPendingDefectCount++;
          }
        });
        var overduePartCount = 0, dueSoonPartCount = 0;
        scheduleRows.forEach(function(s) {
          if (String(s[COL.SCHEDULE.MACHINE_ID]).trim() !== machineId) return;
          var next = _formatDate(s[COL.SCHEDULE.NEXT_DUE]);
          if (!next) return;
          var days = _dateDiffDays(next);
          if (days < 0) overduePartCount++;
          else if (days <= 30) dueSoonPartCount++;
        });
        return {
          machine_id:     String(r[COL.MACHINE.ID]),
          scope:          String(r[COL.MACHINE.SCOPE]),
          machine_name:   String(r[COL.MACHINE.NAME]),
          manufacturer:   String(r[COL.MACHINE.MANUFACTURER]),
          model:          String(r[COL.MACHINE.MODEL]),
          sn:             String(r[COL.MACHINE.SN]),
          range:          String(r[COL.MACHINE.RANGE]),
          operation_date: _formatDate(r[COL.MACHINE.OP_DATE]),
          status:         String(r[COL.MACHINE.STATUS]),
          pending_defect_count: pendingDefectCount,
          critical_pending_defect_count: criticalPendingDefectCount,
          overdue_part_count: overduePartCount,
          due_soon_part_count: dueSoonPartCount
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
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    globalThis.__WRITE_LOCK_HELD__ = true;
    var machineId = _required(data && data.machine_id, 'Machine ID');
    var maintenanceDate = _required(data && data.maintenance_date, 'Maintenance Date');
    var actionType = _required(data && data.action_type, 'Action Type');
    var operator = _required(data && data.operator, 'Operator name');
    var partId = String((data && data.part_id) || '').trim();
    var defectId = String((data && data.defect_id) || '').trim();
    var details = String((data && data.details) || '');
    if (!_machineExists(machineId)) throw new Error('Machine not found: ' + machineId);
    if (!_validateDateString(maintenanceDate)) throw new Error('Maintenance Date must be YYYY-MM-DD.');
    if (!_isAllowed(actionType, ['Repair', 'Part Replacement', 'PM'])) throw new Error('Invalid Action Type.');
    if (actionType === 'Part Replacement' && !partId) throw new Error('Part ID is required for Part Replacement.');
    var partRow = null;
    if (partId) {
      partRow = _partExists(partId);
      if (!partRow) throw new Error('Part not found: ' + partId);
    }
    var defectRow = null;
    if (defectId) {
      defectRow = _getDefectById(defectId);
      if (!defectRow) throw new Error('Defect not found: ' + defectId);
      if (String(defectRow[COL.DEFECT.MACHINE_ID]).trim() !== machineId) throw new Error('Defect does not belong to selected machine.');
      if (String(defectRow[COL.DEFECT.STATUS]).trim() !== 'Pending') throw new Error('Defect must be Pending to resolve.');
    }

    var now             = new Date();
    var timestamp       = _formatDateTime(now);

    // all validations done; write
    _getSheet(SHEETS.LOG).appendRow([
      timestamp, machineId, actionType,
      partId, details, operator, maintenanceDate
    ]);
    var maintenanceId = timestamp + '|' + machineId;
    _logAudit('maintenance_created', 'Maintenance_Log', maintenanceId, null, {
      machine_id: machineId, action_type: actionType, part_id: partId, defect_id: defectId, operator: operator, maintenance_date: maintenanceDate
    });

    if (partId) {
      var lifetimeYears = parseFloat(partRow[COL.PARTS_MASTER.LIFETIME]) || 1;
      _upsertSchedule(machineId, partId, maintenanceDate,
                      _addYears(maintenanceDate, lifetimeYears), 'Active');
    }

    if (actionType === 'Repair' && defectId) {
      var metaCols = _getDefectMetaColumnsOrThrow();
      if (!_resolveDefect(defectId)) throw new Error('Defect not found: ' + defectId);
      var sheet = _getSheet(SHEETS.DEFECT_LOG);
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][COL.DEFECT.ID]).trim() === defectId) {
          sheet.getRange(i + 1, metaCols.root_cause).setValue(String(data.root_cause || ''));
          sheet.getRange(i + 1, metaCols.corrective_action).setValue(String(data.corrective_action || ''));
          sheet.getRange(i + 1, metaCols.resolved_by).setValue(operator);
          sheet.getRange(i + 1, metaCols.linked_maintenance_id).setValue(maintenanceId);
          break;
        }
      }
      _logAudit('defect_resolved', 'Defect_Log', defectId, { status: 'Pending' }, {
        status: 'Resolved', root_cause: String(data.root_cause || ''), corrective_action: String(data.corrective_action || ''), resolved_by: operator
      });
    }

    return { success: true, maintenance_id: maintenanceId };
  } catch (e) { return { success: false, error: e.message }; }
  finally {
    globalThis.__WRITE_LOCK_HELD__ = false;
    try { lock.releaseLock(); } catch (ignore) {}
  }
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
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    globalThis.__WRITE_LOCK_HELD__ = true;
    var machineId = _required(data && data.machine_id, 'Machine ID');
    var dateFound = _required(data && data.date_found, 'Date Found');
    var symptom = _required(data && data.symptom, 'Symptom description');
    var severity = _required(data && data.severity, 'Severity Level');
    var reportedBy = _required(data && data.reported_by, 'Reported By');
    if (!_machineExists(machineId)) throw new Error('Machine not found: ' + machineId);
    if (!_validateDateString(dateFound)) throw new Error('Date Found must be YYYY-MM-DD.');
    if (!_isAllowed(severity, ['Low', 'Medium', 'High', 'Critical'])) throw new Error('Invalid Severity Level.');

    var defectId  = _generateDefectId();
    var timestamp = _formatDateTime(new Date());

    _getSheet(SHEETS.DEFECT_LOG).appendRow([
      defectId, timestamp, machineId,
      dateFound, symptom, severity,
      reportedBy, 'Pending', '', '', '', '', ''
    ]);
    _logAudit('defect_reported', 'Defect_Log', defectId, null, {
      machine_id: machineId, date_found: dateFound, severity: severity, reported_by: reportedBy, status: 'Pending'
    });

    return { success: true, defect_id: defectId };
  } catch (e) { return { success: false, error: e.message }; }
  finally {
    globalThis.__WRITE_LOCK_HELD__ = false;
    try { lock.releaseLock(); } catch (ignore) {}
  }
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
