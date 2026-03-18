/**
 * TriggerService.gs
 * Manages time-based triggers for automated processing and month-end archival.
 */

/**
 * Installs daily (2-3 AM) and weekly Sunday (3-4 AM) triggers.
 * Idempotent: removes existing managed triggers first.
 */
function installTriggers() {
  uninstallTriggers();

  // Daily check — only acts on the last day of the month
  ScriptApp.newTrigger('monthEndTriggerHandler')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();

  ScriptApp.newTrigger('weeklyTriggerHandler')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(3)
    .create();

  Logger.log('TriggerService: triggers instalados (diario 2AM, semanal domingo 3AM).');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Triggers instalados:\n- Mes: revisa diario a las 2 AM, archiva solo el último día\n- Semanal: domingos a las 3 AM, procesa recibos',
    'Recibos', 10
  );
}

/**
 * Removes all triggers that call our handler functions.
 */
function uninstallTriggers() {
  var removed = 0;
  var handlerNames = ['monthEndTriggerHandler', 'weeklyTriggerHandler'];

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (handlerNames.indexOf(trigger.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });

  if (removed > 0) {
    Logger.log('TriggerService: %s triggers eliminados.', removed);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      removed + ' trigger(s) eliminados.', 'Recibos', 5
    );
  }
}

/**
 * Month-end trigger handler (runs daily at 2 AM, acts only on the last day).
 * Archives the month's data into a snapshot tab, then processes pending receipts.
 */
function monthEndTriggerHandler() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('TriggerService: monthEndTriggerHandler no pudo obtener lock, abortando.');
    return;
  }

  try {
    if (_isLastDayOfMonth()) {
      Logger.log('TriggerService: último día del mes, archivando...');
      archiveMonth(new Date());
      _processReceiptsCore(SpreadsheetApp.getActiveSpreadsheet());
    }
  } catch (e) {
    Logger.log('TriggerService: error en monthEndTriggerHandler: %s', e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Weekly trigger handler (runs Sunday 3 AM).
 * Processes all pending receipts and bank statements.
 */
function weeklyTriggerHandler() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('TriggerService: weeklyTriggerHandler no pudo obtener lock, abortando.');
    return;
  }

  try {
    _processReceiptsCore(SpreadsheetApp.getActiveSpreadsheet());
  } catch (e) {
    Logger.log('TriggerService: error en weeklyTriggerHandler: %s', e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Checks if today is the last day of the current month.
 * @returns {boolean}
 */
function _isLastDayOfMonth() {
  var today = new Date();
  var lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  return today.getDate() === lastDay;
}
