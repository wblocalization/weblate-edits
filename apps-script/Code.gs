/**
 * Сервер для страницы «Правки переводов».
 * Каждая правка хранится отдельной строкой на листе «Правки».
 *
 * GET  → JSON-массив всех правок
 * POST → {"action":"upsert","items":[...]}  добавить/обновить правки (по id)
 *        {"action":"delete","ids":[...]}     удалить правки по id
 *        ответ: {"ok":true} или {"ok":false,"error":"..."}
 */

const SHEET_NAME = 'Правки';
const FIELDS = ['id', 'platform', 'component', 'language', 'key', 'oldText', 'newText',
                'comment', 'author', 'date', 'status', 'doneDate'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, FIELDS.length).setValues([FIELDS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readAll_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, FIELDS.length).getValues();
  return values
    .filter(row => row[0] !== '')
    .map(row => {
      const it = {};
      FIELDS.forEach((f, i) => { it[f] = row[i] === '' ? '' : String(row[i]); });
      if (!it.doneDate) it.doneDate = null;
      return it;
    });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_(readAll_(getSheet_()));
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // одновременные сохранения от разных людей идут по очереди
  try {
    const req = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const last = sheet.getLastRow();
    const ids = last >= 2 ? sheet.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];

    if (req.action === 'upsert') {
      const newRows = [];
      (req.items || []).forEach(it => {
        if (!it || !it.id) return;
        // апостроф в начале = «храни как текст»: даты не станут числами, а '=…' — формулой
        const row = FIELDS.map(f => (it[f] == null || it[f] === '') ? '' : "'" + String(it[f]));
        const idx = ids.indexOf(String(it.id));
        if (idx >= 0) sheet.getRange(idx + 2, 1, 1, FIELDS.length).setValues([row]);
        else { newRows.push(row); ids.push(String(it.id)); }
      });
      if (newRows.length) {
        const need = sheet.getLastRow() + newRows.length - sheet.getMaxRows();
        if (need > 0) sheet.insertRowsAfter(sheet.getMaxRows(), need);
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, FIELDS.length).setValues(newRows);
      }
    } else if (req.action === 'delete') {
      const del = new Set((req.ids || []).map(String));
      // удаляем снизу вверх, чтобы номера строк не съезжали
      for (let i = ids.length - 1; i >= 0; i--) {
        if (del.has(ids[i])) sheet.deleteRow(i + 2);
      }
    } else {
      return json_({ ok: false, error: 'неизвестное действие: ' + req.action });
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
