// 優惠儲蓄專案：接收表單 → 存 Word 到雲端 → 寫入名單 → 建立/更新「業務助理行程」行事曆
const FOLDER_ID = '1qNuIBL0uwx9nIIgzKVdzTIGmDfN9gyAa';   // 業務助理客戶資料
const SHEET_ID  = '1N-U2XHz99tkXOLe9cFRr8yMOEPInsiB4yr_nmeUUPtY'; // 優惠儲蓄專案_名單
const CAL_ID    = 'c543d50e3c3202a7f327c26a366e7eb7ffdd199716b7506b1c1d64dc89021124@group.calendar.google.com'; // 業務助理行程
const COLS = ['現在時間','約訪時間','縣市','區域','學校','年','班','F','辦公室','姓名','性別','生日','電話','分機','理財方式','近期有無聽過','理財規劃','一個月1萬是否有壓力','其他勾選','可約時段','備註'];
const HEAD = ['建立時間'].concat(COLS, ['Word檔連結','地點','時長(分)','行程連結','事件ID']);

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const f = d.fields || {};
    const name = (d.filename || '優惠儲蓄專案.docx').replace(/[\\/:*?"<>|]/g, '_');
    const blob = Utilities.newBlob(Utilities.base64Decode(d.docxBase64),
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document', name);
    const file = DriveApp.getFolderById(FOLDER_ID).createFile(blob);

    const sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
    // 找同一位客戶（姓名＋電話）的舊列，用來更新而不重複
    let rowIdx = -1, oldEventId = '';
    const last = sh.getLastRow();
    if (last > 1 && (f['姓名'] || f['電話'])) {
      const data = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
      const iN = HEAD.indexOf('姓名'), iP = HEAD.indexOf('電話'), iE = HEAD.indexOf('事件ID');
      for (let i = data.length - 1; i >= 0; i--) {
        if (String(data[i][iN]) === String(f['姓名'] || '') && String(data[i][iP]) === String(f['電話'] || '')) {
          rowIdx = i + 2; oldEventId = data[i][iE]; break;
        }
      }
    }

    let eventId = oldEventId || '', eventUrl = '';
    if (f['約訪開始']) {
      const cal = CalendarApp.getCalendarById(CAL_ID);
      const start = new Date(f['約訪開始']);
      const end = new Date(start.getTime() + (Number(f['時長']) || 60) * 60000);
      const title = '約訪：' + (f['姓名'] || '未命名') + '（' + [f['學校'], (f['年'] ? f['年'] + '年' : '') + (f['班'] ? f['班'] + '班' : '')].filter(String).join(' ') + '）';
      const desc = [
        '電話：' + (f['電話'] || '') + (f['分機'] ? ' #' + f['分機'] : ''),
        '學校：' + (f['學校'] || '') + '　班級：' + (f['年'] || '') + '年' + (f['班'] || '') + '班 ' + (f['辦公室'] ? f['辦公室'] + '辦公室' : ''),
        '性別：' + (f['性別'] || '') + '　生日：' + (f['生日'] || ''),
        '理財方式：' + (f['理財方式'] || '') + '　近期有無聽過：' + (f['近期有無聽過'] || ''),
        '理財規劃：' + (f['理財規劃'] || '') + '　月存1萬壓力：' + (f['一個月1萬是否有壓力'] || ''),
        '其他：' + (f['其他勾選'] || ''),
        '可約時段：' + (f['可約時段'] || ''),
        '備註：' + (f['備註'] || ''),
        '表單 Word 檔：' + file.getUrl()
      ].join('\n');
      let ev = null;
      if (oldEventId) { try { ev = cal.getEventById(oldEventId); } catch (x) { ev = null; } }
      if (ev) {
        ev.setTitle(title); ev.setTime(start, end); ev.setLocation(f['地點'] || ''); ev.setDescription(desc);
      } else {
        ev = cal.createEvent(title, start, end, { location: f['地點'] || '', description: desc });
      }
      ev.removeAllReminders(); ev.addPopupReminder(30);
      eventId = ev.getId();
      eventUrl = 'https://www.google.com/calendar/event?eid=' +
        Utilities.base64Encode(eventId.split('@')[0] + ' ' + CAL_ID).replace(/=+$/, '');
    }

    const row = [new Date()].concat(COLS.map(k => f[k] || ''), [file.getUrl(), f['地點'] || '', f['時長'] || '', eventUrl, eventId]);
    if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, row.length).setValues([row]); else sh.appendRow(row);

    return out({ ok: true, url: file.getUrl(), name: name, eventUrl: eventUrl, updated: rowIdx > 0 });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}
function doGet() { return out({ ok: true, msg: '優惠儲蓄專案雲端服務運作中' }); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
// 更新後請先在編輯器選 setup 按「執行」，重新授權（含行事曆權限）
function setup() {
  DriveApp.getFolderById(FOLDER_ID).getName();
  SpreadsheetApp.openById(SHEET_ID).getName();
  Logger.log(CalendarApp.getCalendarById(CAL_ID).getName());
}
