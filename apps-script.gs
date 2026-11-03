// 優惠儲蓄專案：接收表單並存到 Google 雲端
const FOLDER_ID = '1qNuIBL0uwx9nIIgzKVdzTIGmDfN9gyAa';
const SHEET_ID  = '1N-U2XHz99tkXOLe9cFRr8yMOEPInsiB4yr_nmeUUPtY';
const COLS = ['現在時間','約訪時間','縣市','區域','學校','年','班','F','辦公室','姓名','性別','生日','電話','分機','理財方式','近期有無聽過','理財規劃','一個月1萬是否有壓力','其他勾選','可約時段','備註'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const name = (d.filename || '優惠儲蓄專案.docx').replace(/[\\/:*?"<>|]/g, '_');
    const blob = Utilities.newBlob(Utilities.base64Decode(d.docxBase64),
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document', name);
    const file = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
    const f = d.fields || {};
    const row = [new Date()].concat(COLS.map(k => f[k] || ''), [file.getUrl()]);
    SpreadsheetApp.openById(SHEET_ID).getSheets()[0].appendRow(row);
    return out({ ok: true, url: file.getUrl(), name: name });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}
function doGet() { return out({ ok: true, msg: '優惠儲蓄專案雲端服務運作中' }); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
// 第一次部署前，請先在編輯器上方選 setup 按「執行」完成授權
function setup() { DriveApp.getFolderById(FOLDER_ID).getName(); SpreadsheetApp.openById(SHEET_ID).getName(); }
