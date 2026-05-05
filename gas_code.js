// GETリクエスト：ロード専用
function doGet(e) {
  const action = e.parameter.action;
  const tripId = e.parameter.trip;

  if (!tripId) return output({ ok: false, error: 'trip id missing' });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (action === 'load') {
    // 旧フォーマット移行: 1行目がJSONの場合 → trip=default として移行
    const firstA = sheet.getRange(1, 1).getValue();
    if (tripId === 'default' && String(firstA).startsWith('{')) {
      const payload = firstA;
      sheet.getRange(1, 1).setValue('default');
      sheet.getRange(1, 2).setValue(payload);
      return output({ ok: true, data: payload });
    }

    // 通常検索
    const values = sheet.getDataRange().getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]) === tripId) {
        return output({ ok: true, data: values[i][1] });
      }
    }
    return output({ ok: false });
  }

  return output({ ok: false, error: 'use POST for save' });
}

// POSTリクエスト：セーブ専用（ペイロードが大きくてもOK）
function doPost(e) {
  const tripId = e.parameter.trip;

  if (!tripId) return output({ ok: false, error: 'trip id missing' });

  const payload = e.postData ? e.postData.getDataAsString() : '';
  if (!payload) return output({ ok: false, error: 'no payload' });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const values = sheet.getDataRange().getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === tripId) {
      sheet.getRange(i + 1, 2).setValue(payload);
      return output({ ok: true });
    }
  }

  sheet.appendRow([tripId, payload]);
  return output({ ok: true });
}

function output(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
