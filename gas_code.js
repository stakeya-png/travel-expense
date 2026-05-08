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

  // スプレッドシート生成（新タブ方式 - CORS不要）
  if (action === 'createSheet') {
    const values = sheet.getDataRange().getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]) === tripId) {
        try {
          const data = JSON.parse(values[i][1]);
          const url = buildSpreadsheet(data);
          // スプレッドシートへリダイレクト
          return HtmlService.createHtmlOutput(
            '<html><head><meta charset="UTF-8"></head><body>' +
            '<p>スプレッドシートを作成しました。</p>' +
            '<a href="' + url + '">クリックして開く</a>' +
            '<script>window.location.href="' + url + '";</script>' +
            '</body></html>'
          );
        } catch(err) {
          return HtmlService.createHtmlOutput('エラー: ' + err.toString());
        }
      }
    }
    return HtmlService.createHtmlOutput('データが見つかりませんでした。先にデータを保存してください。');
  }

  return output({ ok: false, error: 'unknown action' });
}

// POSTリクエスト：セーブ・スプレッドシート生成
function doPost(e) {
  const tripId = e.parameter.trip;
  const action = e.parameter.action || 'save';

  if (!tripId) return output({ ok: false, error: 'trip id missing' });

  const payload = e.postData ? e.postData.getDataAsString() : '';
  if (!payload) return output({ ok: false, error: 'no payload' });

  // スプレッドシート生成
  if (action === 'createSheet') {
    try {
      const data = JSON.parse(payload);
      const url = buildSpreadsheet(data);
      return output({ ok: true, url: url });
    } catch(err) {
      return output({ ok: false, error: err.toString() });
    }
  }

  // デフォルト：データ保存
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

// ===== スプレッドシート生成 =====
function buildSpreadsheet(data) {
  const tripName = data.tripName || '旅行精算';
  const families = data.families || [];
  const expenses = data.expenses || [];
  const rate = data.rate || 160;
  const RATIOS = { adult: 10, teen: 7, child: 5 };

  // スプレッドシートを新規作成
  const fileName = tripName + ' 精算一覧';
  const ss = SpreadsheetApp.create(fileName);
  const sheet = ss.getSheets()[0];
  sheet.setName('精算一覧');

  // メンバー一覧（家族順）
  const allMembers = [];
  families.forEach(f => {
    f.members.forEach(m => {
      allMembers.push({ id: m.id, name: m.name, cat: m.cat, familyId: f.id, familyName: f.name });
    });
  });

  // ===== セクション1: 支払い明細 =====
  const rows = [];

  // ヘッダー行
  rows.push(['日付', '件名', '支払者', '円', '$', ...allMembers.map(m => m.name)]);

  // データ行
  expenses.forEach(e => {
    const payer = allMembers.find(m => m.id === e.payerMemberId);
    const payerLabel = payer ? payer.familyName + ' ' + payer.name : '不明';
    const amountUSD = parseFloat((e.amountJPY / rate).toFixed(2));
    const cols = allMembers.map(m => e.participants.includes(m.id) ? '○' : '');
    rows.push([e.date || '', e.desc, payerLabel, Math.round(e.amountJPY), amountUSD, ...cols]);
  });

  // 空白行×2
  rows.push([]);
  rows.push([]);

  // ===== セクション2: 家族別集計 =====
  const familyPayment = {};
  const familyUsage = {};
  families.forEach(f => { familyPayment[f.id] = 0; familyUsage[f.id] = 0; });

  expenses.forEach(e => {
    const payer = allMembers.find(m => m.id === e.payerMemberId);
    if (payer) familyPayment[payer.familyId] += e.amountJPY;

    const totalRatio = e.participants.reduce((sum, mid) => {
      const m = allMembers.find(x => x.id === mid);
      return sum + (m ? RATIOS[m.cat] : 0);
    }, 0);
    if (totalRatio === 0) return;

    e.participants.forEach(mid => {
      const m = allMembers.find(x => x.id === mid);
      if (!m) return;
      familyUsage[m.familyId] += e.amountJPY * RATIOS[m.cat] / totalRatio;
    });
  });

  rows.push(['家族別集計', ...families.map(f => f.name)]);
  rows.push(['支払合計（円）', ...families.map(f => Math.round(familyPayment[f.id] || 0))]);
  rows.push(['支払合計（$）',  ...families.map(f => parseFloat((familyPayment[f.id] / rate).toFixed(2)))]);
  rows.push(['使用合計（円）', ...families.map(f => Math.round(familyUsage[f.id] || 0))]);
  rows.push(['使用合計（$）',  ...families.map(f => parseFloat((familyUsage[f.id] / rate).toFixed(2)))]);
  rows.push(['差額（円）',     ...families.map(f => Math.round((familyPayment[f.id] || 0) - (familyUsage[f.id] || 0)))]);

  // 全行を一括書き込み
  const maxCols = Math.max(...rows.map(r => r.length));
  const paddedRows = rows.map(r => { while (r.length < maxCols) r.push(''); return r; });
  sheet.getRange(1, 1, paddedRows.length, maxCols).setValues(paddedRows);

  // ===== 書式設定 =====
  const totalCols = maxCols;

  // ヘッダー行：黒背景・白文字・太字
  const hdr = sheet.getRange(1, 1, 1, totalCols);
  hdr.setFontWeight('bold').setBackground('#1a1a18').setFontColor('#ffffff').setHorizontalAlignment('center');

  // データ行：円・ドル列を右寄せ
  if (expenses.length > 0) {
    sheet.getRange(2, 4, expenses.length, 2).setHorizontalAlignment('right');
    sheet.getRange(2, 6, expenses.length, allMembers.length).setHorizontalAlignment('center');
  }

  // 集計セクション：ヘッダー
  const summaryStart = expenses.length + 4; // 1(ヘッダー) + N(データ) + 2(空白) + 1
  const summaryHdr = sheet.getRange(summaryStart, 1, 1, families.length + 1);
  summaryHdr.setFontWeight('bold').setBackground('#e8e5e0').setHorizontalAlignment('center');

  // 数値行を右寄せ
  sheet.getRange(summaryStart + 1, 2, 5, families.length).setHorizontalAlignment('right');

  // 差額行：プラスは緑・マイナスは赤
  families.forEach((f, i) => {
    const diff = Math.round((familyPayment[f.id] || 0) - (familyUsage[f.id] || 0));
    const cell = sheet.getRange(summaryStart + 5, i + 2);
    cell.setFontWeight('bold');
    if (diff > 0) cell.setFontColor('#217a38');
    else if (diff < 0) cell.setFontColor('#b83020');
  });
  sheet.getRange(summaryStart + 5, 1).setFontWeight('bold');

  // 先頭行・先頭3列を固定
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);

  // 列幅自動調整
  sheet.autoResizeColumns(1, totalCols);

  return ss.getUrl();
}

function output(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
