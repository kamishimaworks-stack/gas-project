// ■ 使用モデル設定 (Gemini 3 Flash Preview 固定)
const MODEL_NAME = 'gemini-3-flash-preview';

/**
 * スプレッドシートが開かれたときにメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('★図面座標抽出ツール')
    .addItem('1. APIキーを設定 (初回のみ)', 'showApiKeyDialog')
    .addSeparator()
    .addItem('2. 実行画面を開く', 'showSidebar')
    .addToUi();
}

/**
 * APIキー設定ダイアログ
 */
function showApiKeyDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Gemini APIキー設定',
    'Google AI Studioで取得したAPIキーを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    const key = result.getResponseText().trim();
    if (key) {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
      ui.alert('APIキーを保存しました。');
    } else {
      ui.alert('キーが空のため保存しませんでした。');
    }
  }
}

/**
 * サイドバーを表示
 */
function showSidebar() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) {
    SpreadsheetApp.getUi().alert('先にメニューの「APIキーを設定」からキーを保存してください。');
    return;
  }

  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('図面座標抽出 (機能強化版)')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 【フェーズ1】1ページ分の画像を解析してデータを返す関数
 */
function analyzePage(base64Data, mimeType) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) throw new Error("APIキーが設定されていません。");
    return callGeminiApi(apiKey, base64Data, mimeType);
  } catch (e) {
    throw new Error(e.toString());
  }
}

/**
 * 【フェーズ2】全ページのデータを結合してシートに書き込む関数
 */
function writeMergedData(mergedData, sheetName, options) {
  try {
    return createNewSheetAndWrite(mergedData, sheetName, options);
  } catch (e) {
    throw new Error(e.toString());
  }
}

/**
 * アクティブシートのデータをCSV形式で取得
 * @returns {{csv: string, filename: string}} CSV文字列とファイル名
 */
function exportSheetToCsv() {
  // #region agent log
  try { UrlFetchApp.fetch('http://127.0.0.1:7243/ingest/615b7148-d9f3-476a-b274-902297c2f4e8',{method:'POST',headers:{'Content-Type':'application/json'},payload:JSON.stringify({location:'コード.js:exportSheetToCsv:entry',message:'exportSheetToCsv called',data:{},timestamp:Date.now(),hypothesisId:'H4'})}); } catch(e) {}
  // #endregion
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getDataRange();
  const values = range.getValues();

  if (!values || values.length === 0) {
    throw new Error("データがありません");
  }

  const escapeCsv = function (val) {
    const s = String(val == null ? "" : val);
    if (s.includes(",") || s.includes('"') || s.includes("\n") || s.includes("\r")) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  };

  const lines = values.map(row =>
    row.map(cell => escapeCsv(cell)).join(",")
  );
  const csv = "\uFEFF" + lines.join("\r\n"); // UTF-8 BOM for Excel
  const filename = (sheet.getName() || "export") + ".csv";
  // #region agent log
  try { UrlFetchApp.fetch('http://127.0.0.1:7243/ingest/615b7148-d9f3-476a-b274-902297c2f4e8',{method:'POST',headers:{'Content-Type':'application/json'},payload:JSON.stringify({location:'コード.js:exportSheetToCsv:success',message:'exportSheetToCsv result',data:{csvLen:csv.length,filename:filename,rows:values.length},timestamp:Date.now(),hypothesisId:'H4'})}); } catch(e) {}
  // #endregion
  return { csv: csv, filename: filename };
}

/**
 * 選択中のセルの値を取得（優先リスト用）
 * @returns {string[]|null} 選択セルの値の配列。選択がなければnull
 */
function getSelectedCellValues() {
  // #region agent log
  try { UrlFetchApp.fetch('http://127.0.0.1:7243/ingest/615b7148-d9f3-476a-b274-902297c2f4e8',{method:'POST',headers:{'Content-Type':'application/json'},payload:JSON.stringify({location:'コード.js:getSelectedCellValues:entry',message:'getSelectedCellValues called',data:{},timestamp:Date.now(),hypothesisId:'H5'})}); } catch(e) {}
  // #endregion
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = sheet.getActiveRange();
    if (!range) return null;

    const values = range.getValues();
    const result = [];
    const seen = new Set();

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const val = String(values[r][c] || "").trim();
        if (val && !seen.has(val)) {
          seen.add(val);
          result.push(val);
        }
      }
    }
    // #region agent log
    try { UrlFetchApp.fetch('http://127.0.0.1:7243/ingest/615b7148-d9f3-476a-b274-902297c2f4e8',{method:'POST',headers:{'Content-Type':'application/json'},payload:JSON.stringify({location:'コード.js:getSelectedCellValues:success',message:'getSelectedCellValues result',data:{resultLen:result.length},timestamp:Date.now(),hypothesisId:'H5'})}); } catch(e) {}
    // #endregion
    return result.length > 0 ? result : null;
  } catch (e) {
    // #region agent log
    try { UrlFetchApp.fetch('http://127.0.0.1:7243/ingest/615b7148-d9f3-476a-b274-902297c2f4e8',{method:'POST',headers:{'Content-Type':'application/json'},payload:JSON.stringify({location:'コード.js:getSelectedCellValues:error',message:'getSelectedCellValues exception',data:{error:String(e)},timestamp:Date.now(),hypothesisId:'H5'})}); } catch(e2) {}
    // #endregion
    return null;
  }
}

/**
 * 【追加機能】アクティブシートを手動設定に基づいて並び替え
 * @param {string} priorityInput - 優先項目（カンマ区切り）
 * @param {boolean} doSort - 測点名順に並び替えるか
 */
function sortActiveSheetManually(priorityInput, doSort) {
  // #region agent log
  try { UrlFetchApp.fetch('http://127.0.0.1:7243/ingest/615b7148-d9f3-476a-b274-902297c2f4e8',{method:'POST',headers:{'Content-Type':'application/json'},payload:JSON.stringify({location:'コード.js:sortActiveSheetManually:entry',message:'sortActiveSheetManually called',data:{priorityInput:priorityInput,doSort:doSort},timestamp:Date.now(),hypothesisId:'H3'})}); } catch(e) {}
  // #endregion
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getDataRange();
  const values = range.getValues();
  
  if (values.length < 2) return "データがありません";
  
  const headers = values[0];
  const rows = values.slice(1);
  
  // 共通の整理ロジックを使って並び替え
  const result = organizeData(headers, rows, priorityInput, doSort);
  
  // #region agent log
  try { UrlFetchApp.fetch('http://127.0.0.1:7243/ingest/615b7148-d9f3-476a-b274-902297c2f4e8',{method:'POST',headers:{'Content-Type':'application/json'},payload:JSON.stringify({location:'コード.js:sortActiveSheetManually:beforeWrite',message:'before setValues',data:{resultRowsLen:result.rows.length,resultHeadersLen:result.headers.length,firstRowLen:result.rows[0]?result.rows[0].length:0},timestamp:Date.now(),hypothesisId:'H1'})}); } catch(e) {}
  // #endregion
  
  // シートをクリアして書き直し
  sheet.clearContents();
  
  // ヘッダー書き込み
  sheet.getRange(1, 1, 1, result.headers.length)
    .setValues([result.headers])
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setBorder(true, true, true, true, true, true);
    
  // データ書き込み
  if (result.rows.length > 0) {
    sheet.getRange(2, 1, result.rows.length, result.headers.length).setValues(result.rows);
  }
  
  return `完了: 設定に基づいて並び替えました。\n(優先項目: ${priorityInput || "なし"})`;
}

/**
 * データ整理の共通ロジック（列移動・行移動・ソート）
 */
function organizeData(currentHeaders, currentRows, priorityInput, doSort) {
  let headers = [...currentHeaders];
  let rows = [...currentRows];

  // 1. 優先項目の解析（列名なのか、行の測点名なのかを判別）
  const priorityList = priorityInput ? priorityInput.split(/[,\s]+/).map(s => s.trim()).filter(s => s) : [];
  
  const priorityCols = [];
  const priorityRowKeys = [];

  priorityList.forEach(item => {
    // ヘッダーに部分一致するものがあれば「列」として扱う
    if (headers.some(h => h === item || h.includes(item))) {
      priorityCols.push(item);
    } else {
      // それ以外は「測点名（行）」として扱う
      priorityRowKeys.push(item);
    }
  });

  // 2. 列の並び替え
  if (priorityCols.length > 0) {
    const headerIndices = headers.map((h, i) => ({ name: h, index: i }));
    const priorityIndices = [];
    const otherIndices = [];

    priorityCols.forEach(pName => {
      const found = headerIndices.find(h => h.name === pName || h.name.includes(pName));
      if (found && !priorityIndices.includes(found)) priorityIndices.push(found);
    });

    headerIndices.forEach(h => {
      if (!priorityIndices.includes(h)) otherIndices.push(h);
    });

    const newOrderMap = [...priorityIndices, ...otherIndices].map(h => h.index);
    headers = [...priorityIndices, ...otherIndices].map(h => h.name);
    rows = rows.map(row => newOrderMap.map(idx => row[idx]));
  }

  // 3. 行の並び替え（優先行を上に、残りをオプションに従ってソート）
  const topRows = [];
  const normalRows = [];

  rows.forEach(row => {
    const key = String(row[0]).trim(); // 1列目（測点名）
    if (priorityRowKeys.includes(key)) {
      topRows.push(row);
    } else {
      normalRows.push(row);
    }
  });

  // 優先行を入力された順序に並べる
  topRows.sort((a, b) => {
    const ka = String(a[0]).trim();
    const kb = String(b[0]).trim();
    return priorityRowKeys.indexOf(ka) - priorityRowKeys.indexOf(kb);
  });

  // 残りの行をソート（doSortがtrueの場合）
  if (doSort) {
    const collator = new Intl.Collator("ja", {numeric: true, sensitivity: 'base'});
    normalRows.sort((a, b) => {
      const valA = String(a[0] || "");
      const valB = String(b[0] || "");
      return collator.compare(valA, valB);
    });
  }

  return {
    headers: headers,
    rows: [...topRows, ...normalRows]
  };
}

/**
 * データを整理してシートに書き込む（新規作成用）
 */
function createNewSheetAndWrite(data, sheetNameInput, options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newSheet = ss.insertSheet(ss.getNumSheets());
  
  if (sheetNameInput && sheetNameInput.trim() !== "") {
    let sheetName = sheetNameInput.trim();
    let counter = 2;
    while (ss.getSheetByName(sheetName)) {
      sheetName = `${sheetNameInput.trim()}(${counter})`;
      counter++;
    }
    newSheet.setName(sheetName);
  }

  if (!data || !data.allRows || data.allRows.length === 0) {
    newSheet.getRange(1, 1).setValue("データが見つかりませんでした");
    return "データなし";
  }

  // 列数の正規化
  let headers = data.detectedHeaders || [];
  let maxCols = headers.length;
  if (data.allRows) {
    data.allRows.forEach(r => {
      if (r.values && r.values.length > maxCols) maxCols = r.values.length;
    });
  }
  while (headers.length < maxCols) headers.push("");
  let allRows = data.allRows.map(r => {
    let vals = [...(r.values || [])];
    while (vals.length < maxCols) vals.push("");
    return vals;
  });

  // 面積・計算過程などの除外フィルタ
  const ignoreKeywords = ["面積", "地積", "倍面積", "合計", "計"];
  const calcProcessPattern = /^[XY]\d*n$|^[XY]\d*n\s*-\s*[XY]\d*n$/i; // Xn, Yn, Xn-Yn 等
  const validColIndices = [];
  const filteredHeaders = [];
  headers.forEach((h, i) => {
    const hStr = String(h || "").trim();
    if (hStr === "") return;
    if (ignoreKeywords.some(kw => hStr.includes(kw))) return;
    if (calcProcessPattern.test(hStr)) return; // 計算過程の列を除外
    validColIndices.push(i);
    filteredHeaders.push(h);
  });

  let filteredRows = allRows.map(row => validColIndices.map(i => row[i]));
  let uniqueRows = deduplicateRows(filteredRows).filter(row => {
    const firstCell = String(row[0]).trim();
    if (!firstCell) return false;
    return !ignoreKeywords.some(kw => firstCell.includes(kw));
  });

  // Z座標列の空欄を「0」で埋める（標高なしの測点用）
  const zColKeywords = ["Z", "Z座標", "標高", "高さ", "H"];
  const zColIndices = filteredHeaders
    .map((h, i) => (zColKeywords.some(kw => String(h || "").includes(kw)) ? i : -1))
    .filter(i => i >= 0);
  uniqueRows = uniqueRows.map(row => {
    const newRow = [...row];
    zColIndices.forEach(idx => {
      const val = String(newRow[idx] || "").trim();
      if (val === "") newRow[idx] = "0";
    });
    return newRow;
  });

  if (uniqueRows.length === 0) {
    newSheet.getRange(1, 1).setValue("有効なデータ行がありません");
    return "データなし";
  }

  // ★共通ロジックを使って並び替えを適用
  const organized = organizeData(filteredHeaders, uniqueRows, options.priorityColumns, options.sort);

  // #region agent log
  try { UrlFetchApp.fetch('http://127.0.0.1:7243/ingest/615b7148-d9f3-476a-b274-902297c2f4e8',{method:'POST',headers:{'Content-Type':'application/json'},payload:JSON.stringify({location:'コード.js:createNewSheetAndWrite:beforeWrite',message:'before setValues',data:{organizedRowsLen:organized.rows.length,organizedHeadersLen:organized.headers.length,optionsKeys:Object.keys(options||{})},timestamp:Date.now(),hypothesisId:'H2'})}); } catch(e) {}
  // #endregion

  // 書き込み
  newSheet.getRange(1, 1, 1, organized.headers.length)
    .setValues([organized.headers])
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setBorder(true, true, true, true, true, true);

  newSheet.getRange(2, 1, organized.rows.length, organized.headers.length).setValues(organized.rows);
  newSheet.autoResizeColumns(1, organized.headers.length);
  ss.setActiveSheet(newSheet);

  return `Gemini 3.0: 全${organized.rows.length}行のデータを抽出しました。\n(項目: ${organized.headers.join(", ")})`;
}

/**
 * Gemini API リクエスト処理
 */
function callGeminiApi(apiKey, base64Data, mimeType) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${apiKey}`;
  const promptText = `
    あなたは測量の専門家です。提供された画像の図面から、すべての「求積表（座標リスト）」を見つけ出し、データを抽出してください。
    【抽出ルール (厳守)】
    1. **対象データ**: 表の内側にある「測点」および「座標値（X, Y, Zなど）」、「辺長」などのデータを抽出してください。
    2. **除外データ**: 表の下部や右端にある集計情報（合計、倍面積、面積、地積など）は**絶対に抽出しないでください**。
    3. **無視する情報**: 表の外側にある「地番」「箇所」などのタイトル情報は無視してください。
    4. **面積・計算過程の除外**: 「面積」「地積」「倍面積」などの面積関連列、および「Xn」「Yn」「Xn-Yn」など計算過程・中間計算用の列は**抽出しないでください**。表内の計算過程を示す行（座標差のみの行など）も**抽出しないでください**。
    5. **Z座標（標高）の処理**: 図面に「標高」「高さ」「Z」の記載がある場合は、それをZ座標（高さ）として抽出してください。標高が記載されていない測点・行では、Z座標列に必ず**「0」**を入れてください（空欄禁止。機械がXYZデータを必要とするため）。
    6. **統合**: 複数の表がある場合、それらを統合して一つのリストにまとめてください。
    7. **OCR補正**: O(オー)と0(ゼロ)、l(エル)と1(イチ)などの誤読は補正してください。
  `;
  const responseSchema = {
    type: "OBJECT",
    properties: {
      detectedHeaders: { type: "ARRAY", items: { type: "STRING" } },
      allRows: {
        type: "ARRAY",
        items: {
          type: "OBJECT",
          properties: { values: { type: "ARRAY", items: { type: "STRING" } } }
        }
      }
    },
    required: ["detectedHeaders", "allRows"]
  };
  const payload = {
    contents: [{ parts: [{ text: promptText }, { inline_data: { mime_type: mimeType, data: base64Data } }] }],
    generationConfig: { response_mime_type: "application/json", response_schema: responseSchema }
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const maxRetries = 3;
  for (let i = 0; i < maxRetries; i++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const json = JSON.parse(response.getContentText());
      if (response.getResponseCode() !== 200 || json.error) throw new Error(json.error ? json.error.message : response.getContentText());
      if (!json.candidates || !json.candidates[0].content) throw new Error("データが見つかりませんでした。");
      return JSON.parse(json.candidates[0].content.parts[0].text);
    } catch (e) {
      if (i < maxRetries - 1) Utilities.sleep(2000 * (i + 1));
      else throw e;
    }
  }
}

function deduplicateRows(rows) {
  const seen = new Set();
  const result = [];
  for (const row of rows) {
    const key = row[0] ? String(row[0]).trim() : "";
    if (key && !seen.has(key)) {
      seen.add(key);
      result.push(row);
    }
  }
  return result;
}