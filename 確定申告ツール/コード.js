/**
 * 確定申告ツール v1.0
 * レシート撮影 → AI仕訳判定 → 弥生会計CSV出力
 */

// ── 設定 ──────────────────────────────────────
const scriptProps = PropertiesService.getScriptProperties();
const PROPS = scriptProps ? scriptProps.getProperties() || {} : {};

const CONFIG = {
  API_KEY: PROPS.GEMINI_API_KEY || '',
  RECEIPT_FOLDER_ID: PROPS.RECEIPT_FOLDER_ID || '',
  SPREADSHEET_ID: PROPS.SPREADSHEET_ID || '',
  SHEET_NAME: '経費データ',
  GEMINI_MODEL: 'gemini-3-flash-preview',
  API_ENDPOINT: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent',
};

const SHEET_HEADERS = [
  'ID', 'タイムスタンプ', '取引日付', '店舗名', '金額', '税率',
  '勘定科目', '支払方法', '摘要', '分類', 'ステータス',
  '画像URL', 'ユーザーメモ', 'AI判定理由', 'フォローアップ回答',
  '取引種別'
];

// ── エントリーポイント ──────────────────────────
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('確定申告ツール')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── ユーティリティ ──────────────────────────────
function parseCurrency(val) {
  if (!val) return 0;
  let str = String(val);
  str = str.replace(/[０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  const num = Number(str.replace(/[^0-9.-]+/g, ''));
  return isNaN(num) ? 0 : num;
}

function formatDate(d) {
  try {
    if (!d) return '';
    return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  } catch (e) { return d; }
}

function getSpreadsheet_() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(SHEET_HEADERS);
    sheet.setFrozenRows(1);
    formatSheet_(sheet);
  } else {
    // マイグレーション: 既存シートに「取引種別」列がなければ追加
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headerRow.indexOf('取引種別') === -1) {
      var nextCol = headerRow.length + 1;
      sheet.getRange(1, nextCol).setValue('取引種別');
    }
  }
  return sheet;
}

/**
 * メニューから手動でシート書式を適用（初回 or リセット用）
 */
function setupSheet() {
  const sheet = getOrCreateSheet_();
  formatSheet_(sheet);
  SpreadsheetApp.getUi().alert('シートの書式設定が完了しました。');
}

/**
 * スプレッドシートを開いた時にカスタムメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('確定申告ツール')
    .addItem('シート書式を設定', 'setupSheet')
    .addToUi();
}

/**
 * シートの見た目を整える
 */
function formatSheet_(sheet) {
  const lastCol = SHEET_HEADERS.length; // 15列

  // ── ヘッダー行の書式 ──
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground('#1a56db')       // 濃い青
             .setFontColor('#ffffff')
             .setFontWeight('bold')
             .setFontSize(10)
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle')
             .setWrap(true);
  sheet.setRowHeight(1, 36);

  // ── 列幅の設定 ──
  const colWidths = {
    1: 40,    // A: ID
    2: 140,   // B: タイムスタンプ
    3: 100,   // C: 取引日付
    4: 160,   // D: 店舗名
    5: 100,   // E: 金額
    6: 60,    // F: 税率
    7: 100,   // G: 勘定科目
    8: 160,   // H: 支払方法
    9: 300,   // I: 摘要
    10: 80,   // J: 分類
    11: 80,   // K: ステータス
    12: 200,  // L: 画像URL
    13: 200,  // M: ユーザーメモ
    14: 250,  // N: AI判定理由
    15: 200,  // O: フォローアップ回答
    16: 80    // P: 取引種別
  };
  for (var col in colWidths) {
    sheet.setColumnWidth(Number(col), colWidths[col]);
  }

  // ── データ範囲の書式（1000行分を事前設定）──
  var dataRows = 1000;
  var dataRange = sheet.getRange(2, 1, dataRows, lastCol);
  dataRange.setVerticalAlignment('middle')
           .setFontSize(10);

  // ID列: 中央揃え
  sheet.getRange(2, 1, dataRows, 1).setHorizontalAlignment('center');

  // タイムスタンプ列: 日時書式
  sheet.getRange(2, 2, dataRows, 1).setNumberFormat('yyyy/MM/dd HH:mm');

  // 取引日付列: 日付書式・中央揃え
  sheet.getRange(2, 3, dataRows, 1).setNumberFormat('yyyy/MM/dd')
                                    .setHorizontalAlignment('center');

  // 金額列: カンマ区切り・右揃え
  sheet.getRange(2, 5, dataRows, 1).setNumberFormat('#,##0')
                                    .setHorizontalAlignment('right');

  // 税率・勘定科目・支払方法・分類・ステータス・取引種別: 中央揃え
  [6, 7, 8, 10, 11, 16].forEach(function(c) {
    sheet.getRange(2, c, dataRows, 1).setHorizontalAlignment('center');
  });

  // 摘要列: 折り返し
  sheet.getRange(2, 9, dataRows, 1).setWrap(true);

  // ── 条件付き書式（ステータス列: K列 = 11列目）──
  var rules = sheet.getConditionalFormatRules();

  // 既存の確定申告ツール関連ルールを削除
  rules = rules.filter(function(rule) {
    var ranges = rule.getRanges();
    return !ranges.some(function(r) { return r.getColumn() === 11 || r.getColumn() === 16; });
  });

  var statusRange = sheet.getRange(2, 11, dataRows, 1);

  // 確認済 → 緑背景
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('確認済')
    .setBackground('#d1fae5').setFontColor('#065f46')
    .setRanges([statusRange]).build());

  // 未確認 → 黄背景
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('未確認')
    .setBackground('#fef3c7').setFontColor('#92400e')
    .setRanges([statusRange]).build());

  // 除外 → グレー背景
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('除外')
    .setBackground('#f3f4f6').setFontColor('#6b7280')
    .setRanges([statusRange]).build());

  // 分類列（J列 = 10列目）にも色分け
  var classRange = sheet.getRange(2, 10, dataRows, 1);

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('clear')
    .setBackground('#dbeafe').setFontColor('#1e40af')
    .setRanges([classRange]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('questionable')
    .setBackground('#fef3c7').setFontColor('#92400e')
    .setRanges([classRange]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('not_expense')
    .setBackground('#fee2e2').setFontColor('#991b1b')
    .setRanges([classRange]).build());

  // 取引種別列（P列 = 16列目）の色分け
  var typeRange = sheet.getRange(2, 16, dataRows, 1);

  // 支出 → 赤系
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('支出')
    .setBackground('#fee2e2').setFontColor('#991b1b')
    .setRanges([typeRange]).build());

  // 収入 → 青系
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('収入')
    .setBackground('#dbeafe').setFontColor('#1e40af')
    .setRanges([typeRange]).build());

  sheet.setConditionalFormatRules(rules);

  // ── 交互の背景色（ゼブラストライプ）──
  var bandings = sheet.getBandings();
  bandings.forEach(function(b) { b.remove(); });
  sheet.getRange(1, 1, dataRows + 1, lastCol)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
}

function getReceiptFolder_() {
  if (!CONFIG.RECEIPT_FOLDER_ID) return DriveApp.getRootFolder();
  try {
    return DriveApp.getFolderById(CONFIG.RECEIPT_FOLDER_ID);
  } catch (e) {
    return DriveApp.getRootFolder();
  }
}

// ── Gemini API ──────────────────────────────────
function callGeminiAPI(parts) {
  const payload = {
    contents: [{ parts: parts }],
    generationConfig: {
      response_mime_type: 'application/json',
      temperature: 0.1
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const url = CONFIG.API_ENDPOINT + '?key=' + CONFIG.API_KEY;

  for (let i = 0; i < 3; i++) {
    try {
      const res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200) {
        const json = JSON.parse(res.getContentText());
        let text = json.candidates && json.candidates[0] &&
                   json.candidates[0].content && json.candidates[0].content.parts &&
                   json.candidates[0].content.parts[0] && json.candidates[0].content.parts[0].text;
        if (!text) throw new Error('Geminiからの応答が空です');
        text = text.replace(/^```json\s*/, '').replace(/^```\s*/, '').replace(/\s*```$/, '');
        return JSON.parse(text);
      }
      const errBody = res.getContentText();
      console.warn('Gemini API error (' + res.getResponseCode() + '): ' + errBody);
      Utilities.sleep(1000 * (i + 1));
    } catch (e) {
      console.warn('Retry ' + i + ': ' + e.message);
      if (i === 2) throw e;
      Utilities.sleep(1000 * (i + 1));
    }
  }
}

// ── AIプロンプト ─────────────────────────────────
function getReceiptAnalysisPrompt_(userMemo) {
  let prompt = `あなたは日本の個人事業主（IT/AIツール開発業）の経理アシスタントです。
レシート画像を解析し、以下のJSON形式で情報を抽出してください。

【事業概要】
- 業種: IT/AIツール開発（個人事業主）
- 主な経費: SaaS利用料、開発機器、打ち合わせ費用、書籍、交通費

【勘定科目ルール】
- 通信費: インターネット、サーバー、SaaS（Cursor, GitHub, AWS等）、携帯電話
- 会議費: 1人あたり5,000円以下の飲食を伴う打ち合わせ
- 消耗品費: 10万円未満の備品・機器・文房具
- 新聞図書費: 書籍、電子書籍、技術雑誌、オンライン学習サービス
- 接待交際費: 取引先との会食（1人5,000円超）、贈答品
- 旅費交通費: 電車、タクシー、出張費用
- 車両掛費: ガソリン代、車検費用、自動車保険、駐車場代、高速道路料金、オイル交換、タイヤ交換、車両修理費など車・バイクに関する費用
- 雑費: 上記に該当しないもの

【税率判定ルール】
- 8%: 飲食料品（軽減税率対象）、テイクアウト
- 10%: 一般的な商品・サービス
- 非課税: 切手、印紙、保険料

【分類ルール】
- clear: 明確に事業経費と判断できる
- questionable: 事業用途か不明確（確認が必要）
- not_expense: 経費にならない可能性が高い

【摘要の書式】
- 形式: [カテゴリ] 説明（最大64全角文字）
- 例: [開発環境] Cursor Pro版 月額利用料（自社開発用）
- 例: [打合せ] ○○氏とカフェで開発方針打合せ`;

  if (userMemo) {
    prompt += '\n\n【ユーザーメモ】\n' + userMemo;
  }

  prompt += `

【出力JSON形式】
{
  "storeName": "店舗名またはサービス名",
  "date": "YYYY/MM/DD",
  "amount": 1500,
  "taxRate": "10%",
  "classification": "clear",
  "accountTitle": "通信費",
  "description": "[カテゴリ] 摘要テキスト",
  "followUpQuestions": [],
  "confidence": 0.95,
  "reasoning": "判定理由の説明"
}

【followUpQuestions について】
classificationが"questionable"の場合、用途を明確にするための質問を生成してください。
例:
- SaaS/APIレシート → "このサービスは特定案件用ですか？自社環境用ですか？"
- レストラン/カフェ → "どなたとの打合せでしたか？議題は何でしたか？"
- 高額機器 → "どの開発・テストに使用しますか？"

各質問は文字列の配列として返してください。`;

  return prompt;
}

function getTransferSlipPrompt_(userMemo, transferPurpose) {
  let prompt = `あなたは日本の個人事業主（IT/AIツール開発業）の経理アシスタントです。
振込用紙（振込控え・払込受領証）の画像を解析し、以下のJSON形式で情報を抽出してください。

【事業概要】
- 業種: IT/AIツール開発（個人事業主）

【振込用紙から読み取る情報】
- 振込先名（受取人）→ storeNameに設定
- 振込日付 → dateに設定
- 振込金額（手数料を含まない本体金額）→ amountに設定
- 振込手数料（記載がある場合）→ transferFeeに設定（なければ0）
- 用途・通信欄の内容

【勘定科目ルール（振込用）】
- 地代家賃: 事務所家賃、駐車場代、倉庫賃借料
- 水道光熱費: 電気料金、ガス料金、水道料金
- 通信費: インターネット、固定電話、携帯電話、サーバー、クラウド、SaaS
- 外注費: 業務委託費、外注加工費、デザイン外注
- 広告宣伝費: 広告掲載料、Web広告、販促費
- 消耗品費: 10万円未満の備品・機器・事務用品
- 新聞図書費: 書籍、電子書籍、技術雑誌、オンライン学習
- 研修費: セミナー参加費、研修受講料
- 諸会費: 協会年会費、組合費、商工会議所会費
- 保険料: 火災保険、賠償責任保険、事業用損害保険
- 租税公課: 事業税、消費税、固定資産税、印紙税、自動車税
- 車両掛費: 車検費用、自動車保険、駐車場代、高速道路料金
- 支払手数料: 振込手数料（transferFeeに該当）
- 事業主貸: 所得税、住民税、国民健康保険料、国民年金保険料（経費ではなく事業主個人の支出）
- 雑費: 上記に該当しないもの

【税率判定ルール】
- 10%: 一般的な商品・サービス、家賃
- 非課税: 保険料、振込手数料、税金、国民健康保険、国民年金

【分類ルール】
- clear: 明確に事業経費と判断できる
- questionable: 事業用途か不明確（確認が必要）
- not_expense: 経費にならない可能性が高い（個人的な支払いなど）

【摘要の書式】
- 形式: [振込] 〇〇へ△△の支払い（最大64全角文字）
- 例: [振込] NTTファイナンスへ通信費の支払い
- 例: [振込] AWS Japanへサーバー利用料の支払い
- 例: [振込] ○○市役所へ固定資産税の支払い`;

  if (transferPurpose) {
    prompt += '\n\n【ユーザーが選択した振込用途】\n' + transferPurpose + '\nこの用途を参考に勘定科目を判定してください。';
  }

  if (userMemo) {
    prompt += '\n\n【ユーザーメモ】\n' + userMemo;
  }

  prompt += `

【出力JSON形式】
{
  "storeName": "振込先名（受取人）",
  "date": "YYYY/MM/DD",
  "amount": 振込金額（手数料を含まない本体金額の数値）,
  "transferFee": 振込手数料（数値、記載なしは0）,
  "taxRate": "10%",
  "classification": "clear",
  "accountTitle": "勘定科目",
  "description": "[振込] 摘要テキスト",
  "followUpQuestions": [],
  "confidence": 0.95,
  "reasoning": "判定理由の説明"
}

【重要】
- amountには振込手数料を含めないでください。手数料はtransferFeeに分けて記載してください。
- transferFeeが0より大きい場合、振込手数料は別途「支払手数料」として計上可能です。

【followUpQuestions について】
classificationが"questionable"の場合に質問を生成してください。
例:
- 用途不明 → "この振込は何の支払いですか？（サーバー費、外注費など）"
- 個人/事業不明 → "この振込は事業用の支払いですか？"

各質問は文字列の配列として返してください。`;

  return prompt;
}

function getBankStatementPrompt_(userMemo) {
  var currentYear = new Date().getFullYear();
  var prompt = `あなたは日本の個人事業主（IT/AIツール開発業）の経理アシスタントです。
通帳ページまたはネットバンキングの明細画像を解析し、**入金取引のみ**を抽出してください。
出金取引は無視してください。

【事業概要】
- 業種: IT/AIツール開発（個人事業主）
- 主な収入: 開発案件の報酬、ツール利用料、コンサルティング費

【勘定科目ルール（収入）】
- 売上高: 本業（IT/AIツール開発）に関連する入金、開発報酬、案件報酬
- 雑収入: 本業以外の収入（利息、還付金、その他の入金）

【税率判定ルール】
- 10%: 一般的なサービス提供（売上高）
- 非課税: 利息、保険金、還付金

【抽出ルール】
- 入金（振込入金・入金・利息等）のみ抽出する
- 出金・引落・振替は無視する
- 同じ画像に複数の入金取引がある場合はすべて抽出する
- 年が記載されていない場合は ` + currentYear + ` 年として扱う
- 日付は必ず YYYY/MM/DD 形式で返す`;

  if (userMemo) {
    prompt += '\n\n【ユーザーメモ】\n' + userMemo;
  }

  prompt += `

【出力JSON形式】
必ず配列で返してください（1件でも配列）。入金取引が0件の場合は空配列 [] を返してください。
[
  {
    "date": "YYYY/MM/DD",
    "storeName": "振込元名（振込人名義）",
    "amount": 100000,
    "taxRate": "10%",
    "accountTitle": "売上高",
    "description": "[入金] 摘要テキスト（最大64全角文字）",
    "confidence": 0.9,
    "reasoning": "判定理由の説明"
  }
]

【重要】
- 必ずJSON配列として返してください
- 入金が見つからない場合は空配列 [] を返してください`;

  return prompt;
}

function getTextExpensePrompt_(userMemo, date, amount) {
  let prompt = `あなたは日本の個人事業主（IT/AIツール開発業）の経理アシスタントです。
以下のテキスト情報から経費の勘定科目を判定してください。

【事業概要】
- 業種: IT/AIツール開発（個人事業主）

【勘定科目ルール】
- 通信費: インターネット、サーバー、SaaS（Cursor, GitHub, AWS等）、携帯電話
- 会議費: 1人あたり5,000円以下の飲食を伴う打ち合わせ
- 消耗品費: 10万円未満の備品・機器・文房具
- 新聞図書費: 書籍、電子書籍、技術雑誌、オンライン学習サービス
- 接待交際費: 取引先との会食（1人5,000円超）、贈答品
- 旅費交通費: 電車、タクシー、出張費用
- 車両掛費: ガソリン代、車検費用、自動車保険、駐車場代、高速道路料金、オイル交換、タイヤ交換、車両修理費など車・バイクに関する費用
- 雑費: 上記に該当しないもの

【税率判定ルール】
- 8%: 飲食料品（軽減税率対象）
- 10%: 一般的な商品・サービス
- 非課税: 切手、印紙、保険料

【分類ルール】
- clear: 明確に事業経費と判断できる
- questionable: 事業用途か不明確
- not_expense: 経費にならない可能性が高い

【入力情報】
メモ: ` + userMemo;

  if (date) prompt += '\n日付: ' + date;
  if (amount) prompt += '\n金額: ' + amount + '円';

  prompt += `

上記の情報を元に、店舗名・金額・日付をテキストから推測し、以下のJSON形式で回答してください。
金額が入力されていればそれを使い、テキストに金額が含まれていればそこから抽出してください。
日付が入力されていればそれを使ってください。

【摘要の書式】
- 形式: [カテゴリ] 説明（最大64全角文字）

{
  "storeName": "店舗名またはサービス名",
  "date": "YYYY/MM/DD",
  "amount": 数値,
  "taxRate": "10%",
  "classification": "clear",
  "accountTitle": "勘定科目",
  "description": "[カテゴリ] 摘要テキスト",
  "followUpQuestions": [],
  "confidence": 0.8,
  "reasoning": "判定理由"
}`;

  return prompt;
}

function getFollowUpPrompt_(originalResult, answers) {
  return `あなたは日本の個人事業主（IT/AIツール開発業）の経理アシスタントです。
以下のレシート解析の初回結果に対して、ユーザーが追加情報を提供しました。
この情報を踏まえて、勘定科目と摘要を再判定してください。

【初回AI解析結果】
` + JSON.stringify(originalResult, null, 2) + `

【ユーザー回答】
` + answers + `

【勘定科目ルール】（同上）
- 通信費 / 会議費 / 消耗品費 / 新聞図書費 / 接待交際費 / 旅費交通費 / 車両掛費 / 雑費

回答の内容を摘要に反映し（相手先名・用途など）、再判定結果を以下のJSON形式で返してください。
摘要は最大64全角文字です。

{
  "storeName": "店舗名",
  "date": "YYYY/MM/DD",
  "amount": 数値,
  "taxRate": "10%",
  "classification": "clear",
  "accountTitle": "勘定科目",
  "description": "[カテゴリ] 更新された摘要",
  "followUpQuestions": [],
  "confidence": 0.95,
  "reasoning": "再判定理由"
}`;
}

// ── レシート解析（画像あり）─────────────────────
function analyzeReceipt(formData) {
  const folder = getReceiptFolder_();
  const timestamp = new Date();
  let imageUrl = '';

  // 画像をDriveに保存
  if (formData.imageData) {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(formData.imageData),
      formData.mimeType || 'image/jpeg',
      'RECEIPT_' + timestamp.getTime() + '.jpg'
    );
    const file = folder.createFile(blob);
    imageUrl = file.getUrl();
  }

  // Geminiへリクエスト
  const parts = [];
  parts.push({ text: getReceiptAnalysisPrompt_(formData.userMemo || '') });
  if (formData.imageData) {
    parts.push({
      inline_data: {
        mime_type: formData.mimeType || 'image/jpeg',
        data: formData.imageData
      }
    });
  }

  const result = callGeminiAPI(parts);
  result.imageUrl = imageUrl;
  return JSON.stringify(result);
}

// ── 振込用紙解析（画像あり）──────────────────────
function analyzeTransferSlip(formData) {
  const folder = getReceiptFolder_();
  const timestamp = new Date();
  let imageUrl = '';

  // 画像をDriveに保存
  if (formData.imageData) {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(formData.imageData),
      formData.mimeType || 'image/jpeg',
      'TRANSFER_' + timestamp.getTime() + '.jpg'
    );
    const file = folder.createFile(blob);
    imageUrl = file.getUrl();
  }

  // Geminiへリクエスト
  const parts = [];
  parts.push({ text: getTransferSlipPrompt_(formData.userMemo || '', formData.transferPurpose || '') });
  if (formData.imageData) {
    parts.push({
      inline_data: {
        mime_type: formData.mimeType || 'image/jpeg',
        data: formData.imageData
      }
    });
  }

  const result = callGeminiAPI(parts);
  result.imageUrl = imageUrl;
  return JSON.stringify(result);
}

// ── 通帳/明細解析（画像あり）──────────────────────
function analyzeBankStatement(formData) {
  const folder = getReceiptFolder_();
  const timestamp = new Date();
  let imageUrl = '';

  // 画像をDriveに保存
  if (formData.imageData) {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(formData.imageData),
      formData.mimeType || 'image/jpeg',
      'BANK_' + timestamp.getTime() + '.jpg'
    );
    const file = folder.createFile(blob);
    imageUrl = file.getUrl();
  }

  // Geminiへリクエスト
  const parts = [];
  parts.push({ text: getBankStatementPrompt_(formData.userMemo || '') });
  if (formData.imageData) {
    parts.push({
      inline_data: {
        mime_type: formData.mimeType || 'image/jpeg',
        data: formData.imageData
      }
    });
  }

  const result = callGeminiAPI(parts);
  // 配列で正規化
  var transactions = Array.isArray(result) ? result : [result];
  return JSON.stringify({ transactions: transactions, imageUrl: imageUrl });
}

// ── テキスト経費解析（画像なし）──────────────────
function analyzeTextExpense(formData) {
  const prompt = getTextExpensePrompt_(
    formData.userMemo,
    formData.date || '',
    formData.amount || ''
  );

  const parts = [{ text: prompt }];
  const result = callGeminiAPI(parts);
  result.imageUrl = '';
  return JSON.stringify(result);
}

// ── フォローアップ再判定 ─────────────────────────
function analyzeFollowUp(formData) {
  const prompt = getFollowUpPrompt_(formData.originalResult, formData.answers);
  const parts = [{ text: prompt }];
  const result = callGeminiAPI(parts);
  result.imageUrl = formData.originalResult.imageUrl || '';
  return JSON.stringify(result);
}

// ── CRUD関数 ─────────────────────────────────────
function saveExpense(data) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  try {
    lockAcquired = lock.tryLock(10000);
    if (!lockAcquired) throw new Error('排他制御タイムアウト');

    const sheet = getOrCreateSheet_();
    const lastRow = sheet.getLastRow();
    const newId = lastRow < 1 ? 1 : lastRow; // ヘッダー行を考慮

    const now = new Date();
    sheet.appendRow([
      newId,
      now,
      data.date || '',
      data.storeName || '',
      parseCurrency(data.amount),
      data.taxRate || '10%',
      data.accountTitle || '',
      data.paymentMethod || '',
      data.description || '',
      data.classification || 'clear',
      data.status || '確認済',
      data.imageUrl || '',
      data.userMemo || '',
      data.reasoning || '',
      data.followUpAnswers || '',
      data.transactionType || '支出'
    ]);

    // キャッシュ無効化
    try { CacheService.getScriptCache().remove('expenses_data'); } catch(e) {}

    return JSON.stringify({ success: true, id: newId });
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

function saveBatchIncome(dataArray) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  try {
    lockAcquired = lock.tryLock(30000);
    if (!lockAcquired) throw new Error('排他制御タイムアウト');

    const sheet = getOrCreateSheet_();
    var savedCount = 0;

    for (var i = 0; i < dataArray.length; i++) {
      var data = dataArray[i];
      var lastRow = sheet.getLastRow();
      var newId = lastRow < 1 ? 1 : lastRow;
      var now = new Date();

      sheet.appendRow([
        newId,
        now,
        data.date || '',
        data.storeName || '',
        parseCurrency(data.amount),
        data.taxRate || '10%',
        data.accountTitle || '売上高',
        '普通預金',
        data.description || '',
        'clear',
        data.status || '確認済',
        data.imageUrl || '',
        data.userMemo || '',
        data.reasoning || '',
        '',
        '収入'
      ]);
      savedCount++;
    }

    // キャッシュ無効化
    try { CacheService.getScriptCache().remove('expenses_data'); } catch(e) {}

    return JSON.stringify({ success: true, count: savedCount });
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

function getExpenses(filters) {
  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return JSON.stringify([]);

  var colCount = Math.max(SHEET_HEADERS.length, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, lastRow - 1, colCount).getValues();
  let expenses = data.map(function(row, idx) {
    return {
      rowIndex: idx + 2,
      id: row[0],
      timestamp: row[1],
      date: formatDate(row[2]),
      storeName: row[3],
      amount: row[4],
      taxRate: row[5],
      accountTitle: row[6],
      paymentMethod: row[7],
      description: row[8],
      classification: row[9],
      status: row[10],
      imageUrl: row[11],
      userMemo: row[12],
      reasoning: row[13],
      followUpAnswers: row[14],
      transactionType: row[15] || '支出'
    };
  });

  // フィルタ適用
  if (filters) {
    if (filters.year && filters.month) {
      expenses = expenses.filter(function(e) {
        if (!e.date) return false;
        var parts = e.date.split('/');
        return parts[0] === String(filters.year) && parseInt(parts[1]) === parseInt(filters.month);
      });
    } else if (filters.year) {
      expenses = expenses.filter(function(e) {
        if (!e.date) return false;
        return e.date.split('/')[0] === String(filters.year);
      });
    }
    if (filters.status) {
      expenses = expenses.filter(function(e) { return e.status === filters.status; });
    }
    if (filters.transactionType) {
      expenses = expenses.filter(function(e) { return e.transactionType === filters.transactionType; });
    }
  }

  // 日付降順ソート
  expenses.sort(function(a, b) {
    return (b.date || '').localeCompare(a.date || '');
  });

  return JSON.stringify(expenses);
}

function updateExpense(rowIndex, data) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  try {
    lockAcquired = lock.tryLock(10000);
    if (!lockAcquired) throw new Error('排他制御タイムアウト');

    const sheet = getOrCreateSheet_();
    const row = parseInt(rowIndex);
    if (row < 2) throw new Error('無効な行番号');

    // 更新可能フィールド
    if (data.date !== undefined) sheet.getRange(row, 3).setValue(data.date);
    if (data.storeName !== undefined) sheet.getRange(row, 4).setValue(data.storeName);
    if (data.amount !== undefined) sheet.getRange(row, 5).setValue(parseCurrency(data.amount));
    if (data.taxRate !== undefined) sheet.getRange(row, 6).setValue(data.taxRate);
    if (data.accountTitle !== undefined) sheet.getRange(row, 7).setValue(data.accountTitle);
    if (data.paymentMethod !== undefined) sheet.getRange(row, 8).setValue(data.paymentMethod);
    if (data.description !== undefined) sheet.getRange(row, 9).setValue(data.description);
    if (data.classification !== undefined) sheet.getRange(row, 10).setValue(data.classification);
    if (data.status !== undefined) sheet.getRange(row, 11).setValue(data.status);
    if (data.transactionType !== undefined) sheet.getRange(row, 16).setValue(data.transactionType);

    try { CacheService.getScriptCache().remove('expenses_data'); } catch(e) {}
    return JSON.stringify({ success: true });
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

function deleteExpense(rowIndex) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  try {
    lockAcquired = lock.tryLock(10000);
    if (!lockAcquired) throw new Error('排他制御タイムアウト');

    const sheet = getOrCreateSheet_();
    const row = parseInt(rowIndex);
    if (row < 2) throw new Error('無効な行番号');

    sheet.deleteRow(row);
    try { CacheService.getScriptCache().remove('expenses_data'); } catch(e) {}
    return JSON.stringify({ success: true });
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

// ── 弥生会計CSV出力 ─────────────────────────────
function exportYayoiCsv(filters) {
  const expenses = JSON.parse(getExpenses(filters));
  // 確認済のみ、除外を除く
  const targets = expenses.filter(function(e) {
    return e.status === '確認済';
  });

  if (targets.length === 0) {
    return JSON.stringify({ success: false, message: '出力対象のデータがありません（ステータス「確認済」のデータが必要です）' });
  }

  var csvRows = [];
  targets.forEach(function(e) {
    csvRows.push(buildYayoiRow_(e));
  });

  var csv = csvRows.map(function(row) {
    return row.map(function(cell) {
      var s = String(cell);
      if (s.indexOf(',') >= 0 || s.indexOf('"') >= 0 || s.indexOf('\n') >= 0) {
        return '"' + s.replace(/"/g, '""') + '"';
      }
      return s;
    }).join(',');
  }).join('\r\n');

  var expenseTargets = targets.filter(function(e) { return e.transactionType !== '収入'; });
  var incomeTargets = targets.filter(function(e) { return e.transactionType === '収入'; });

  return JSON.stringify({
    success: true,
    csv: csv,
    count: targets.length,
    totalAmount: targets.reduce(function(sum, e) { return sum + (Number(e.amount) || 0); }, 0),
    expenseCount: expenseTargets.length,
    expenseTotal: expenseTargets.reduce(function(sum, e) { return sum + (Number(e.amount) || 0); }, 0),
    incomeCount: incomeTargets.length,
    incomeTotal: incomeTargets.reduce(function(sum, e) { return sum + (Number(e.amount) || 0); }, 0)
  });
}

function buildYayoiRow_(expense) {
  // 弥生会計25列フォーマット
  var row = new Array(25).fill('');
  var isIncome = expense.transactionType === '収入';
  var amount = String(Math.floor(Number(expense.amount) || 0));

  // 1. 識別フラグ
  row[0] = '2000';

  // 4. 取引日付 (index 3)
  row[3] = expense.date || '';

  if (isIncome) {
    // 収入: Dr. 普通預金 / Cr. 売上高等
    row[4] = getCreditAccountForIncome_(expense.paymentMethod);  // 借方: 普通預金
    row[7] = '対象外';                                            // 借方税区分
    row[8] = amount;                                              // 借方金額
    row[10] = expense.accountTitle || '売上高';                   // 貸方: 売上高等
    row[13] = getIncomeTaxCategory_(expense.taxRate);             // 貸方税区分
    row[14] = amount;                                             // 貸方金額
  } else {
    // 支出: Dr. 勘定科目 / Cr. 支払方法→貸方科目
    row[4] = expense.accountTitle || '';                           // 借方: 勘定科目
    row[7] = getTaxCategory_(expense.taxRate);                    // 借方税区分
    row[8] = amount;                                              // 借方金額
    row[10] = getCreditAccount_(expense.paymentMethod);           // 貸方: 支払方法
    row[13] = '対象外';                                           // 貸方税区分
    row[14] = amount;                                             // 貸方金額
  }

  // 17. 摘要 (index 16)
  row[16] = expense.description || '';

  // 20. タイプ (index 19)
  row[19] = '0';

  // 23. 付箋1 (index 22)
  row[22] = '0';

  // 24. 付箋2 (index 23)
  row[23] = '0';

  // 25. 調整 (index 24)
  row[24] = '0';

  return row;
}

function getTaxCategory_(taxRate) {
  var rate = String(taxRate || '');
  if (rate === '8%' || rate === '8') return '課対仕入8%（軽）';
  if (rate === '10%' || rate === '10') return '課対仕入10%';
  if (rate === '非課税') return '対象外';
  return '課対仕入10%';
}

function getIncomeTaxCategory_(taxRate) {
  var rate = String(taxRate || '');
  if (rate === '10%' || rate === '10') return '課税売上10%';
  if (rate === '8%' || rate === '8') return '課税売上8%';
  if (rate === '非課税') return '対象外';
  return '課税売上10%';
}

function getCreditAccountForIncome_(paymentMethod) {
  var method = String(paymentMethod || '');
  if (method === '当座預金') return '当座預金';
  return '普通預金';
}

function getCreditAccount_(paymentMethod) {
  var method = String(paymentMethod || '');
  if (method === '現金') return '現金';
  if (method === '事業用クレジットカード') return '未払金';
  if (method === '個人立替') return '事業主借';
  return '事業主借';
}

// ── 重複チェック ─────────────────────────────────
function checkDuplicate(data) {
  var sheet = getOrCreateSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return JSON.stringify({ isDuplicate: false });

  var colCount = Math.max(SHEET_HEADERS.length, sheet.getLastColumn());
  var allData = sheet.getRange(2, 1, lastRow - 1, colCount).getValues();
  var targetDate = String(data.date || '');
  var targetAmount = parseCurrency(data.amount);
  var targetStore = String(data.storeName || '');
  var targetType = data.transactionType || '支出';

  for (var i = 0; i < allData.length; i++) {
    var row = allData[i];
    var rowDate = formatDate(row[2]);
    var rowAmount = Number(row[4]) || 0;
    var rowStore = String(row[3] || '');
    var rowType = row[15] || '支出';

    if (rowDate === targetDate && rowAmount === targetAmount && rowStore === targetStore && rowType === targetType) {
      return JSON.stringify({
        isDuplicate: true,
        existingId: row[0],
        message: '同じ日付・金額・店舗名のデータが既に登録されています（ID: ' + row[0] + '）'
      });
    }
  }
  return JSON.stringify({ isDuplicate: false });
}
