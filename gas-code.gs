// ============================================================
//  聴読解練習テスト — Google Apps Script
//  Google Spreadsheet に紐づけてデプロイしてください
// ============================================================

// ---- 正解データ ----
const CORRECT_ANSWERS = {
  s1: [4, 3, 1, 2, 3],
  s2: [3, 3, 1, 4, 2, 1, 4, 3, 3, 4],
  s3: [3, 1, 4, 3, 2, 4, 1, 3, 3, 3]
};

// ---- 色設定 ----
const COLOR_CORRECT = "#d9ead3";  // 薄緑（正解）
const COLOR_WRONG   = "#fce8e6";  // 薄赤（不正解）
const COLOR_HEADER  = "#6b6ef9";  // 紫（ヘッダー）
const COLOR_SCORE   = "#e8eaff";  // 薄紫（得点列）

// ============================================================
//  POST リクエストを受け取ってシートに書き込む
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const s1 = data.answers[0] || [];
    const s2 = data.answers[1] || [];
    const s3 = data.answers[2] || [];
    const allAnswers = [s1, s2, s3];
    const correctKeys = [CORRECT_ANSWERS.s1, CORRECT_ANSWERS.s2, CORRECT_ANSWERS.s3];

    // スコア計算
    const scores = allAnswers.map((ans, si) =>
      correctKeys[si].reduce((sum, c, qi) => sum + (ans[qi] === c ? 1 : 0), 0)
    );
    const total = scores.reduce((a, b) => a + b, 0);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // ヘッダーがなければ自動追加
    if (sheet.getLastRow() === 0) {
      setupHeaders(sheet);
    }

    // ---- 行データを組み立て ----
    // 各問: "回答 ○" or "回答 ✗"  の形式
    const buildCells = (ans, correct) =>
      correct.map((c, i) => {
        const a = ans[i];
        if (a === undefined || a === null) return "－";
        return a + (a === c ? " ○" : " ✗");
      });

    const s1Cells = buildCells(s1, CORRECT_ANSWERS.s1);
    const s2Cells = buildCells(s2, CORRECT_ANSWERS.s2);
    const s3Cells = buildCells(s3, CORRECT_ANSWERS.s3);

    const row = [
      new Date(),          // 提出日時
      data.studentId,      // 学号
      data.name,           // 姓名
      ...s1Cells,          // S1-Q1〜Q5 (5列)
      `${scores[0]}/5`,    // S1得点
      ...s2Cells,          // S2-Q1〜Q10 (10列)
      `${scores[1]}/10`,   // S2得点
      ...s3Cells,          // S3-Q1〜Q10 (10列)
      `${scores[2]}/10`,   // S3得点
      `${total}/25`        // 合計得点
    ];

    sheet.appendRow(row);

    // ---- セルの色付け ----
    const lastRow = sheet.getLastRow();
    colorizeRow(sheet, lastRow, allAnswers, correctKeys, scores, total);

    return buildResponse({
      status: "success",
      score: { s1: scores[0], s2: scores[1], s3: scores[2], total: total }
    });

  } catch (err) {
    return buildResponse({ status: "error", message: err.toString() });
  }
}

// ============================================================
//  色付け処理
// ============================================================
function colorizeRow(sheet, row, allAnswers, correctKeys, scores, total) {
  // 列インデックス（1始まり）
  // A=1:提出日時, B=2:学号, C=3:姓名
  // D〜H (4〜8): S1-Q1〜Q5, I(9): S1得点
  // J〜S (10〜19): S2-Q1〜Q10, T(20): S2得点
  // U〜AD (21〜30): S3-Q1〜Q10, AE(31): S3得点
  // AF(32): 合計得点

  const sections = [
    { startCol: 4,  answers: allAnswers[0], correct: correctKeys[0], scoreCol: 9  },
    { startCol: 10, answers: allAnswers[1], correct: correctKeys[1], scoreCol: 20 },
    { startCol: 21, answers: allAnswers[2], correct: correctKeys[2], scoreCol: 31 }
  ];

  sections.forEach(sec => {
    sec.answers.forEach((ans, qi) => {
      const col = sec.startCol + qi;
      const isOk = ans === sec.correct[qi];
      const cell = sheet.getRange(row, col);
      cell.setBackground(isOk ? COLOR_CORRECT : COLOR_WRONG);
      cell.setHorizontalAlignment("center");
    });
    // 得点列
    sheet.getRange(row, sec.scoreCol)
      .setBackground(COLOR_SCORE)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
  });

  // 合計得点列
  sheet.getRange(row, 32)
    .setBackground("#d0d3fc")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // 提出日時を読みやすく
  sheet.getRange(row, 1)
    .setNumberFormat("yyyy/MM/dd HH:mm:ss");
}

// ============================================================
//  ヘッダー行を設定する
// ============================================================
function setupHeaders(sheet) {
  const headers = [
    "提出日時", "学号", "姓名",
    "S1-Q1", "S1-Q2", "S1-Q3", "S1-Q4", "S1-Q5", "S1得点",
    "S2-Q1", "S2-Q2", "S2-Q3", "S2-Q4", "S2-Q5",
    "S2-Q6", "S2-Q7", "S2-Q8", "S2-Q9", "S2-Q10", "S2得点",
    "S3-Q1", "S3-Q2", "S3-Q3", "S3-Q4", "S3-Q5",
    "S3-Q6", "S3-Q7", "S3-Q8", "S3-Q9", "S3-Q10", "S3得点",
    "合計得点"
  ];

  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight("bold");
  range.setBackground(COLOR_HEADER);
  range.setFontColor("#ffffff");
  range.setHorizontalAlignment("center");

  // 列幅
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 90);
  for (let i = 4; i <= headers.length; i++) {
    sheet.setColumnWidth(i, 68);
  }

  // ヘッダー行を固定
  sheet.setFrozenRows(1);
}

// ---- GET: 動作確認用 ----
function doGet(e) {
  return ContentService.createTextOutput("✅ GAS is running!");
}

// ---- JSON レスポンス ----
function buildResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
