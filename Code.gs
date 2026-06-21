// =====================================================================
// 見積書自動生成システム
// 目的: 手作業による転記ミス・作業時間(数時間)を削減し、
//       ボタン一つで統一フォーマットの見積書PDFを自動生成する
// =====================================================================

// ----- シート名・フォルダ名 -----
const SHEET_CUSTOMERS  = 'Customers';
const SHEET_ATTENDANCE = 'Attendance';
const SHEET_QUOTE      = 'QuoteFormat';
const PDF_FOLDER_NAME  = '見積書';

// ----- 自社情報(変更時はここだけ修正) -----
const COMPANY_NAME     = '株式会社〇〇〇〇';
const COMPANY_ZIP      = '〒000-0000';
const COMPANY_ADDRESS  = '都道府県市区町村番地';
const COMPANY_BUILDING = 'ビル名 階数';

// ----- QuoteFormatシートの固定行番号(変更時はここだけ修正) -----
const ROW_HEADER       = 1;  // 見積番号・発行日・自社情報
const ROW_TITLE        = 5;  // 「見積書」タイトル
const ROW_CUSTOMER     = 7;  // 顧客名・件名
const ROW_DETAIL_START = 10; // 明細テーブル開始行


// =====================================================================
// カスタムメニューをスプレッドシートに追加する
// =====================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('見積書作成')
    .addItem('全顧客分を生成', 'generateAllQuotes')
    .addToUi();
}


// =====================================================================
// メイン処理: 全顧客分の見積書を一括生成してPDF保存する
// =====================================================================
function generateAllQuotes() {
  const ui = SpreadsheetApp.getUi();

  // 対象月をダイアログで取得
  const response = ui.prompt('対象月を入力してください（例: 2026-04）');
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const targetMonth = response.getResponseText().trim();

  if (!targetMonth) {
    ui.alert('対象月が入力されていません');
    return;
  }

  // QuoteFormatシートの存在確認
  const quoteSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUOTE);
  if (!quoteSheet) {
    ui.alert(`「${SHEET_QUOTE}」シートが見つかりません。シート名を確認してください。`);
    return;
  }

  const customers     = getCustomers();
  const allAttendance = getAllAttendance(); // Attendanceを1回だけ読み込む
  const pdfFolder     = getOrCreatePdfFolder();

  let generatedCount = 0;

  for (const customer of customers) {

    // 必須項目チェック: 空欄があれば即中断
    if (!validateRequiredFields(customer)) {
      ui.alert('必須項目が入力されていません');
      return;
    }

    // 対象月の出勤実績がない顧客はスキップ(メモリ上でフィルタ)
    const attendance = filterAttendance(allAttendance, customer.customer_code, targetMonth);
    if (attendance.length === 0) continue;

    // QuoteFormatシートにレイアウトごと書き込む
    fillQuoteFormat(quoteSheet, customer, attendance, targetMonth);

    // 書き込みをシートに確定させてからPDF出力する
    SpreadsheetApp.flush();

    // QuoteFormatシートをPDF化してDriveに保存
    const fileName = `見積書_${customer.customer_name}_${targetMonth}`;
    exportSheetAsPdf(quoteSheet, fileName, pdfFolder);

    // 連続リクエストによるレート制限(429)を避けるため待機
    // 429発生時はリトライが自動対処するため500msで運用
    Utilities.sleep(500);

    generatedCount++;
  }

  ui.alert(`${generatedCount}件の見積書を生成しました`);
}


// =====================================================================
// Customersシートを全件取得してオブジェクト配列で返す
// 列の順番に依存せず、1行目のヘッダー名でマッピングする
// =====================================================================
function getCustomers() {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CUSTOMERS);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];

  // ヘッダー行をキーにして各行をオブジェクト化
  return data.slice(1).map(row => {
    const record = {};
    headers.forEach((header, i) => { record[header] = row[i]; });
    return record;
  });
}


// =====================================================================
// Attendanceシートを全件取得してオブジェクト配列で返す(1回だけ呼ぶ)
// =====================================================================
function getAllAttendance() {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ATTENDANCE);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];

  return data.slice(1).map(row => {
    const record = {};
    headers.forEach((header, i) => { record[header] = row[i]; });
    return record;
  });
}


// =====================================================================
// 読み込み済みのAttendance配列を顧客コード・対象月でフィルタして返す
// work_month がDateオブジェクトとして格納されている場合も考慮する
// =====================================================================
function filterAttendance(allAttendance, customerCode, targetMonth) {
  return allAttendance.filter(row =>
    row.customer_code === customerCode &&
    formatMonth(row.work_month) === targetMonth
  );
}


// =====================================================================
// 見積書に必須の項目が入力されているか確認する
// 空欄の場合は false を返して呼び出し元で処理を中断させる
// =====================================================================
function validateRequiredFields(customer) {
  const hasName  = customer.customer_name !== '' && customer.customer_name != null;
  const hasPrice = customer.unit_price    !== '' && customer.unit_price    != null;
  return hasName && hasPrice;
}


// =====================================================================
// QuoteFormatシートをクリアしてレイアウト・スタイルごと構築する
// 出勤実績のスタッフ数に応じて行数が変わるため毎回クリアしてから描画する
// =====================================================================
function fillQuoteFormat(sheet, customer, attendanceRows, targetMonth) {
  sheet.clearContents();
  sheet.clearFormats();

  const today      = new Date();
  const issueDate  = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy年M月d日');
  const quoteNo    = `Q-${customer.customer_code}-${targetMonth.replace('-', '')}`;
  const [year, month] = targetMonth.split('-');
  const monthLabel = `${year}年${Number(month)}月`;

  // ----- 列幅・行高の設定 -----
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 200);
  sheet.setRowHeight(ROW_TITLE, 55);

  // ----- ヘッダー: 見積番号・発行日 -----
  sheet.getRange(ROW_HEADER,     1).setValue(`見積番号: ${quoteNo}`);
  sheet.getRange(ROW_HEADER + 1, 1).setValue(`発行日: ${issueDate}`);
  sheet.getRange(ROW_HEADER, 1, 2, 1)
    .setFontSize(9)
    .setFontColor('#666666');

  // ----- 自社情報（右上） -----
  sheet.getRange(ROW_HEADER,     4).setValue(COMPANY_NAME).setFontWeight('bold').setFontSize(12);
  sheet.getRange(ROW_HEADER + 1, 4).setValue(COMPANY_ZIP);
  sheet.getRange(ROW_HEADER + 2, 4).setValue(COMPANY_ADDRESS);
  sheet.getRange(ROW_HEADER + 3, 4).setValue(COMPANY_BUILDING);
  sheet.getRange(ROW_HEADER, 4, 4, 1)
    .setHorizontalAlignment('right')
    .setFontSize(10)
    .setFontColor('#1a1a2e');

  // ----- タイトル -----
  sheet.getRange(ROW_TITLE, 1, 1, 4).merge();
  sheet.getRange(ROW_TITLE, 1)
    .setValue('見　積　書')
    .setFontSize(22)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontColor('#1a3c5e');

  // ----- 顧客情報 -----
  sheet.getRange(ROW_CUSTOMER, 1)
    .setValue(`${customer.customer_name} 様`)
    .setFontSize(13)
    .setFontWeight('bold');
  sheet.getRange(ROW_CUSTOMER + 1, 1)
    .setValue(`ご担当: ${customer.contact_person} 様`)
    .setFontSize(10);
  sheet.getRange(ROW_CUSTOMER, 3)
    .setValue(`件名: 業務委託費(${monthLabel}分)`)
    .setFontSize(10);
  sheet.getRange(ROW_CUSTOMER + 1, 3)
    .setValue('見積有効期限: 発行日より30日間')
    .setFontSize(9)
    .setFontColor('#666666');

  // 顧客情報下の区切り線
  sheet.getRange(ROW_CUSTOMER + 1, 1, 1, 4)
    .setBorder(false, false, true, false, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

  // ----- 明細テーブル -----
  const tableHeaderStyle = sheet.getRange(ROW_DETAIL_START, 1, 1, 3);
  tableHeaderStyle.setValues([['項目', '数量', '金額']]);
  tableHeaderStyle
    .setBackground('#1a3c5e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#1a3c5e', SpreadsheetApp.BorderStyle.SOLID);

  const amount   = Number(customer.unit_price);
  const amountRow = sheet.getRange(ROW_DETAIL_START + 1, 1, 1, 3);
  amountRow.setValues([[`月額利用料(${Number(month)}月分)`, '1', `¥${amount.toLocaleString()}`]]);
  amountRow.setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(ROW_DETAIL_START + 1, 2, 1, 2).setHorizontalAlignment('right');

  // ----- 出勤実績明細(スタッフ数に応じて動的に行を追加) -----
  let currentRow = ROW_DETAIL_START + 3;

  sheet.getRange(currentRow, 1)
    .setValue(`出勤実績明細(${monthLabel})`)
    .setFontWeight('bold')
    .setFontColor('#1a3c5e');
  currentRow++;

  // 実績テーブルヘッダー
  const attendHeader = sheet.getRange(currentRow, 1, 1, 3);
  attendHeader.setValues([['担当者', '出勤日数', '勤務時間']]);
  attendHeader
    .setBackground('#e8ecf0')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#aaaaaa', SpreadsheetApp.BorderStyle.SOLID);
  currentRow++;

  let totalDays = 0, totalHours = 0;
  for (const rec of attendanceRows) {
    const staffRow = sheet.getRange(currentRow, 1, 1, 3);
    staffRow.setValues([[rec.staff_name, `${rec.attendance_days}日`, `${rec.work_hours}h`]]);
    staffRow.setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(currentRow, 2, 1, 2).setHorizontalAlignment('center');
    totalDays  += Number(rec.attendance_days);
    totalHours += Number(rec.work_hours);
    currentRow++;
  }

  // 合計行
  const totalRow = sheet.getRange(currentRow, 1, 1, 3);
  totalRow.setValues([['合計', `${totalDays}日`, `${totalHours}h`]]);
  totalRow
    .setFontWeight('bold')
    .setBackground('#e8ecf0')
    .setBorder(true, true, true, true, true, true, '#aaaaaa', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(currentRow, 2, 1, 2).setHorizontalAlignment('center');
  currentRow += 2;

  // ----- 金額サマリー -----
  const taxAmount = Math.floor(amount * Number(customer.tax_rate) / 100);
  const total     = amount + taxAmount;

  sheet.getRange(currentRow, 2).setValue('小計').setHorizontalAlignment('right');
  sheet.getRange(currentRow, 3).setValue(`¥${amount.toLocaleString()}`).setHorizontalAlignment('right');
  sheet.getRange(currentRow, 2, 1, 2)
    .setBorder(false, false, true, false, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(currentRow + 1, 2).setValue(`消費税(${customer.tax_rate}%)`).setHorizontalAlignment('right');
  sheet.getRange(currentRow + 1, 3).setValue(`¥${taxAmount.toLocaleString()}`).setHorizontalAlignment('right');
  sheet.getRange(currentRow + 1, 2, 1, 2)
    .setBorder(false, false, true, false, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

  // 合計金額（強調）
  sheet.getRange(currentRow + 2, 2)
    .setValue('合計金額')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('right');
  sheet.getRange(currentRow + 2, 3)
    .setValue(`¥${total.toLocaleString()}`)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('right');
  sheet.getRange(currentRow + 2, 2, 1, 2)
    .setBackground('#1a3c5e')
    .setFontColor('#ffffff')
    .setBorder(true, true, true, true, false, false, '#1a3c5e', SpreadsheetApp.BorderStyle.SOLID);

  currentRow += 4;

  // ----- フッター -----
  sheet.getRange(currentRow, 1, 1, 4)
    .setBorder(true, false, false, false, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(currentRow,     1)
    .setValue('支払条件: 月末締め翌月末払い')
    .setFontSize(9)
    .setFontColor('#444444');
  sheet.getRange(currentRow + 1, 1)
    .setValue('備考: ご不明点がございましたら担当者までご連絡ください')
    .setFontSize(9)
    .setFontColor('#444444');
}


// =====================================================================
// QuoteFormatシートをPDF化し、指定フォルダに保存する
// 429(レート制限)が返った場合は待機してリトライする
// =====================================================================
function exportSheetAsPdf(sheet, fileName, folder) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheetId = sheet.getSheetId();

  const exportUrl = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export`
    + `?format=pdf&gid=${sheetId}&portrait=true&fitw=true`;

  const token = ScriptApp.getOAuthToken();

  // 最大3回リトライ(待機時間: 3秒 → 6秒 → 9秒)
  let response;
  for (let attempt = 1; attempt <= 3; attempt++) {
    response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) break;

    if (response.getResponseCode() === 429) {
      Utilities.sleep(attempt * 3000);
    } else {
      throw new Error(`PDF出力に失敗しました（エラー: ${response.getResponseCode()}）`);
    }
  }

  const pdfBlob = response.getBlob().setName(`${fileName}.pdf`);
  folder.createFile(pdfBlob);
}


// =====================================================================
// スプレッドシートと同じ親フォルダ内の「見積書」フォルダを返す
// フォルダが存在しない場合は自動作成する
// =====================================================================
function getOrCreatePdfFolder() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();

  // 既存の「見積書」フォルダを検索
  const existingFolders = parentFolder.getFoldersByName(PDF_FOLDER_NAME);
  if (existingFolders.hasNext()) return existingFolders.next();

  // なければ新規作成
  return parentFolder.createFolder(PDF_FOLDER_NAME);
}


// =====================================================================
// work_month の値をシートの格納形式にかかわらず "YYYY-MM" 文字列に統一する
// Google Sheetsは日付をDateオブジェクトとして扱う場合があるため
// =====================================================================
function formatMonth(value) {
  if (value instanceof Date) {
    const year  = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, '0');
    return `${year}-${month}`;
  }
  return String(value);
}
