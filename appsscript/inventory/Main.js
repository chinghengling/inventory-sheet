/**
 * ============================================================
 *  主進入點：選單、設定、手動執行、觸發器安裝
 * ============================================================
 */

/**
 * 建立「消耗比較」選單的實際內容。
 * 拆成獨立函式（而不是直接寫在 onOpen 裡）是為了保留彈性：
 * 如果之後這個專案又加了別的工具、也需要自己的 onOpen 邏輯，
 * 只要把各自的初始化呼叫合併寫進同一個 onOpen() 就好，一個專案只能有一個 onOpen()。
 */
function ccBuildMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('消耗比較')
    .addItem('▶ 執行今日比較', 'ccRunToday')
    .addItem('▶ 執行指定日期…', 'ccRunForDatePrompt')
    .addSeparator()
    .addItem('⚙ 初始化（安裝觸發器＋加完成勾選框）', 'ccSetup')
    .addItem('＋ 只在目前分頁加完成勾選框', 'ccAddCheckboxToActiveSheet')
    .addItem('🧹 重置今日「已跑過」旗標', 'ccResetTodayFlag')
    .addToUi();
}

/**
 * 已確認掃碼盤點後端（doGet/doPost 純 Web App）沒有自己的 onOpen()，
 * 所以這裡直接建立，不會衝突。
 * ⚠️ 如果未來這個專案又加了其他也需要 onOpen 的工具，記得把新工具的初始化呼叫
 *    也合併加進這個函式，一個專案只能有一個 onOpen()。
 */
function onOpen() {
  ccBuildMenu_();
}

/** 核心：對指定日期跑完整流程並寫出。 */
function ccRunComparisonFor_(date) {
  const result = ccBuildComparison_(date);
  const tab = ccWriteComparison_(result);
  return tab;
}

/** 手動：跑今天（依營業日換日點判斷，早上 8 點前算前一個營業日）。 */
function ccRunToday() {
  const tab = ccRunComparisonFor_(ccBusinessDate_());
  ccSafeAlert_('完成', '已在比較試算表產生分頁：' + tab);
}

/** 手動：跳出輸入框問日期（MMDD 或 yyyy/MM/dd）。 */
function ccRunForDatePrompt() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('執行指定日期', '請輸入日期（0709 或 2025/07/09）：', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const raw = res.getResponseText().trim();
  var date = /^\d{2,4}$/.test(raw) ? ccDateFromTabName_(raw) : new Date(raw.replace(/-/g, '/'));
  if (!date || isNaN(date.getTime())) { ui.alert('無法解析日期：' + raw); return; }
  const tab = ccRunComparisonFor_(date);
  ui.alert('已產生分頁：' + tab);
}

/** 初始化：安裝觸發器 +（CHECKBOX 模式才需要）幫範本與今日分頁加勾選框。 */
function ccSetup() {
  ccInstallOnEditTrigger_();

  if (CC_CONFIG.COMPLETION_MODE === 'CHECKBOX') {
    const ss = ccGetSourceSS_();
    ['『範本』', '範本(含進銷存)'].forEach(function (n) {
      const s = ss.getSheetByName(n);
      if (s) ccAddCheckbox_(s);
    });
    const todayTab = ccResolveTabName_(ccBusinessDate_(), ccListSourceTabNames_(ss));
    if (todayTab) ccAddCheckbox_(ss.getSheetByName(todayTab));
    ccSafeAlert_('初始化完成',
      '1) 已安裝完成偵測觸發器\n' +
      '2) 已在範本與今日分頁加上「盤點完成」勾選框（' + CC_CONFIG.COMPLETION_CELL + '）\n\n' +
      '提醒：請確認 Config.gs 裡兩個試算表 ID 已填好。');
  } else {
    ccSafeAlert_('初始化完成',
      '1) 已安裝完成偵測觸發器\n' +
      '2) 目前模式：' + CC_CONFIG.COMPLETION_MODE + '，儲存格 ' + CC_CONFIG.COMPLETION_CELL +
      ' 只要被輸入內容就會自動觸發，不需要額外設定\n\n' +
      '提醒：請確認 Config.gs 裡兩個試算表 ID 已填好。');
  }
}

/** 安裝式 onEdit 觸發器（避免重複安裝）。 */
function ccInstallOnEditTrigger_() {
  const exists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'ccOnEditInstalled';
  });
  if (exists) return;
  ScriptApp.newTrigger('ccOnEditInstalled')
    .forSpreadsheet(CC_CONFIG.SOURCE_SPREADSHEET_ID)
    .onEdit()
    .create();
}

/** 在某分頁指定位置放「盤點完成」勾選框與說明。 */
function ccAddCheckbox_(sheet) {
  if (!sheet) return;
  const labelRange = sheet.getRange(CC_CONFIG.COMPLETION_LABEL_CELL);
  labelRange.setValue('盤點完成→').setFontWeight('bold').setHorizontalAlignment('right');
  const cbRange = sheet.getRange(CC_CONFIG.COMPLETION_CELL);
  cbRange.insertCheckboxes();
  cbRange.setValue(false);
}

/** 選單用：只對目前開著的分頁加勾選框。 */
function ccAddCheckboxToActiveSheet() {
  ccAddCheckbox_(SpreadsheetApp.getActiveSheet());
  ccSafeAlert_('完成', '已加上勾選框於 ' + CC_CONFIG.COMPLETION_CELL);
}

/** 重置今日旗標（改資料後想重跑時用；依營業日換日點判斷）。 */
function ccResetTodayFlag() {
  const tab = ccResolveTabName_(ccBusinessDate_(), ccListSourceTabNames_(ccGetSourceSS_()));
  if (tab) PropertiesService.getScriptProperties().deleteProperty('done_' + tab);
  ccSafeAlert_('完成', '已重置 ' + tab + ' 的旗標，可重新觸發。');
}

function ccSafeAlert_(title, msg) {
  try { SpreadsheetApp.getUi().alert(title, msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { console.log(title + '：' + msg); }
}
