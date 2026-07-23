/**
 * ============================================================
 *  完成偵測：盤點完成 → 觸發比較
 * ============================================================
 *  三種模式（在 Config.COMPLETION_MODE 切換）：
 *   CELL_INPUT ：指定儲存格(預設 B5)被輸入內容 → 跑（目前使用中）
 *   CHECKBOX   ：每日分頁的勾選框被打勾 → 跑
 *   ALL_FILLED ：三大區所有品項的「當日消耗」都是數字 → 跑
 * ------------------------------------------------------------
 */

/**
 * 安裝式 onEdit 觸發器的進入點。
 * （安裝式而非簡單觸發，才有權限讀寫「另一張」試算表。）
 */
function ccOnEditInstalled(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    const tabName = sheet.getName();

    // 只處理「看起來像日期」的盤點分頁
    if (!CC_CONFIG.DATE_TAB_REGEX.test(tabName)) return;

    if (CC_CONFIG.COMPLETION_MODE === 'CELL_INPUT') {
      ccHandleCellInputEdit_(e, sheet, tabName);
    } else if (CC_CONFIG.COMPLETION_MODE === 'CHECKBOX') {
      ccHandleCheckboxEdit_(e, sheet, tabName);
    } else {
      ccHandleAllFilledEdit_(e, sheet, tabName);
    }
  } catch (err) {
    console.error('ccOnEditInstalled 失敗：' + err.stack);
  }
}

/** CELL_INPUT 模式：指定儲存格被輸入內容就觸發；清空則允許之後重跑。 */
function ccHandleCellInputEdit_(e, sheet, tabName) {
  const a1 = e.range.getA1Notation();
  if (a1 !== CC_CONFIG.COMPLETION_CELL) return;

  const val = e.range.getValue();
  const hasContent = !(val === '' || val === null || val === undefined);
  const doneKey = 'done_' + tabName;
  const props = PropertiesService.getScriptProperties();

  if (!hasContent) {               // 清空 → 允許之後重跑
    props.deleteProperty(doneKey);
    return;
  }
  if (props.getProperty(doneKey) === '1') return; // 今天已跑過

  const date = ccDateFromTabName_(tabName);
  if (!date) return;
  ccRunComparisonFor_(date);
  props.setProperty(doneKey, '1');
  ccToast_('「' + tabName + '」偵測到 ' + CC_CONFIG.COMPLETION_CELL + ' 輸入 → 已更新消耗比較表 ✅');
}

/** CHECKBOX 模式：只在「完成勾選框」被改動時反應。 */
function ccHandleCheckboxEdit_(e, sheet, tabName) {
  const a1 = e.range.getA1Notation();
  if (a1 !== CC_CONFIG.COMPLETION_CELL) return;

  const checked = e.range.getValue() === true;
  const doneKey = 'done_' + tabName;
  const props = PropertiesService.getScriptProperties();

  if (!checked) {                 // 取消勾選 → 允許之後重跑
    props.deleteProperty(doneKey);
    return;
  }
  if (props.getProperty(doneKey) === '1') return; // 今天已跑過

  const date = ccDateFromTabName_(tabName);
  if (!date) return;
  ccRunComparisonFor_(date);
  props.setProperty(doneKey, '1');
  ccToast_('「' + tabName + '」盤點完成 → 已更新消耗比較表 ✅');
}

/** ALL_FILLED 模式：任何編輯後檢查三大區是否都算得出消耗。 */
function ccHandleAllFilledEdit_(e, sheet, tabName) {
  const doneKey = 'done_' + tabName;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(doneKey) === '1') return;

  const grid = sheet.getDataRange().getValues();
  const parsed = ccParseDailySheet_(grid);

  // 三大 DIRECT 區塊：每個品項的 value 都要是數字才算完成
  var allFilled = true;
  CC_CONFIG.SECTIONS.forEach(function (sec) {
    if (sec.mode !== 'DIRECT') return;
    const s = parsed[sec.key];
    if (!s || !s.order.length) { allFilled = false; return; }
    s.order.forEach(function (k) { if (!ccIsNum_(s.items[k].value)) allFilled = false; });
  });
  if (!allFilled) return;

  const date = ccDateFromTabName_(tabName);
  if (!date) return;
  ccRunComparisonFor_(date);
  props.setProperty(doneKey, '1');
  ccToast_('「' + tabName + '」全部填妥 → 已更新消耗比較表 ✅');
}

function ccToast_(msg) {
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, '消耗比較', 6); } catch (e) {}
}
