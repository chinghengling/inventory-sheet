/**
 * ============================================================
 *  輸出層：依「當日消耗報告-2.xlsx」樣板寫出當天分頁
 * ============================================================
 *  欄位：品名 / 當日消耗 / (每個 CC_CONFIG.COMPARE_COLUMNS 各出一組「比較值／評級」) / 今日備量(淡) / 今日備量(旺)
 *  版面：
 *    列1：表頭
 *    列2起：逐 zone 出現 →「JZ zone」這種標題列（A欄上色）→ 品項列
 * ------------------------------------------------------------
 */

/** 將比較結果寫入輸出試算表。分頁名 = 來源當天日期(如 0711)，同名分頁重跑時清空重寫。 */
function ccWriteComparison_(result) {
  const ss = ccGetOutputSS_();
  const tabName = ccOutputTabName_(result.date);

  var sheet = ss.getSheetByName(tabName);
  if (sheet) sheet.clear();
  else sheet = ss.insertSheet(tabName, 0); // 新分頁插到最前面

  const columns = CC_CONFIG.COMPARE_COLUMNS; // 例 [{header:'-1w 比較',weeks:[1]}, {header:'-3w 比較',weeks:[1,2,3]}]

  // 表頭：品名 / 當日消耗 / (比較值, 評級) x N / 今日備量(淡) / 今日備量(旺)
  const headers = ['品名', '當日消耗'];
  columns.forEach(function (col) { headers.push(col.header, '評級'); });
  headers.push('今日備量(淡)', '今日備量(旺)');
  const nCols = headers.length;

  // 每個比較欄對應的「比較值欄」「評級欄」在整張表的欄位編號（1-indexed），依 COMPARE_COLUMNS 順序（index-based）
  const colPositions = columns.map(function (_, i) { return { value: 3 + i * 2, grade: 4 + i * 2 }; });
  const prepMildCol = nCols - 1;
  const prepPeakCol = nCols;

  // ── 表頭列 ──────────────────────────────────────────────
  sheet.getRange(1, 1, 1, nCols).setValues([headers]);
  sheet.getRange(1, 1, 1, nCols)
    .setFontFamily(CC_CONFIG.FONT).setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange(1, 1).setHorizontalAlignment('left');
  sheet.getRange(1, 2).setBackground('#' + CC_CONFIG.COL_B_FILL);

  // 小備註（不影響主表版面，放在表頭右側外的欄位）
  const now = Utilities.formatDate(new Date(), CC_CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm');
  var note = '產生時間 ' + now;
  if (result.missingTabs && result.missingTabs.length) {
    note += '　|　找不到分頁：' + result.missingTabs.join('、');
  }
  sheet.getRange(1, nCols + 2).setValue(note)
    .setFontFamily(CC_CONFIG.FONT).setFontSize(9).setFontColor('#666666');

  // ── 逐 zone 寫入 ────────────────────────────────────────
  var r = 2;
  CC_CONFIG.SECTIONS.forEach(function (sec) {
    const secRows = result.rows.filter(function (row) { return row.section === sec.key; });
    if (!secRows.length) return;

    // zone 標題列
    sheet.getRange(r, 1).setValue(sec.key + ' zone');
    sheet.getRange(r, 1)
      .setFontFamily(CC_CONFIG.FONT).setFontSize(14).setFontWeight('bold')
      .setBackground('#' + (CC_CONFIG.ZONE_COLORS[sec.key] || 'CCCCCC'));
    sheet.getRange(r, 2).setBackground('#' + CC_CONFIG.COL_B_FILL);
    r++;

    // 品項列
    secRows.forEach(function (row) {
      const rowVals = new Array(nCols).fill('');
      rowVals[0] = row.name;
      rowVals[1] = ccIsNum_(row.today) ? ccRound2_(row.today) : '';
      row.cols.forEach(function (c, i) {
        rowVals[colPositions[i].value - 1] = ccIsNum_(c.value) ? ccRound2_(c.value) : '';
        rowVals[colPositions[i].grade - 1] = c.grade;
      });
      rowVals[prepMildCol - 1] = row.prepMild == null ? '' : row.prepMild;
      rowVals[prepPeakCol - 1] = row.prepPeak == null ? '' : row.prepPeak;

      sheet.getRange(r, 1, 1, nCols).setValues([rowVals]);

      // A 欄：品名，粗體 14pt（照樣板）
      sheet.getRange(r, 1).setFontFamily(CC_CONFIG.FONT).setFontSize(14).setFontWeight('bold');
      // B 欄：淡黃底（照樣板），非粗體、置中
      sheet.getRange(r, 2)
        .setFontFamily(CC_CONFIG.FONT).setFontWeight('normal')
        .setBackground('#' + CC_CONFIG.COL_B_FILL).setHorizontalAlignment('center');

      // 每個比較欄的比較值欄 + 評級欄
      row.cols.forEach(function (c, i) {
        const vCell = sheet.getRange(r, colPositions[i].value);
        vCell.setFontFamily(CC_CONFIG.FONT).setFontSize(10).setFontWeight('normal')
          .setHorizontalAlignment('center');
        const gCell = sheet.getRange(r, colPositions[i].grade);
        gCell.setFontFamily(CC_CONFIG.FONT).setFontSize(10).setFontWeight('bold')
          .setHorizontalAlignment('center');
        if (CC_CONFIG.APPLY_GRADE_COLOR && CC_CONFIG.GRADE_COLORS[c.grade]) {
          gCell.setBackground('#' + CC_CONFIG.GRADE_COLORS[c.grade]);
        }
      });

      // 今日備量（淡／旺）：純數字，置中
      sheet.getRange(r, prepMildCol, 1, 2)
        .setFontFamily(CC_CONFIG.FONT).setFontSize(10).setFontWeight('normal')
        .setHorizontalAlignment('center');

      r++;
    });
  });

  // ── 版面收尾 ────────────────────────────────────────────
  sheet.setColumnWidth(1, CC_CONFIG.NAME_COL_WIDTH);
  for (var c = 2; c <= nCols; c++) sheet.setColumnWidth(c, CC_CONFIG.DATA_COL_WIDTH);
  sheet.setFrozenRows(1);

  return tabName;
}

/** 依 OUTPUT_TAB_PATTERN 產生分頁名。目前支援 'MMDD'。 */
function ccOutputTabName_(date) {
  if (CC_CONFIG.OUTPUT_TAB_PATTERN === 'MMDD') {
    return Utilities.formatDate(date, CC_CONFIG.TIMEZONE, 'MMdd');
  }
  return Utilities.formatDate(date, CC_CONFIG.TIMEZONE, CC_CONFIG.OUTPUT_TAB_PATTERN);
}
