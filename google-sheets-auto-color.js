/**
 * 學生簽到系統 - Google Sheets 自動變色腳本
 * 
 * 使用方式：
 * 1. 打開 Google Sheets
 * 2. 選單 → 擴充功能 → Apps Script
 * 3. 刪除預設內容，貼上此腳本
 * 4. 點擊「執行」按鈕（選擇 setupAllConditionalFormatting 函數）
 * 5. 授權後即可自動設定所有條件式格式
 */

// 顏色定義
const COLORS = {
  GREEN: '#d4edda',   // 淡綠色
  YELLOW: '#fff3cd',  // 淡黃色
  RED: '#f8d7da'      // 淡紅色
};

/**
 * 主函數：設定所有工作表的條件式格式
 */
function setupAllConditionalFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 簽到紀錄
  setupCheckinRecordFormatting(ss);
  
  // 2. 請假紀錄
  setupLeaveRecordFormatting(ss);
  
  // 3. 出席統計
  setupAttendanceStatsFormatting(ss);
  
  // 4. 調代課紀錄
  setupSubstituteFormatting(ss);
  
  SpreadsheetApp.getUi().alert('✅ 條件式格式設定完成！\n\n已設定：\n• 簽到紀錄\n• 請假紀錄\n• 出席統計\n• 調代課紀錄');
}

/**
 * 簽到紀錄 - 條件式格式
 * 狀態欄位：已報到(綠)、遲到(黃)、缺席(紅)
 */
function setupCheckinRecordFormatting(ss) {
  const sheet = ss.getSheetByName('簽到紀錄');
  if (!sheet) {
    Logger.log('找不到「簽到紀錄」工作表');
    return;
  }
  
  // 清除現有條件式格式
  sheet.clearConditionalFormatRules();
  
  // 找到「狀態」欄位的索引
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('狀態') + 1;
  
  if (statusCol === 0) {
    Logger.log('簽到紀錄：找不到「狀態」欄位');
    return;
  }
  
  // 設定範圍（整個資料區域）
  const range = sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getLastColumn());
  const statusColLetter = columnToLetter(statusCol);
  
  const rules = [];
  
  // 已報到 - 淡綠色
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + statusColLetter + '2="已報到"')
    .setBackground(COLORS.GREEN)
    .setRanges([range])
    .build());
  
  // 遲到 - 淡黃色
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + statusColLetter + '2="遲到"')
    .setBackground(COLORS.YELLOW)
    .setRanges([range])
    .build());
  
  // 缺席 - 淡紅色
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + statusColLetter + '2="缺席"')
    .setBackground(COLORS.RED)
    .setRanges([range])
    .build());
  
  sheet.setConditionalFormatRules(rules);
  Logger.log('✅ 簽到紀錄格式設定完成');
}

/**
 * 請假紀錄 - 條件式格式
 * 狀態欄位：已核准(綠)、待審核(黃)、已駁回(紅)
 */
function setupLeaveRecordFormatting(ss) {
  const sheet = ss.getSheetByName('請假紀錄');
  if (!sheet) {
    Logger.log('找不到「請假紀錄」工作表');
    return;
  }
  
  // 清除現有條件式格式
  sheet.clearConditionalFormatRules();
  
  // 找到「狀態」欄位的索引
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('狀態') + 1;
  
  if (statusCol === 0) {
    Logger.log('請假紀錄：找不到「狀態」欄位');
    return;
  }
  
  // 設定範圍
  const range = sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getLastColumn());
  const statusColLetter = columnToLetter(statusCol);
  
  const rules = [];
  
  // 已核准 - 淡綠色
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + statusColLetter + '2="已核准"')
    .setBackground(COLORS.GREEN)
    .setRanges([range])
    .build());
  
  // 待審核 - 淡黃色
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + statusColLetter + '2="待審核"')
    .setBackground(COLORS.YELLOW)
    .setRanges([range])
    .build());
  
  // 已駁回 - 淡紅色
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + statusColLetter + '2="已駁回"')
    .setBackground(COLORS.RED)
    .setRanges([range])
    .build());
  
  sheet.setConditionalFormatRules(rules);
  Logger.log('✅ 請假紀錄格式設定完成');
}

/**
 * 出席統計 - 條件式格式
 * 完整支援三種格式：
 * - 文字格式："100%", "0%", "24%"
 * - 數字格式：100, 0, 24
 * - 百分比格式：1, 0, 0.24
 */
function setupAttendanceStatsFormatting(ss) {
  const sheet = ss.getSheetByName('出席統計');
  if (!sheet) {
    Logger.log('找不到「出席統計」工作表');
    return;
  }
  
  // 清除現有條件式格式
  sheet.clearConditionalFormatRules();
  
  // 出席率在 G 欄
  const col = 'G';
  
  // 設定範圍
  const range = sheet.getRange('A2:H500');
  
  const rules = [];
  
  // ========== 規則1: 100% - 淡綠色 ==========
  // 文字 "100%" | 數字 100 | 百分比 1
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($' + col + '2="100%", $' + col + '2=100, AND(ISNUMBER($' + col + '2), $' + col + '2=1))')
    .setBackground('#d4edda')
    .setRanges([range])
    .build());
  
  // ========== 規則2: 0% - 淡紅色 ==========
  // 文字 "0%" | 數字 0（排除空白列）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(LEN($A2)>0, OR($' + col + '2="0%", AND(ISNUMBER($' + col + '2), $' + col + '2=0)))')
    .setBackground('#f8d7da')
    .setRanges([range])
    .build());
  
  // ========== 規則3: 1%~59% - 淡黃色 ==========
  // 文字格式：提取數字判斷
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(LEN($A2)>0, ISNUMBER(VALUE(SUBSTITUTE($' + col + '2,"%",""))), VALUE(SUBSTITUTE($' + col + '2,"%",""))>0, VALUE(SUBSTITUTE($' + col + '2,"%",""))<60)')
    .setBackground('#fff3cd')
    .setRanges([range])
    .build());
  
  // 數字格式：1~59
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(LEN($A2)>0, ISNUMBER($' + col + '2), $' + col + '2>=1, $' + col + '2<60)')
    .setBackground('#fff3cd')
    .setRanges([range])
    .build());
  
  // 百分比格式：0.01~0.59
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(LEN($A2)>0, ISNUMBER($' + col + '2), $' + col + '2>0, $' + col + '2<0.6)')
    .setBackground('#fff3cd')
    .setRanges([range])
    .build());
  
  sheet.setConditionalFormatRules(rules);
  Logger.log('✅ 出席統計格式設定完成 - 共 ' + rules.length + ' 條規則');
}

/**
 * 欄位數字轉字母（1=A, 2=B, ...）
 */
function columnToLetter(column) {
  let letter = '';
  while (column > 0) {
    const temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = Math.floor((column - temp - 1) / 26);
  }
  return letter;
}

/**
 * 建立選單（可選）
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎨 自動變色')
    .addItem('設定所有條件式格式', 'setupAllConditionalFormatting')
    .addSeparator()
    .addItem('只設定簽到紀錄', 'setupCheckinOnly')
    .addItem('只設定請假紀錄', 'setupLeaveOnly')
    .addItem('只設定出席統計', 'setupStatsOnly')
    .addItem('只設定調代課紀錄', 'setupSubstituteOnly')
    .addToUi();
}

function setupCheckinOnly() {
  setupCheckinRecordFormatting(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('✅ 簽到紀錄格式設定完成！');
}

function setupLeaveOnly() {
  setupLeaveRecordFormatting(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('✅ 請假紀錄格式設定完成！');
}

function setupStatsOnly() {
  setupAttendanceStatsFormatting(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('✅ 出席統計格式設定完成！');
}

function setupSubstituteOnly() {
  setupSubstituteFormatting(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('✅ 調代課紀錄格式設定完成！');
}

/**
 * 調代課紀錄 - 條件式格式
 * 類型欄位：調課(淡藍色)、代課(淡橘色)
 */
function setupSubstituteFormatting(ss) {
  const sheet = ss.getSheetByName('調代課紀錄');
  if (!sheet) {
    Logger.log('找不到「調代課紀錄」工作表');
    return;
  }
  
  // 清除現有條件式格式
  sheet.clearConditionalFormatRules();
  
  // 類型在 B 欄
  const col = 'B';
  
  // 設定範圍
  const range = sheet.getRange('A2:K500');
  
  const rules = [];
  
  // 調課 - 淡藍色
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + col + '2="調課"')
    .setBackground('#e3f2fd')
    .setRanges([range])
    .build());
  
  // 代課 - 淡橘色
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + col + '2="代課"')
    .setBackground('#fff3e0')
    .setRanges([range])
    .build());
  
  sheet.setConditionalFormatRules(rules);
  Logger.log('✅ 調代課紀錄格式設定完成');
}
