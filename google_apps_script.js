/**
 * ==========================================================================
 * 芳香閣 芳療個案身心諮詢表 - Google Apps Script 雲端自動記錄程式
 * ==========================================================================
 * 
 * 【功能說明】：
 * 1. 接收來自網頁 (index.html) 身心諮詢表送出的 POST 請求。
 * 2. 自動在 Google Sheet 建立並美化欄位表頭（深綠品牌色、置中對齊、凍結首列）。
 * 3. 將個案諮詢資料（時間、姓名、年齡、主訴、安全禁忌、香氣偏好等）自動寫入新列。
 * 4. 支援 GET 請求做為服務存活測試。
 */

// 表頭欄位名稱定義
var HEADERS = [
  "填表時間",
  "姓名 / 稱呼",
  "生理性別",
  "年齡區間",
  "聯絡電話",
  "電子信箱",
  "職業生活型態",
  "主要身心主訴",
  "詳細狀況描述",
  "懷孕 / 哺乳狀態",
  "懷孕週數",
  "病史與安全禁忌",
  "過敏史 (含基底油)",
  "常規用藥情況",
  "偏好香氣調性",
  "排斥氣味雷區",
  "知情同意確認"
];

// 處理 POST 請求 (接收網頁表單資料)
function doPost(e) {
  var lock = LockService.getScriptLock();
  // 加上鎖定避免多人同時送出造成欄位覆蓋
  lock.tryLock(10000);

  try {
    var sheet = getOrCreateSheet();
    var data = {};

    // 解析傳入的 JSON 內容
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    // 依序對應欄位值
    var row = [
      data.timestamp || new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei', hour12: false }),
      data.clientName || "",
      data.gender || "",
      data.clientAge || "",
      data.clientPhone || "",
      data.clientEmail || "",
      data.lifestyle || "",
      data.concerns || "",
      data.symptomDetail || "",
      data.pregnancyStatus || "",
      data.pregnancyWeeks || "",
      data.healthConditions || "",
      data.allergyDetails || "",
      data.medicationDetails || "",
      data.scentFavorites || "",
      data.scentDislikes || "",
      data.consent || ""
    ];

    // 新增資料至試算表最後一列
    sheet.appendRow(row);

    // 回傳成功 JSON 回應
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "諮詢資料已成功儲存至 Google 試算表",
      timestamp: data.timestamp
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

// 處理 GET 請求 (測試連線用)
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    message: "芳香閣 個案諮詢表 Google Apps Script 雲端服務正常運作中！",
    time: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

// 取得或自動建立美化工作表
function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "個案諮詢紀錄";
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // 若工作表第一列為空，自動填入表頭並美化
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);

    // 美化表頭樣式 (芳香閣森林深綠品牌色系 #4A6B53)
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground("#4A6B53");
    headerRange.setFontColor("#FFFFFF");
    headerRange.setFontWeight("bold");
    headerRange.setHorizontalAlignment("center");
    headerRange.setVerticalAlignment("middle");
    sheet.setRowHeight(1, 38);
    sheet.setFrozenRows(1); // 凍結第一列
  }

  return sheet;
}
