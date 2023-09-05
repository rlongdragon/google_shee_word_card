function showChinese() {
  let ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName("字卡");
  // sheet.getRange(7, 3).setBackgroundRGB(0,0,0)
  sheet.getRange(8, 4).setFontColor("black");
}

function next() {
  let ss = SpreadsheetApp.getActive();
  let wordsSheet = ss.getSheetByName("單字表");
  let wordList = ss.getSheetByName("測驗單字表");
  let wordCard = ss.getSheetByName("字卡");
  let global = ss.getSheetByName("環境變數");

  if (global.getRange(4, 2).getValue() == 0) {
    wordCard.getRange(6, 4).setFontColor("#34a853");
  }

  // 隱藏答案
  wordCard.getRange(8, 4).setFontColor("#FFF");

  // 隨機抓取單字
  let length = global.getRange(4, 2).getValue();
  let word = Math.floor(Math.random() * length)

  global.getRange(1, 2).setValue(word);
  Logger.log(word)
}

function start() {
  let ss = SpreadsheetApp.getActive();
  let wordsSheet = ss.getSheetByName("單字表");
  let wordList = ss.getSheetByName("測驗單字表");
  let wordCard = ss.getSheetByName("字卡");
  let global = ss.getSheetByName("環境變數");

  let top = global.getRange(3, 2).getValue();
  let bottom = global.getRange(2, 2).getValue();

  // 初始設定所有單字的狀態
  let count = wordCard.getRange(3, 4).getValue();
  wordList.getRange(`C1:C${top - bottom}`).setValue(count);

  // 從單字庫中複製要用的單字到測驗單字表
  wordList.getRange(`A1:B${top - bottom}`).setValues(wordsSheet.getRange(`A${bottom + 2}:B${top + 1}`).getValues());

  global.getRange(4, 2).setValue(top - bottom);

  wordCard.getRange(6, 4).setFontColor("#FFF");
  next()
}

function knowed() {
  let ss = SpreadsheetApp.getActive();
  let wordList = ss.getSheetByName("測驗單字表");
  let global = ss.getSheetByName("環境變數");

  let word = global.getRange(1, 2).getValue()

  wordList.getRange(word + 1, 3).setValue(wordList.getRange(word + 1, 3).getValue() - 1);

  if (wordList.getRange(word + 1, 3).getValue() == 0) {
    wordList.deleteRow(word + 1);
    global.getRange(4, 2).setValue(global.getRange(4, 2).getValue() - 1);
  }

  next();
}

function skilled() {
  let ss = SpreadsheetApp.getActive();
  let wordList = ss.getSheetByName("測驗單字表");
  let global = ss.getSheetByName("環境變數");

  let word = global.getRange(1, 2).getValue();

  wordList.getRange(word + 1, 3).setValue(wordList.getRange(word + 1, 3).getValue() - 2);

  if (wordList.getRange(word + 1, 3).getValue() <= 0) {
    wordList.deleteRow(word + 1);
    global.getRange(4, 2).setValue(global.getRange(4, 2).getValue() - 1);
  }

  next();
}

function knowlsee() {
  let ss = SpreadsheetApp.getActive();
  let wordCard = ss.getSheetByName("字卡");
  let wordList = ss.getSheetByName("測驗單字表");
  let global = ss.getSheetByName("環境變數");

  let word = global.getRange(1, 2).getValue();

  wordList.getRange(word + 1, 3).setValue(wordList.getRange(word + 1, 3).getValue() + 1);

  if (wordList.getRange(word + 1, 3).getValue() == wordCard.getRange(3, 4).getValue() + 1) {
    wordList.getRange(word + 1, 3).setValue(wordList.getRange(word + 1, 3).getValue() - 1);
  }

  next();
}