// get Spreadsheet
let ss = SpreadsheetApp.openById('1T6n7roA_O4jVxwBrBaEf0Qpxh1FfI07WpGiYjD9VoKU');
let sheet = ss.getSheetByName('全部人'); // sheet name
//let sheet = ss.getSheetByName('工作表4'); // sheet name
let NAME = 2;
let SCORE = 19;
let WIN = 27;
let EMAIL = 1;
let SEND = 28;
let participants = [];
let dataRange = sheet.getDataRange();
let values = dataRange.getValues();

function test() {}

function randomPicker() {
  // get the participants' name & score filter
  for (let i = 1; i < values.length; i++) { // from row2 to the end
    sheet.getRange(i+1, WIN+1).setValue(0); //winner = 0
  }

  let winners = [];
  let cnt = 0;
  while (winners.length < 30) {
    let index = Math.floor(Math.random() * values.length); //random index
    if (index == 0 || values[index][SCORE] < 2) continue;
    if (winners.indexOf(index) == -1) {
      winners.push(index);
      sheet.getRange(index+1, WIN+1).setValue(1);
      Logger.log("Winner #" + (cnt+1) + " is " + values[index][NAME]);
      cnt = cnt+1;
    }
  }

  let winsheet = ss.getSheetByName('Participants');
  if (winsheet) {
    ss.deleteSheet(winsheet);
    ss.insertSheet("Participants");
  } else {
    ss.insertSheet("Participants");
  }
  winsheet = ss.getSheetByName('Participants');
  winsheet.appendRow(values[0]);

  // winner's value to new sheet
  for (let i = 0; i < winners.length; i++) {
    //Logger.log("Winner #" + (i+1) + " is " + values[winners[i]][NAME]);
    winsheet.appendRow(values[winners[i]]);
    winsheet.getRange(i+2, WIN+1).setValue(1);
  }
}

function findAlternate() {
  let winsheet = ss.getSheetByName('Participants');
  let alternates = [];
  while (alternates.length < 1) {
    let index = Math.floor(Math.random() * values.length); //random index
    if (alternates.indexOf(index) == -1 && !values[index][WIN] && values[index][SCORE]>0) {
      alternates.push(index);
      sheet.getRange(index+1, WIN+1).setValue(2);
      winsheet.appendRow(values[index]);
      winsheet.getRange(winsheet.getLastRow(), WIN+1).setValue(2);
    }
  }

  //Logger.log("Alternative" + " is " + values[alternate][NAME]);
}

function sendSuccessEmail() {
  var templ = HtmlService.createTemplateFromFile('successEmail');
  var message = templ.evaluate().getContent();
  let winvalues = ss.getSheetByName('Participants').getDataRange().getValues();

  let Main = "【正取通知】5/5你需姐姐下午茶極限斜槓場次正取通知";
  let Main2 = "【請確認出席】【正取通知】5/5你需姐姐下午茶極限斜槓場次第二次正取通知";

  for (let i = 1; i < values.length; i++) {
    if (values[i][WIN] == 1) {
      let Email = values[i][EMAIL];
      console.log(Email);
      MailApp.sendEmail({
        to: Email,
        subject: Main2,
        htmlBody: message
      });
    }
  }
}

function sendFailEmail() {
  var templ = HtmlService.createTemplateFromFile('failEmail');
  var message = templ.evaluate().getContent();
  let cnt = 0;
  let Main = "【未錄取通知】5/5你需姊姊下午茶極限斜槓場次未錄取通知";
  let Main2 = "【測試】5/5你需姊姊下午茶極限斜槓場次未錄取通知";
  for (let i = 1; i < values.length; i++) {
    if (values[i][WIN] == 0) {
      cnt++;
      sheet.getRange(i+1, SEND+1).setValue(1);
      let Email = values[i][EMAIL];
      console.log(Email);
      MailApp.sendEmail({
        to: Email,
        subject: Main,
        htmlBody: message
      });
    }
  }
  console.log(cnt);
}

function sendAlternateEmail() {
  var templ = HtmlService.createTemplateFromFile('alternateEmail');
  var message = templ.evaluate().getContent();
  let winvalues = ss.getSheetByName('Participants').getDataRange().getValues();

  let Main = "【請確認出席】【候補成功通知】5/5你需姐姐下午茶極限斜槓場次候補正取通知";
  for (let i = 1; i < winvalues.length; i++) {
    if (winvalues[i][WIN] == 2) {
      let Email = winvalues[i][EMAIL];
      console.log(Email);
      MailApp.sendEmail({
        to: Email,
        subject: Main,
        htmlBody: message
      });
    }
  }
}

function sendNotPartEmail() {
  let doc = DocumentApp.openById("1OsSm1_DzwEGKlQ4DCKcJ_Oeit2l5fIicaYBxfoBYntw");
  let message = doc.getBody().getText();
  let winsheet = ss.getSheetByName('Participants');
  let winvalues = ss.getSheetByName('Participants').getDataRange().getValues();

  console.log(message);

  let Main = "【取消資格】5/5你需姐姐下午茶極限斜槓場次候補取消資格通知";
  for (let i = 1; i < winvalues.length; i++) {
    if (winvalues[i][WIN] == 1 && winvalues[i][PART] == 0) {
      winsheet.getRange(i+1, WIN+1).setValue(3);
      let Email = winvalues[i][EMAIL];
      let Name = winvalues[i][NAME];
      console.log(Name+Email);
      GmailApp.sendEmail(
        Email,
        Main,
        message
      );
    }
  }
}


