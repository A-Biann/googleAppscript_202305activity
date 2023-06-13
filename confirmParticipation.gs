let PART = 29;
let PHONEre = 6;
let PHONE = 28;
function test() {}

function confirmPart() {
  let responseS = SpreadsheetApp.openById('1gntdonK9r66kX_sk95uU-XqD4fEJCzmxLN5fzBS61a4');
  let responseRange = responseS.getDataRange();
  let response = responseRange.getValues();
  let winsheet = ss.getSheetByName('Participants');
  let winvalues = ss.getSheetByName('Participants').getDataRange().getValues();

  for (let i = 0; i < response.length; i++) {
    let j = 1;
    for (; j < winvalues.length; j++) {
      if (response[i][NAME] == winvalues[j][NAME]) {
        winvalues[j][PART] = 1;
        winsheet.getRange(j+1, PART+1).setValue(1);
        winsheet.getRange(j+1, PHONE+1).setValue("\'"+response[i][PHONEre]);
        console.log("Participants #" + (j+1) + " " + winvalues[j][NAME] + " has been confirmed!");
        break;
      }
    }
  }
}