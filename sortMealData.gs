function myFunction(submission) {
  //let ss = SpreadsheetApp.getActive();
  let ss = submission.source;
  let sheet = ss.getSheetByName("表單回應 1");
  let Drinksheet = ss.getSheetByName("飲料");
  let Mealsheet = ss.getSheetByName("餐盒");

  Meals = ["原味堅果(素)", "梅花豬肉", "雪花牛肉"];
  Drinks = ["春芽冷露", "熟成紅茶(微糖)", "熟成紅茶(無糖)", "不要手搖"];

  let Last = sheet.getLastRow();
  let Name = sheet.getRange(Last, 3).getValue();
  let Email = sheet.getRange(Last, 2).getValue();
  let Meal = sheet.getRange(Last, 4).getValue();
  let Drink = sheet.getRange(Last, 5).getValue();
  let Phone = sheet.getRange(Last, 7).getValue();

  //drink
  for (let i = 0; i < Drinks.length; i++) {
    if (Drink == Drinks[i]) {
      for (let j = 2; j < 70; j++) {
        if (Drinksheet.getRange(j, 2*i+1).getValue() === "") {
          Drinksheet.getRange(j, 2*i+1).setValue(Name);
          break;
        }
      }
      break;
    }
  }
  //meal
  for (let i = 0; i < Meals.length; i++) {
    if (Meal == Meals[i]) {
      for (let j = 2; j < 70; j++) {
        if (Mealsheet.getRange(j, 2*i+1).getValue() === "") {
          Mealsheet.getRange(j, 2*i+1).setValue(Name);
          break;
        }
      }
      break;
    }
  }
}