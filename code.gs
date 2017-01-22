function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu().addItem('Open wordTool',
    'showSidebar').addToUi();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var thisDoc = SpreadsheetApp.getActive();
  //   var classCode = thisDoc.getName().substring(0,12);
  var ui = HtmlService.createTemplateFromFile('Sidebar').evaluate().setTitle(
    'Reminder').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
}

function makeCall(word) {
  var link = "http://words.bighugelabs.com/api/2/a806ad9ebb59e564fc5f33826381088f/" +
    word + "/json";
  var obj = UrlFetchApp.fetch(link);
  var result = JSON.parse(obj);
    //  Logger.log(result.noun.syn)
  return result;
}

function addClass() {
  // Log information about the data-validation rule for cell A1.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('sheet3');
//  var scourceCell = sheet.getRange('A4');
//  var targetCell = sheet.getRange('B4');
  //Check to where the last row is
  var lastRow = sheet.getLastRow() + 1;
  // Check to see if th ere is a word in the class column
  for (var i = 1; i < lastRow; i++) {
    var content = sheet.getRange(i, 2);
    var contentVal = content.getValue();
    if (contentVal === "") {
      var word = sheet.getRange(i, 1).getValue();
      var call = makeCall(word);
      var text = Object.keys(call);
      var rule = SpreadsheetApp.newDataValidation().requireValueInList(text,
        false).build();
      content.setDataValidation(rule);
    }
  }
  Logger.log(text);
}

function addSynonyms() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('sheet3');
  var lastRow = sheet.getLastRow() + 1;
  for(var x = 2; x< lastRow; x ++){
  var word = sheet.getRange(x,1).getValue();
  var wClass = sheet.getRange(x,2).getValue();
  var i = 0;
  var text = makeCall(word);
  Logger.log(text);
  while (i < 5) {
    var targetCell = sheet.getRange(x, 3 + i, 1, 1);
    targetCell.setValue(text[wClass].syn[i]);
    i++;
  }
  }
}

  function appendItems() {
     var arr = ["a","b","c"];
     var ss = SpreadsheetApp.getActiveSpreadsheet();
     var sheet = ss.getSheetByName('Entities');
     sheet.appendRow(arr);
}


function cleanWords() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('sheet4');
  var scourceCell = sheet.getRange('A4');
  var targetCell = sheet.getRange('B4');
  //Check to where the last row is
  var lastRow = sheet.getLastRow() + 1;
  // Check to see if there is a word in the class column
  for (var i = 1; i < lastRow; i++) {
    var content = sheet.getRange(i, 1);
    var contentVal = content.getValue();
    var result = contentVal.toLowerCase().replace(/-/g, ' ').replace(/[0-9]/g,
      '').replace(/ /g, '');
    sheet.getRange(i, 5).setValue(result);
    Logger.log(result);
  }
}



function toUnique(a,b,c){//array,placeholder,placeholder
 b=a.length;
 while(c=--b)while(c--)a[b]!==a[c]||a.splice(c,1);
 return a ;// not needed ;);
}




function findHighlighted(id) {
  var results = [];
 
  var document =  DriveApp.getFileById(id).getId();
  var body = DocumentApp.openById(document).getBody(),
    bodyTextElement = body.editAsText(),
    bodyString = bodyTextElement.getText(),
    char, len;
  for (char = 0, len = bodyString.length; char < len; char++) {
    if (bodyTextElement.getBackgroundColor(char) !== null) // Is any hight light
      results.push([char, bodyString.charAt(char)]);
  }
  return results;
}

function getTabs() {
    var sheetName = [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    for(var item in sheets){
    sheetName.push(sheets[item].getName());
    }
    return sheetName;
}







function getWords(id,tab) {
  var arr = findHighlighted(id);
  var wordList = [];
  var holding = [];
  var nextNum, sum;
  for (var i = 0; i < arr.length; i++) {
    if (arr[i + 1] === undefined) {
      nextNum = 0;
    } else {
      nextNum = arr[i + 1][0];
    }
      sum = (Number(arr[i][0]) + 1);
    if (nextNum === sum) {
      holding.push(arr[i][1].toLowerCase())
    } else {
      holding.push(arr[i][1]);
      wordList.push(holding.join(""));
      holding = [];
    }
  }
  Logger.log(wordList);
  var uniqueWordlist = toUnique(wordList);
  appendEntities(uniqueWordlist,tab);
}

function appendEntities(wordList,tab){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(tab);
 for (var items in wordList){
  sheet.appendRow([wordList[items]]);
  }
}
