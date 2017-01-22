function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu().addItem('Open Extraction Tool',
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
    'Word Extractor').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
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
