function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Side Bar").addItem("Sidebar Form","showFormInSidebar").addToUi();
}
//OPEN THE FORM IN SIDEBAR 
function showFormInSidebar() {      
  var form = HtmlService.createTemplateFromFile('index').evaluate().setTitle('Looping My Sheet');
  SpreadsheetApp.getUi().showSidebar(form);
}
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index');
}
// Connecting css.html file to index.html
function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent()
}

function loopingsheet(studentClass, rangeOfName, templateName, classHeader, hiddingRow){
  var sheetName = studentClass;
  if (sheetName == "") {
    sheetName = SpreadsheetApp.getActiveSpreadsheet()
                .getSheetByName('Ambil Range')
                .getRange(rangeOfName)
                .getValues()
                .toString()
                .split(',');
  }

  for (i=0; i<=(sheetName.length - 1); i++) {  
    // Active sheet for duplicated at proccess of looping
    SpreadsheetApp.getActiveSpreadsheet().getSheets();

    const templateOfSheet = SpreadsheetApp.getActiveSpreadsheet();

    // Format for looping sheet
    templateOfSheet.insertSheet(sheetName[i],1,{template: templateOfSheet.getSheetByName(templateName)});

    // Set value into custome cell
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName[i]).getRange(classHeader).setValue(sheetName[i]);

    // Number of row that will be hidden
    let notZero = () => {
    ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName[i]).getRange('A9:A48').getValues()
    filterNotZero = ss.filter((x) => x > 0)
    startNumber = 9
    return (filterNotZero.length) + startNumber
    }
    
    // Hide row if any empty cell
    if(hiddingRow == 'hide') {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName[i]).hideRow(
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName[i]).getRange(`A${notZero()}:A43`)
        )
    }
  

  }
}
