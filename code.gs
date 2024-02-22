function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Side Bar").addItem("Sidebar Form","showFormInSidebar").addToUi();
}
//OPEN THE FORM IN SIDEBAR 
function showFormInSidebar() {      
  var form = HtmlService.createTemplateFromFile('index').evaluate().setTitle('Input Data Sidebar');
  SpreadsheetApp.getUi().showSidebar(form);
}
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index');
}

function loopingsheet(kelas, ambilRange, namaTemplate, classHeader){

var namaSheet = kelas;
if (namaSheet == "") {namaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ambil Range').getRange(ambilRange).getValues().toString().split(',');}
var berapaKali = namaSheet.length - 1;
  for (i=0; i<=berapaKali; i++) {  
    let sheetYgAktif = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var jumlahSheet = namaSheet[i];
    const sheetTemplate = SpreadsheetApp.getActiveSpreadsheet();
    var formatLooping = sheetTemplate.insertSheet(jumlahSheet,1,{template: sheetTemplate.getSheetByName(namaTemplate)});

    // Set value into custome cell
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet[i]).getRange(classHeader).setValue(namaSheet[i]);

    // Number of row that will be hidden
    let notZero = () => {
    ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet[i]).getRange('A9:A43').getValues()
    filterNotZero = ss.filter((x) => x > 0)
    startNumber = 9
    return (filterNotZero.length) + startNumber
    }

    // Hide row if any an empty cell
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet[i]).hideRow(
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet[i]).getRange(`A${notZero()}:A43`)
    )

  }
}
