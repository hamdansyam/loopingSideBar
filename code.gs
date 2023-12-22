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

function loopingsheet(kelas, ambilRange, namaTemplate){
var namaSheet = kelas;
if (namaSheet == "") {namaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ambil Range').getRange(ambilRange).getValues().toString().split(',');}
var berapaKali = namaSheet.length - 1;
  for (i=0; i<=berapaKali; i++) {  
    let sheetYgAktif = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var jumlahSheet = namaSheet[i];
    const sheetTemplate = SpreadsheetApp.getActiveSpreadsheet();
    var formatLooping = sheetTemplate.insertSheet(jumlahSheet,1,{template: sheetTemplate.getSheetByName(namaTemplate)});
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet[i]).getRange('a6').setValue(namaSheet[i]);
    }
}
