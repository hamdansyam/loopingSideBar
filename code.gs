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
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = (sheet) => ss.getSheetByName(sheet)
  let range = (s, r) => sheet(s).getRange(r).getValue()
  let ranges = (s, r) => sheet(s).getRange(r).getValues()
  let value = (s, r, v) => sheet(s).getRange(r).setValue(v) 
  let values = (s, r, v) => sheet(s).getRange(r).setValue(v) 

  var namaSheet = kelas;
  if (namaSheet == "") {
    namaSheet = ranges('Ambil Range', ambilRange).toString().split(',');
    }
  var berapaKali = namaSheet.length - 1;
    for (i=0; i<=berapaKali; i++) {  
      let ss = SpreadsheetApp.getActiveSpreadsheet()
      let sheet = (sheet) => ss.getSheetByName(sheet)
      let range = (s, r) => sheet(s).getRange(r).getValue()
      let ranges = (s, r) => sheet(s).getRange(r).getValues()
      let value = (s, r, v) => sheet(s).getRange(r).setValue(v) 
      let values = (s, r, v) => sheet(s).getRange(r).setValue(v) 

      /* The active sheet */
      ss.getSheets();

      var jumlahSheet = namaSheet[i];
      
      /* Format Looping */
      var formatLooping = ss.insertSheet(jumlahSheet,1,{template: sheet(namaTemplate)});

      // Set value into custome cell
      value(namaSheet[i], classHeader, namaSheet[i])

      // Number of row that will be hidden
      let notZero = () => {
        filterNotZero = ranges(namaSheet[i],'A9:A43').filter((x) => x > 0)
        startNumber = 9
        return (filterNotZero.length) + startNumber
      }

      // Hide row if any an empty cell
      sheet(namaSheet[i]).hideRow(
        sheet(namaSheet[i]).getRange(`A${notZero()}:A43`)
      )

    }
}
