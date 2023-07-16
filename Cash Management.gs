function onOpen(e) {

  SpreadsheetApp.getUi()
    .createMenu('Refresh Projects List')
    .addItem('Refresh Sheet', 'copyData')
    .addToUi();
  try {
    budgetfreez()
  }
  catch { }
}

function insteadImport() {

  ImRange()

}


function copyData() {

  const SOURCE_SPREADSHEET_ID = "1YH81k0aVtIszac7klfbRU79hlavwlE3PWh3PCwTGSUo";
  const TARGET_SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const SOURCE_SHEET_NAME = "PROJECTS LIST";
  const TARGET_SHEET_NAME = "PROJECTS LIST";
  const NAME_COL_INDEX = 15;
  const NAME = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('General Settings').getRange('B2').getValue();
  Logger.log(NAME)

  const sourceSheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID).getSheetByName(SOURCE_SHEET_NAME);
  const sourceValues = sourceSheet.getRange("E4:AT200").getValues();
  const targetValues = sourceValues.filter((row, i) => row[NAME_COL_INDEX - 1] == NAME);
  const targetSheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID).getSheetByName(TARGET_SHEET_NAME);
  targetSheet.getRange("E4:AT200").clearContent();
  targetSheet.getRange(4, 5, targetValues.length, targetValues[0].length).setValues(targetValues);
  try {
    copyInvoicesWithID()
  }
  catch {}
  try {
    getRequestsStatuses()
  }
  catch {}
}


function budgetfreez() {

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COST CENTER_BUDGET").getRange("F1:AB1").getValues();
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COST CENTER_BUDGET");
  var me = Session.getEffectiveUser();
  var protections = s.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  for (var b = 0; b < protections.length; b++) {

    protections[b].remove()
  }


  Logger.log(sh)
  Logger.log(sh[0].length)

  for (var i = 0; sh[0].length >= i; i++) {
    if (sh[0][i] == true) {
      s.getRange(1, 6 + i, s.getLastRow(), 1).setFontColor("#999999");
      var protection = s.getRange(1, 6 + i, s.getLastRow(), 1).protect().setDescription('Approved monthly budget range');
      protection.addEditor(me);
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
      }
    }


    else {
      s.getRange(1, 6 + i, s.getLastRow(), 1).setFontColor("black")

    }

  }
}


function ImRange() {


  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var settings = ss.getSheetByName('Import Settings')
  var ids = settings.getRange(2, 1, settings.getLastRow() - 1, settings.getLastColumn()).getValues()
  function importRange(sourceSSid, sourcesheetid, sourcerange, targetsheetid, targetrange) {
    var source = getSheetBySpreadsheetsId(sourceSSid, sourcesheetid).getRange(sourcerange).getValues()
    var row = getSheetById(targetsheetid).getRange(targetrange).getRow()
    var col = getSheetById(targetsheetid).getRange(targetrange).getColumn()
    var target = getSheetById(targetsheetid).getRange(row, col, source.length, source[0].length)
    target.setValues(source)
  }
  for (var i = 0; i < ids.length; i++) {
    var sourceSSid = ids[i][1]
    Logger.log(sourceSSid)
    var sourcesheetid = ids[i][2]
    var sourcerange = ids[i][3]
    var targetsheetid = ids[i][4]
    var targetrange = ids[i][5]
    importRange(sourceSSid, sourcesheetid, sourcerange, targetsheetid, targetrange)
  }
}

//\\ Get sheet by spreadsheet id and sheet id//\\

function getSheetBySpreadsheetsId(spreadsheetsid, sheetid) {
  return SpreadsheetApp.openById(spreadsheetsid).getSheets().filter(
    function (s) { return s.getSheetId() == sheetid }
  )[0]
}

//\\ Get sheet from active spreadsheet by id//\\

function getSheetById(id) {
  return SpreadsheetApp.getActive().getSheets().filter(
    function (s) { return s.getSheetId() == id }
  )[0]
}

//\\ Get range from row and column//\\
function get_cell_rc(row, column) {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(row, column)
}


//\\ Identify the last row with the data according to input range length //\\

function getLastRowSpecial(range) {
  var rowNum = 0;
  var blank = false;
  for (var row = 0; row < range.length; row++) {

    if (range[row][0] === "" && !blank) {
      rowNum = row;
      blank = true;

    } else if (range[row][0] !== "") {
      blank = false;
    };
  };
  return rowNum;
};



function Rows() {

  var lastRow1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QB: input vs Database").getRange("A1:A").getValues();
  Logger.log(getLastRowSpecial(lastRow1))


}


