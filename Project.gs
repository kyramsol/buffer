function sendData(e) {
  try {
    var spreadsheetName = SpreadsheetApp.getActiveSheet().getName();
    if (spreadsheetName == "Working tab") {
      saveWorkingTabActions(e);
    }
    else if (spreadsheetName == "REQUESTS") {
      saveRequestsActions();
    }
    else if (spreadsheetName == "REPORT")
      reportSheet(e)
    else
      Logger.log(e);
  }
  catch (error) {
    BugTracker.storeBugData(error, "Qub", "Serhii");
    throw error;
  }
}



function reportSheet(e) {

  //FuelFinanceLibrary1.project_reportfilters(e)
  var activesheet = e.source.getActiveSheet();
  var cell = e.range;


  if (activesheet.getName() == "REPORT") {

    var projecthead = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PROJECTS LIST").getRange("D3:AT3").getValues();
    var index = projecthead[0].indexOf(e.value);
    Logger.log(index)
    var selfeature = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PROJECTS LIST").getRange(4, index + 4, SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PROJECTS LIST").getLastRow(), 1).getValues().filter((v, i, a) => a.indexOf(v) === i);

    if (cell.getColumn() == 12 && cell.getRow() == 6) {
      cell.offset(1, 0).setValue("Loading...")
      Logger.log("Ok")

      cell.offset(1, 0).clearContent().clearDataValidations();

      cell.offset(1, 3).clearContent().clearDataValidations();

      cell.offset(0, 3).clearContent();

      cell.offset(1, 6).clearContent().clearDataValidations();

      cell.offset(0, 6).clearContent();

      cell.offset(1, 9).clearContent().clearDataValidations();

      cell.offset(0, 9).clearContent();

      cell.offset(1, 12).clearContent().clearDataValidations();

      cell.offset(0, 12).clearContent();



      var Rule = SpreadsheetApp.newDataValidation().requireValueInList(selfeature).build()
      cell.offset(1, 0).setDataValidation(Rule);
    }

    else if ((cell.getColumn() == 15 || cell.getColumn() == 18 || cell.getColumn() == 21 || cell.getColumn() == 24) && cell.getRow() == 6) {
      cell.offset(1, 0).setValue("Loading...")


      var projectlist = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PROJECTS LIST").getRange(4, 4, SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PROJECTS LIST").getLastRow() - 4, SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PROJECTS LIST").getLastColumn() - 4).getValues();
      var firstfeature = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REPORT").getRange("L6").getValue();
      var firstfeaturevalue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REPORT").getRange("L7").getValue();

      //  var projecthead = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PROJECTS LIST").getRange("D3:AD3").getValues;
      var featueind = projecthead[0].indexOf(firstfeature);
      Logger.log(featueind)
      var featurtargind = projecthead[0].indexOf(e.value)

      var rulearr = projectlist.reduce((x, i) => {

        if (i[featueind] == firstfeaturevalue) {

          x.push(i[featurtargind])

        }
        return x;

      }, [])

      Logger.log(rulearr)

      cell.offset(1, 0).clearContent().clearDataValidations();
      var Rule = SpreadsheetApp.newDataValidation().requireValueInList(rulearr).build();
      cell.offset(1, 0).setDataValidation(Rule);

    }

  }
  //FuelFinanceLibrary1.costcentreproject_ids_update_addnewrow(e)


  var sh = e.source.getActiveSheet();
  var newrowbudg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Colors Handbook").getRange("A13:AC13");


  if (sh.getName() == "COST CENTER_BUDGET") {

    if (e.range.rowStart == sh.getLastRow() && sh.getRange(e.range.rowStart, 2) !== "Add сategory") {
      // sh.appendRow(["", "Add category", "", "", "","Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount", "Add amount"]) 

      sh.appendRow(["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""])
      newrowbudg.copyTo(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COST CENTER_BUDGET").getRange(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COST CENTER_BUDGET").getLastRow(), 1, 1, 29))

    }
  }

  else
    Logger.log("Not Report sheet")

}

function edit_filt(e) {

  cc();
  filt();
  //FuelFinanceLibrary1.project_reportfilters(e)   
  //FuelFinanceLibrary1.costcentreproject_ids_update_addnewrow(e)

}

