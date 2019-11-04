
function TimesheetRenderer(fetchTimesheet) {

  this.fetchTimesheet = fetchTimesheet;

  this.render = function(workspaceId, timesheetDate) {

    var timesheet = this.fetchTimesheet.execute(workspaceId, timesheetDate);

    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Toggl - " + formatYYYYMM(timesheetDate);

    var sheet = activeSpreadsheet.getSheetByName(sheetName);
    if (sheet) {
      activeSpreadsheet.deleteSheet(sheet);
    }

    var sheet = activeSpreadsheet.insertSheet(sheetName, activeSpreadsheet.getSheets().length);

    var titles = sheet.getRange(1, 1, 1, 4);
    titles.setValues([["Date", "Customer", "Duration", "Duration in Hours"]]);
    titles.setFontWeights([["bold", "bold", "bold","bold"]]);

    var row = 2

    var timesheetIterator = timesheet.iterator();

    for (var timesheetDay = timesheetIterator.next(); !timesheetDay.done; timesheetDay = timesheetIterator.next()) {

      var start = timesheetDay.value.date();
      var durationInHours = 0;

      var clientsIterator = timesheetDay.value.iterator();

      for(var item = clientsIterator.next(); !item.done; item = clientsIterator.next()) {
        var duration = millisToDuration(item.value.duration);
        durationInHours = durationInHours + millisToDecimalHours(item.value.duration);

        sheet.getRange(row, 1, 1, 4).setValues([[start, item.value.clientName, duration, millisToDecimalHours(item.value.duration)]]);
        sheet.getRange(row, 1).setNumberFormat("MM/dd/yyyy")
        ++row;
      }
    }

    sheet.autoResizeColumn(1);
    sheet.autoResizeColumn(2);
    sheet.autoResizeColumn(3);
    sheet.autoResizeColumn(4);
  };
}
