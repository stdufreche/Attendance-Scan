function onOpen(e) 
{
  var menu = SpreadsheetApp.getUi().createMenu("Attendance");
  if (e && e.authMode == ScriptApp.AuthMode.NONE) 
  {
    // Add a normal menu item (works in all authorization modes).
    menu.addItem('Insert New Week', 'showDateInputDialog');
  } 
  else 
  {
    // Add a menu item based on properties (doesn't work in AuthMode.NONE).
    menu.addItem('Insert New Week', 'showDateInputDialog');
  }
  menu.addToUi();
}

function doPost(e) 
{
  Logger.log("doPost(e)");

  switch (e.parameter.tripType) 
  {
    case  "MultiScan":
      processMultiScan(e);
      return ContentService.createTextOutput("MultiScan: " + e.parameter.tripType);
      break;

    default:
      processTrip(e);
      return ContentService.createTextOutput("default: " + e.parameter.tripType);
  }
}