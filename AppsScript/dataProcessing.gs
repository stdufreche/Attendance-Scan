//Process scanned ID and time list from AHK program
function processMultiScan(e) {
  Logger.log("processMultiScan(e)");
  try {
    const sheet = SpreadsheetApp.getActive();
    const sheetID = sheet.getId();
    const tripType = "Attendance";
    const data = e.parameters;
    
    var sheetScanLog = sheet.getSheetByName('Attendance Log');
    var sheetConfig = sheet.getSheetByName('Config');
    var currentConfig = sheetConfig.getRange(1,1,sheetConfig.getLastRow(),sheetConfig.getLastColumn()).getValues();
    var scanValues = [];
    var scanPeriod = [];

    for (const key in data) { // Loop through all Time-SIDNO data pairs
      console.log('Key: %s, Value: %s', key, data[key]);
      var scanDateTime = new Date(key);
      var studentName = RosterRecall.getStudentBySIDNO(sheetID, data[key]); // Pull names from master roster through RosterRecall library

      if (key != "tripType") { // Process if not tripType parameter
        scanPeriod = GetScanPeriod(currentConfig, scanDateTime);
        scanValues.push([scanDateTime.toLocaleString(),studentName[1],data[key],scanPeriod[0],tripType,scanPeriod[1]]);
      }
    }
  } 
  catch (err)
  {
    console.log('Failed with error %s', err.message);
    return ContentService.createTextOutput("processMultScan Failed: " + err);
  }

  currentValueRange = sheetScanLog.getRange(sheetScanLog.getLastRow()+1,1,scanValues.length,6);
  currentValueRange.setValues(scanValues);

  return ContentService.createTextOutput("processMultScan: " + scanValues);
}

//Process scanned ID and tripType from Hall Pass program
function processTrip(e) 
{
  Logger.log("processTrip(e): " + e.tripType);
  const sheet = SpreadsheetApp.getActive();
  const sheetID = sheet.getId();
  var sheetScanLog = sheet.getSheetByName('Trip Log');
  var sheetConfig = sheet.getSheetByName('Config');
  var currentConfig = sheetConfig.getRange(1,1,sheetConfig.getLastRow(),sheetConfig.getLastColumn()).getValues();
  var scanSIDNO = parseInt(e.parameter.SIDNO);
  var scanDateTime = new Date();

  if(!e.parameter.tripType)
    e.parameter.tripType = "Bathroom";

  var studentName = RosterRecall.getStudentBySIDNO(sheetID, scanSIDNO);
  var scanValues = [];
  var scanPeriod = [];

  //Pull in previous 30 rows of scan history
  var prevScanValues = sheetScanLog.getRange(2, 1, 50, 6).getValues();

  //Loop to cycle through previous scan history and match SIDNO
  for (var r = 0; r<30; r++) {
    if (prevScanValues[r][4] == scanSIDNO) {
      var deltaTime = (Date.parse(scanDateTime.toLocaleString())-Date.parse(prevScanValues[r][0]));
      if (prevScanValues[r][2]=='' && deltaTime<3600000){
        var deltaValues = [['=TIMEVALUE(R[0]C[1])-TIMEVALUE(R[0]C[-2])',scanDateTime.toLocaleString()]];
        var deltaValueRange = sheetScanLog.getRange(2+r,3,1,2);
        deltaValueRange.setValues(deltaValues);
        return ContentService.createTextOutput("Checked In");
      };
      break;
    };
  };

  scanPeriod = GetScanPeriod(currentConfig, scanDateTime);

  //Set and format data export matrix from WebApp
  scanValues.push([scanDateTime.toLocaleString(),studentName[1],"","",scanSIDNO,scanPeriod[0],e.parameter.tripType]);

  //Move existing data down by 1 row and write new values
  sheetScanLog.insertRowBefore(2);
  currentValueRange = sheetScanLog.getRange(2,1,1,7);
  currentValueRange.setValues(scanValues);
  currentValueRange.offset(0, 3).setNumberFormat('hh":"mm":"ss" "am/pm');
  currentValueRange.offset(0, 4).setNumberFormat("0");
  return ContentService.createTextOutput("Checked Out");
}

function addWeek(selectedDate) 
{
  Logger.log("addWeek(" + selectedDate + ")");
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  for (var r = 0; r<=sheets.length; r++) {
    Logger.log("Sheet: " + sheets[r].getName());
    var sheetCell = sheets[r].getRange(1, 2).getValue();
    Logger.log(sheetCell);
    if (sheetCell != "Period") {
      Logger.log("No match, skip");
      return;
    }
    var sheetPeriod = sheets[r].getRange(1,3).getValue();

    const lastRow = sheets[r].getLastRow()-3;
    const lastCol = sheets[r].getLastColumn();
    var formulaValues = [];
    var formulaTemp = [];
    formulaValues.push([selectedDate,"=R[0]C[-1]+1","=R[0]C[-1]+1","=R[0]C[-1]+1","=R[0]C[-1]+1"]);
    formulaValues.push(["=R[-1]C[0]","=R[-1]C[0]","=R[-1]C[0]","=R[-1]C[0]","=R[-1]C[0]"]);
    for (var i=0; i<lastRow; i++) {
      var rowOffset = -2-i;
      var currentRow = i+4;
      var periodOffset = sheetPeriod + 1;
      for (var j=0; j<5; j++) 
      {
        formulaTemp[j] = "=if(DATEVALUE(R["+ rowOffset +"]C[0])<=TODAY(), ifs(IFNA(ROWS(FILTER(\'Attendance Log\'!$A:$F,DATEVALUE(\'Attendance Log\'!$A:$A)=DATEVALUE(R["+ rowOffset +"]C[0]),\'Attendance Log\'!$C:$C=$B"+currentRow+",\'Attendance Log\'!$D:$D=$C$1,\'Attendance Log\'!$F:$F>0)))>0,\"T\",IFNA(ROWS(FILTER(\'Attendance Log\'!$A:$D,DATEVALUE(\'Attendance Log\'!$A:$A)=DATEVALUE(R["+ rowOffset +"]C[0]),\'Attendance Log\'!$C:$C=$B"+currentRow+",\'Attendance Log\'!$D:$D=$C$1)))>0,\"P\", 1,\"U\"), \"\")";
      }
      formulaValues.push([formulaTemp[0],formulaTemp[1],formulaTemp[2],formulaTemp[3],formulaTemp[4]]);
    }

    currentValueRange = sheets[r].getRange(2,lastCol+1,formulaValues.length,5);
    currentValueRange.setValues(formulaValues);

  }

}

function showDateInputDialog() 
{
  var htmlOutput = HtmlService.createHtmlOutputFromFile('DateDialog')
      .setWidth(300)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Enter Date');
}

function processDate(dateString) 
{
  if (dateString) {
    Logger.log("Date String: " + dateString);
    addWeek(dateString);
    } else {
    Logger.log('No date was selected.');
  }
}

function GetScanPeriod(currentConfig, scanDateTime) {
    const dimensions = [ currentConfig.length, currentConfig[0].length ];
    Logger.log(currentConfig);
    var schedule = currentConfig[0][3];
    Logger.log(schedule);

    for (var i = 0; i<dimensions[0]; i++) {
      Logger.log("currentConfig["+ i + "][0]: " + currentConfig[i][0]);
      if (currentConfig[i][0] == schedule) {
        var scheduleStart = i;
        Logger.log("Schedule Found: " + schedule + " (Row " + i + ")");
        break;
      }
    }

    for (var i = scheduleStart+1; i<scheduleStart+9; i++) {
      Logger.log(i + " CurrentTime: " + scanDateTime + " " + scanDateTime.getHours() + ":" + scanDateTime.getMinutes() + " ConfigTime: " + currentConfig[i][3] + " " + currentConfig[i][3].getHours() + ":" + currentConfig[i][3].getMinutes());
      var configTime = currentConfig[i][3].getHours() + currentConfig[i][3].getMinutes()/60;
      var tardyTime = currentConfig[i][4].getHours() + currentConfig[i][4].getMinutes()/60;
      var scanTime = scanDateTime.getHours() + scanDateTime.getMinutes()/60;
      Logger.log("configTime: " + configTime + " tardyTime: " + " scanTime: " + scanTime);
      if (scanTime <= configTime) {
        Logger.log("timehit");
        var tardyDuration = scanTime - tardyTime;
        var scanPeriod = [currentConfig[i][1], tardyDuration];
        Logger.log(scanPeriod);
        break;
      }
    }
  return scanPeriod;
}


function formatAll() {
  var spreadsheet = SpreadsheetApp.getActive();
  //var sheet = spreadsheet.getSheetByName('Roster');
  //var range = sheet.getRange(2, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  //range.sort(1);
  //sheet.sort (1, true);

  var logsheet = spreadsheet.getSheetByName('Trip Log');
  logsheet.getRange('A:A').setNumberFormat('m"/"dd"/"yyyy"  "hh":"mm":"ss" "am/pm');
  logsheet.getRange('B:B').setNumberFormat('@');
  logsheet.getRange('C:C').setNumberFormat('[mm]":"ss');
  logsheet.getRange('D:D').setNumberFormat('hh":"mm":"ss" "am/pm');
  logsheet.getRange('E:E').setNumberFormat('0');
  logsheet.getRange('F:F').setNumberFormat('@');
  logsheet.getRange('G:G').setNumberFormat('@');
};

function TestFunction() {
  var scanDateTime = new Date("05/07/2026 10:41:29 AM");
  const sheet = SpreadsheetApp.getActive();
  var output = GetScanPeriod(scanDateTime, sheet);
  Logger.log(output);
}
