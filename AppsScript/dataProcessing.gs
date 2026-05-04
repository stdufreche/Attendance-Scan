//Process scanned ID and time list from AHK program
function processMultiScan(e) {
  Logger.log("processMultiScan(e)");
  try {
    const sheet = SpreadsheetApp.getActive();
    const sheetID = sheet.getId();
    const tripType = "Attendance";
    const data = e.parameters;
    
    var sheetScanLog = sheet.getSheetByName('Attendance Log');
    var scanValues = [];
    
    for (const key in data) {
      console.log('Key: %s, Value: %s', key, data[key]);

      var scanDateTime = new Date(key);

      var studentName = RosterRecall.getStudentBySIDNO(sheetID, data[key]);

      if (key != "tripType")
        scanValues.push([scanDateTime.toLocaleString(),studentName[1],data[key],"=IFS(TIMEVALUE(R[0]C[-3])<TIMEVALUE(Config!C$2), 1, TIMEVALUE(R[0]C[-3])<TIMEVALUE(Config!C$3), 2, TIMEVALUE(R[0]C[-3])<TIMEVALUE(Config!C$4), 3, TIMEVALUE(R[0]C[-3])<TIMEVALUE(Config!C$5), 4, TIMEVALUE(R[0]C[-3])<TIMEVALUE(Config!C$6), 5, TIMEVALUE(R[0]C[-3])<TIMEVALUE(Config!C$7), 6, TIMEVALUE(R[0]C[-3])<TIMEVALUE(Config!C$8), 7, TIMEVALUE(R[0]C[-3])<TIMEVALUE(Config!C$9), 8, true, 9)",tripType]);
    }
  } 
  catch (err)
  {
    Logger.log('Failed with error %s', err.message);
    return ContentService.createTextOutput("processMultScan Failed: " + err);
  }

  currentValueRange = sheetScanLog.getRange(sheetScanLog.getLastRow()+1,1,scanValues.length,5);
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
  var scanSIDNO = parseInt(e.parameter.SIDNO);
  var scanDateTime = new Date();

  if(!e.parameter.tripType)
    e.parameter.tripType = "Bathroom";

  var studentName = RosterRecall.getStudentBySIDNO(sheetID, scanSIDNO);
  var scanValues = [];

  //Set and format data export matrix from WebAp
  scanValues.push([scanDateTime.toLocaleString(),studentName[1],"","",scanSIDNO,"=IFS(TIMEVALUE(R[0]C[-5])<TIMEVALUE(Config!C$2), 1, TIMEVALUE(R[0]C[-5])<TIMEVALUE(Config!C$3), 2, TIMEVALUE(R[0]C[-5])<TIMEVALUE(Config!C$4), 3, TIMEVALUE(R[0]C[-5])<TIMEVALUE(Config!C$5), 4, TIMEVALUE(R[0]C[-5])<TIMEVALUE(Config!C$6), 5, TIMEVALUE(R[0]C[-5])<TIMEVALUE(Config!C$7), 6, TIMEVALUE(R[0]C[-5])<TIMEVALUE(Config!C$8), 7, true, 8)",e.parameter.tripType]);

  //Pull in previous 30 rows of scan history
  var prevScanValues = sheetScanLog.getRange(2, 1, 50, 6).getValues();

  //Loop to cycle through previous scan history and match SIDNO
  for (var r = 0; r<30; r++) {
    if (prevScanValues[r][4] == scanValues[0][4]) {
      var deltaTime = (Date.parse(scanValues[0][0])-Date.parse(prevScanValues[r][0]));
      if (prevScanValues[r][2]=='' && deltaTime<3600000){
        var deltaValues = [['=TIMEVALUE(R[0]C[1])-TIMEVALUE(R[0]C[-2])',scanValues[0]]];
        var deltaValueRange = sheetScanLog.getRange(2+r,3,1,2);
        deltaValueRange.setValues(deltaValues);
        return ContentService.createTextOutput("Checked In");
      };
      break;
    };
  };

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
        formulaTemp[j] = "=if(DATEVALUE(R["+ rowOffset +"]C[0])<=TODAY(), ifs(IFNA(ROWS(FILTER(\'Attendance Log\'!$A:$D,DATEVALUE(\'Attendance Log\'!$A:$A)=DATEVALUE(R["+ rowOffset +"]C[0]),\'Attendance Log\'!$C:$C=$B"+currentRow+",\'Attendance Log\'!$D:$D=$C$1,TIMEVALUE(\'Attendance Log\'!$A:$A)<TIMEVALUE(Config!$D$"+periodOffset+"))))>0,\"P\",IFNA(ROWS(FILTER(\'Attendance Log\'!$A:$D,DATEVALUE(\'Attendance Log\'!$A:$A)=DATEVALUE(R["+ rowOffset +"]C[0]),\'Attendance Log\'!$C:$C=$B"+currentRow+",\'Attendance Log\'!$D:$D=$C$1,TIMEVALUE(\'Attendance Log\'!$A:$A)>TIMEVALUE(Config!$D$"+periodOffset+"))))>0,\"T\", 1,\"U\"), \"\")";
      }
      formulaValues.push([formulaTemp[0],formulaTemp[1],formulaTemp[2],formulaTemp[3],formulaTemp[4]]);
    }

    currentValueRange = sheets[r].getRange(2,lastCol+1,formulaValues.length,5);
    currentValueRange.setValues(formulaValues);

  }

}

function showDateInputDialog() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('DateDialog')
      .setWidth(300)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Enter Date');
}

function processDate(dateString) {
  if (dateString) {
    Logger.log("Date String: " + dateString);
    addWeek(dateString);
    } else {
    Logger.log('No date was selected.');
  }
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