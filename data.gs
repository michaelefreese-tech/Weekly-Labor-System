// data.gs

function importData() {

  if (!ValidateImport()) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName("IMPORT");
  const coverSheet = ss.getSheetByName("COVER");

  const weekIndicator = importSheet.getRange("X3").getValue().toString().trim();

  const sourceDriverHours = importSheet.getRange("R16").getValue();
  const sourceDriverVariance = importSheet.getRange("S16").getValue();
  const sourceInsiderHours = importSheet.getRange("U16").getValue();
  const sourceInsiderVariance = importSheet.getRange("V16").getValue();

  const destinationMap = {
    "WEEK 1": {
      driverHours: "I8",
      insiderHours: "I10",
      driverVariance: "N8",
      insiderVariance: "N10"
    },
    "WEEK 2": {
      driverHours: "I15",
      insiderHours: "I17",
      driverVariance: "N15",
      insiderVariance: "N17"
    },
    "WEEK 3": {
      driverHours: "I22",
      insiderHours: "I24",
      driverVariance: "N22",
      insiderVariance: "N24"
    },
    "WEEK 4": {
      driverHours: "I29",
      insiderHours: "I31",
      driverVariance: "N29",
      insiderVariance: "N31"
    }
  };

  const target = destinationMap[weekIndicator];

  if (!target) {
    SpreadsheetApp.getUi().alert("Invalid week indicator: " + weekIndicator);
    return;
  }

  coverSheet.getRange(target.driverHours).setValue(sourceDriverHours);
  coverSheet.getRange(target.insiderHours).setValue(sourceInsiderHours);
  coverSheet.getRange(target.driverVariance).setValue(sourceDriverVariance);
  coverSheet.getRange(target.insiderVariance).setValue(sourceInsiderVariance);

  Logger.log("Imported data for " + weekIndicator);
}

// Clears COVER data entry cells for testing
function ClearCoverInputs() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("COVER");
  if (!sheet) return;

  const colors = Info._META?.COLORS || {};
  const white = colors["white"] || "#ffffff";

  const inputCells = [
    "I8", "I10", "N8", "N10", "O9",
    "I15", "I17", "N15", "N17", "O16",
    "I22", "I24", "N22", "N24", "O23",
    "I29", "I31", "N29", "N31", "O30"
  ];

  inputCells.forEach(a1 => {
    sheet.getRange(a1)
      .clearContent()
      .setBackground(white);
  });
}

//----------------------------------------------------------------------------------------------------
// FUNCTIONS FOR SETTING WEEK INDICATOR
//----------------------------------------------------------------------------------------------------

function UpdateWeekLabel(index) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("IMPORT");

  sheet.getRange("X3").setValue("WEEK " + (index + 1));
}


function SetActiveWeek(index) {

  const props = PropertiesService.getDocumentProperties();
  props.setProperty("currentWeekIndex", index.toString());

  UpdateWeekView(index);
  UpdateWeekLabel(index);
}


function SelectWeek1() {
  SetActiveWeek(0);
}


function SelectWeek2() {
  SetActiveWeek(1);
}


function SelectWeek3() {
  SetActiveWeek(2);
}


function SelectWeek4() {
  SetActiveWeek(3);
}
