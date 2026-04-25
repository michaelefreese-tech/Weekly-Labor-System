// main.gs

//----------------------------------------------------------------------------------------------------
// DICTIONARIES AND VARIABLES
//----------------------------------------------------------------------------------------------------

// Main configuration dictionary
const Info = {
  "_META": {
    "REQUIRED_SHEETS":[
      "COVER",
      "IMPORT",
      "INFO"],
    "COLORS" : {
      "week_1": "#c0a9d9",
      "week_2": "#8ed835",
      "week_3": "#f1cb65",
      "week_4": "#4285f4",
      "ptd": "#00ffff",
      "background": "#1155cc",
      "orange": "#FFA500",
      "light blue": "#ADD8E6",
      "red": "#FF0000",
      "yellow": "#FFFF00",
      "purple": "#800080",
      "pink": "#FFC0CB",
      "brown": "#A52A2A",
      "black": "#000000",
      "white": "#ffffff"}},
  "COVER": {
    "STRING": {
      "C3": "WEEKLY LABOR",
      "P3": "Date : ",
      "P4": "Store :",
      "P5": "Period :",
      "D9": "WEEK 1", "G8": "Ideal Driver HRS", "G10": "Ideal inside HRS", "J8": "Total", "L8": "Driver Variance", "L10": "Inside Variance", "O8": "Training hrs", "P8": "Total", "Q8": "labor %",
      "D16": "WEEK 2", "G15": "Ideal Driver HRS", "G17": "Ideal inside HRS", "J15": "Total", "L15": "Driver Variance", "L17": "Inside Variance", "O15": "Training hrs", "P15": "Total", "Q15": "labor %",
      "D23": "WEEK 3", "G22": "Ideal Driver HRS", "G24": "Ideal inside HRS", "J22": "Total", "L22": "Driver Variance", "L24": "Inside Variance", "O22": "Training hrs", "P22": "Total", "Q22": "labor %",
      "D30": "WEEK 4", "G29": "Ideal Driver HRS", "G31": "Ideal inside HRS", "J29": "Total", "L29": "Driver Variance", "L31": "Inside Variance", "O29": "Training hrs", "P29": "Total", "Q29": "labor %",
      "D37": "PTD", "G36": "Ideal Driver HRS", "G39": "Ideal inside HRS", "J36": "Total", "L36": "Driver Variance", "L39": "Inside Variance", "O36": "Training hrs", "P36": "Total", "Q36": "labor %"},
    "FORMULAS": {
      "D8":'=TEXT(INFO!G10,"MM/DD/YYYY") & " - " & TEXT(INFO!I10,"MM/DD/YYYY")',
      "D15": '=TEXT(INFO!G11,"MM/DD/YYYY") & " - " & TEXT(INFO!I11,"MM/DD/YYYY")',
      "D22": '=TEXT(INFO!G12,"MM/DD/YYYY") & " - " & TEXT(INFO!I12,"MM/DD/YYYY")',
      "D29": '=TEXT(INFO!G13,"MM/DD/YYYY") & " - " & TEXT(INFO!I13,"MM/DD/YYYY")',
      "D36": '=TEXT(INFO!G10,"MM/DD/YYYY") & " - " & TEXT(INFO!I13,"MM/DD/YYYY")',
      "J9": "=SUM(I8:I11)", "P9": "=SUM(N8,N10,(O9))", "Q9": "=SUM(P9/J9)",
      "J16": "=SUM(I15:I17)", "P16": "=SUM(N15,N17,(O15))", "Q16": "=SUM(P16/J16)",
      "J23": "=SUM(I22:I24)", "P23": "=SUM(N22,N24,(O23))", "Q23": "=SUM(P23/J23)",
      "J30": "=SUM(I29:I31)", "P30": "=SUM(N29,N31,(O30))", "Q30": "=SUM(P30/J30)", 
      "I36": "=SUM(I8,I15,I22,I29)", "I38": "=SUM(I10,I17,I24,I31)", "J37": "=SUM(I36:I38)", "N36": "=SUM(N8,N15,N22,N29)", "N38": "=SUM(N10,N17,N24,N31)", "O37": "=SUM(O9,O16,O23,O30)", "P37": "=SUM(N36,N38,(O37))", "Q37": "=SUM(P37/J37)"},
    "FORMAT": {},
    "COLORS": {
      "background": ["B2:S41"],
      "week_1" : ["C7:R12"],
      "week_2": ["C14:R19"],
      "week_3": ["C21:R26"],
      "week_4": ["C28:R33"],
      "ptd": ["C35:R40"],
      "yellow": [
        "G8:H11", "G15:H18", "G22:H25", "G29:H32", "G36:H39", 
        "L8:M11", "L15:M18", "L22:M25", "L29:M32", "L36:M39"],
      "white": [
        "C3:L5", "Q3:R5", 
        "D9:E11", "I8:I11", "J9:J11", "N8:N11", "O9:P11",
        "D16:E18", "I15:I18", "J16:J18", "N15:N18", "O16:P18", 
        "D23:E25", "I22:I25", "J23:J25", "N22:N25", "O23:P25", 
        "D30:E32", "I29:I32", "J30:J32", "N29:N32", "O30:P32",
        "D37:E39", "I36:I39", "J37:J39", "N36:N39", "O37:P39"],
      "orange": [
        "D8:E8", "J8", "O8:Q8", 
        "D15:E15", "J15", "O15:Q15", 
        "D22:E22", "J22", "O22:Q22",
        "D29:E29", "J29", "O29:Q29",
        "D36:E36", "J36", "O36:Q36",
        "P3:P5"],
      },
    "VARIABLES": {},
    "OUTLINEMED": [
      "B2:S41", "C3:L5", "P3", "P4", "P5", "Q3", "Q4", "Q5", 
      "C7:R12", "D8:E8", "D9:E11", "G8:H9", "G10:H11", "I8:I9", "I10:I11", "J8", "J9:J11", "L8:M9", "L10:M11", "N8:N9", "N10:N11", "O8", "O9:O11", "P8", "P9:P11", "Q8", "Q9:Q11",
      "C14:R19", "D15:E15", "D16:E18", "G15:H16", "G17:H18", "I15:I16", "I17:I18", "J15", "J16:J18", "L15:M16", "L17:M18", "N15:N16", "N17:N18", "O15", "O16:O18", 
        "P15", "P16:P18", "Q15", "Q16:Q18",
      "C21:R26", "D22:E22", "D23:E25", "G22:H23", "G24:H25", "I22:I23", "I24:I25", "J22", "J23:J25", "L22:M23", "L24:M25", "N22:N23", "N24:N25", "O22", "O23:O25", 
        "P22", "P23:P25", "Q22", "Q23:Q25", 
      "C28:R33", "D29:E29", "D30:E32", "G29:H30", "G31:H32", "I29:I30", "I31:I32", "J29", "J30:J32", "L29:M30", "L31:M32", "N29:N30", "N31:N32", "O29", "O30:O32", 
        "P29", "P30:P32", "Q29", "Q30:Q32", 
      "C35:R40", "D36:E36", "D37:E39", "G36:H37", "G38:H39", "I36:I37", "I38:I39", "J36", "J37:J39", "L36:M37", "L38:M39", "N36:N37", "N38:N39", "O36", "O37:O39", 
        "P36", "P37:P39", "Q36", "Q37:Q39"],
    "OUTLINETHIN": [],
    "MERGE": [],
    "RIGHT": [
      "P3", "P4", "P5"],
    "CENTER": [
      "J8", "O8", "P8", "Q8", 
      "J15", "O15", "P15", "Q15", 
      "J22", "O22", "P22", "Q22",
      "J29", "O29", "P29", "Q29",
      "J36", "O36", "P36", "Q36"],
    "MERGERIGHT": [],
    "MERGELEFT": [],
    "MERGECENTER": [
      "C3:L5", "Q3:R3", "Q4:R4", "Q5:R5", 
      "D8:E8", "D9:E11", "G8:H9", "G10:H11", "I8:I9", "I10:I11", "J9:J11", "L8:M9", "L10:M11", "N8:N9", "N10:N11", "O9:O11", "P9:P11", "Q9:Q11", 
      "D15:E15", "D16:E18", "G15:H16", "G17:H18", "I15:I16", "I17:I18", "J16:J18", "L15:M16", "L17:M18", "N15:N16", "N17:N18", "O16:O18", "P16:P18", "Q16:Q18",
      "D22:E22", "D23:E25", "G22:H23", "G24:H25", "I22:I23", "I24:I25", "J23:J25", "L22:M23", "L24:M25", "N22:N23", "N24:N25", "O23:O25", "P23:P25", "Q23:Q25",
      "D29:E29", "D30:E32", "G29:H30", "G31:H32", "I29:I30", "I31:I32", "J30:J32", "L29:M30", "L31:M32", "N29:N30", "N31:N32", "O30:O32", "P30:P32", "Q30:Q32",
      "D36:E36", "D37:E39", "G36:H37", "G38:H39", "I36:I37", "I38:I39", "J37:J39", "L36:M37", "L38:M39", "N36:N37", "N38:N39", "O37:O39", "P37:P39", "Q37:Q39"],
    "ALLOWEDCELLS": [
      "I8","I10", "N8", "N10", "O9",
      "I15", "I17", "N15", "N17", "O16",
      "I22", "I24", "N22", "N24", "O23",
      "I29", "I31", "N29", "N31", "O30"],
    "DATACELLS": [],},

  "IMPORT": {
    "STRING": {
      "C3": "IMPORT WEEKLY LABOR DETAIL REPORT"},

    "FORMULAS": {},
    "FORMAT": {},
    "COLORS": {
      "background": ["B2:AF20"],
      "white": ["C3:V3", "X3:AE7", "C9:AE19"]},

    "VARIABLES": {
      "WEEKINDICATOR": "X3"},
    "OUTLINEMED": [
        "B2:AF20", "C3:V3", "X3:AE7", "C9:AE19", "R16", "S16", "U16", "V16"],

    "OUTLINETHIN": [],
    "MERGE": [],
    "RIGHT": [],
    "MERGERIGHT": [],
    "MERGELEFT": [],
    "MERGECENTER": [
      "C3:V3", "X3:AE7"],
    "ALLOWEDCELLS": [
      "C9", "D9", "E9", "F9", "G9", "H9", "I9", "J9", "K9", "L9", "M9", "N9", "O9", "P9", "Q9", "R9", "S9", "T9", "U9", "V9", "W9", "X9", "Y9", "Z9", "AA9", "AB9", "AC9", "AD9", "AE9",
      "C10", "D10", "E10", "F10", "G10", "H10", "I10", "J10", "K10", "L10", "M10", "N10", "O10", "P10", "Q10", "R10", "S10", "T10", "U10", "V10", "W10", "X10", "Y10", "Z10", "AA10", "AB10", "AC10", "AD10", "AE10",
      "C11", "D11", "E11", "F11", "G11", "H11", "I11", "J11", "K11", "L11", "M11", "N11", "O11", "P11", "Q11", "R11", "S11", "T11", "U11", "V11", "W11", "X11", "Y11", "Z11", "AA11", "AB11", "AC11", "AD11", "AE11",
      "C12", "D12", "E12", "F12", "G12", "H12", "I12", "J12", "K12", "L12", "M12", "N12", "O12", "P12", "Q12", "R12", "S12", "T12", "U12", "V12", "W12", "X12", "Y12", "Z12", "AA12", "AB12", "AC12", "AD12", "AE12",
      "C13", "D13", "E13", "F13", "G13", "H13", "I13", "J13", "K13", "L13", "M13", "N13", "O13", "P13", "Q13", "R13", "S13", "T13", "U13", "V13", "W13", "X13", "Y13", "Z13", "AA13", "AB13", "AC13", "AD13", "AE13",
      "C14", "D14", "E14", "F14", "G14", "H14", "I14", "J14", "K14", "L14", "M14", "N14", "O14", "P14", "Q14", "R14", "S14", "T14", "U14", "V14", "W14", "X14", "Y14", "Z14", "AA14", "AB14", "AC14", "AD14", "AE14",
      "C15", "D15", "E15", "F15", "G15", "H15", "I15", "J15", "K15", "L15", "M15", "N15", "O15", "P15", "Q15", "R15", "S15", "T15", "U15", "V15", "W15", "X15", "Y15", "Z15", "AA15", "AB15", "AC15", "AD15", "AE15",
      "C16", "D16", "E16", "F16", "G16", "H16", "I16", "J16", "K16", "L16", "M16", "N16", "O16", "P16", "Q16", "R16", "S16", "T16", "U16", "V16", "W16", "X16", "Y16", "Z16", "AA16", "AB16", "AC16", "AD16", "AE16",
      "C17", "D17", "E17", "F17", "G17", "H17", "I17", "J17", "K17", "L17", "M17", "N17", "O17", "P17", "Q17", "R17", "S17", "T17", "U17", "V17", "W17", "X17", "Y17", "Z17", "AA17", "AB17", "AC17", "AD17", "AE17",
      "C18", "D18", "E18", "F18", "G18", "H18", "I18", "J18", "K18", "L18", "M18", "N18", "O18", "P18", "Q18", "R18", "S18", "T18", "U18", "V18", "W18", "X18", "Y18", "Z18", "AA18", "AB18", "AC18", "AD18", "AE18",
      "C19", "D19", "E19", "F19", "G19", "H19", "I19", "J19", "K19", "L19", "M19", "N19", "O19", "P19", "Q19", "R19", "S19", "T19", "U19", "V19", "W19", "X19", "Y19", "Z19", "AA19", "AB19", "AC19", "AD19", "AE19"],
    "DATACELLS": []},

  "INFO": {
    "STRING": {
      "C3": "CONFIGURATION",
      "C6": "STORE LIST:",
      "E6": "Start of Period",
      "G6": "Period number:",
      "I6": "Labor %: ",
      "E9": "Week Ranges",
      "E15": "Current Week:"},
    "FORMULAS": {},
    "COLORS": {
      "background": ["B2:J19"],
      "orange": ["C3:I4", "C6", "E6", "G6", "I6", "E9:I9", "E15:F15"],
      "white": ["C7:C18", "E7", "G7", "I7", "E10:I13", "G15:I15"]},
    "VARIABLES": {},
    "OUTLINEMED": [
      "B2:J19", "C3:I4", "C6", "E6", "G6", "I6", "E7", "G7", "I7", 
      "C7:C18", "E9:I9", "E10:I13", "E15:F15", "G15:I15"],
    "OUTLINETHIN": [],
    "MERGE": [],
    "RIGHT": [],
    "CENTER": ["C6", "E6", "G6", "I6"],
    "MERGERIGHT": ["E15:F15"],
    "MERGELEFT": [],
    "MERGECENTER": ["C3:I4", "E9:I9"],
    "ALLOWEDCELLS": ["C7", "C8", "C9", "C10", "C11", "C12", "C13", "C14", "C15", "C16", "C17", "C18", "I7"],
    "DATACELLS": []}
};

const TempCellBackup = {};

//----------------------------------------------------------------------------------------------------
// FUNCTIONS FOR OPENING THE FILE
//----------------------------------------------------------------------------------------------------

// Runs when the spreadsheet is opened
function onOpen() {

  BuildMenus();

  if (!IsFileInitialized()) {
    BuildSetupMenu();
  }

  // Rebuild file state
  GetWeekInfo();
}


function BuildMenus() {

  const props = PropertiesService.getDocumentProperties();
  const isLocked = props.getProperty("docLocked") === "true";

  const security = SpreadsheetApp.getUi().createMenu("🔒 Security");

  if (isLocked) {
    security.addItem("🔓 Unlock Document", "Unlock");
  } else {
    security.addItem("🔐 Lock Document", "Lock");
  }

  security
    .addSeparator()
    .addItem("🛠 Set/Change Password", "SetupPassword")
    .addItem("❓ Recover Password", "RecoverPassword")
    .addToUi();

  const tools = SpreadsheetApp.getUi().createMenu("🛠 Script Tools");

  if (!isLocked) {
    tools
      .addItem("Set New Version", "PromptNewVersion")
      .addSeparator();
  }

  tools
    .addItem("Show Current Version", "ShowVersion")
    .addItem("View Changelog", "ViewChangeLog")
    .addToUi();
}


function IsFileInitialized() {

  const props = PropertiesService.getDocumentProperties();
  return props.getProperty("FILE_INITIALIZED") === "true";
}


function BuildSetupMenu() {

  SpreadsheetApp.getUi().createMenu("⚙️ Setup")
    .addItem("Initialize This File", "InitializeThisFile")
    .addToUi();
}


function InitializeThisFile() {

  EnsureRequiredSheetsExist();

  SetupPassword();
  ApplySheetProtections();
  InstallChangeTrigger();

  GetWeekInfo();
  Restore();

  const props = PropertiesService.getDocumentProperties();
  props.setProperty("FILE_INITIALIZED", "true");

  SpreadsheetApp.getUi().alert("Setup complete.");
}


function EnsureRequiredSheetsExist() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = Info._META?.REQUIRED_SHEETS || [];

  requiredSheets.forEach(sheetName => {
    if (!ss.getSheetByName(sheetName)) {
      ss.insertSheet(sheetName);
    }
  });

}

