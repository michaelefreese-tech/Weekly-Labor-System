// editTrigger.gs

function onEdit(e) {

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const range = e.range;

  // COVER validation listener runs even if file is unlocked
  if (sheetName === "COVER") {
    if (IsCoverValidationEdit(range)) {
      ValidateCoverInputs();
    }
  }

  const props = PropertiesService.getDocumentProperties();
  const isLocked = props.getProperty("docLocked") === "true";

  // If document is not locked, allow edit after validation listener
  if (!isLocked) return;

  let editedCell = range.getA1Notation();

  if (range.isPartOfMerge()) {
    const mergedRanges = range.getMergedRanges();
    if (mergedRanges.length > 0) {
      editedCell = mergedRanges[0].getCell(1, 1).getA1Notation();
    }
  }

  // ================================
  // SHEET: COVER
  // ================================
  if (sheetName === "COVER") {

    const allowedCells = Info["COVER"]?.ALLOWEDCELLS || [];

    if (!allowedCells.includes(editedCell)) {
      SpreadsheetApp.getUi().alert("⛔ You don't have permission to edit this COVER cell: " + editedCell);
      Restore();
      return;
    }

    return;
  }

  // ================================
  // SHEET: IMPORT
  // ================================
  if (sheetName === "IMPORT") {

    const allowedCells = Info["IMPORT"]?.ALLOWEDCELLS || [];

    if (!allowedCells.includes(editedCell)) {
      SpreadsheetApp.getUi().alert("⛔ You don't have permission to edit this IMPORT cell: " + editedCell);
      Restore();
      return;
    }

    return;
  }

  // ================================
  // SHEET: INFO
  // ================================
  if (sheetName === "INFO") {

    const allowedCells = Info["INFO"]?.ALLOWEDCELLS || [];

    if (!allowedCells.includes(editedCell)) {
      SpreadsheetApp.getUi().alert("⛔ You don't have permission to edit this INFO cell: " + editedCell);
      Restore();
      GetWeekInfo();
      return;
    }

    if (editedCell === "E7" || editedCell === "G7") {
      GetWeekInfo();
      Restore();
      return;
    }

    return;
  }

  SpreadsheetApp.getUi().alert("⛔ This sheet is locked.");
  Restore();
  GetWeekInfo();
}

// Structural trigger — restore missing required sheets
function onChange(e) {

  const props = PropertiesService.getDocumentProperties();
  const isLocked = props.getProperty("docLocked") === "true";
  if (!isLocked) return;

  EnsureRequiredSheetsExist();

  Restore();
  GetWeekInfo();
  ApplySheetProtections();
}

function InstallChangeTrigger() {

  const ss = SpreadsheetApp.getActive();

  const existing = ScriptApp.getProjectTriggers();
  existing.forEach(t => {
    if (t.getHandlerFunction() === "onChange") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("onChange")
    .forSpreadsheet(ss)
    .onChange()
    .create();
}

function EmergencyUnlock() {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty("docLocked", "false");
}
