// validation.gs

function IsBlank(value) {
  return value === "" || value === null || value === undefined;
}

function IsCoverValidationEdit(range) {

  const row = range.getRow();
  const col = range.getColumn();

  const validationBoxes = [
    { col: 9, rows: [8, 9] },
    { col: 9, rows: [10, 11] },
    { col: 14, rows: [8, 9] },
    { col: 14, rows: [10, 11] },

    { col: 9, rows: [15, 16] },
    { col: 9, rows: [17, 18] },
    { col: 14, rows: [15, 16] },
    { col: 14, rows: [17, 18] },

    { col: 9, rows: [22, 23] },
    { col: 9, rows: [24, 25] },
    { col: 14, rows: [22, 23] },
    { col: 14, rows: [24, 25] },

    { col: 9, rows: [29, 30] },
    { col: 9, rows: [31, 32] },
    { col: 14, rows: [29, 30] },
    { col: 14, rows: [31, 32] }
  ];

  return validationBoxes.some(box =>
    col === box.col &&
    row >= box.rows[0] &&
    row <= box.rows[1]
  );
}

function ValidateCoverInputs() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COVER");
  if (!sheet) return;

  const colors = Info._META?.COLORS || {};
  const white = colors["white"] || "#ffffff";
  const red = colors["red"] || "#ff0000";

  const weekGroups = [
    [
      { valueCell: "I8", range: "I8:I9" },
      { valueCell: "I10", range: "I10:I11" },
      { valueCell: "N8", range: "N8:N9" },
      { valueCell: "N10", range: "N10:N11" }
    ],
    [
      { valueCell: "I15", range: "I15:I16" },
      { valueCell: "I17", range: "I17:I18" },
      { valueCell: "N15", range: "N15:N16" },
      { valueCell: "N17", range: "N17:N18" }
    ],
    [
      { valueCell: "I22", range: "I22:I23" },
      { valueCell: "I24", range: "I24:I25" },
      { valueCell: "N22", range: "N22:N23" },
      { valueCell: "N24", range: "N24:N25" }
    ],
    [
      { valueCell: "I29", range: "I29:I30" },
      { valueCell: "I31", range: "I31:I32" },
      { valueCell: "N29", range: "N29:N30" },
      { valueCell: "N31", range: "N31:N32" }
    ]
  ];

  weekGroups.forEach(group => {

    const values = group.map(item => sheet.getRange(item.valueCell).getValue());

    const started = values.some(value => !IsBlank(value));
    const complete = values.every(value => !IsBlank(value));

    group.forEach((item, index) => {

      if (!started || complete) {
        sheet.getRange(item.range).setBackground(white);
        return;
      }

      if (IsBlank(values[index])) {
        sheet.getRange(item.range).setBackground(red);
      } else {
        sheet.getRange(item.range).setBackground(white);
      }

    });

  });
}

function ValidateImport() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("IMPORT");
  if (!sheet) return false;

  const colors = Info._META?.COLORS || {};
  const white = colors["white"] || "#ffffff";
  const red = colors["red"] || "#ff0000";

  const indicatorCell = Info["IMPORT"]?.VARIABLES?.WEEKINDICATOR || "X3";

  const requiredFields = [
    { cell: "R16", label: "Driver Ideal Hours" },
    { cell: "S16", label: "Driver Variance" },
    { cell: "U16", label: "Insider Ideal Hours" },
    { cell: "V16", label: "Insider Variance" }
  ];

  // Reset colors
  sheet.getRange(indicatorCell).setBackground(white);
  requiredFields.forEach(f => sheet.getRange(f.cell).setBackground(white));

  const errors = [];

  const weekIndicator = sheet.getRange(indicatorCell).getValue().toString().trim();

  if (!weekIndicator) {
    sheet.getRange(indicatorCell).setBackground(red);
    errors.push("Week Indicator");
  }

  requiredFields.forEach(field => {

    const cell = sheet.getRange(field.cell);
    const value = cell.getValue();

    if (IsBlank(value)) {
      cell.setBackground(red);
      errors.push(field.label);
      return;
    }

    if (isNaN(Number(value))) {
      cell.setBackground(red);
      errors.push(field.label + " must be a number");
    }

  });

  if (errors.length > 0) {
    SpreadsheetApp.getUi().alert(
      "⛔ Import blocked.\n\nMissing or invalid fields:\n" + errors.join("\n")
    );
    return false;
  }

  return true;
}
