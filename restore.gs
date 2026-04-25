// restore.gs

function Restore() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const ColorDict = Info._META?.COLORS || {};

  sheets.forEach(sheet => {

    const sheetName = sheet.getName();
    const sheetInfo = Info[sheetName];
    if (!sheetInfo) return;

    const allMerges = [
      ...(sheetInfo.MERGE || []),
      ...(sheetInfo.MERGERIGHT || []),
      ...(sheetInfo.MERGELEFT || []),
      ...(sheetInfo.MERGECENTER || [])
    ];

    // Use FULL sheet range, not getDataRange()
    // getDataRange() can cut through merged ranges and cause breakApart errors
    const fullRange = sheet.getRange(
      1,
      1,
      sheet.getMaxRows(),
      sheet.getMaxColumns()
    );

    // Break all merges safely
    try {
      fullRange.breakApart();
    } catch (err) {
      Logger.log("Full breakApart failed on " + sheetName + ": " + err);
    }

    // // Reset base style
    // fullRange
    //   .clearFormat()
    //   .setFontWeight("normal")
    //   .setFontStyle("normal")
    //   .setFontSize(10)
    //   .setHorizontalAlignment("left")
    //   .setVerticalAlignment("middle");

    // Reapply merges
    allMerges.forEach(rangeStr => {
      try {
        sheet.getRange(rangeStr).merge();
      } catch (err) {
        Logger.log("Merge failed " + sheetName + "!" + rangeStr + ": " + err);
      }
    });

    // Alignment
    (sheetInfo.RIGHT || []).forEach(r => {
      sheet.getRange(r).setHorizontalAlignment("right");
    });

    (sheetInfo.CENTER || []).forEach(r => {
      sheet.getRange(r).setHorizontalAlignment("center");
    });

    (sheetInfo.MERGECENTER || []).forEach(r => {
      sheet.getRange(r).setHorizontalAlignment("center");
    });

    // Colors
    if (sheetInfo.COLORS) {
      for (const key in sheetInfo.COLORS) {
        const hex = ColorDict[key.toLowerCase()];
        if (!hex) continue;

        sheetInfo.COLORS[key].forEach(r => {
          sheet.getRange(r).setBackground(hex);
        });
      }
    }

    // Borders
    (sheetInfo.OUTLINETHIN || []).forEach(r => {
      sheet.getRange(r).setBorder(
        true, true, true, true, false, false,
        "black",
        SpreadsheetApp.BorderStyle.SOLID
      );
    });

    (sheetInfo.OUTLINEMED || []).forEach(r => {
      sheet.getRange(r).setBorder(
        true, true, true, true, false, false,
        "black",
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      );
    });

    // Strings
    if (sheetInfo.STRING) {
      for (const cell in sheetInfo.STRING) {
        sheet.getRange(cell)
          .setValue(sheetInfo.STRING[cell])
          .setFontFamily("Arial");
      }
    }

    // Formulas
    if (sheetInfo.FORMULAS) {
      for (const cell in sheetInfo.FORMULAS) {
        sheet.getRange(cell).setFormula(sheetInfo.FORMULAS[cell]);
      }
    }

    // Number formats
    if (sheetInfo.FORMAT) {
      for (const type in sheetInfo.FORMAT) {

        let formatString = null;

        if (type === "currency") formatString = "$#,##0.00";
        if (type === "currency_round") formatString = "$#,##0";
        if (type === "number") formatString = "0.00";
        if (type === "date") formatString = "MM/dd/yyyy";
        if (type === "datetime") formatString = "MM/dd/yyyy hh:mm AM/PM";

        if (!formatString) continue;

        sheetInfo.FORMAT[type].forEach(r => {
          sheet.getRange(r).setNumberFormat(formatString);
        });
      }
    }

  });
}
