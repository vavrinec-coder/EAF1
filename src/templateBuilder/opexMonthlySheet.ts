/* global Excel */

import { MonthlySheetSpec } from "./monthlySheet";

const DEFAULT_SHEET_RANGE = "A1:ZZ200";
const COLUMN_HIDE_LIMIT = 200;
const DEFAULT_TIMELINE_COLUMN_WIDTH = 12;
const BASE_CONSTANTS_COLUMNS = 4;
const MAX_EXCEL_COLUMNS = 16384;
const MAX_EXCEL_ROWS = 1048576;

export async function createOpexMonthlySheet(
  spec: MonthlySheetSpec,
  lineItemsAddress: string
): Promise<void> {
  let forecastRowCount = 0;

  await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    const controlsSheet = worksheets.getItemOrNullObject("Controls");
    controlsSheet.load("name");
    let sheet = worksheets.getItemOrNullObject("Opex Monthly");
    sheet.load("name");
    await context.sync();

    if (controlsSheet.isNullObject) {
      throw new Error('Controls sheet not found. Create the "Controls" sheet first.');
    }

    const controlsHeaderCell = controlsSheet.getRange("A7");
    controlsHeaderCell.load("format/fill/color");
    await context.sync();

    if (sheet.isNullObject) {
      sheet = worksheets.add("Opex Monthly");
    } else {
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("address");
      await context.sync();
      if (!usedRange.isNullObject) {
        usedRange.clear();
      }
    }

    const totalModelColumns = spec.constantsColumns + spec.timelineColumns;
    const unitColumnIndex = Math.max(0, spec.constantsColumns - 2);
    const unitColumnLetter = columnIndexToLetters(unitColumnIndex);
    sheet.tabColor = spec.tabColor;
    sheet.showGridlines = false;

    const baseRange = sheet.getRange(DEFAULT_SHEET_RANGE);
    baseRange.format.font.name = spec.font.name;
    baseRange.format.font.size = spec.font.size;
    baseRange.format.font.color = spec.font.color;

    applyColumnWidths(sheet, spec.columnWidths, spec.constantsColumns);
    applyTimelineColumnWidths(sheet, spec.constantsColumns, spec.timelineColumns);
    applyOpexSpecificColumnWidths(sheet);

    const visibleColumnCount = Math.min(totalModelColumns, COLUMN_HIDE_LIMIT);
    if (visibleColumnCount > 0) {
      const visibleRange = sheet.getRangeByIndexes(0, 0, 1, visibleColumnCount);
      visibleRange.format.columnHidden = false;
    }

    if (totalModelColumns + 1 <= COLUMN_HIDE_LIMIT) {
      const hiddenStart = totalModelColumns + 1;
      const hiddenCount = COLUMN_HIDE_LIMIT - totalModelColumns;
      const hiddenRange = sheet.getRangeByIndexes(0, hiddenStart - 1, 1, hiddenCount);
      hiddenRange.format.columnHidden = true;
    }

    const headerRange = sheet.getRangeByIndexes(0, 0, spec.headerRows, totalModelColumns);
    headerRange.format.fill.color = spec.headerFillColor;
    headerRange.format.font.color = "#FFFFFF";

    const sectionRange = sheet.getRangeByIndexes(6, 0, 1, totalModelColumns);
    sectionRange.format.fill.color = spec.sectionColor;
    sheet.getRange("A7").values = [["MODEL FLAGS"]];
    sheet.getRange("A7").format.font.color = "#FFFFFF";
    sheet.getRange("B7").values = [[""]];

    sheet.getRange("C1:C5").formulas = [
      ["=Controls!B20"],
      ["=Controls!B21"],
      ["=Controls!B22"],
      ["=Controls!B23"],
      ["=Controls!B24"],
    ];

    const timelineStartColIndex = spec.constantsColumns;
    const timelineFormulaRangeStart = sheet.getRangeByIndexes(
      0,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeEnd = sheet.getRangeByIndexes(
      1,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeType = sheet.getRangeByIndexes(
      2,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeCounter = sheet.getRangeByIndexes(
      3,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeYear = sheet.getRangeByIndexes(
      4,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );

    const controlsTimelineStart = controlsSheet.getRangeByIndexes(
      19,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const controlsTimelineEnd = controlsSheet.getRangeByIndexes(
      20,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    controlsTimelineStart.load("numberFormat");
    controlsTimelineEnd.load("numberFormat");

    const startDateFormulas: string[] = [];
    const endDateFormulas: string[] = [];
    const periodTypeFormulas: string[] = [];
    const periodCounterFormulas: string[] = [];
    const financialYearFormulas: string[] = [];

    for (let i = 0; i < spec.timelineColumns; i += 1) {
      const columnIndex = timelineStartColIndex + i;
      const columnLetter = columnIndexToLetters(columnIndex);
      startDateFormulas.push(`=Controls!${columnLetter}20`);
      endDateFormulas.push(`=Controls!${columnLetter}21`);
      periodTypeFormulas.push(`=Controls!${columnLetter}22`);
      periodCounterFormulas.push(`=Controls!${columnLetter}23`);
      financialYearFormulas.push(`=Controls!${columnLetter}24`);
    }

    timelineFormulaRangeStart.formulas = [startDateFormulas];
    timelineFormulaRangeEnd.formulas = [endDateFormulas];
    timelineFormulaRangeType.formulas = [periodTypeFormulas];
    timelineFormulaRangeCounter.formulas = [periodCounterFormulas];
    timelineFormulaRangeYear.formulas = [financialYearFormulas];

    sheet.getRangeByIndexes(0, timelineStartColIndex, 5, spec.timelineColumns).format.horizontalAlignment =
      "Right";

    await context.sync();
    timelineFormulaRangeStart.numberFormat = controlsTimelineStart.numberFormat;
    timelineFormulaRangeEnd.numberFormat = controlsTimelineEnd.numberFormat;

    const controlsFlagLabelRange = sheet.getRange("B9:B21");
    const controlsFlagUnitRange = sheet.getRange(`${unitColumnLetter}9:${unitColumnLetter}21`);
    const controlsFlagTimelineRange = sheet.getRangeByIndexes(
      8,
      timelineStartColIndex,
      13,
      spec.timelineColumns
    );

    const flagLabelFormulas: string[][] = [];
    const flagUnitFormulas: string[][] = [];
    const flagTimelineFormulas: string[][] = [];
    for (let rowOffset = 0; rowOffset < 13; rowOffset += 1) {
      const controlsRow = 44 + rowOffset;
      flagLabelFormulas.push([`=Controls!B${controlsRow}`]);
      flagUnitFormulas.push([`=Controls!${unitColumnLetter}${controlsRow}`]);

      const rowFormulas: string[] = [];
      for (let columnOffset = 0; columnOffset < spec.timelineColumns; columnOffset += 1) {
        const columnIndex = timelineStartColIndex + columnOffset;
        const columnLetter = columnIndexToLetters(columnIndex);
        rowFormulas.push(`=Controls!${columnLetter}${controlsRow}`);
      }
      flagTimelineFormulas.push(rowFormulas);
    }

    controlsFlagLabelRange.formulas = flagLabelFormulas;
    controlsFlagUnitRange.formulas = flagUnitFormulas;
    controlsFlagTimelineRange.formulas = flagTimelineFormulas;
    controlsFlagUnitRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;

    const variablesHeaderRange = sheet.getRangeByIndexes(23, 0, 1, totalModelColumns);
    variablesHeaderRange.format.fill.color = spec.sectionColor;
    sheet.getRange("A24").values = [["VARIABLES FOR OPEX FORECASTING"]];
    sheet.getRange("A24").format.font.color = "#FFFFFF";

    const normalizedLineItemsAddress = normalizeRangeAddress(lineItemsAddress);
    if (!normalizedLineItemsAddress) {
      throw new Error("Opex line items range is required.");
    }

    const lineItemsRange = getRangeFromAddress(context, normalizedLineItemsAddress);
    lineItemsRange.load(["rowCount", "columnCount"]);
    await context.sync();

    forecastRowCount = lineItemsRange.rowCount;
    const lineItemsEndRow =
      lineItemsRange.rowCount > 0 ? 57 + lineItemsRange.rowCount : 57;
    const opexAccountSource =
      lineItemsRange.rowCount > 0 ? `=$B$58:$B$${lineItemsEndRow}` : "";

    const variablesItems = [
      ["Revenue"],
      ["New Revenue"],
      ["Payroll"],
      ["New Payroll"],
      ["Headcount"],
      ["New Headcount"],
      ["Number of Customers"],
      ["Number of New Customers"],
      ["Placeholder"],
      ["Placeholder"],
      ["Placeholder"],
    ];
    sheet.getRangeByIndexes(25, 1, variablesItems.length, 1).values = variablesItems;

    const linkedHeaderRange = sheet.getRangeByIndexes(38, 0, 1, totalModelColumns);
    linkedHeaderRange.format.fill.color = spec.sectionColor;
    sheet.getRange("A39").values = [["OPEX LINKED CALCULATIONS"]];
    sheet.getRange("A39").format.font.color = "#FFFFFF";

    sheet.getRange("B41").values = [["Description"]];
    sheet.getRange("B41").format.font.bold = true;
    sheet.getRange("G41").values = [["Map to Opex account:"]];
    sheet.getRange("G41").format.font.bold = true;
    sheet.getRange("G41").format.font.color = "#000000";
    sheet.getRange("G41").format.fill.clear();
    sheet.getRange("G41").format.borders.getItem(Excel.BorderIndex.edgeTop).style =
      Excel.BorderLineStyle.none;
    sheet.getRange("G41").format.borders.getItem(Excel.BorderIndex.edgeBottom).style =
      Excel.BorderLineStyle.none;
    sheet.getRange("G41").format.borders.getItem(Excel.BorderIndex.edgeLeft).style =
      Excel.BorderLineStyle.none;
    sheet.getRange("G41").format.borders.getItem(Excel.BorderIndex.edgeRight).style =
      Excel.BorderLineStyle.none;

    const linkedPlaceholders = Array.from({ length: 10 }, () => ["Placeholder"]);
    sheet.getRangeByIndexes(41, 1, linkedPlaceholders.length, 1).values = linkedPlaceholders;

    const forecastHeaderRange = sheet.getRangeByIndexes(53, 0, 1, totalModelColumns);
    forecastHeaderRange.format.fill.color = spec.sectionColor;
    sheet.getRange("A54").values = [["OPEX FORECAST"]];
    sheet.getRange("A54").format.font.color = "#FFFFFF";
    sheet.getRange("P54:Q54").clear(Excel.ClearApplyTo.contents);
    sheet.getRange("G57").values = [["Driver"]];
    sheet.getRange("G57").format.font.bold = true;
    const opexHeaderRange = sheet.getRange("H57:O57");
    opexHeaderRange.values = [
      ["Month(s)", "Y1", "Y2", "Y3+", "$ per Month", "ID", "StartDate", "EndDate"],
    ];
    opexHeaderRange.format.font.bold = true;
    opexHeaderRange.format.horizontalAlignment = Excel.HorizontalAlignment.right;
    sheet.getRange("H57:Q57").format.horizontalAlignment = Excel.HorizontalAlignment.right;
    sheet.getRange("H5:L5").formulas = [["=H57", "=I57", "=J57", "=K57", "=L57"]];
    sheet.getRange("I56").values = [["Annual % change / % of"]];
    sheet.getRange("I56").format.font.bold = true;

    if (lineItemsRange.rowCount > 0 && lineItemsRange.columnCount > 0) {
      const targetRange = sheet.getRangeByIndexes(
        57,
        1,
        lineItemsRange.rowCount,
        lineItemsRange.columnCount
      );
      targetRange.copyFrom(lineItemsRange, Excel.RangeCopyType.all, false, false);
    }

    if (lineItemsRange.rowCount > 0) {
      const driverRange = sheet.getRangeByIndexes(57, 6, lineItemsRange.rowCount, 1);
      driverRange.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: "='Controls'!$B$74:$B$90",
        },
      };
      const defaultDrivers = Array.from({ length: lineItemsRange.rowCount }, () => ["No Forecast"]);
      driverRange.values = defaultDrivers;
      driverRange.format.font.color = "#3333FF";
      driverRange.format.fill.color = "#FFFFAB";
      applyHairlineBorders(driverRange);

      if (opexAccountSource) {
        const mapRange = sheet.getRange("G42:G51");
        mapRange.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: opexAccountSource,
          },
        };
        mapRange.format.font.color = "#3333FF";
        mapRange.format.fill.color = "#FFFFAB";
        applyHairlineBorders(mapRange);
      }

      const defaultRows = Array.from({ length: lineItemsRange.rowCount }, () => [0]);
      const applyForecastColumn = (
        columnIndex: number,
        numberFormat: string,
        fontColor?: string
      ) => {
        const range = sheet.getRangeByIndexes(57, columnIndex, lineItemsRange.rowCount, 1);
        range.values = defaultRows;
        range.numberFormat = Array.from({ length: lineItemsRange.rowCount }, () => [numberFormat]);
        if (fontColor) {
          range.format.font.color = fontColor;
        }
        applyHairlineBorders(range);
      };

      applyForecastColumn(7, '#,##0;[Red]-#,##0;"-"', "#3333FF");
      applyForecastColumn(8, '0.0%;[Red]-0.0%;"-"', "#3333FF");
      applyForecastColumn(9, '0.0%;[Red]-0.0%;"-"', "#3333FF");
      applyForecastColumn(10, '0.0%;[Red]-0.0%;"-"', "#3333FF");
      applyForecastColumn(11, '#,##0;[Red]-#,##0;"-"', "#3333FF");
      applyForecastColumn(12, '#,##0;[Red]-#,##0;"-"');
      applyForecastColumn(13, '[$-en-US]d/mmm/yy;[$-en-US]d/mmm/yy;"-";@');
      applyForecastColumn(14, '[$-en-US]d/mmm/yy;[$-en-US]d/mmm/yy;"-";@');

      const rowOffsetBase = 58;
      const matchFormulas = Array.from({ length: lineItemsRange.rowCount }, (_, index) => {
        const rowNumber = rowOffsetBase + index;
        return [`=IFNA(MATCH(G${rowNumber},Controls!$B$74:$B$90,0),0)`];
      });
      const matchRangeM = sheet.getRangeByIndexes(57, 12, lineItemsRange.rowCount, 1);
      matchRangeM.formulas = matchFormulas;

      const matchRangeN = sheet.getRangeByIndexes(57, 13, lineItemsRange.rowCount, 1);
      const nFormulas = Array.from({ length: lineItemsRange.rowCount }, (_, index) => {
        const rowNumber = rowOffsetBase + index;
        return [
          `=IFERROR(IF(M${rowNumber}=9,EOMONTH(MAX(Controls!$C$9,EDATE(Controls!$C$11,-$H${rowNumber})),0),0),0)`,
        ];
      });
      matchRangeN.formulas = nFormulas;

      const opexDateFormulas = Array.from({ length: lineItemsRange.rowCount }, (_, index) => {
        const rowNumber = rowOffsetBase + index;
        return [`=IFERROR(IF(M${rowNumber}=9,Controls!$C$10,0),0)`];
      });
      const matchRangeO = sheet.getRangeByIndexes(57, 14, lineItemsRange.rowCount, 1);
      matchRangeO.formulas = opexDateFormulas;

    }

    const detailsHeaderRow = lineItemsEndRow + 3;
    const detailsHeaderRange = sheet.getRangeByIndexes(
      detailsHeaderRow - 1,
      0,
      1,
      totalModelColumns
    );
    detailsHeaderRange.format.fill.color = spec.sectionColor;
    sheet.getRange(`A${detailsHeaderRow}`).values = [["LINE ITEM DETAILS / VENDORS"]];
    sheet.getRange(`A${detailsHeaderRow}`).format.font.color = "#FFFFFF";

    const detailLabelsRow = detailsHeaderRow + 2;
    sheet.getRange(`B${detailLabelsRow}`).values = [["Line Item Detail / Vendor"]];
    sheet.getRange(`B${detailLabelsRow}`).format.font.bold = true;
    sheet.getRange(`G${detailLabelsRow}`).values = [["Map to Opex account:"]];
    sheet.getRange(`G${detailLabelsRow}`).format.font.bold = true;

    if (lineItemsRange.rowCount > 0 && opexAccountSource) {
      const detailPlaceholderStartRow = detailLabelsRow + 1;
      const detailPlaceholderCount = 100;
      const detailPlaceholders = Array.from({ length: detailPlaceholderCount }, () => ["PLACEHOLDER"]);
      sheet
        .getRangeByIndexes(detailPlaceholderStartRow - 1, 1, detailPlaceholderCount, 1)
        .values = detailPlaceholders;

      const detailMapRange = sheet.getRangeByIndexes(
        detailPlaceholderStartRow - 1,
        6,
        detailPlaceholderCount,
        1
      );
      detailMapRange.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: opexAccountSource,
        },
      };
      detailMapRange.format.font.color = "#3333FF";
      detailMapRange.format.fill.color = "#FFFFAB";
      applyHairlineBorders(detailMapRange);
    }

    const controlsHeaderRange = controlsSheet.getRangeByIndexes(70, 0, 1, totalModelColumns);
    controlsHeaderRange.format.fill.color = controlsHeaderCell.format.fill.color;
    controlsSheet.getRange("B71").values = [["OPEX DRIVERS"]];

    const controlsHeaderLabels = controlsSheet.getRange("B73:C73");
    controlsHeaderLabels.values = [["Opex drivers", "Driver ID"]];
    controlsHeaderLabels.format.font.bold = true;

    const driversLabelRange = controlsSheet.getRange("B74:B90");
    const driversIdRange = controlsSheet.getRange("C74:C90");
    driversLabelRange.values = [
      ["% of Revenue"],
      ["% of New Revenue"],
      ["% of Payroll"],
      ["% of New Payroll"],
      ["Cost per New Customer"],
      ["Cost per Customer"],
      ["Cost per Headcount"],
      ["Cost per New Headcount"],
      ["Trailing Avg [+] Growth"],
      ["Lookback [+] Growth"],
      ["Fixed Value [+] Inflation"],
      ["LID / Vendor"],
      ["Linked Calculation"],
      ["Placeholder"],
      ["Placeholder"],
      ["Placeholder"],
      ["No Forecast"],
    ];
    driversIdRange.values = [
      [1],
      [2],
      [3],
      [4],
      [5],
      [6],
      [7],
      [8],
      [9],
      [10],
      [11],
      [12],
      [13],
      [13],
      [14],
      [15],
      [16],
    ];
    driversLabelRange.format.font.color = "#3333FF";
    driversLabelRange.format.fill.color = "#FFFFAB";
    applyHairlineBorders(driversLabelRange);
    driversIdRange.format.font.color = "#3333FF";
    driversIdRange.format.fill.clear();
    applyHairlineBorders(driversIdRange);

    const bodyRowStart = 5;
    const bodyRowCount = 995;
    const bodyRange = sheet.getRangeByIndexes(bodyRowStart, 0, bodyRowCount, totalModelColumns);
    bodyRange.format.font.name = spec.font.name;
    bodyRange.format.font.size = spec.font.size;
    if (spec.timelineColumns > 0) {
      const timelineBodyRange = sheet.getRangeByIndexes(
        bodyRowStart,
        timelineStartColIndex,
        bodyRowCount,
        spec.timelineColumns
      );
      timelineBodyRange.numberFormat = Array.from({ length: bodyRowCount }, () =>
        Array.from({ length: spec.timelineColumns }, () => '#,##0;[Red]-#,##0;"-"')
      );
    }

    if (totalModelColumns < MAX_EXCEL_COLUMNS) {
      const clearColumnCount = MAX_EXCEL_COLUMNS - totalModelColumns;
      const clearRange = sheet.getRangeByIndexes(0, totalModelColumns, MAX_EXCEL_ROWS, clearColumnCount);
      clearRange.clear(Excel.ClearApplyTo.all);
    }

    sheet.activate();
    await context.sync();
  });

  if (forecastRowCount > 0) {
    await applyOpexForecastConditionalFormat(forecastRowCount);
  }
}

function applyColumnWidths(
  sheet: Excel.Worksheet,
  widths: MonthlySheetSpec["columnWidths"],
  constantsColumns: number
): void {
  const columns: Array<[string, number]> = [
    ["A", widths.A],
    ["B", widths.B],
    ["C", widths.C],
    ["D", widths.D],
  ];

  columns.forEach(([column, width]) => {
    sheet.getRange(`${column}:${column}`).format.columnWidth = toColumnWidthPoints(width);
  });

  const additionalColumns = Math.max(0, constantsColumns - BASE_CONSTANTS_COLUMNS);
  if (additionalColumns <= 0) {
    return;
  }

  const lastIndex = BASE_CONSTANTS_COLUMNS + additionalColumns - 1;
  const nextToLastIndex = lastIndex - 1;
  for (let offset = 0; offset < additionalColumns; offset += 1) {
    const columnIndex = BASE_CONSTANTS_COLUMNS + offset;
    let width = widths.otherAdditionalColumnsWidth;
    if (additionalColumns === 1) {
      width = widths.lastColumnWidth;
    } else if (additionalColumns === 2) {
      width = offset === 0 ? widths.nextToLastColumnWidth : widths.lastColumnWidth;
    } else if (columnIndex === nextToLastIndex) {
      width = widths.nextToLastColumnWidth;
    } else if (columnIndex === lastIndex) {
      width = widths.lastColumnWidth;
    }

    const columnLetter = columnIndexToLetters(columnIndex);
    sheet.getRange(`${columnLetter}:${columnLetter}`).format.columnWidth = toColumnWidthPoints(width);
  }
}

function applyTimelineColumnWidths(
  sheet: Excel.Worksheet,
  constantsColumns: number,
  timelineColumns: number
): void {
  if (timelineColumns <= 0) {
    return;
  }

  const startIndex = constantsColumns;
  const range = sheet.getRangeByIndexes(0, startIndex, 1, timelineColumns);
  range.format.columnWidth = toColumnWidthPoints(DEFAULT_TIMELINE_COLUMN_WIDTH);
}

function applyOpexSpecificColumnWidths(sheet: Excel.Worksheet): void {
  const widths: Array<[string, number]> = [
    ["C", 2],
    ["D", 2],
    ["G", 20],
    ["H", 11],
    ["I", 8],
    ["J", 8],
    ["K", 8],
    ["L", 14],
    ["M", 5],
    ["N", 12],
    ["O", 12],
    ["P", 1],
    ["Q", 1],
  ];

  widths.forEach(([column, width]) => {
    sheet.getRange(`${column}:${column}`).format.columnWidth = toColumnWidthPoints(width);
  });
}

function applyHairlineBorders(range: Excel.Range): void {
  const borders = range.format.borders;
  const borderItems = [
    Excel.BorderIndex.edgeTop,
    Excel.BorderIndex.edgeBottom,
    Excel.BorderIndex.edgeLeft,
    Excel.BorderIndex.edgeRight,
    Excel.BorderIndex.insideHorizontal,
    Excel.BorderIndex.insideVertical,
  ];

  borderItems.forEach((borderIndex) => {
    const border = borders.getItem(borderIndex);
    border.style = Excel.BorderLineStyle.continuous;
    border.weight = Excel.BorderWeight.hairline;
  });
}

async function applyOpexForecastConditionalFormat(rowCount: number): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject("Opex Monthly");
    const application = context.workbook.application;
    sheet.load("name");
    application.load("calculationMode");
    await context.sync();

    if (sheet.isNullObject) {
      return;
    }

    const originalCalcMode = application.calculationMode;
    const conditionalRange = sheet.getRangeByIndexes(57, 7, rowCount, 1);
    conditionalRange.conditionalFormats.clearAll();
    const conditionalFormat = conditionalRange.conditionalFormats.add(
      Excel.ConditionalFormatType.custom
    );
    conditionalFormat.custom.format.fill.color = "#D9D9D9";
    conditionalFormat.custom.rule.formula = "=FALSE";
    await context.sync();
    conditionalFormat.custom.rule.formula = "=NOT(OR($M58=9,$M58=10))";

    conditionalRange.setDirty();
    const driverIdRange = sheet.getRangeByIndexes(57, 12, rowCount, 1);
    driverIdRange.setDirty();

    application.calculationMode = Excel.CalculationMode.manual;
    await context.sync();
    application.calculate(Excel.CalculationType.fullRebuild);
    application.calculationMode = originalCalcMode;

    await context.sync();
  });
}

function toColumnWidthPoints(width: number): number {
  const pixels = Math.floor(width * 7 + 5);
  return pixels * 0.75;
}

function columnIndexToLetters(index: number): string {
  let value = index + 1;
  let letters = "";

  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }

  return letters;
}

function normalizeRangeAddress(value: string): string {
  const trimmed = value.trim();
  if (!trimmed || trimmed === "=") {
    return "";
  }

  return trimmed.startsWith("=") ? trimmed.slice(1).trim() : trimmed;
}

type ParsedSheetAddress = {
  sheetName: string | null;
  address: string;
};

function parseSheetAddress(value: string): ParsedSheetAddress {
  if (!value.includes("!")) {
    return { sheetName: null, address: value };
  }

  const quotedMatch = /^'(.+)'!(.+)$/.exec(value);
  if (quotedMatch) {
    return {
      sheetName: quotedMatch[1].replace(/''/g, "'"),
      address: quotedMatch[2],
    };
  }

  const parts = value.split("!");
  const sheetName = parts[0];
  const address = parts.slice(1).join("!");
  return { sheetName, address };
}

function getRangeFromAddress(
  context: Excel.RequestContext,
  address: string
): Excel.Range {
  const parsed = parseSheetAddress(address);
  const worksheet = parsed.sheetName
    ? context.workbook.worksheets.getItem(parsed.sheetName)
    : context.workbook.getActiveWorksheet();
  return worksheet.getRange(parsed.address);
}
