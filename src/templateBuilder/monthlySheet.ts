/* global Excel */

export type MonthlySheetSpec = {
  constantsColumns: number;
  timelineColumns: number;
  tabColor: string;
  sectionColor: string;
  font: {
    name: string;
    color: string;
    size: number;
  };
  headerRows: number;
  headerFillColor: string;
  columnWidths: {
    A: number;
    B: number;
    C: number;
    D: number;
    E: number;
    F: number;
    G: number;
    H: number;
    I: number;
    J: number;
  };
};

const DEFAULT_SHEET_RANGE = "A1:ZZ200";
const COLUMN_HIDE_LIMIT = 200;
const DEFAULT_TIMELINE_COLUMN_WIDTH = 12;
const MAX_EXCEL_COLUMNS = 16384;
const MAX_EXCEL_ROWS = 1048576;

export async function createMonthlySheet(spec: MonthlySheetSpec): Promise<void> {
  await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    const controlsSheet = worksheets.getItemOrNullObject("Controls");
    controlsSheet.load("name");
    let sheet = worksheets.getItemOrNullObject("Monthly");
    sheet.load("name");
    await context.sync();

    if (controlsSheet.isNullObject) {
      throw new Error('Controls sheet not found. Create the "Controls" sheet first.');
    }

    if (sheet.isNullObject) {
      sheet = worksheets.add("Monthly");
    } else {
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("address");
      await context.sync();
      if (!usedRange.isNullObject) {
        usedRange.clear();
      }
    }

    const totalModelColumns = spec.constantsColumns + spec.timelineColumns;
    sheet.tabColor = spec.tabColor;
    sheet.showGridlines = false;

    const baseRange = sheet.getRange(DEFAULT_SHEET_RANGE);
    baseRange.format.font.name = spec.font.name;
    baseRange.format.font.size = spec.font.size;
    baseRange.format.font.color = spec.font.color;

    applyColumnWidths(sheet, spec.columnWidths);
    applyTimelineColumnWidths(sheet, spec.constantsColumns, spec.timelineColumns);

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
    sheet.getRange("B7").values = [["MODEL FLAGS"]];
    sheet.getRange("B7").format.font.color = "#FFFFFF";

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
    const controlsFlagUnitRange = sheet.getRange("I9:I21");
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
      flagUnitFormulas.push([`=Controls!I${controlsRow}`]);

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

    if (totalModelColumns < MAX_EXCEL_COLUMNS) {
      const clearColumnCount = MAX_EXCEL_COLUMNS - totalModelColumns;
      const clearRange = sheet.getRangeByIndexes(0, totalModelColumns, MAX_EXCEL_ROWS, clearColumnCount);
      clearRange.clear(Excel.ClearApplyTo.all);
    }

    sheet.activate();
    await context.sync();
  });
}

function applyColumnWidths(
  sheet: Excel.Worksheet,
  widths: MonthlySheetSpec["columnWidths"]
): void {
  const columns: Array<[string, number]> = [
    ["A", widths.A],
    ["B", widths.B],
    ["C", widths.C],
    ["D", widths.D],
    ["E", widths.E],
    ["F", widths.F],
    ["G", widths.G],
    ["H", widths.H],
    ["I", widths.I],
    ["J", widths.J],
  ];

  columns.forEach(([column, width]) => {
    sheet.getRange(`${column}:${column}`).format.columnWidth = toColumnWidthPoints(width);
  });
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
