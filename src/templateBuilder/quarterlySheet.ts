/* global Excel */

export type QuarterlySheetSpec = {
  constantsColumns: number;
  timelineColumns: number;
  tabColor: string;
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

export async function createQuarterlySheet(spec: QuarterlySheetSpec): Promise<void> {
  await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    let sheet = worksheets.getItemOrNullObject("Quarterly");
    sheet.load("name");
    await context.sync();

    if (sheet.isNullObject) {
      sheet = worksheets.add("Quarterly");
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

    sheet.getRange("C1:C4").formulas = [
      ["=Controls!B28"],
      ["=Controls!B29"],
      ["=Controls!B30"],
      ["=Controls!B31"],
    ];
    sheet.getRange("C1:C4").format.font.color = "#FFFFFF";

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
    const timelineFormulaRangeCounter = sheet.getRangeByIndexes(
      2,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeQuarter = sheet.getRangeByIndexes(
      3,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );

    const controlsTimelineStart = controlsSheet.getRangeByIndexes(
      27,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const controlsTimelineEnd = controlsSheet.getRangeByIndexes(
      28,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    controlsTimelineStart.load("numberFormat");
    controlsTimelineEnd.load("numberFormat");

    const startDateFormulas: string[] = [];
    const endDateFormulas: string[] = [];
    const counterFormulas: string[] = [];
    const quarterFormulas: string[] = [];

    for (let i = 0; i < spec.timelineColumns; i += 1) {
      const columnIndex = timelineStartColIndex + i;
      const columnLetter = columnIndexToLetters(columnIndex);
      startDateFormulas.push(`=Controls!${columnLetter}28`);
      endDateFormulas.push(`=Controls!${columnLetter}29`);
      counterFormulas.push(`=Controls!${columnLetter}30`);
      quarterFormulas.push(`=Controls!${columnLetter}31`);
    }

    timelineFormulaRangeStart.formulas = [startDateFormulas];
    timelineFormulaRangeEnd.formulas = [endDateFormulas];
    timelineFormulaRangeCounter.formulas = [counterFormulas];
    timelineFormulaRangeQuarter.formulas = [quarterFormulas];

    sheet.getRangeByIndexes(0, timelineStartColIndex, 4, spec.timelineColumns).format.horizontalAlignment =
      "Right";

    await context.sync();
    timelineFormulaRangeStart.numberFormat = controlsTimelineStart.numberFormat;
    timelineFormulaRangeEnd.numberFormat = controlsTimelineEnd.numberFormat;

    sheet.activate();
    await context.sync();
  });
}

function applyColumnWidths(
  sheet: Excel.Worksheet,
  widths: QuarterlySheetSpec["columnWidths"]
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
