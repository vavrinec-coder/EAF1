/* global Excel */

export type ControlsSheetSpec = {
  constantsColumns: number;
  timelineColumns: number;
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
  timeHeader: {
    startCell: string;
    title: string;
    rows: number;
    columns: number;
    fillColor: string;
    fontColor: string;
  };
  constantsBlock: {
    startRow: number;
    timelineStartDate: Date;
    actualsEndDate: Date;
    timelineLength: number;
    financialYearEndMonth: number;
    historicalPeriod: string;
    forecastPeriod: string;
  };
};

const DEFAULT_SHEET_RANGE = "A1:ZZ200";
const COLUMN_HIDE_LIMIT = 200;
const DATE_NUMBER_FORMAT = "[$-en-US]d/mmm/yy;@";
const DEFAULT_TIMELINE_COLUMN_WIDTH = 12;

export async function createControlsSheet(spec: ControlsSheetSpec): Promise<void> {
  await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    let sheet = worksheets.getItemOrNullObject("Controls");
    sheet.load("name");
    await context.sync();

    if (sheet.isNullObject) {
      sheet = worksheets.add("Controls");
    } else {
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("address");
      await context.sync();
      if (!usedRange.isNullObject) {
        usedRange.clear();
      }
    }

    const totalModelColumns = spec.constantsColumns + spec.timelineColumns;

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

    const timeAnchor = sheet.getRange(spec.timeHeader.startCell);
    const timeRange = timeAnchor.getResizedRange(
      spec.timeHeader.rows - 1,
      spec.timeHeader.columns - 1
    );
    timeRange.format.fill.color = spec.timeHeader.fillColor;
    timeRange.format.font.name = spec.font.name;
    timeRange.format.font.color = spec.timeHeader.fontColor;
    const timeValues: string[][] = Array.from({ length: spec.timeHeader.rows }, (_, rowIndex) =>
      Array.from({ length: spec.timeHeader.columns }, (_, columnIndex) =>
        rowIndex === 0 && columnIndex === 0 ? spec.timeHeader.title : ""
      )
    );
    timeRange.values = timeValues;

    const constantsStartRow = spec.constantsBlock.startRow;
    const labelRange = sheet.getRangeByIndexes(constantsStartRow - 1, 1, 9, 1);
    labelRange.values = [
      ["Timeline Start Date"],
      ["Actuals End Date"],
      ["Forecast Start Date"],
      ["Timeline length"],
      ["Forecast End Date"],
      ["Financial Year End (month)"],
      ["Last Actual Period Column number"],
      ["Historical period"],
      ["Forecast period"],
    ];
    labelRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;

    const unitRange = sheet.getRangeByIndexes(constantsStartRow - 1, 8, 9, 1);
    unitRange.values = [
      ["Date"],
      ["Date"],
      ["Date"],
      ["#months"],
      ["Date"],
      ["month"],
      ["#"],
      ["Label"],
      ["Label"],
    ];
    unitRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;

    const valueRange = sheet.getRangeByIndexes(constantsStartRow - 1, 2, 9, 1);
    valueRange.values = [
      [toExcelDateSerial(spec.constantsBlock.timelineStartDate)],
      [toExcelDateSerial(spec.constantsBlock.actualsEndDate)],
      [null],
      [spec.constantsBlock.timelineLength],
      [null],
      [spec.constantsBlock.financialYearEndMonth],
      [null],
      [spec.constantsBlock.historicalPeriod],
      [spec.constantsBlock.forecastPeriod],
    ];
    valueRange.format.horizontalAlignment = Excel.HorizontalAlignment.right;

    const row1 = constantsStartRow;
    const row2 = constantsStartRow + 1;
    const row3 = constantsStartRow + 2;
    const row4 = constantsStartRow + 3;
    const row5 = constantsStartRow + 4;
    const row6 = constantsStartRow + 5;
    const row7 = constantsStartRow + 6;
    const row8 = constantsStartRow + 7;
    const row9 = constantsStartRow + 8;
    sheet.getRange(`C${row3}`).formulas = [[`=C${row2}+1`]];
    sheet.getRange(`C${row5}`).formulas = [[`=EOMONTH(C${row1},C${row4}-1)`]];
    sheet.getRange(`C${row7}`).formulas = [[`=ROUND((C${row2}-C${row1})/30,0)`]];

    sheet.getRange(`C${row1}`).format.font.color = "#3333FF";
    sheet.getRange(`C${row2}`).format.font.color = "#3333FF";
    sheet.getRange(`C${row4}`).format.font.color = "#3333FF";
    sheet.getRange(`C${row6}`).format.font.color = "#3333FF";
    sheet.getRange(`C${row8}`).format.font.color = "#3333FF";
    sheet.getRange(`C${row9}`).format.font.color = "#3333FF";
    sheet.getRange(`C${row3}`).format.font.color = "#000000";
    sheet.getRange(`C${row5}`).format.font.color = "#000000";

    sheet.getRange(`C${row1}`).numberFormat = [[DATE_NUMBER_FORMAT]];
    sheet.getRange(`C${row2}`).numberFormat = [[DATE_NUMBER_FORMAT]];
    sheet.getRange(`C${row3}`).numberFormat = [[DATE_NUMBER_FORMAT]];
    sheet.getRange(`C${row5}`).numberFormat = [[DATE_NUMBER_FORMAT]];

    sheet.getRange("B19").values = [["Monthly timeline"]];
    sheet.getRange("B19").format.font.bold = true;
    sheet.getRange("B27").values = [["Quarterly timeline"]];
    sheet.getRange("B27").format.font.bold = true;

    sheet.getRange("B20:B25").values = [
      ["Start Date"],
      ["End Date"],
      ["Period type"],
      ["Period counter"],
      ["Financial Year"],
      ["Financial Quarter"],
    ];
    sheet.getRange("B28:B31").values = [
      ["Start Date"],
      ["End Date"],
      ["Period counter"],
      ["Financial Quarter"],
    ];
    sheet.getRange("I20:I25").values = [
      ["Date"],
      ["Date"],
      ["Label"],
      ["Counter"],
      ["Year"],
      ["Label"],
    ];
    sheet.getRange("I28:I31").values = [["Date"], ["Date"], ["Counter"], ["Label"]];

    const timelineStartColIndex = spec.constantsColumns;
    const timelineFormulaRangeStart = sheet.getRangeByIndexes(
      19,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeEnd = sheet.getRangeByIndexes(
      20,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeType = sheet.getRangeByIndexes(
      21,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeCounter = sheet.getRangeByIndexes(
      22,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeYear = sheet.getRangeByIndexes(
      23,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeQuarter = sheet.getRangeByIndexes(
      24,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const quarterlyColumns = Math.max(1, Math.floor(spec.timelineColumns / 3));
    const timelineFormulaRangeQuarterlyStart = sheet.getRangeByIndexes(
      27,
      timelineStartColIndex,
      1,
      quarterlyColumns
    );
    const timelineFormulaRangeQuarterlyEnd = sheet.getRangeByIndexes(
      28,
      timelineStartColIndex,
      1,
      quarterlyColumns
    );
    const timelineFormulaRangeQuarterlyCounter = sheet.getRangeByIndexes(
      29,
      timelineStartColIndex,
      1,
      quarterlyColumns
    );
    const timelineFormulaRangeQuarterlyLabel = sheet.getRangeByIndexes(
      30,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );

    const startDateFormulas: string[] = [];
    const endDateFormulas: string[] = [];
    const periodTypeFormulas: string[] = [];
    const periodCounterFormulas: string[] = [];
    const financialYearFormulas: string[] = [];
    const financialQuarterFormulas: string[] = [];
    const quarterlyStartFormulas: string[] = [];
    const quarterlyEndFormulas: string[] = [];
    const quarterlyCounterFormulas: string[] = [];
    const quarterlyLabelFormulas: string[] = [];
    const timelineStartColumnLetter = columnIndexToLetters(timelineStartColIndex);
    const timelineEndColumnLetter = columnIndexToLetters(
      timelineStartColIndex + spec.timelineColumns - 1
    );
    const timelineQuarterRange = `$${timelineStartColumnLetter}25:$${timelineEndColumnLetter}25`;
    const timelineMatchRange = `$${timelineStartColumnLetter}21:$${timelineEndColumnLetter}21`;
    for (let i = 0; i < spec.timelineColumns; i += 1) {
      const columnIndex = timelineStartColIndex + i;
      const columnLetter = columnIndexToLetters(columnIndex);
      const prevColumnLetter = columnIndexToLetters(columnIndex - 1);
      startDateFormulas.push(
        `=IF(ISBLANK(${prevColumnLetter}20),$C$9,${prevColumnLetter}21+1)`
      );
      endDateFormulas.push(`=EOMONTH(${columnLetter}20,0)`);
      periodTypeFormulas.push(`=IF(${columnLetter}21>$C$10,$C$17,$C$16)`);
      periodCounterFormulas.push(`=IF(ISBLANK(${prevColumnLetter}23),1,${prevColumnLetter}23+1)`);
      financialYearFormulas.push(
        `=IF(MONTH(${columnLetter}21)>$C$14,YEAR(${columnLetter}21)+1,YEAR(${columnLetter}21))`
      );
      financialQuarterFormulas.push(
        `=CONCAT(CHOOSE(INT(MOD(MONTH(${columnLetter}21)-$C$14-1,12)/3)+1,"Q1","Q2","Q3","Q4")," ",${columnLetter}24)`
      );
      quarterlyStartFormulas.push(
        `=IF(ISBLANK(${prevColumnLetter}28),$C$9,${prevColumnLetter}29+1)`
      );
      quarterlyEndFormulas.push(`=EOMONTH(${columnLetter}28,2)`);
      quarterlyCounterFormulas.push(
        `=IF(ISBLANK(${prevColumnLetter}30),1,${prevColumnLetter}30+1)`
      );
      quarterlyLabelFormulas.push(
        `=INDEX(${timelineQuarterRange},MATCH(${columnLetter}29,${timelineMatchRange},0))`
      );
    }

    timelineFormulaRangeStart.formulas = [startDateFormulas];
    timelineFormulaRangeEnd.formulas = [endDateFormulas];
    timelineFormulaRangeType.formulas = [periodTypeFormulas];
    timelineFormulaRangeCounter.formulas = [periodCounterFormulas];
    timelineFormulaRangeYear.formulas = [financialYearFormulas];
    timelineFormulaRangeQuarter.formulas = [financialQuarterFormulas];
    timelineFormulaRangeQuarterlyStart.formulas = [quarterlyStartFormulas.slice(0, quarterlyColumns)];
    timelineFormulaRangeQuarterlyEnd.formulas = [quarterlyEndFormulas.slice(0, quarterlyColumns)];
    timelineFormulaRangeQuarterlyCounter.formulas = [
      quarterlyCounterFormulas.slice(0, quarterlyColumns),
    ];
    timelineFormulaRangeQuarterlyLabel.formulas = [quarterlyLabelFormulas];
    timelineFormulaRangeType.format.horizontalAlignment = "Right";
    timelineFormulaRangeQuarter.format.horizontalAlignment = "Right";
    timelineFormulaRangeStart.numberFormat = [
      Array.from({ length: spec.timelineColumns }, () => DATE_NUMBER_FORMAT),
    ];
    timelineFormulaRangeEnd.numberFormat = [
      Array.from({ length: spec.timelineColumns }, () => DATE_NUMBER_FORMAT),
    ];
    timelineFormulaRangeQuarterlyStart.numberFormat = [
      Array.from({ length: quarterlyColumns }, () => DATE_NUMBER_FORMAT),
    ];
    timelineFormulaRangeQuarterlyEnd.numberFormat = [
      Array.from({ length: quarterlyColumns }, () => DATE_NUMBER_FORMAT),
    ];
    sheet.getRangeByIndexes(21, timelineStartColIndex, 1, spec.timelineColumns).format.horizontalAlignment =
      "Right";
    sheet.getRangeByIndexes(24, timelineStartColIndex, 1, spec.timelineColumns).format.horizontalAlignment =
      "Right";

    sheet.activate();
    await context.sync();
  });
}

function applyColumnWidths(
  sheet: Excel.Worksheet,
  widths: ControlsSheetSpec["columnWidths"]
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

function toExcelDateSerial(date: Date): number {
  const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  const excelEpoch = Date.UTC(1899, 11, 30);
  return (utc - excelEpoch) / 86400000;
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
