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
  };
};

const DEFAULT_SHEET_RANGE = "A1:ZZ200";
const COLUMN_HIDE_LIMIT = 200;
const DATE_NUMBER_FORMAT = "[$-en-US]d/mmm/yy;@";

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
    const labelRange = sheet.getRangeByIndexes(constantsStartRow - 1, 1, 5, 1);
    labelRange.values = [
      ["Timeline Start Date"],
      ["Actuals End Date"],
      ["Forecast Start Date"],
      ["Timeline length"],
      ["Forecast End Date"],
    ];
    labelRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;

    const unitRange = sheet.getRangeByIndexes(constantsStartRow - 1, 8, 5, 1);
    unitRange.values = [["Date"], ["Date"], ["Date"], ["#months"], ["Date"]];
    unitRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;

    const valueRange = sheet.getRangeByIndexes(constantsStartRow - 1, 2, 5, 1);
    valueRange.values = [
      [toExcelDateSerial(spec.constantsBlock.timelineStartDate)],
      [toExcelDateSerial(spec.constantsBlock.actualsEndDate)],
      [null],
      [spec.constantsBlock.timelineLength],
      [null],
    ];
    valueRange.format.horizontalAlignment = Excel.HorizontalAlignment.right;

    const row1 = constantsStartRow;
    const row2 = constantsStartRow + 1;
    const row3 = constantsStartRow + 2;
    const row4 = constantsStartRow + 3;
    const row5 = constantsStartRow + 4;
    sheet.getRange(`C${row3}`).formulas = [[`=C${row2}+1`]];
    sheet.getRange(`C${row5}`).formulas = [[`=EOMONTH(C${row1},C${row4}-1)`]];

    sheet.getRange(`C${row1}`).format.font.color = "#3333FF";
    sheet.getRange(`C${row2}`).format.font.color = "#3333FF";
    sheet.getRange(`C${row4}`).format.font.color = "#3333FF";
    sheet.getRange(`C${row3}`).format.font.color = "#000000";
    sheet.getRange(`C${row5}`).format.font.color = "#000000";

    sheet.getRange(`C${row1}`).numberFormat = [[DATE_NUMBER_FORMAT]];
    sheet.getRange(`C${row2}`).numberFormat = [[DATE_NUMBER_FORMAT]];
    sheet.getRange(`C${row3}`).numberFormat = [[DATE_NUMBER_FORMAT]];
    sheet.getRange(`C${row5}`).numberFormat = [[DATE_NUMBER_FORMAT]];

    applyThinOutline(sheet.getRangeByIndexes(constantsStartRow - 1, 2, 5, 1));
    applyThinOutline(sheet.getRange(`C${row3}`));
    applyThinOutline(sheet.getRange(`C${row5}`));

    sheet.activate();
    await context.sync();
  });
}

function applyColumnWidths(
  sheet: Excel.Worksheet,
  widths: ControlsSheetSpec["columnWidths"]
): void {
  const entries: number[] = [
    widths.A,
    widths.B,
    widths.C,
    widths.D,
    widths.E,
    widths.F,
    widths.G,
    widths.H,
    widths.I,
    widths.J,
  ];

  entries.forEach((width, index) => {
    sheet.getRangeByIndexes(0, index, 1, 1).columnWidth = width;
  });
}

function applyThinOutline(range: Excel.Range): void {
  const edges = [
    Excel.BorderIndex.edgeTop,
    Excel.BorderIndex.edgeBottom,
    Excel.BorderIndex.edgeLeft,
    Excel.BorderIndex.edgeRight,
  ];

  edges.forEach((edge) => {
    const border = range.format.borders.getItem(edge);
    border.style = Excel.BorderLineStyle.continuous;
    border.weight = Excel.BorderWeight.thin;
    border.color = "#000000";
  });
}

function toExcelDateSerial(date: Date): number {
  const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  const excelEpoch = Date.UTC(1899, 11, 30);
  return (utc - excelEpoch) / 86400000;
}
