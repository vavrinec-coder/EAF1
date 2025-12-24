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
};

const DEFAULT_SHEET_RANGE = "A1:ZZ200";
const COLUMN_HIDE_LIMIT = 200;

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

    const timeRange = sheet.getRangeByIndexes(5, 0, 1, totalModelColumns);
    timeRange.format.fill.color = "#FFD966";
    timeRange.format.font.name = spec.font.name;
    timeRange.format.font.color = "#000000";
    const timeValues = [Array.from({ length: totalModelColumns }, (_, index) => (index === 0 ? "TIME" : ""))];
    timeRange.values = timeValues;

    const labelRange = sheet.getRange("B8:B12");
    labelRange.values = [
      ["Timeline Start Date"],
      ["Actuals End Date"],
      ["Forecast Start Date"],
      ["Timeline length"],
      ["Forecast End Date"],
    ];
    labelRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;

    const unitRange = sheet.getRange("I8:I12");
    unitRange.values = [["Date"], ["Date"], ["Date"], ["#months"], ["Date"]];
    unitRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;

    const valueRange = sheet.getRange("C8:C12");
    valueRange.values = [
      [new Date(2023, 0, 1)],
      [null],
      [null],
      [spec.timelineColumns],
      [null],
    ];
    valueRange.format.horizontalAlignment = Excel.HorizontalAlignment.right;

    sheet.getRange("C9").formulas = [["=C8+1"]];
    sheet.getRange("C10").formulas = [["=C9+1"]];
    sheet.getRange("C12").formulas = [["=EOMONTH(C8,C11-1)"]];

    sheet.getRange("C8").format.font.color = "#3333FF";
    sheet.getRange("C9").format.font.color = "#3333FF";
    sheet.getRange("C11").format.font.color = "#3333FF";
    sheet.getRange("C10").format.font.color = "#000000";
    sheet.getRange("C12").format.font.color = "#000000";

    sheet.getRange("C8:C10").numberFormat = Array.from({ length: 3 }, () => ["m/d/yyyy"]);
    sheet.getRange("C12").numberFormat = [["m/d/yyyy"]];

    applyThinOutline(sheet.getRange("C8:C12"));
    applyThinOutline(sheet.getRange("C10"));
    applyThinOutline(sheet.getRange("C12"));

    sheet.activate();
    await context.sync();
  });
}

function applyColumnWidths(
  sheet: Excel.Worksheet,
  widths: ControlsSheetSpec["columnWidths"]
): void {
  const entries: Array<[string, number]> = [
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

  entries.forEach(([column, width]) => {
    sheet.getRange(`${column}:${column}`).columnWidth = width;
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
