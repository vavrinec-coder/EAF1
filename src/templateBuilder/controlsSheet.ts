/* global Excel */

export type ControlsSheetSpec = {
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
  timeHeader: {
    startCell: string;
    title: string;
    rows: number;
    columns: number;
    fillColor: string;
    fontColor: string;
  };
  flagsHeader: {
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
const MAX_EXCEL_COLUMNS = 16384;
const MAX_EXCEL_ROWS = 1048576;

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
    if (totalModelColumns < MAX_EXCEL_COLUMNS) {
      const clearColumnCount = MAX_EXCEL_COLUMNS - totalModelColumns;
      const clearRange = sheet.getRangeByIndexes(0, totalModelColumns, MAX_EXCEL_ROWS, clearColumnCount);
      clearRange.clear(Excel.ClearApplyTo.all);
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

    const flagsAnchor = sheet.getRange(spec.flagsHeader.startCell);
    const flagsRange = flagsAnchor.getResizedRange(
      spec.flagsHeader.rows - 1,
      spec.flagsHeader.columns - 1
    );
    flagsRange.format.fill.color = spec.flagsHeader.fillColor;
    flagsRange.format.font.name = spec.font.name;
    flagsRange.format.font.color = spec.flagsHeader.fontColor;
    const flagsValues: string[][] = Array.from({ length: spec.flagsHeader.rows }, (_, rowIndex) =>
      Array.from({ length: spec.flagsHeader.columns }, (_, columnIndex) =>
        rowIndex === 0 && columnIndex === 0 ? spec.flagsHeader.title : ""
      )
    );
    flagsRange.values = flagsValues;

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
    sheet.getRange("B43").values = [["Monthly flags"]];
    sheet.getRange("B43").format.font.bold = true;
    sheet.getRange("B58").values = [["Quarterly flags"]];
    sheet.getRange("B58").format.font.bold = true;
    sheet.getRange("B61").values = [["Annual flags"]];
    sheet.getRange("B61").format.font.bold = true;

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
    sheet.getRange("B33").values = [["Annual timeline"]];
    sheet.getRange("B33").format.font.bold = true;
    sheet.getRange("B34:B38").values = [
      ["Start Date"],
      ["End Date"],
      ["Months of Forecast in Year"],
      ["Period type"],
      ["Period counter"],
    ];
    sheet.getRange("B44:B53").values = [
      ["Month"],
      ["Days"],
      ["Actuals Flag"],
      ["Forecast Flag"],
      ["First Forecast Flag"],
      ["Forecast Counter"],
      ["Last Actuals Flag"],
      ["Year counter"],
      ["Financial Year End Flag"],
      ["Quarter counter"],
    ];
    sheet.getRange("B59").values = [["Quarter End Column on monthly timeline"]];
    sheet.getRange("B62").values = [["Year End Column on monthly timeline"]];
    sheet.getRange("B54:B56").values = [["Placeholder"], ["Placeholder"], ["Placeholder"]];
    sheet.getRange("I20:I25").values = [
      ["Date"],
      ["Date"],
      ["Label"],
      ["Counter"],
      ["Year"],
      ["Label"],
    ];
    sheet.getRange("I28:I31").values = [["Date"], ["Date"], ["Counter"], ["Label"]];
    sheet.getRange("I34:I38").values = [["Date"], ["Date"], ["#"], ["Label"], ["Counter"]];
    sheet.getRange("I44:I53").values = [
      ["#"],
      ["#"],
      ["'1/0"],
      ["'1/0"],
      ["'1/0"],
      ["#"],
      ["'1/0"],
      ["#"],
      ["'1/0"],
      ["#"],
    ];
    sheet.getRange("I59").values = [["#"]];
    sheet.getRange("I62").values = [["#"]];
    sheet.getRange("I44:I53").numberFormat = Array.from({ length: 10 }, () => ["@"]);
    sheet.getRange("I44:I53").format.horizontalAlignment = Excel.HorizontalAlignment.left;

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
      quarterlyColumns
    );
    const annualColumns = Math.max(1, Math.floor(spec.timelineColumns / 12));
    const timelineFormulaRangeAnnualStart = sheet.getRangeByIndexes(
      33,
      timelineStartColIndex,
      1,
      annualColumns
    );
    const timelineFormulaRangeAnnualEnd = sheet.getRangeByIndexes(
      34,
      timelineStartColIndex,
      1,
      annualColumns
    );
    const timelineFormulaRangeAnnualMonths = sheet.getRangeByIndexes(
      35,
      timelineStartColIndex,
      1,
      annualColumns
    );
    const timelineFormulaRangeAnnualType = sheet.getRangeByIndexes(
      36,
      timelineStartColIndex,
      1,
      annualColumns
    );
    const timelineFormulaRangeAnnualCounter = sheet.getRangeByIndexes(
      37,
      timelineStartColIndex,
      1,
      annualColumns
    );
    const timelineFormulaRangeQuarterlyFlagMatch = sheet.getRangeByIndexes(
      58,
      timelineStartColIndex,
      1,
      quarterlyColumns
    );
    const timelineFormulaRangeAnnualFlagMatch = sheet.getRangeByIndexes(
      61,
      timelineStartColIndex,
      1,
      annualColumns
    );
    const timelineFormulaRangeFlagsMonth = sheet.getRangeByIndexes(
      43,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeFlagsDays = sheet.getRangeByIndexes(
      44,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeFlagsActuals = sheet.getRangeByIndexes(
      45,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeFlagsForecast = sheet.getRangeByIndexes(
      46,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeFlagsFirstForecast = sheet.getRangeByIndexes(
      47,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeFlagsForecastCounter = sheet.getRangeByIndexes(
      48,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeFlagsLastActuals = sheet.getRangeByIndexes(
      49,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeFlagsYearCounter = sheet.getRangeByIndexes(
      50,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeFlagsYearEndFlag = sheet.getRangeByIndexes(
      51,
      timelineStartColIndex,
      1,
      spec.timelineColumns
    );
    const timelineFormulaRangeFlagsQuarterCounter = sheet.getRangeByIndexes(
      52,
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
    const annualStartFormulas: string[] = [];
    const annualEndFormulas: string[] = [];
    const annualMonthsFormulas: string[] = [];
    const annualTypeFormulas: string[] = [];
    const annualCounterFormulas: string[] = [];
    const quarterlyFlagMatchFormulas: string[] = [];
    const annualFlagMatchFormulas: string[] = [];
    const flagsMonthFormulas: string[] = [];
    const flagsDaysFormulas: string[] = [];
    const flagsActualsFormulas: string[] = [];
    const flagsForecastFormulas: string[] = [];
    const flagsFirstForecastFormulas: string[] = [];
    const flagsForecastCounterFormulas: string[] = [];
    const flagsLastActualsFormulas: string[] = [];
    const flagsYearCounterFormulas: string[] = [];
    const flagsYearEndFlagFormulas: string[] = [];
    const flagsQuarterCounterFormulas: string[] = [];
    const timelineStartColumnLetter = columnIndexToLetters(timelineStartColIndex);
    const timelineEndColumnLetter = columnIndexToLetters(
      timelineStartColIndex + spec.timelineColumns - 1
    );
    const timelineQuarterRange = `$${timelineStartColumnLetter}25:$${timelineEndColumnLetter}25`;
    const timelineMatchRange = `$${timelineStartColumnLetter}21:$${timelineEndColumnLetter}21`;
    const timelineMatchRangeAbsolute = `$${timelineStartColumnLetter}$21:$${timelineEndColumnLetter}$21`;
    for (let i = 0; i < spec.timelineColumns; i += 1) {
      const columnIndex = timelineStartColIndex + i;
      const columnLetter = columnIndexToLetters(columnIndex);
      const prevColumnLetter = columnIndexToLetters(columnIndex - 1);
      const nextColumnLetter = columnIndexToLetters(columnIndex + 1);
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
      annualStartFormulas.push(
        `=IF(ISBLANK(${prevColumnLetter}34),$C$9,${prevColumnLetter}35+1)`
      );
      annualEndFormulas.push(`=EOMONTH(${columnLetter}34,11)`);
      annualMonthsFormulas.push(
        `=LET(yrStart,${columnLetter}34,yrEnd,${columnLetter}35,actEnd,$C$10,fStart,MAX(yrStart,EOMONTH(actEnd,0)+1),MAX(0,(YEAR(yrEnd)*12+MONTH(yrEnd))-(YEAR(fStart)*12+MONTH(fStart))+1))`
      );
      annualTypeFormulas.push(
        `=IF(${columnLetter}36=0,$C$16,IF(${columnLetter}36=12,$C$17,CONCAT($C$16,12-${columnLetter}36,"+",$C$17,${columnLetter}36)))`
      );
      annualCounterFormulas.push(
        `=IF(ISBLANK(${prevColumnLetter}38),1,${prevColumnLetter}38+1)`
      );
      quarterlyFlagMatchFormulas.push(
        `=MATCH(${columnLetter}29,${timelineMatchRangeAbsolute},0)`
      );
      annualFlagMatchFormulas.push(`=MATCH(${columnLetter}35,${timelineMatchRangeAbsolute},0)`);
      flagsMonthFormulas.push(`=MONTH(${columnLetter}20)`);
      flagsDaysFormulas.push(`=DAY(${columnLetter}21)`);
      flagsActualsFormulas.push(`=IF(${columnLetter}22=$C$16,1,0)`);
      flagsForecastFormulas.push(`=1-${columnLetter}46`);
      flagsFirstForecastFormulas.push(
        `=IF(AND(${columnLetter}47=1,${prevColumnLetter}47=0),1,0)`
      );
      flagsForecastCounterFormulas.push(
        `=(${prevColumnLetter}49+1)*${columnLetter}47`
      );
      flagsLastActualsFormulas.push(
        `=IF(AND(${columnLetter}46=1,${nextColumnLetter}47=1),1,0)`
      );
      flagsYearCounterFormulas.push(`=${columnLetter}24-YEAR($C$9)+1`);
      flagsYearEndFlagFormulas.push(`=IF(${columnLetter}44=$C$14,1,0)`);
      flagsQuarterCounterFormulas.push(
        `=IF(ISBLANK(${prevColumnLetter}53),1,IF(${columnLetter}25=${prevColumnLetter}25,${prevColumnLetter}53,${prevColumnLetter}53+1))`
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
    timelineFormulaRangeQuarterlyLabel.formulas = [
      quarterlyLabelFormulas.slice(0, quarterlyColumns),
    ];
    timelineFormulaRangeAnnualStart.formulas = [annualStartFormulas.slice(0, annualColumns)];
    timelineFormulaRangeAnnualEnd.formulas = [annualEndFormulas.slice(0, annualColumns)];
    timelineFormulaRangeAnnualMonths.formulas = [annualMonthsFormulas.slice(0, annualColumns)];
    timelineFormulaRangeAnnualType.formulas = [annualTypeFormulas.slice(0, annualColumns)];
    timelineFormulaRangeAnnualCounter.formulas = [annualCounterFormulas.slice(0, annualColumns)];
    timelineFormulaRangeQuarterlyFlagMatch.formulas = [
      quarterlyFlagMatchFormulas.slice(0, quarterlyColumns),
    ];
    timelineFormulaRangeAnnualFlagMatch.formulas = [annualFlagMatchFormulas.slice(0, annualColumns)];
    timelineFormulaRangeFlagsMonth.formulas = [flagsMonthFormulas];
    timelineFormulaRangeFlagsDays.formulas = [flagsDaysFormulas];
    timelineFormulaRangeFlagsActuals.formulas = [flagsActualsFormulas];
    timelineFormulaRangeFlagsForecast.formulas = [flagsForecastFormulas];
    timelineFormulaRangeFlagsFirstForecast.formulas = [flagsFirstForecastFormulas];
    timelineFormulaRangeFlagsForecastCounter.formulas = [flagsForecastCounterFormulas];
    timelineFormulaRangeFlagsLastActuals.formulas = [flagsLastActualsFormulas];
    timelineFormulaRangeFlagsYearCounter.formulas = [flagsYearCounterFormulas];
    timelineFormulaRangeFlagsYearEndFlag.formulas = [flagsYearEndFlagFormulas];
    timelineFormulaRangeFlagsQuarterCounter.formulas = [flagsQuarterCounterFormulas];
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
    timelineFormulaRangeAnnualStart.numberFormat = [
      Array.from({ length: annualColumns }, () => DATE_NUMBER_FORMAT),
    ];
    timelineFormulaRangeAnnualEnd.numberFormat = [
      Array.from({ length: annualColumns }, () => DATE_NUMBER_FORMAT),
    ];
    sheet.getRangeByIndexes(21, timelineStartColIndex, 1, spec.timelineColumns).format.horizontalAlignment =
      "Right";
    sheet.getRangeByIndexes(24, timelineStartColIndex, 1, spec.timelineColumns).format.horizontalAlignment =
      "Right";
    sheet.getRangeByIndexes(20, timelineStartColIndex, 6, spec.timelineColumns).format.horizontalAlignment =
      "Right";
    sheet.getRangeByIndexes(27, timelineStartColIndex, 4, quarterlyColumns).format.horizontalAlignment =
      "Right";
    sheet.getRangeByIndexes(33, timelineStartColIndex, 5, annualColumns).format.horizontalAlignment =
      "Right";
    sheet.getRangeByIndexes(43, timelineStartColIndex, 10, spec.timelineColumns).format.horizontalAlignment =
      "Right";
    sheet.getRangeByIndexes(58, timelineStartColIndex, 1, quarterlyColumns).format.horizontalAlignment =
      "Right";
    sheet.getRangeByIndexes(61, timelineStartColIndex, 1, annualColumns).format.horizontalAlignment =
      "Right";

    sheet.activate();
    sheet.freezePanes.freezeAt(sheet.getRange("L6"));
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
