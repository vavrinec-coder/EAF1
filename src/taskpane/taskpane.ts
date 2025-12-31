/* global document, Excel, Office */

import { createControlsSheet, ControlsSheetSpec } from "../templateBuilder/controlsSheet";
import { createMonthlySheet, MonthlySheetSpec } from "../templateBuilder/monthlySheet";
import { createFsMonthlySheet, FsMonthlySheetSpec } from "../templateBuilder/fsMonthlySheet";
import { createQuarterlySheet, QuarterlySheetSpec } from "../templateBuilder/quarterlySheet";
import { createAnnualSheet, AnnualSheetSpec } from "../templateBuilder/annualSheet";

const MAX_EXCEL_ROWS = 1048576;
const MAX_EXCEL_COLUMNS = 16384;
const DEFAULT_CONSTANTS_COLUMNS = 10;
const DEFAULT_TIME_HEADER_START_CELL = "A7";
const DEFAULT_TIME_HEADER_ROWS = 1;
const DEFAULT_CONSTANTS_START_ROW = 9;
const DEFAULT_FLAGS_HEADER_START_CELL = "A41";
const DEFAULT_FLAGS_HEADER_ROWS = 1;

let selectedRangeEl: HTMLSpanElement;
let headerListEl: HTMLDivElement;
let statusEl: HTMLDivElement;
let timelineColumnsInputEl: HTMLInputElement;
let modelFontNameInputEl: HTMLInputElement;
let modelFontColorInputEl: HTMLInputElement;
let modelFontSizeInputEl: HTMLInputElement;
let modelHeaderRowsInputEl: HTMLInputElement;
let modelHeaderBackgroundInputEl: HTMLInputElement;
let widthColAInputEl: HTMLInputElement;
let widthColBInputEl: HTMLInputElement;
let widthColCInputEl: HTMLInputElement;
let widthColDInputEl: HTMLInputElement;
let widthColEInputEl: HTMLInputElement;
let widthColFInputEl: HTMLInputElement;
let widthColGInputEl: HTMLInputElement;
let widthColHInputEl: HTMLInputElement;
let widthColIInputEl: HTMLInputElement;
let widthColJInputEl: HTMLInputElement;
let createControlsButtonEl: HTMLButtonElement;
let controlsTabColorInputEl: HTMLInputElement;
let createMonthlyButtonEl: HTMLButtonElement;
let monthlyTabColorInputEl: HTMLInputElement;
let monthlySectionColorInputEl: HTMLInputElement;
let createFsMonthlyButtonEl: HTMLButtonElement;
let fsMonthlyTabColorInputEl: HTMLInputElement;
let fsMonthlySectionColorInputEl: HTMLInputElement;
let createQuarterlyButtonEl: HTMLButtonElement;
let quarterlyTabColorInputEl: HTMLInputElement;
let quarterlySectionColorInputEl: HTMLInputElement;
let createAnnualButtonEl: HTMLButtonElement;
let annualTabColorInputEl: HTMLInputElement;
let annualSectionColorInputEl: HTMLInputElement;
let simplifyTextButtonEl: HTMLButtonElement;
let loadLineItemCategoriesButtonEl: HTMLButtonElement;
let updateFsFormattingButtonEl: HTMLButtonElement;
let timeHeaderTitleInputEl: HTMLInputElement;
let timeHeaderFillInputEl: HTMLInputElement;
let timeHeaderFontColorInputEl: HTMLInputElement;
let constantsTimelineLengthInputEl: HTMLInputElement;
let constantsStartDateInputEl: HTMLInputElement;
let constantsActualsEndInputEl: HTMLInputElement;
let constantsYearEndMonthInputEl: HTMLSelectElement;
let constantsHistoricalPeriodInputEl: HTMLInputElement;
let constantsForecastPeriodInputEl: HTMLInputElement;
let flagsHeaderTitleInputEl: HTMLInputElement;
let flagsHeaderFillInputEl: HTMLInputElement;
let flagsHeaderFontColorInputEl: HTMLInputElement;

let timelineLengthDirty = false;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    return;
  }

  const sideloadMessage = document.getElementById("sideload-msg");
  const appBody = document.getElementById("app-body");
  if (sideloadMessage) {
    sideloadMessage.style.display = "none";
  }
  if (appBody) {
    appBody.style.display = "flex";
  }

  selectedRangeEl = document.getElementById("selected-range") as HTMLSpanElement;
  headerListEl = document.getElementById("header-list") as HTMLDivElement;
  statusEl = document.getElementById("status") as HTMLDivElement;

  const loadButton = document.getElementById("load-selection") as HTMLButtonElement;
  const unpivotButton = document.getElementById("unpivot") as HTMLButtonElement;
  simplifyTextButtonEl = document.getElementById("simplify-text") as HTMLButtonElement;
  loadLineItemCategoriesButtonEl = document.getElementById(
    "load-line-item-categories"
  ) as HTMLButtonElement;
  updateFsFormattingButtonEl = document.getElementById(
    "update-fs-formatting"
  ) as HTMLButtonElement;

  timelineColumnsInputEl = document.getElementById("model-timeline-columns") as HTMLInputElement;
  modelFontNameInputEl = document.getElementById("model-font-name") as HTMLInputElement;
  modelFontColorInputEl = document.getElementById("model-font-color") as HTMLInputElement;
  modelFontSizeInputEl = document.getElementById("model-font-size") as HTMLInputElement;
  modelHeaderRowsInputEl = document.getElementById("model-header-rows") as HTMLInputElement;
  modelHeaderBackgroundInputEl = document.getElementById("model-header-background") as HTMLInputElement;
  widthColAInputEl = document.getElementById("width-col-a") as HTMLInputElement;
  widthColBInputEl = document.getElementById("width-col-b") as HTMLInputElement;
  widthColCInputEl = document.getElementById("width-col-c") as HTMLInputElement;
  widthColDInputEl = document.getElementById("width-col-d") as HTMLInputElement;
  widthColEInputEl = document.getElementById("width-col-e") as HTMLInputElement;
  widthColFInputEl = document.getElementById("width-col-f") as HTMLInputElement;
  widthColGInputEl = document.getElementById("width-col-g") as HTMLInputElement;
  widthColHInputEl = document.getElementById("width-col-h") as HTMLInputElement;
  widthColIInputEl = document.getElementById("width-col-i") as HTMLInputElement;
  widthColJInputEl = document.getElementById("width-col-j") as HTMLInputElement;
  createControlsButtonEl = document.getElementById("create-controls-sheet") as HTMLButtonElement;
  controlsTabColorInputEl = document.getElementById(
    "controls-tab-color"
  ) as HTMLInputElement;
  createMonthlyButtonEl = document.getElementById("create-monthly-sheet") as HTMLButtonElement;
  monthlyTabColorInputEl = document.getElementById("monthly-tab-color") as HTMLInputElement;
  monthlySectionColorInputEl = document.getElementById(
    "monthly-section-color"
  ) as HTMLInputElement;
  createFsMonthlyButtonEl = document.getElementById(
    "create-fs-monthly-sheet"
  ) as HTMLButtonElement;
  fsMonthlyTabColorInputEl = document.getElementById(
    "fs-monthly-tab-color"
  ) as HTMLInputElement;
  fsMonthlySectionColorInputEl = document.getElementById(
    "fs-monthly-section-color"
  ) as HTMLInputElement;
  createQuarterlyButtonEl = document.getElementById(
    "create-quarterly-sheet"
  ) as HTMLButtonElement;
  quarterlyTabColorInputEl = document.getElementById("quarterly-tab-color") as HTMLInputElement;
  quarterlySectionColorInputEl = document.getElementById(
    "quarterly-section-color"
  ) as HTMLInputElement;
  createAnnualButtonEl = document.getElementById("create-annual-sheet") as HTMLButtonElement;
  annualTabColorInputEl = document.getElementById("annual-tab-color") as HTMLInputElement;
  annualSectionColorInputEl = document.getElementById("annual-section-color") as HTMLInputElement;
  timeHeaderTitleInputEl = document.getElementById("time-header-title") as HTMLInputElement;
  timeHeaderFillInputEl = document.getElementById("time-header-fill") as HTMLInputElement;
  timeHeaderFontColorInputEl = document.getElementById("time-header-font-color") as HTMLInputElement;
  constantsTimelineLengthInputEl = document.getElementById(
    "constants-timeline-length"
  ) as HTMLInputElement;
  constantsStartDateInputEl = document.getElementById("constants-start-date") as HTMLInputElement;
  constantsActualsEndInputEl = document.getElementById("constants-actuals-end") as HTMLInputElement;
  constantsYearEndMonthInputEl = document.getElementById(
    "constants-year-end-month"
  ) as HTMLSelectElement;
  constantsHistoricalPeriodInputEl = document.getElementById(
    "constants-historical-period"
  ) as HTMLInputElement;
  constantsForecastPeriodInputEl = document.getElementById(
    "constants-forecast-period"
  ) as HTMLInputElement;
  flagsHeaderTitleInputEl = document.getElementById("flags-header-title") as HTMLInputElement;
  flagsHeaderFillInputEl = document.getElementById("flags-header-fill") as HTMLInputElement;
  flagsHeaderFontColorInputEl = document.getElementById(
    "flags-header-font-color"
  ) as HTMLInputElement;

  loadButton.addEventListener("click", () => {
    void handleLoadSelection();
  });
  unpivotButton.addEventListener("click", () => {
    void handleUnpivot();
  });
  simplifyTextButtonEl.addEventListener("click", () => {
    void handleSimplifyText();
  });
  loadLineItemCategoriesButtonEl.addEventListener("click", () => {
    void handleLoadLineItemCategories();
  });
  updateFsFormattingButtonEl.addEventListener("click", () => {
    void handleUpdateFsFormatting();
  });
  createControlsButtonEl.addEventListener("click", () => {
    void handleCreateControlsSheet();
  });
  createMonthlyButtonEl.addEventListener("click", () => {
    void handleCreateMonthlySheet();
  });
  createFsMonthlyButtonEl.addEventListener("click", () => {
    void handleCreateFsMonthlySheet();
  });
  createQuarterlyButtonEl.addEventListener("click", () => {
    void handleCreateQuarterlySheet();
  });
  createAnnualButtonEl.addEventListener("click", () => {
    void handleCreateAnnualSheet();
  });
  constantsTimelineLengthInputEl.addEventListener("input", () => {
    timelineLengthDirty = constantsTimelineLengthInputEl.value.trim().length > 0;
  });
  timelineColumnsInputEl.addEventListener("input", () => {
    syncDerivedDefaults();
  });

  renderHeaders([]);
  setStatus("Ready. Select a range and click Load Selection.", "info");
  syncDerivedDefaults(true);
});

type SelectionData = {
  address: string;
  values: Excel.RangeValueType[][];
  rowCount: number;
  columnCount: number;
};

async function readSelection(): Promise<SelectionData> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["values", "address", "rowCount", "columnCount"]);
    await context.sync();

    return {
      address: range.address,
      values: range.values,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
    };
  });
}

async function handleLoadSelection(): Promise<void> {
  setStatus("Loading selection...", "info");

  try {
    const selection = await readSelection();
    if (selection.rowCount < 2) {
      setStatus("Select a range with a header row and at least one data row.", "error");
      return;
    }

    const headers = normalizeHeaders(selection.values[0]);
    selectedRangeEl.textContent = selection.address;
    renderHeaders(headers);
    setStatus(`Loaded ${selection.rowCount - 1} data rows.`, "info");
  } catch (error) {
    setStatus(getErrorMessage(error), "error");
  }
}

async function handleUnpivot(): Promise<void> {
  const checkboxes = getHeaderCheckboxes();
  if (checkboxes.length === 0) {
    setStatus("Load a selection first.", "error");
    return;
  }

  const idColumnIndices = checkboxes
    .filter((checkbox) => checkbox.checked)
    .map(getCheckboxIndex)
    .filter((index) => index >= 0);
  const unpivotColumnIndices = checkboxes
    .filter((checkbox) => !checkbox.checked)
    .map(getCheckboxIndex)
    .filter((index) => index >= 0);

  if (unpivotColumnIndices.length === 0) {
    setStatus("Select at least one column to unpivot.", "error");
    return;
  }

  setStatus("Unpivoting...", "info");

  try {
    const result = await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "rowCount", "columnCount"]);

      const worksheets = context.workbook.worksheets;
      worksheets.load("items/name");

      await context.sync();

      if (range.rowCount < 2) {
        throw new Error("Select a range with a header row and at least one data row.");
      }

      const headers = normalizeHeaders(range.values[0]);
      if (headers.length !== checkboxes.length) {
        throw new Error("Selection changed. Click Load Selection again.");
      }

      if (
        idColumnIndices.some((index) => index >= headers.length) ||
        unpivotColumnIndices.some((index) => index >= headers.length)
      ) {
        throw new Error("Selection changed. Click Load Selection again.");
      }

      const outputRowCount = 1 + (range.rowCount - 1) * unpivotColumnIndices.length;
      if (outputRowCount > MAX_EXCEL_ROWS) {
        throw new Error(
          `Result has ${outputRowCount} rows, which exceeds Excel's limit of ${MAX_EXCEL_ROWS}.`
        );
      }

      const values = range.values;
      const output: Excel.RangeValueType[][] = [];
      const idHeaders = idColumnIndices.map((index) => headers[index]);
      output.push([...idHeaders, "Attribute", "Value"]);

      for (let rowIndex = 1; rowIndex < range.rowCount; rowIndex++) {
        for (const columnIndex of unpivotColumnIndices) {
          const row: Excel.RangeValueType[] = [];
          for (const idIndex of idColumnIndices) {
            row.push(normalizeUnpivotIdValue(values[rowIndex][idIndex]));
          }
          row.push(headers[columnIndex], values[rowIndex][columnIndex]);
          output.push(row);
        }
      }

      const existingNames = worksheets.items.map((sheet) => sheet.name);
      const targetName = getUniqueSheetName(existingNames, "Unpivot");
      const targetSheet = worksheets.add(targetName);
      const outputRange = targetSheet.getRangeByIndexes(0, 0, output.length, output[0].length);
      outputRange.values = output;
      targetSheet.activate();

      return { sheetName: targetName, rowCount: output.length };
    });

    setStatus(`Created "${result.sheetName}" with ${result.rowCount} rows.`, "info");
  } catch (error) {
    setStatus(getErrorMessage(error), "error");
  }
}

async function handleSimplifyText(): Promise<void> {
  const fontName = modelFontNameInputEl.value.trim();
  if (!fontName) {
    setStatus("Font name is required.", "error");
    return;
  }

  const fontColor = modelFontColorInputEl.value.trim();
  if (!isValidHexColor(fontColor)) {
    setStatus("Font color must be a valid hex value (e.g., #000000).", "error");
    return;
  }

  const fontSize = parsePositiveNumber(modelFontSizeInputEl.value);
  if (fontSize === null) {
    setStatus("Font size must be a number greater than 0.", "error");
    return;
  }

  setStatus("Simplifying text...", "info");

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values"]);
      await context.sync();

      range.values = range.values.map((row) =>
        row.map((value) => (typeof value === "string" ? value.trim() : value))
      );

      range.format.font.name = fontName;
      range.format.font.color = fontColor;
      range.format.font.size = fontSize;
      range.format.font.bold = false;
      range.format.horizontalAlignment = Excel.HorizontalAlignment.left;
      range.format.indentLevel = 0;
      range.format.fill.clear();
      clearBorders(range.format.borders);

      await context.sync();
    });

    setStatus("Simplified text.", "info");
  } catch (error) {
    setStatus(getErrorMessage(error), "error");
  }
}

async function handleLoadLineItemCategories(): Promise<void> {
  setStatus("Loading line item categories...", "info");

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
      await context.sync();

      if (range.columnCount !== 1 || range.columnIndex !== 1) {
        throw new Error("Select a single-column range in column B.");
      }

      if (range.rowCount < 1) {
        throw new Error("Select at least one row in column B.");
      }

      const targetRange = range.getOffsetRange(0, 3);
      const defaultValues = Array.from({ length: range.rowCount }, () => ["Detail"]);
      targetRange.values = defaultValues;
      targetRange.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: "Detail,SubTotal,Total,GrandTotal,SubTotal Header,Total Header,GrandTotal Header",
        },
      };
      targetRange.format.font.color = "#FFFFFF";
      targetRange.format.fill.color = "#782170";

      await context.sync();
    });

    setStatus("Loaded line item categories.", "info");
  } catch (error) {
    setStatus(getErrorMessage(error), "error");
  }
}

async function handleUpdateFsFormatting(): Promise<void> {
  const timelineColumns = parsePositiveInt(timelineColumnsInputEl.value);
  if (timelineColumns === null) {
    setStatus("Columns for timeline must be a whole number of 1 or greater.", "error");
    return;
  }

  const totalColumns = DEFAULT_CONSTANTS_COLUMNS + timelineColumns;
  if (totalColumns > MAX_EXCEL_COLUMNS) {
    setStatus("Total model columns exceed Excel column limits.", "error");
    return;
  }

  const rulesResult = getFsFormattingRules();
  if (!rulesResult.ok) {
    setStatus(rulesResult.error, "error");
    return;
  }

  setStatus("Updating FS formatting...", "info");

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
      const sheet = range.worksheet;
      await context.sync();

      if (range.columnCount !== 1 || range.columnIndex !== 1) {
        throw new Error("Select a single-column range in column B.");
      }

      if (range.rowCount < 1) {
        throw new Error("Select at least one row in column B.");
      }

      const categoryRange = sheet.getRangeByIndexes(range.rowIndex, 4, range.rowCount, 1);
      categoryRange.load("values");
      await context.sync();

      const rules = rulesResult.rules;
      const unknownCategories: string[] = [];
      for (let rowOffset = 0; rowOffset < range.rowCount; rowOffset += 1) {
        const rawValue = categoryRange.values[rowOffset][0];
        const category = rawValue === null || rawValue === undefined ? "" : String(rawValue).trim();
        const rule = rules.get(category);
        if (!rule) {
          const label = category.length > 0 ? category : "(blank)";
          if (!unknownCategories.includes(label)) {
            unknownCategories.push(label);
          }
          continue;
        }

        const rowIndex = range.rowIndex + rowOffset;
        applyFsFormattingForRow(sheet, {
          rowIndex,
          totalColumns,
          timelineStartColumn: DEFAULT_CONSTANTS_COLUMNS,
          timelineColumns,
          rule,
        });
      }

      if (unknownCategories.length > 0) {
        throw new Error(
          `Unknown line item categories in column E: ${unknownCategories.join(", ")}.`
        );
      }

      await context.sync();
    });

    setStatus("Updated FS formatting.", "info");
  } catch (error) {
    setStatus(getErrorMessage(error), "error");
  }
}

async function handleCreateControlsSheet(): Promise<void> {
  const result = getControlsSheetSpecFromForm();
  if (!result.ok) {
    setStatus(result.error, "error");
    return;
  }

  setStatus('Creating "Controls" sheet...', "info");

  try {
    await createControlsSheet(result.spec);
    setStatus('Created "Controls" sheet.', "info");
  } catch (error) {
    setStatus(getErrorMessage(error), "error");
  }
}

async function handleCreateMonthlySheet(): Promise<void> {
  const result = getMonthlySheetSpecFromForm();
  if (!result.ok) {
    setStatus(result.error, "error");
    return;
  }

  setStatus('Creating "Monthly" sheet...', "info");

  try {
    await createMonthlySheet(result.spec);
    setStatus('Created "Monthly" sheet.', "info");
  } catch (error) {
    setStatus(getErrorMessage(error), "error");
  }
}

async function handleCreateFsMonthlySheet(): Promise<void> {
  const result = getFsMonthlySheetSpecFromForm();
  if (!result.ok) {
    setStatus(result.error, "error");
    return;
  }

  setStatus('Creating "FS Monthly" sheet...', "info");

  try {
    await createFsMonthlySheet(result.spec);
    setStatus('Created "FS Monthly" sheet.', "info");
  } catch (error) {
    setStatus(getErrorMessage(error), "error");
  }
}

async function handleCreateQuarterlySheet(): Promise<void> {
  const result = getQuarterlySheetSpecFromForm();
  if (!result.ok) {
    setStatus(result.error, "error");
    return;
  }

  setStatus('Creating "Quarterly" sheet...', "info");

  try {
    await createQuarterlySheet(result.spec);
    setStatus('Created "Quarterly" sheet.', "info");
  } catch (error) {
    setStatus(getErrorMessage(error), "error");
  }
}

async function handleCreateAnnualSheet(): Promise<void> {
  const result = getAnnualSheetSpecFromForm();
  if (!result.ok) {
    setStatus(result.error, "error");
    return;
  }

  setStatus('Creating "Annual" sheet...', "info");

  try {
    await createAnnualSheet(result.spec);
    setStatus('Created "Annual" sheet.', "info");
  } catch (error) {
    setStatus(getErrorMessage(error), "error");
  }
}

type ControlsSheetFormResult =
  | { ok: true; spec: ControlsSheetSpec }
  | { ok: false; error: string };

function getControlsSheetSpecFromForm(): ControlsSheetFormResult {
  const constantsColumns = DEFAULT_CONSTANTS_COLUMNS;

  const timelineColumns = parsePositiveInt(timelineColumnsInputEl.value);
  if (timelineColumns === null) {
    return { ok: false, error: "Columns for timeline must be a whole number of 1 or greater." };
  }

  const totalColumns = constantsColumns + timelineColumns;
  if (totalColumns > MAX_EXCEL_COLUMNS) {
    return { ok: false, error: "Total model columns exceed Excel column limits." };
  }

  const fontName = modelFontNameInputEl.value.trim();
  if (!fontName) {
    return { ok: false, error: "Font name is required." };
  }

  const fontColor = modelFontColorInputEl.value.trim();
  if (!isValidHexColor(fontColor)) {
    return { ok: false, error: "Font color must be a valid hex value (e.g., #000000)." };
  }

  const fontSize = parsePositiveNumber(modelFontSizeInputEl.value);
  if (fontSize === null) {
    return { ok: false, error: "Font size must be a number greater than 0." };
  }

  const headerRows = parsePositiveInt(modelHeaderRowsInputEl.value);
  if (headerRows === null) {
    return { ok: false, error: "Number of rows for sheet header must be 1 or greater." };
  }

  const headerFillColor = modelHeaderBackgroundInputEl.value.trim();
  if (!isValidHexColor(headerFillColor)) {
    return { ok: false, error: "Sheet header background must be a valid hex value." };
  }

  const timeHeaderStartCell = DEFAULT_TIME_HEADER_START_CELL;

  const parsedStartCell = parseA1CellAddress(timeHeaderStartCell);
  if (!parsedStartCell) {
    return { ok: false, error: "TIME header start cell must be a single A1 address like A7." };
  }

  const timeHeaderTitle = timeHeaderTitleInputEl.value.trim();
  if (!timeHeaderTitle) {
    return { ok: false, error: "TIME header title is required." };
  }

  const timeHeaderRows = DEFAULT_TIME_HEADER_ROWS;

  const timeHeaderColumns = totalColumns;

  const timeHeaderEndRow = parsedStartCell.row + timeHeaderRows - 1;
  const timeHeaderEndColumn = parsedStartCell.column + timeHeaderColumns - 1;
  if (timeHeaderEndRow > MAX_EXCEL_ROWS || timeHeaderEndColumn > MAX_EXCEL_COLUMNS) {
    return { ok: false, error: "TIME header range exceeds worksheet limits." };
  }

  const timeHeaderFill = timeHeaderFillInputEl.value.trim();
  if (!isValidHexColor(timeHeaderFill)) {
    return { ok: false, error: "TIME header background must be a valid hex value." };
  }

  const timeHeaderFontColor = timeHeaderFontColorInputEl.value.trim();
  if (!isValidHexColor(timeHeaderFontColor)) {
    return { ok: false, error: "TIME header font color must be a valid hex value." };
  }

  const flagsHeaderTitle = flagsHeaderTitleInputEl.value.trim();
  if (!flagsHeaderTitle) {
    return { ok: false, error: "FLAGS header title is required." };
  }

  const flagsHeaderFill = flagsHeaderFillInputEl.value.trim();
  if (!isValidHexColor(flagsHeaderFill)) {
    return { ok: false, error: "FLAGS header background must be a valid hex value." };
  }

  const flagsHeaderFontColor = flagsHeaderFontColorInputEl.value.trim();
  if (!isValidHexColor(flagsHeaderFontColor)) {
    return { ok: false, error: "FLAGS header font color must be a valid hex value." };
  }

  const constantsStartRow = DEFAULT_CONSTANTS_START_ROW;
  if (constantsStartRow + 8 > MAX_EXCEL_ROWS) {
    return { ok: false, error: "Constants block exceeds worksheet row limits." };
  }

  const constantsTimelineLength = parsePositiveInt(constantsTimelineLengthInputEl.value);
  if (constantsTimelineLength === null) {
    return { ok: false, error: "Timeline length must be a whole number of 1 or greater." };
  }

  const financialYearEndMonth = parsePositiveInt(constantsYearEndMonthInputEl.value);
  if (financialYearEndMonth === null || financialYearEndMonth > 12) {
    return { ok: false, error: "Financial Year End month must be between 1 and 12." };
  }

  const timelineStartDate = parseDateInput(constantsStartDateInputEl.value);
  if (!timelineStartDate) {
    return { ok: false, error: "Timeline Start Date must be a valid date." };
  }

  const actualsEndDate = parseDateInput(constantsActualsEndInputEl.value);
  if (!actualsEndDate) {
    return { ok: false, error: "Actuals End Date must be a valid date." };
  }

  const historicalPeriod = constantsHistoricalPeriodInputEl.value.trim();
  const forecastPeriod = constantsForecastPeriodInputEl.value.trim();

  const widthA = parseNonNegativeNumber(widthColAInputEl.value);
  const widthB = parseNonNegativeNumber(widthColBInputEl.value);
  const widthC = parseNonNegativeNumber(widthColCInputEl.value);
  const widthD = parseNonNegativeNumber(widthColDInputEl.value);
  const widthE = parseNonNegativeNumber(widthColEInputEl.value);
  const widthF = parseNonNegativeNumber(widthColFInputEl.value);
  const widthG = parseNonNegativeNumber(widthColGInputEl.value);
  const widthH = parseNonNegativeNumber(widthColHInputEl.value);
  const widthI = parseNonNegativeNumber(widthColIInputEl.value);
  const widthJ = parseNonNegativeNumber(widthColJInputEl.value);

  if (
    widthA === null ||
    widthB === null ||
    widthC === null ||
    widthD === null ||
    widthE === null ||
    widthF === null ||
    widthG === null ||
    widthH === null ||
    widthI === null ||
    widthJ === null
  ) {
    return { ok: false, error: "Column widths must be numbers of 0 or greater." };
  }

  const tabColor = controlsTabColorInputEl.value.trim();
  if (!isValidHexColor(tabColor)) {
    return { ok: false, error: "Tab color must be a valid hex value." };
  }

  return {
    ok: true,
    spec: {
      constantsColumns,
      timelineColumns,
      tabColor,
      font: {
        name: fontName,
        color: fontColor,
        size: fontSize,
      },
      headerRows,
      headerFillColor,
      columnWidths: {
        A: widthA,
        B: widthB,
        C: widthC,
        D: widthD,
        E: widthE,
        F: widthF,
        G: widthG,
        H: widthH,
        I: widthI,
        J: widthJ,
      },
      timeHeader: {
        startCell: timeHeaderStartCell,
        title: timeHeaderTitle,
        rows: timeHeaderRows,
        columns: timeHeaderColumns,
        fillColor: timeHeaderFill,
        fontColor: timeHeaderFontColor,
      },
      flagsHeader: {
        startCell: DEFAULT_FLAGS_HEADER_START_CELL,
        title: flagsHeaderTitle,
        rows: DEFAULT_FLAGS_HEADER_ROWS,
        columns: totalColumns,
        fillColor: flagsHeaderFill,
        fontColor: flagsHeaderFontColor,
      },
      constantsBlock: {
        startRow: constantsStartRow,
        timelineStartDate,
        actualsEndDate,
        timelineLength: constantsTimelineLength,
        financialYearEndMonth,
        historicalPeriod,
        forecastPeriod,
      },
    },
  };
}

type MonthlySheetFormResult =
  | { ok: true; spec: MonthlySheetSpec }
  | { ok: false; error: string };

function getMonthlySheetSpecFromForm(): MonthlySheetFormResult {
  const constantsColumns = DEFAULT_CONSTANTS_COLUMNS;

  const timelineColumns = parsePositiveInt(timelineColumnsInputEl.value);
  if (timelineColumns === null) {
    return { ok: false, error: "Columns for timeline must be a whole number of 1 or greater." };
  }

  const totalColumns = constantsColumns + timelineColumns;
  if (totalColumns > MAX_EXCEL_COLUMNS) {
    return { ok: false, error: "Total model columns exceed Excel column limits." };
  }

  const fontName = modelFontNameInputEl.value.trim();
  if (!fontName) {
    return { ok: false, error: "Font name is required." };
  }

  const fontColor = modelFontColorInputEl.value.trim();
  if (!isValidHexColor(fontColor)) {
    return { ok: false, error: "Font color must be a valid hex value (e.g., #000000)." };
  }

  const fontSize = parsePositiveNumber(modelFontSizeInputEl.value);
  if (fontSize === null) {
    return { ok: false, error: "Font size must be a number greater than 0." };
  }

  const headerRows = parsePositiveInt(modelHeaderRowsInputEl.value);
  if (headerRows === null) {
    return { ok: false, error: "Number of rows for sheet header must be 1 or greater." };
  }

  const headerFillColor = modelHeaderBackgroundInputEl.value.trim();
  if (!isValidHexColor(headerFillColor)) {
    return { ok: false, error: "Sheet header background must be a valid hex value." };
  }

  const widthA = parseNonNegativeNumber(widthColAInputEl.value);
  const widthB = parseNonNegativeNumber(widthColBInputEl.value);
  const widthC = parseNonNegativeNumber(widthColCInputEl.value);
  const widthD = parseNonNegativeNumber(widthColDInputEl.value);
  const widthE = parseNonNegativeNumber(widthColEInputEl.value);
  const widthF = parseNonNegativeNumber(widthColFInputEl.value);
  const widthG = parseNonNegativeNumber(widthColGInputEl.value);
  const widthH = parseNonNegativeNumber(widthColHInputEl.value);
  const widthI = parseNonNegativeNumber(widthColIInputEl.value);
  const widthJ = parseNonNegativeNumber(widthColJInputEl.value);

  if (
    widthA === null ||
    widthB === null ||
    widthC === null ||
    widthD === null ||
    widthE === null ||
    widthF === null ||
    widthG === null ||
    widthH === null ||
    widthI === null ||
    widthJ === null
  ) {
    return { ok: false, error: "Column widths must be numbers of 0 or greater." };
  }

  const tabColor = monthlyTabColorInputEl.value.trim();
  if (!isValidHexColor(tabColor)) {
    return { ok: false, error: "Tab color must be a valid hex value." };
  }

  const sectionColor = monthlySectionColorInputEl.value.trim();
  if (!isValidHexColor(sectionColor)) {
    return { ok: false, error: "Section color must be a valid hex value." };
  }

  return {
    ok: true,
    spec: {
      constantsColumns,
      timelineColumns,
      tabColor,
      sectionColor,
      font: {
        name: fontName,
        color: fontColor,
        size: fontSize,
      },
      headerRows,
      headerFillColor,
      columnWidths: {
        A: widthA,
        B: widthB,
        C: widthC,
        D: widthD,
        E: widthE,
        F: widthF,
        G: widthG,
        H: widthH,
        I: widthI,
        J: widthJ,
      },
    },
  };
}

type FsMonthlySheetFormResult =
  | { ok: true; spec: FsMonthlySheetSpec }
  | { ok: false; error: string };

function getFsMonthlySheetSpecFromForm(): FsMonthlySheetFormResult {
  const constantsColumns = DEFAULT_CONSTANTS_COLUMNS;

  const timelineColumns = parsePositiveInt(timelineColumnsInputEl.value);
  if (timelineColumns === null) {
    return { ok: false, error: "Columns for timeline must be a whole number of 1 or greater." };
  }

  const totalColumns = constantsColumns + timelineColumns;
  if (totalColumns > MAX_EXCEL_COLUMNS) {
    return { ok: false, error: "Total model columns exceed Excel column limits." };
  }

  const fontName = modelFontNameInputEl.value.trim();
  if (!fontName) {
    return { ok: false, error: "Font name is required." };
  }

  const fontColor = modelFontColorInputEl.value.trim();
  if (!isValidHexColor(fontColor)) {
    return { ok: false, error: "Font color must be a valid hex value (e.g., #000000)." };
  }

  const fontSize = parsePositiveNumber(modelFontSizeInputEl.value);
  if (fontSize === null) {
    return { ok: false, error: "Font size must be a number greater than 0." };
  }

  const headerRows = parsePositiveInt(modelHeaderRowsInputEl.value);
  if (headerRows === null) {
    return { ok: false, error: "Number of rows for sheet header must be 1 or greater." };
  }

  const headerFillColor = modelHeaderBackgroundInputEl.value.trim();
  if (!isValidHexColor(headerFillColor)) {
    return { ok: false, error: "Sheet header background must be a valid hex value." };
  }

  const widthA = parseNonNegativeNumber(widthColAInputEl.value);
  const widthB = parseNonNegativeNumber(widthColBInputEl.value);
  const widthC = parseNonNegativeNumber(widthColCInputEl.value);
  const widthD = parseNonNegativeNumber(widthColDInputEl.value);
  const widthE = parseNonNegativeNumber(widthColEInputEl.value);
  const widthF = parseNonNegativeNumber(widthColFInputEl.value);
  const widthG = parseNonNegativeNumber(widthColGInputEl.value);
  const widthH = parseNonNegativeNumber(widthColHInputEl.value);
  const widthI = parseNonNegativeNumber(widthColIInputEl.value);
  const widthJ = parseNonNegativeNumber(widthColJInputEl.value);

  if (
    widthA === null ||
    widthB === null ||
    widthC === null ||
    widthD === null ||
    widthE === null ||
    widthF === null ||
    widthG === null ||
    widthH === null ||
    widthI === null ||
    widthJ === null
  ) {
    return { ok: false, error: "Column widths must be numbers of 0 or greater." };
  }

  const tabColor = fsMonthlyTabColorInputEl.value.trim();
  if (!isValidHexColor(tabColor)) {
    return { ok: false, error: "Tab color must be a valid hex value." };
  }

  const sectionColor = fsMonthlySectionColorInputEl.value.trim();
  if (!isValidHexColor(sectionColor)) {
    return { ok: false, error: "Section color must be a valid hex value." };
  }

  return {
    ok: true,
    spec: {
      constantsColumns,
      timelineColumns,
      tabColor,
      sectionColor,
      font: {
        name: fontName,
        color: fontColor,
        size: fontSize,
      },
      headerRows,
      headerFillColor,
      columnWidths: {
        A: widthA,
        B: widthB,
        C: widthC,
        D: widthD,
        E: widthE,
        F: widthF,
        G: widthG,
        H: widthH,
        I: widthI,
        J: widthJ,
      },
    },
  };
}

type QuarterlySheetFormResult =
  | { ok: true; spec: QuarterlySheetSpec }
  | { ok: false; error: string };

function getQuarterlySheetSpecFromForm(): QuarterlySheetFormResult {
  const constantsColumns = DEFAULT_CONSTANTS_COLUMNS;

  const modelTimelineColumns = parsePositiveInt(timelineColumnsInputEl.value);
  if (modelTimelineColumns === null) {
    return { ok: false, error: "Columns for timeline must be a whole number of 1 or greater." };
  }

  const timelineColumns = Math.max(1, Math.floor(modelTimelineColumns / 3));
  const totalColumns = constantsColumns + timelineColumns;
  if (totalColumns > MAX_EXCEL_COLUMNS) {
    return { ok: false, error: "Total model columns exceed Excel column limits." };
  }

  const fontName = modelFontNameInputEl.value.trim();
  if (!fontName) {
    return { ok: false, error: "Font name is required." };
  }

  const fontColor = modelFontColorInputEl.value.trim();
  if (!isValidHexColor(fontColor)) {
    return { ok: false, error: "Font color must be a valid hex value (e.g., #000000)." };
  }

  const fontSize = parsePositiveNumber(modelFontSizeInputEl.value);
  if (fontSize === null) {
    return { ok: false, error: "Font size must be a number greater than 0." };
  }

  const headerRows = parsePositiveInt(modelHeaderRowsInputEl.value);
  if (headerRows === null) {
    return { ok: false, error: "Number of rows for sheet header must be 1 or greater." };
  }

  const headerFillColor = modelHeaderBackgroundInputEl.value.trim();
  if (!isValidHexColor(headerFillColor)) {
    return { ok: false, error: "Sheet header background must be a valid hex value." };
  }

  const widthA = parseNonNegativeNumber(widthColAInputEl.value);
  const widthB = parseNonNegativeNumber(widthColBInputEl.value);
  const widthC = parseNonNegativeNumber(widthColCInputEl.value);
  const widthD = parseNonNegativeNumber(widthColDInputEl.value);
  const widthE = parseNonNegativeNumber(widthColEInputEl.value);
  const widthF = parseNonNegativeNumber(widthColFInputEl.value);
  const widthG = parseNonNegativeNumber(widthColGInputEl.value);
  const widthH = parseNonNegativeNumber(widthColHInputEl.value);
  const widthI = parseNonNegativeNumber(widthColIInputEl.value);
  const widthJ = parseNonNegativeNumber(widthColJInputEl.value);

  if (
    widthA === null ||
    widthB === null ||
    widthC === null ||
    widthD === null ||
    widthE === null ||
    widthF === null ||
    widthG === null ||
    widthH === null ||
    widthI === null ||
    widthJ === null
  ) {
    return { ok: false, error: "Column widths must be numbers of 0 or greater." };
  }

  const tabColor = quarterlyTabColorInputEl.value.trim();
  if (!isValidHexColor(tabColor)) {
    return { ok: false, error: "Tab color must be a valid hex value." };
  }

  const sectionColor = quarterlySectionColorInputEl.value.trim();
  if (!isValidHexColor(sectionColor)) {
    return { ok: false, error: "Section color must be a valid hex value." };
  }

  return {
    ok: true,
    spec: {
      constantsColumns,
      timelineColumns,
      tabColor,
      sectionColor,
      font: {
        name: fontName,
        color: fontColor,
        size: fontSize,
      },
      headerRows,
      headerFillColor,
      columnWidths: {
        A: widthA,
        B: widthB,
        C: widthC,
        D: widthD,
        E: widthE,
        F: widthF,
        G: widthG,
        H: widthH,
        I: widthI,
        J: widthJ,
      },
    },
  };
}

type AnnualSheetFormResult =
  | { ok: true; spec: AnnualSheetSpec }
  | { ok: false; error: string };

function getAnnualSheetSpecFromForm(): AnnualSheetFormResult {
  const constantsColumns = DEFAULT_CONSTANTS_COLUMNS;

  const modelTimelineColumns = parsePositiveInt(timelineColumnsInputEl.value);
  if (modelTimelineColumns === null) {
    return { ok: false, error: "Columns for timeline must be a whole number of 1 or greater." };
  }

  const timelineColumns = Math.max(1, Math.floor(modelTimelineColumns / 12));
  const totalColumns = constantsColumns + timelineColumns;
  if (totalColumns > MAX_EXCEL_COLUMNS) {
    return { ok: false, error: "Total model columns exceed Excel column limits." };
  }

  const fontName = modelFontNameInputEl.value.trim();
  if (!fontName) {
    return { ok: false, error: "Font name is required." };
  }

  const fontColor = modelFontColorInputEl.value.trim();
  if (!isValidHexColor(fontColor)) {
    return { ok: false, error: "Font color must be a valid hex value (e.g., #000000)." };
  }

  const fontSize = parsePositiveNumber(modelFontSizeInputEl.value);
  if (fontSize === null) {
    return { ok: false, error: "Font size must be a number greater than 0." };
  }

  const headerRows = parsePositiveInt(modelHeaderRowsInputEl.value);
  if (headerRows === null) {
    return { ok: false, error: "Number of rows for sheet header must be 1 or greater." };
  }

  const headerFillColor = modelHeaderBackgroundInputEl.value.trim();
  if (!isValidHexColor(headerFillColor)) {
    return { ok: false, error: "Sheet header background must be a valid hex value." };
  }

  const widthA = parseNonNegativeNumber(widthColAInputEl.value);
  const widthB = parseNonNegativeNumber(widthColBInputEl.value);
  const widthC = parseNonNegativeNumber(widthColCInputEl.value);
  const widthD = parseNonNegativeNumber(widthColDInputEl.value);
  const widthE = parseNonNegativeNumber(widthColEInputEl.value);
  const widthF = parseNonNegativeNumber(widthColFInputEl.value);
  const widthG = parseNonNegativeNumber(widthColGInputEl.value);
  const widthH = parseNonNegativeNumber(widthColHInputEl.value);
  const widthI = parseNonNegativeNumber(widthColIInputEl.value);
  const widthJ = parseNonNegativeNumber(widthColJInputEl.value);

  if (
    widthA === null ||
    widthB === null ||
    widthC === null ||
    widthD === null ||
    widthE === null ||
    widthF === null ||
    widthG === null ||
    widthH === null ||
    widthI === null ||
    widthJ === null
  ) {
    return { ok: false, error: "Column widths must be numbers of 0 or greater." };
  }

  const tabColor = annualTabColorInputEl.value.trim();
  if (!isValidHexColor(tabColor)) {
    return { ok: false, error: "Tab color must be a valid hex value." };
  }

  const sectionColor = annualSectionColorInputEl.value.trim();
  if (!isValidHexColor(sectionColor)) {
    return { ok: false, error: "Section color must be a valid hex value." };
  }

  return {
    ok: true,
    spec: {
      constantsColumns,
      timelineColumns,
      tabColor,
      sectionColor,
      font: {
        name: fontName,
        color: fontColor,
        size: fontSize,
      },
      headerRows,
      headerFillColor,
      columnWidths: {
        A: widthA,
        B: widthB,
        C: widthC,
        D: widthD,
        E: widthE,
        F: widthF,
        G: widthG,
        H: widthH,
        I: widthI,
        J: widthJ,
      },
    },
  };
}

function parsePositiveInt(value: string): number | null {
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < 1 || !Number.isInteger(parsed)) {
    return null;
  }

  return parsed;
}

function parsePositiveNumber(value: string): number | null {
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return null;
  }

  return parsed;
}

function parseNonNegativeNumber(value: string): number | null {
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < 0) {
    return null;
  }

  return parsed;
}

function parseNonNegativeInt(value: string): number | null {
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < 0 || !Number.isInteger(parsed)) {
    return null;
  }

  return parsed;
}

type FsFormattingRule = {
  bold: boolean;
  indent: number;
  fontColor: string;
  fillColor: string | null;
  borderLine: FsBorderLine;
  numberFormat: string;
};

type FsBorderLine = "none" | "top-single" | "top-double" | "top-double-bottom-double";

type FsFormattingRulesResult =
  | { ok: true; rules: Map<string, FsFormattingRule> }
  | { ok: false; error: string };

function getFsFormattingRules(): FsFormattingRulesResult {
  const ruleEntries: Array<[string, { prefix: string; label: string }]> = [
    ["Detail", { prefix: "fs-detail", label: "Detail" }],
    ["SubTotal", { prefix: "fs-subtotal", label: "SubTotal" }],
    ["Total", { prefix: "fs-total", label: "Total" }],
    ["GrandTotal", { prefix: "fs-grandtotal", label: "GrandTotal" }],
    ["SubTotal Header", { prefix: "fs-subtotal-header", label: "SubTotal Header" }],
    ["Total Header", { prefix: "fs-total-header", label: "Total Header" }],
    ["GrandTotal Header", { prefix: "fs-grandtotal-header", label: "GrandTotal Header" }],
  ];

  const rules = new Map<string, FsFormattingRule>();
  for (const [category, { prefix, label }] of ruleEntries) {
    const result = getFsFormattingRule(prefix, label);
    if (!result.ok) {
      return result;
    }
    rules.set(category, result.rule);
  }

  return { ok: true, rules };
}

type FsFormattingRuleResult =
  | { ok: true; rule: FsFormattingRule }
  | { ok: false; error: string };

function getFsFormattingRule(prefix: string, label: string): FsFormattingRuleResult {
  const boldValue = (document.getElementById(`${prefix}-bold`) as HTMLSelectElement).value;
  const bold = boldValue === "true";

  const indentValue = (document.getElementById(`${prefix}-indent`) as HTMLInputElement).value;
  const indent = parseNonNegativeInt(indentValue);
  if (indent === null) {
    return { ok: false, error: `${label} Ident must be a whole number of 0 or greater.` };
  }

  const fontColor = (document.getElementById(`${prefix}-font-color`) as HTMLInputElement).value.trim();
  if (!isValidHexColor(fontColor)) {
    return { ok: false, error: `${label} font color must be a valid hex value (e.g., #000000).` };
  }

  const fillColorValue = (
    document.getElementById(`${prefix}-fill-color`) as HTMLInputElement
  ).value.trim();
  if (fillColorValue.length > 0 && !isValidHexColor(fillColorValue)) {
    return { ok: false, error: `${label} fill color must be empty or a valid hex value.` };
  }

  const borderLine = (document.getElementById(
    `${prefix}-border-line`
  ) as HTMLSelectElement).value as FsBorderLine;
  if (!isValidFsBorderLine(borderLine)) {
    return { ok: false, error: `${label} border line must be a supported option.` };
  }

  const numberFormat = (document.getElementById(
    `${prefix}-number-format`
  ) as HTMLInputElement).value;

  return {
    ok: true,
    rule: {
      bold,
      indent,
      fontColor,
      fillColor: fillColorValue.length > 0 ? fillColorValue : null,
      borderLine,
      numberFormat,
    },
  };
}

function isValidFsBorderLine(value: string): value is FsBorderLine {
  return (
    value === "none" ||
    value === "top-single" ||
    value === "top-double" ||
    value === "top-double-bottom-double"
  );
}

type FsFormattingRowInput = {
  rowIndex: number;
  totalColumns: number;
  timelineStartColumn: number;
  timelineColumns: number;
  rule: FsFormattingRule;
};

function applyFsFormattingForRow(sheet: Excel.Worksheet, input: FsFormattingRowInput): void {
  const { rowIndex, totalColumns, timelineStartColumn, timelineColumns, rule } = input;
  const rowRange = sheet.getRangeByIndexes(rowIndex, 1, 1, totalColumns - 1);
  rowRange.format.font.bold = rule.bold;
  rowRange.format.font.color = rule.fontColor;

  if (rule.fillColor) {
    rowRange.format.fill.color = rule.fillColor;
  } else {
    rowRange.format.fill.clear();
  }

  applyFsBorderLine(rowRange.format.borders, rule.borderLine);

  const indentCell = sheet.getRangeByIndexes(rowIndex, 1, 1, 1);
  indentCell.format.indentLevel = rule.indent;

  const timelineRange = sheet.getRangeByIndexes(rowIndex, timelineStartColumn, 1, timelineColumns);
  timelineRange.numberFormat = [Array.from({ length: timelineColumns }, () => rule.numberFormat)];
}

function applyFsBorderLine(
  borders: Excel.RangeBorderCollection,
  borderLine: FsBorderLine
): void {
  const topBorder = borders.getItem(Excel.BorderIndex.edgeTop);
  const bottomBorder = borders.getItem(Excel.BorderIndex.edgeBottom);

  topBorder.style = Excel.BorderLineStyle.none;
  bottomBorder.style = Excel.BorderLineStyle.none;

  if (borderLine === "top-single") {
    topBorder.style = Excel.BorderLineStyle.continuous;
  } else if (borderLine === "top-double") {
    topBorder.style = Excel.BorderLineStyle.double;
  } else if (borderLine === "top-double-bottom-double") {
    topBorder.style = Excel.BorderLineStyle.double;
    bottomBorder.style = Excel.BorderLineStyle.double;
  }
}

type CellAddress = {
  row: number;
  column: number;
};

function parseA1CellAddress(value: string): CellAddress | null {
  if (value.includes(":") || value.includes("!")) {
    return null;
  }

  const match = /^([A-Z]{1,3})([1-9][0-9]*)$/.exec(value);
  if (!match) {
    return null;
  }

  const column = columnLettersToNumber(match[1]);
  const row = Number(match[2]);
  if (!Number.isFinite(row) || row < 1) {
    return null;
  }

  if (column < 1 || column > MAX_EXCEL_COLUMNS || row > MAX_EXCEL_ROWS) {
    return null;
  }

  return { row, column };
}

function columnLettersToNumber(letters: string): number {
  let value = 0;
  for (let i = 0; i < letters.length; i += 1) {
    value = value * 26 + (letters.charCodeAt(i) - 64);
  }

  return value;
}

function parseDateInput(value: string): Date | null {
  if (!value) {
    return null;
  }

  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) {
    return null;
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
    return null;
  }

  return new Date(year, month - 1, day);
}

function syncDerivedDefaults(force = false): void {
  const timelineColumns = parsePositiveInt(timelineColumnsInputEl.value);
  if (timelineColumns === null) {
    return;
  }

  const totalColumns = DEFAULT_CONSTANTS_COLUMNS + timelineColumns;
  if (force || !timelineLengthDirty) {
    constantsTimelineLengthInputEl.value = timelineColumns.toString();
    timelineLengthDirty = false;
  }
}

function isValidHexColor(value: string): boolean {
  return /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(value);
}

function normalizeHeaders(headerRow: Excel.RangeValueType[]): string[] {
  return headerRow.map((value, index) => {
    if (value === null || value === undefined) {
      return `Column ${index + 1}`;
    }

    const text = String(value).trim();
    return text.length > 0 ? text : `Column ${index + 1}`;
  });
}

function renderHeaders(headers: string[]): void {
  headerListEl.innerHTML = "";

  if (headers.length === 0) {
    const emptyState = document.createElement("div");
    emptyState.className = "empty-state";
    emptyState.textContent = "Load a selection to see headers.";
    headerListEl.appendChild(emptyState);
    return;
  }

  const headerCounts = new Map<string, number>();
  headers.forEach((header) => {
    headerCounts.set(header, (headerCounts.get(header) ?? 0) + 1);
  });

  headers.forEach((header, index) => {
    const row = document.createElement("label");
    row.className = "checkbox-row";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.className = "header-checkbox";
    checkbox.dataset.columnIndex = index.toString();

    const isRepeated = (headerCounts.get(header) ?? 0) > 1;
    checkbox.checked = index === 0 && !isRepeated;

    const labelText = document.createElement("span");
    labelText.textContent = header;

    row.appendChild(checkbox);
    row.appendChild(labelText);
    headerListEl.appendChild(row);
  });
}

function getHeaderCheckboxes(): HTMLInputElement[] {
  return Array.from(headerListEl.querySelectorAll<HTMLInputElement>("input.header-checkbox"));
}

function getCheckboxIndex(checkbox: HTMLInputElement): number {
  const indexValue = checkbox.dataset.columnIndex ?? "";
  const index = Number.parseInt(indexValue, 10);
  return Number.isNaN(index) ? -1 : index;
}

function getUniqueSheetName(existingNames: string[], baseName: string): string {
  const existingLower = existingNames.map((name) => name.toLowerCase());
  const baseLower = baseName.toLowerCase();

  if (!existingLower.includes(baseLower)) {
    return baseName;
  }

  let suffix = 2;
  while (existingLower.includes(`${baseName} ${suffix}`.toLowerCase())) {
    suffix += 1;
  }

  return `${baseName} ${suffix}`;
}

function normalizeUnpivotIdValue(value: Excel.RangeValueType): Excel.RangeValueType {
  if (typeof value === "string") {
    return value.trim();
  }

  return value;
}

function clearBorders(borders: Excel.RangeBorderCollection): void {
  const borderIndexes: Excel.BorderIndex[] = [
    Excel.BorderIndex.edgeTop,
    Excel.BorderIndex.edgeBottom,
    Excel.BorderIndex.edgeLeft,
    Excel.BorderIndex.edgeRight,
    Excel.BorderIndex.insideHorizontal,
    Excel.BorderIndex.insideVertical,
    Excel.BorderIndex.diagonalDown,
    Excel.BorderIndex.diagonalUp,
  ];

  borderIndexes.forEach((index) => {
    const border = borders.getItem(index);
    border.style = Excel.BorderLineStyle.none;
  });
}

function setStatus(message: string, kind: "info" | "error"): void {
  statusEl.textContent = message;
  statusEl.classList.remove("status--info", "status--error");

  if (message.length === 0) {
    return;
  }

  statusEl.classList.add(kind === "error" ? "status--error" : "status--info");
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return "Something went wrong. Please try again.";
}
