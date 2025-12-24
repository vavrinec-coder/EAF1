/* global document, Excel, Office */

import { createControlsSheet, ControlsSheetSpec } from "../templateBuilder/controlsSheet";

const MAX_EXCEL_ROWS = 1048576;
const MAX_EXCEL_COLUMNS = 16384;

let selectedRangeEl: HTMLSpanElement;
let headerListEl: HTMLDivElement;
let statusEl: HTMLDivElement;
let constantsColumnsInputEl: HTMLInputElement;
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

  constantsColumnsInputEl = document.getElementById("model-constants-columns") as HTMLInputElement;
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

  loadButton.addEventListener("click", () => {
    void handleLoadSelection();
  });
  unpivotButton.addEventListener("click", () => {
    void handleUnpivot();
  });
  createControlsButtonEl.addEventListener("click", () => {
    void handleCreateControlsSheet();
  });

  renderHeaders([]);
  setStatus("Ready. Select a range and click Load Selection.", "info");
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
            row.push(values[rowIndex][idIndex]);
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

type ControlsSheetFormResult =
  | { ok: true; spec: ControlsSheetSpec }
  | { ok: false; error: string };

function getControlsSheetSpecFromForm(): ControlsSheetFormResult {
  const constantsColumns = parsePositiveInt(constantsColumnsInputEl.value);
  if (constantsColumns === null) {
    return { ok: false, error: "Columns for constants must be a whole number of 1 or greater." };
  }

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

  return {
    ok: true,
    spec: {
      constantsColumns,
      timelineColumns,
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
