/* global document, Excel, Office */

import { insertSectionHeader, SectionHeaderSpec } from "../templateBuilder/sectionHeader";

const MAX_EXCEL_ROWS = 1048576;
const MAX_EXCEL_COLUMNS = 16384;

let selectedRangeEl: HTMLSpanElement;
let headerListEl: HTMLDivElement;
let statusEl: HTMLDivElement;
let sectionHeaderFormEl: HTMLFormElement;
let sectionHeaderToggleEl: HTMLButtonElement;
let sectionStartCellInputEl: HTMLInputElement;
let sectionStartCellErrorEl: HTMLDivElement;
let sectionTitleInputEl: HTMLInputElement;
let sectionRowsInputEl: HTMLInputElement;
let sectionColumnsInputEl: HTMLInputElement;
let sectionFillColorInputEl: HTMLInputElement;
let sectionFontNameInputEl: HTMLInputElement;
let sectionFontSizeInputEl: HTMLInputElement;
let sectionFontBoldInputEl: HTMLInputElement;
let sectionFontColorInputEl: HTMLInputElement;
let sectionHorizontalAlignEl: HTMLSelectElement;
let sectionVerticalAlignEl: HTMLSelectElement;
let sectionBorderEl: HTMLSelectElement;

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

  sectionHeaderToggleEl = document.getElementById("toggle-section-header") as HTMLButtonElement;
  sectionHeaderFormEl = document.getElementById("section-header-form") as HTMLFormElement;
  sectionStartCellInputEl = document.getElementById("section-start-cell") as HTMLInputElement;
  sectionStartCellErrorEl = document.getElementById("section-start-error") as HTMLDivElement;
  sectionTitleInputEl = document.getElementById("section-title") as HTMLInputElement;
  sectionRowsInputEl = document.getElementById("section-rows") as HTMLInputElement;
  sectionColumnsInputEl = document.getElementById("section-columns") as HTMLInputElement;
  sectionFillColorInputEl = document.getElementById("section-fill-color") as HTMLInputElement;
  sectionFontNameInputEl = document.getElementById("section-font-name") as HTMLInputElement;
  sectionFontSizeInputEl = document.getElementById("section-font-size") as HTMLInputElement;
  sectionFontBoldInputEl = document.getElementById("section-font-bold") as HTMLInputElement;
  sectionFontColorInputEl = document.getElementById("section-font-color") as HTMLInputElement;
  sectionHorizontalAlignEl = document.getElementById("section-horizontal-align") as HTMLSelectElement;
  sectionVerticalAlignEl = document.getElementById("section-vertical-align") as HTMLSelectElement;
  sectionBorderEl = document.getElementById("section-border") as HTMLSelectElement;

  loadButton.addEventListener("click", () => {
    void handleLoadSelection();
  });
  unpivotButton.addEventListener("click", () => {
    void handleUnpivot();
  });
  sectionHeaderToggleEl.addEventListener("click", () => {
    toggleSectionHeaderForm();
  });
  sectionHeaderFormEl.addEventListener("submit", (event) => {
    event.preventDefault();
    void handleInsertSectionHeader();
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

function toggleSectionHeaderForm(): void {
  sectionHeaderFormEl.hidden = !sectionHeaderFormEl.hidden;

  if (!sectionHeaderFormEl.hidden) {
    sectionStartCellInputEl.focus();
  }
}

async function handleInsertSectionHeader(): Promise<void> {
  const result = getSectionHeaderSpecFromForm();
  if (!result.ok) {
    setStatus(result.error, "error");
    return;
  }

  setStatus("Adding section header...", "info");

  try {
    await insertSectionHeader(result.spec);
    setStatus(
      `Added "${result.spec.title}" section header at ${result.spec.startCell}.`,
      "info"
    );
  } catch (error) {
    setStatus(getErrorMessage(error), "error");
  }
}

type SectionHeaderFormResult =
  | { ok: true; spec: SectionHeaderSpec }
  | { ok: false; error: string };

function getSectionHeaderSpecFromForm(): SectionHeaderFormResult {
  setStartCellError("");

  const startCellInput = sectionStartCellInputEl.value.trim().toUpperCase();
  if (!startCellInput) {
    setStartCellError("Enter a start cell like B7.");
    return { ok: false, error: "Start cell is required." };
  }

  const parsedStartCell = parseA1CellAddress(startCellInput);
  if (!parsedStartCell) {
    setStartCellError("Use a single-cell A1 address like B7.");
    return { ok: false, error: "Start cell must be a single-cell address like B7." };
  }

  const title = sectionTitleInputEl.value.trim();
  if (!title) {
    return { ok: false, error: "Enter a title for the section header." };
  }

  const rows = parsePositiveInt(sectionRowsInputEl.value);
  if (rows === null) {
    return { ok: false, error: "Rows must be a whole number of 1 or greater." };
  }

  const columns = parsePositiveInt(sectionColumnsInputEl.value);
  if (columns === null) {
    return { ok: false, error: "Columns must be a whole number of 1 or greater." };
  }

  const endRow = parsedStartCell.row + rows - 1;
  const endColumn = parsedStartCell.column + columns - 1;
  if (endRow > MAX_EXCEL_ROWS || endColumn > MAX_EXCEL_COLUMNS) {
    return {
      ok: false,
      error: "Section header would exceed worksheet limits. Adjust the start cell or size.",
    };
  }

  const fillColor = sectionFillColorInputEl.value.trim();
  if (!isValidHexColor(fillColor)) {
    return { ok: false, error: "Background color must be a valid hex value (e.g., #FFD966)." };
  }

  const fontName = sectionFontNameInputEl.value.trim();
  if (!fontName) {
    return { ok: false, error: "Font name is required." };
  }

  const fontSize = parsePositiveNumber(sectionFontSizeInputEl.value);
  if (fontSize === null) {
    return { ok: false, error: "Font size must be a number greater than 0." };
  }

  const fontColor = sectionFontColorInputEl.value.trim();
  if (!isValidHexColor(fontColor)) {
    return { ok: false, error: "Font color must be a valid hex value (e.g., #000000)." };
  }

  const horizontalValue = sectionHorizontalAlignEl.value;
  const horizontalAlignment = horizontalValue === "center" ? "center" : "left";

  const verticalValue = sectionVerticalAlignEl.value;
  const verticalAlignment = verticalValue === "center" ? "center" : "center";

  const borderValue = sectionBorderEl.value;
  const border = borderValue === "none" ? "none" : "thin";

  return {
    ok: true,
    spec: {
      startCell: startCellInput,
      title,
      rows,
      columns,
      fillColor,
      font: {
        name: fontName,
        size: fontSize,
        bold: sectionFontBoldInputEl.checked,
        color: fontColor,
      },
      alignment: {
        horizontal: horizontalAlignment,
        vertical: verticalAlignment,
      },
      border,
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

function setStartCellError(message: string): void {
  sectionStartCellErrorEl.textContent = message;
  sectionStartCellErrorEl.hidden = message.length === 0;
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
