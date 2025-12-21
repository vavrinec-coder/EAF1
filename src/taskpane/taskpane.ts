/* global document, Excel, Office */

const MAX_EXCEL_ROWS = 1048576;

let selectedRangeEl: HTMLSpanElement;
let headerListEl: HTMLDivElement;
let statusEl: HTMLDivElement;

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

  loadButton.addEventListener("click", () => {
    void handleLoadSelection();
  });
  unpivotButton.addEventListener("click", () => {
    void handleUnpivot();
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
