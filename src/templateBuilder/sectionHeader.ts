/* global Excel */

export type SectionHeaderSpec = {
  startCell: string;
  title: string;
  rows: number;
  columns: number;
  fillColor: string;
  font: {
    name: string;
    size: number;
    bold: boolean;
    color: string;
  };
  alignment: {
    horizontal: "left" | "center";
    vertical: "center";
  };
  border: "none" | "thin";
};

export async function insertSectionHeader(spec: SectionHeaderSpec): Promise<void> {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const anchorCell = worksheet.getRange(spec.startCell);
    const targetRange = anchorCell.getResizedRange(spec.rows - 1, spec.columns - 1);

    targetRange.format.fill.color = spec.fillColor;
    targetRange.format.font.name = spec.font.name;
    targetRange.format.font.size = spec.font.size;
    targetRange.format.font.bold = spec.font.bold;
    targetRange.format.font.color = spec.font.color;

    targetRange.format.horizontalAlignment =
      spec.alignment.horizontal === "center"
        ? Excel.HorizontalAlignment.center
        : Excel.HorizontalAlignment.left;
    targetRange.format.verticalAlignment = Excel.VerticalAlignment.center;

    const values: string[][] = Array.from({ length: spec.rows }, (_, rowIndex) =>
      Array.from({ length: spec.columns }, (_, columnIndex) =>
        rowIndex === 0 && columnIndex === 0 ? spec.title : ""
      )
    );
    targetRange.values = values;

    const edges = [
      Excel.BorderIndex.edgeTop,
      Excel.BorderIndex.edgeBottom,
      Excel.BorderIndex.edgeLeft,
      Excel.BorderIndex.edgeRight,
    ];

    if (spec.border === "thin") {
      edges.forEach((edge) => {
        const border = targetRange.format.borders.getItem(edge);
        border.style = Excel.BorderLineStyle.continuous;
        border.weight = Excel.BorderWeight.thin;
        border.color = "#000000";
      });
    } else {
      edges.forEach((edge) => {
        const border = targetRange.format.borders.getItem(edge);
        border.style = Excel.BorderLineStyle.none;
      });
    }

    await context.sync();
  });
}
