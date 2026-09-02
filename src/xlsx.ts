// xlsx.ts
import * as XLSX from "xlsx";
import type { SheetDocument, CellData, SheetMeta } from "./types";
import { indexToColumnLetter, columnLetterToIndex } from "./chart";

/**
 * Parses an Excel (.xlsx, .xls, .csv) file buffer into an AntSheet SheetDocument
 */
export function parseXLSX(
  buffer: ArrayBuffer | Uint8Array | Buffer | Array<number>,
  sheetIndexOrName: number | string = 0
): SheetDocument {
  const workbook = XLSX.read(buffer, { type: "buffer", cellStyles: true, cellFormula: true });
  
  let sheetName: string;
  if (typeof sheetIndexOrName === "number") {
    sheetName = workbook.SheetNames[sheetIndexOrName] || "Sheet1";
  } else {
    sheetName = sheetIndexOrName;
  }

  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    return {
      name: sheetName,
      meta: {},
      cells: {},
      merges: [],
    };
  }

  const meta: SheetMeta = {};
  const cells: Record<string, CellData> = {};
  const merges: string[] = [];

  // 1. Column Widths
  if (worksheet["!cols"]) {
    const colWidths: Record<string, number> = {};
    worksheet["!cols"].forEach((col, idx) => {
      if (col && (col.wpx || col.wch)) {
        const colLetter = indexToColumnLetter(idx);
        colWidths[colLetter] = col.wpx ?? Math.round((col.wch ?? 10) * 8);
      }
    });
    if (Object.keys(colWidths).length > 0) meta.columnWidth = colWidths;
  }

  // 2. Row Heights
  if (worksheet["!rows"]) {
    const rowHeights: Record<string, number> = {};
    worksheet["!rows"].forEach((row, idx) => {
      if (row && (row.hpx || row.hpt)) {
        rowHeights[String(idx + 1)] = row.hpx ?? Math.round((row.hpt ?? 20) * 1.33);
      }
    });
    if (Object.keys(rowHeights).length > 0) meta.rowHeight = rowHeights;
  }

  // 3. Merges
  if (worksheet["!merges"]) {
    worksheet["!merges"].forEach((m) => {
      const startRef = `${indexToColumnLetter(m.s.c)}${m.s.r + 1}`;
      const endRef = `${indexToColumnLetter(m.e.c)}${m.e.r + 1}`;
      merges.push(startRef === endRef ? startRef : `${startRef}:${endRef}`);
    });
  }

  // 4. Cells
  for (const key of Object.keys(worksheet)) {
    if (key.startsWith("!")) continue;
    const cell = worksheet[key];
    if (!cell) continue;

    let value: string | number | boolean | null = null;
    if (cell.f) {
      value = `=${cell.f}`;
    } else if (cell.v !== undefined) {
      value = cell.v;
    }

    cells[key.toUpperCase()] = {
      value,
    };
  }

  return {
    name: sheetName,
    meta,
    cells,
    merges,
  };
}

/**
 * Exports an AntSheet SheetDocument into Excel (.xlsx) file binary bytes (Uint8Array)
 */
export function exportToXLSX(doc: SheetDocument): Uint8Array {
  const workbook = XLSX.utils.book_new();
  const worksheet: XLSX.WorkSheet = {};

  let minRow = Infinity;
  let maxRow = -1;
  let minCol = Infinity;
  let maxCol = -1;

  // 1. Populate Cells
  for (const [ref, cell] of Object.entries(doc.cells)) {
    const match = ref.match(/^([A-Z]+)(\d+)$/);
    if (!match) continue;

    const colIndex = columnLetterToIndex(match[1]);
    const rowIndex = parseInt(match[2], 10) - 1;

    minRow = Math.min(minRow, rowIndex);
    maxRow = Math.max(maxRow, rowIndex);
    minCol = Math.min(minCol, colIndex);
    maxCol = Math.max(maxCol, colIndex);

    const cellObj: XLSX.CellObject = { v: "", t: "s" };


    if (cell.value === null || cell.value === undefined) {
      cellObj.v = "";
      cellObj.t = "s";
    } else if (typeof cell.value === "string" && cell.value.startsWith("=")) {
      cellObj.f = cell.value.slice(1);
    } else if (typeof cell.value === "number") {
      cellObj.v = cell.value;
      cellObj.t = "n";
    } else if (typeof cell.value === "boolean") {
      cellObj.v = cell.value;
      cellObj.t = "b";
    } else {
      cellObj.v = String(cell.value);
      cellObj.t = "s";
    }

    worksheet[ref] = cellObj;
  }

  // 2. Range Dimension
  if (maxRow >= 0 && maxCol >= 0) {
    worksheet["!ref"] = XLSX.utils.encode_range({
      s: { r: minRow === Infinity ? 0 : minRow, c: minCol === Infinity ? 0 : minCol },
      e: { r: maxRow, c: maxCol },
    });
  } else {
    worksheet["!ref"] = "A1:A1";
  }

  // 3. Merges
  if (doc.merges.length > 0) {
    worksheet["!merges"] = doc.merges.map((rangeStr) => {
      const [start, end] = rangeStr.split(":");
      const startMatch = start.match(/^([A-Z]+)(\d+)$/)!;
      const endMatch = (end || start).match(/^([A-Z]+)(\d+)$/)!;

      return {
        s: { r: parseInt(startMatch[2], 10) - 1, c: columnLetterToIndex(startMatch[1]) },
        e: { r: parseInt(endMatch[2], 10) - 1, c: columnLetterToIndex(endMatch[1]) },
      };
    });
  }

  // 4. Column Widths
  if (doc.meta.columnWidth) {
    const cols: XLSX.ColInfo[] = [];
    for (const [colLetter, width] of Object.entries(doc.meta.columnWidth)) {
      const cIdx = columnLetterToIndex(colLetter);
      cols[cIdx] = { wpx: width };
    }
    worksheet["!cols"] = cols;
  }

  // 5. Row Heights
  if (doc.meta.rowHeight) {
    const rows: XLSX.RowInfo[] = [];
    for (const [rowStr, height] of Object.entries(doc.meta.rowHeight)) {
      const rIdx = parseInt(rowStr, 10) - 1;
      rows[rIdx] = { hpx: height };
    }
    worksheet["!rows"] = rows;
  }

  XLSX.utils.book_append_sheet(workbook, worksheet, doc.name || "Sheet1");

  const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  return new Uint8Array(wbout);
}
