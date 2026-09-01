// converter.ts
import type { SheetDocument, CellStyle } from "./types";

// ─────────────────────────────────────────────
// Univer types (simplified)
// ─────────────────────────────────────────────
export interface UniverCell {
  v?: string | number | boolean | null;
  t?: number; // 1 = string, 2 = number, 3 = boolean
  s?: IStyleData;
}

export interface IStyleData {
  ff?: string;          // font family
  fs?: number;          // font size
  bl?: number;          // bold (1 = true)
  it?: number;          // italic (1 = true)
  ul?: { s: number };   // underline
  st?: { s: number };   // strikethrough
  cl?: { rgb?: string }; // text color
  bg?: { rgb?: string }; // background
  ht?: number;          // horizontal align (1=left, 2=center, 3=right)
  vt?: number;          // vertical align (1=top, 2=middle, 3=bottom)
  n?: { pattern: string }; // number format
  bd?: IBorderData;
}

export interface IBorderData {
  t?: IBorderStyleData;
  r?: IBorderStyleData;
  b?: IBorderStyleData;
  l?: IBorderStyleData;
}

export interface IBorderStyleData {
  s: number; // style (1 = thin, dll)
  cl?: { rgb?: string };
}

export interface IRange {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

export interface UniverSheetData {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  cellData: Record<number, Record<number, UniverCell>>;
  mergeData: IRange[];
  rowData: Record<number, { h?: number }>;
  columnData: Record<number, { w?: number }>;
}

// ─────────────────────────────────────────────
// Main Converter
// ─────────────────────────────────────────────
export function convertToUniver(doc: SheetDocument): UniverSheetData {
  const cellData: Record<number, Record<number, UniverCell>> = {};
  const mergeData: IRange[] = [];
  const rowData: Record<number, { h?: number }> = {};
  const columnData: Record<number, { w?: number }> = {};

  // 1. Column Width
  if (doc.meta.columnWidth) {
    for (const [colLetter, width] of Object.entries(doc.meta.columnWidth)) {
      const colIndex = columnLetterToIndex(colLetter);
      columnData[colIndex] = { w: width };
    }
  }

  // 2. Row Height
  if (doc.meta.rowHeight) {
    for (const [rowStr, height] of Object.entries(doc.meta.rowHeight)) {
      const rowIndex = parseInt(rowStr, 10) - 1;
      rowData[rowIndex] = { h: height };
    }
  }

  // 3. Cells
  for (const [ref, cell] of Object.entries(doc.cells)) {
    const { row, col } = parseCellRef(ref);

    if (!cellData[row]) cellData[row] = {};

    cellData[row][col] = {
      v: cell.value,
      t: getCellType(cell.value),
      s: cell.style ? convertStyle(cell.style) : undefined,
    };
  }

  // 4. Merges
  for (const range of doc.merges) {
    mergeData.push(parseRange(range));
  }

  return {
    id: "sheet-1",
    name: doc.name,
    rowCount: 100,
    columnCount: 26,
    cellData,
    mergeData,
    rowData,
    columnData,
  };
}

// ─────────────────────────────────────────────
// Style Mapping
// ─────────────────────────────────────────────
function convertStyle(style: CellStyle): IStyleData {
  const s: IStyleData = {};

  // Font
  if (style.fontFamily) s.ff = style.fontFamily;
  if (style.fontSize) s.fs = style.fontSize;

  if (
    style.fontWeight === "bold" ||
    style.fontWeight === 700 ||
    style.fontWeight === "700"
  ) {
    s.bl = 1;
  }

  if (style.fontStyle === "italic") {
    s.it = 1;
  }

  if (style.textDecoration?.includes("underline")) {
    s.ul = { s: 1 };
  }
  if (style.textDecoration?.includes("lineThrough")) {
    s.st = { s: 1 };
  }

  // Color
  if (style.color) s.cl = { rgb: style.color };
  if (style.backgroundColor) s.bg = { rgb: style.backgroundColor };

  // Align
  if (style.textAlign === "left") s.ht = 1;
  if (style.textAlign === "center") s.ht = 2;
  if (style.textAlign === "right") s.ht = 3;

  if (style.verticalAlign === "top") s.vt = 1;
  if (style.verticalAlign === "middle") s.vt = 2;
  if (style.verticalAlign === "bottom") s.vt = 3;

  // Format
  if (style.format) {
    s.n = { pattern: mapFormatToPattern(style.format) };
  }

  // Border
  const border = convertBorder(style);
  if (border) s.bd = border;

  return s;
}

function convertBorder(style: CellStyle): IBorderData | undefined {
  const bd: IBorderData = {};

  const apply = (
    side: "t" | "r" | "b" | "l",
    value?: string | number
  ) => {
    if (value === undefined) return;

    if (typeof value === "number") {
      bd[side] = {
        s: value >= 2 ? 2 : 1,
        cl: style.borderColor ? { rgb: style.borderColor } : undefined,
      };
    } else {
      // contoh: "2 solid #0f172a"
      const parts = value.split(" ");
      const width = parseInt(parts[0], 10) || 1;
      const color = parts[2] || style.borderColor;

      bd[side] = {
        s: width >= 2 ? 2 : 1,
        cl: color ? { rgb: color } : undefined,
      };
    }
  };

  // border: all
  if (style.border === "all") {
    const width = style.borderWidth || 1;
    const color = style.borderColor;
    const borderStyle = {
      s: width >= 2 ? 2 : 1,
      cl: color ? { rgb: color } : undefined,
    };
    bd.t = borderStyle;
    bd.r = borderStyle;
    bd.b = borderStyle;
    bd.l = borderStyle;
  }

  // border: bottom / top / left / right
  if (style.border === "bottom") apply("b", style.borderWidth || 1);
  if (style.border === "top") apply("t", style.borderWidth || 1);
  if (style.border === "left") apply("l", style.borderWidth || 1);
  if (style.border === "right") apply("r", style.borderWidth || 1);

  // Detail per sisi
  apply("t", style.borderTop);
  apply("r", style.borderRight);
  apply("b", style.borderBottom);
  apply("l", style.borderLeft);

  return Object.keys(bd).length > 0 ? bd : undefined;
}

// ─────────────────────────────────────────────
// Format Mapping
// ─────────────────────────────────────────────
function mapFormatToPattern(format: string): string {
  switch (format) {
    case "currency":
    case "currency:IDR":
      return '"Rp"#,##0';
    case "currency:USD":
      return '"$"#,##0.00';
    case "currency:EUR":
      return '"€"#,##0.00';
    case "currency:JPY":
      return '"¥"#,##0';
    case "percent":
      return "0%";
    case "percent:1":
      return "0.0%";
    case "percent:2":
      return "0.00%";
    case "number":
      return "#,##0.###";
    case "number:2":
      return "#,##0.00";
    case "number:0":
      return "#,##0";
    case "compact":
      return "0.0,," ; // sederhana
    case "date":
      return "dd/mm/yyyy";
    case "date:long":
      return "dd mmmm yyyy";
    case "date:iso":
      return "yyyy-mm-dd";
    case "datetime":
      return "dd/mm/yyyy hh:mm";
    case "time":
      return "hh:mm";
    default:
      return format;
  }
}

// ─────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────
function getCellType(value: string | number | boolean | null): number {
  if (typeof value === "number") return 2;
  if (typeof value === "boolean") return 3;
  return 1; // string
}

function columnLetterToIndex(letter: string): number {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index - 1;
}

function parseCellRef(ref: string): { row: number; col: number } {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) throw new Error(`Invalid cell ref: ${ref}`);

  return {
    col: columnLetterToIndex(match[1]),
    row: parseInt(match[2], 10) - 1,
  };
}

function parseRange(range: string): IRange {
  const [start, end] = range.split(":");
  const startPos = parseCellRef(start);
  const endPos = end ? parseCellRef(end) : startPos;

  return {
    startRow: startPos.row,
    startColumn: startPos.col,
    endRow: endPos.row,
    endColumn: endPos.col,
  };
}