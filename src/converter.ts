// converter.ts
import type { SheetDocument, SheetMeta, CellData, CellStyle } from "./types";
import { indexToColumnLetter } from "./chart";

// ─────────────────────────────────────────────
// Univer types (simplified)
// ─────────────────────────────────────────────
export interface UniverCell {
  v?: string | number | boolean | null;
  t?: number; // 1 = string, 2 = number, 3 = boolean
  f?: string; // Formula (e.g. "=SUM(A1:A5)")
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
  tb?: number;          // text wrap (1 = overflow, 2 = wrap, 3 = clip)
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
  s: number; // style (1 = thin, 2 = medium/thick)
  cl?: { rgb?: string };
}

export interface IRange {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

export interface IFreezeData {
  startRow?: number;
  startColumn?: number;
  xSplit?: number;
  ySplit?: number;
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
  freeze?: IFreezeData;
}

// ─────────────────────────────────────────────
// AntSheet -> Univer Converter
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

    const isFormula = typeof cell.value === "string" && cell.value.startsWith("=");

    cellData[row][col] = {
      v: isFormula ? null : cell.value,
      f: isFormula ? (cell.value as string) : undefined,
      t: isFormula ? undefined : getCellType(cell.value),
      s: cell.style ? convertStyle(cell.style) : undefined,
    };
  }

  // 4. Merges
  for (const range of doc.merges) {
    mergeData.push(parseRange(range));
  }

  // 5. Freeze Panes
  let freeze: IFreezeData | undefined = undefined;
  if (doc.meta.freezeRow !== undefined || doc.meta.freezeColumn !== undefined) {
    freeze = {
      startRow: doc.meta.freezeRow ?? 0,
      startColumn: doc.meta.freezeColumn ?? 0,
      ySplit: doc.meta.freezeRow ?? 0,
      xSplit: doc.meta.freezeColumn ?? 0,
    };
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
    freeze,
  };
}

// ─────────────────────────────────────────────
// Univer -> AntSheet Converter
// ─────────────────────────────────────────────
export function convertFromUniver(univerData: UniverSheetData): SheetDocument {
  const name = univerData.name || "Untitled";
  const meta: SheetMeta = {};
  const cells: Record<string, CellData> = {};
  const merges: string[] = [];

  // 1. Column Widths
  if (univerData.columnData) {
    const colWidths: Record<string, number> = {};
    for (const [colIdxStr, colInfo] of Object.entries(univerData.columnData)) {
      if (colInfo?.w !== undefined) {
        const colLetter = indexToColumnLetter(parseInt(colIdxStr, 10));
        colWidths[colLetter] = colInfo.w;
      }
    }
    if (Object.keys(colWidths).length > 0) meta.columnWidth = colWidths;
  }

  // 2. Row Heights
  if (univerData.rowData) {
    const rowHeights: Record<string, number> = {};
    for (const [rowIdxStr, rowInfo] of Object.entries(univerData.rowData)) {
      if (rowInfo?.h !== undefined) {
        const rowNum = parseInt(rowIdxStr, 10) + 1;
        rowHeights[String(rowNum)] = rowInfo.h;
      }
    }
    if (Object.keys(rowHeights).length > 0) meta.rowHeight = rowHeights;
  }

  // 3. Freeze Panes
  if (univerData.freeze) {
    const freezeRow = univerData.freeze.startRow ?? univerData.freeze.ySplit;
    const freezeCol = univerData.freeze.startColumn ?? univerData.freeze.xSplit;
    if (freezeRow && freezeRow > 0) meta.freezeRow = freezeRow;
    if (freezeCol && freezeCol > 0) meta.freezeColumn = freezeCol;
  }

  // 4. Cells
  if (univerData.cellData) {
    for (const [rStr, cols] of Object.entries(univerData.cellData)) {
      const r = parseInt(rStr, 10);
      if (!cols) continue;

      for (const [cStr, cell] of Object.entries(cols)) {
        const c = parseInt(cStr, 10);
        if (!cell) continue;

        const ref = `${indexToColumnLetter(c)}${r + 1}`;
        const value = cell.f ? cell.f : (cell.v ?? null);
        const style = cell.s ? convertUniverStyle(cell.s) : undefined;

        if (value !== null || style) {
          cells[ref] = { value, style };
        }
      }
    }
  }

  // 5. Merges
  if (univerData.mergeData) {
    for (const m of univerData.mergeData) {
      const startRef = `${indexToColumnLetter(m.startColumn)}${m.startRow + 1}`;
      const endRef = `${indexToColumnLetter(m.endColumn)}${m.endRow + 1}`;
      merges.push(startRef === endRef ? startRef : `${startRef}:${endRef}`);
    }
  }

  return {
    name,
    meta,
    cells,
    merges,
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

  // Text Wrapping
  if (style.textWrap === "wrap" || style.wrapText === true) {
    s.tb = 2;
  } else if (style.textWrap === "clip") {
    s.tb = 3;
  } else if (style.textWrap === "overflow" || style.wrapText === false) {
    s.tb = 1;
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

function convertUniverStyle(s: IStyleData): CellStyle | undefined {
  const style: CellStyle = {};

  if (s.ff) style.fontFamily = s.ff;
  if (s.fs) style.fontSize = s.fs;
  if (s.bl === 1) style.fontWeight = "bold";
  if (s.it === 1) style.fontStyle = "italic";
  if (s.ul?.s === 1) style.textDecoration = "underline";
  if (s.st?.s === 1) style.textDecoration = "lineThrough";
  if (s.cl?.rgb) style.color = s.cl.rgb;
  if (s.bg?.rgb) style.backgroundColor = s.bg.rgb;

  if (s.ht === 1) style.textAlign = "left";
  if (s.ht === 2) style.textAlign = "center";
  if (s.ht === 3) style.textAlign = "right";

  if (s.vt === 1) style.verticalAlign = "top";
  if (s.vt === 2) style.verticalAlign = "middle";
  if (s.vt === 3) style.verticalAlign = "bottom";

  if (s.tb === 2) {
    style.textWrap = "wrap";
  } else if (s.tb === 3) {
    style.textWrap = "clip";
  } else if (s.tb === 1) {
    style.textWrap = "overflow";
  }

  return Object.keys(style).length > 0 ? style : undefined;
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

  // Detail per side
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
      return "0.0,,";
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