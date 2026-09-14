// converter.ts
import type { SheetDocument, WorkbookDocument, CellStyle } from "./types";
import {
  colLetterToIndex,
  indexToColLetter,
  parseCellRef,
  parseRangeCoordinates,
} from "./coords";

// ─────────────────────────────────────────────
// Univer Types (simplified)
// ─────────────────────────────────────────────

export interface UniverCell {
  v?: string | number | boolean | null;
  t?: number; // 1 = string, 2 = number, 3 = boolean
  f?: string; // Formula (e.g. "=SUM(A1:A5)")
  s?: IStyleData;
}

export interface IStyleData {
  ff?: string; // font family
  fs?: number; // font size
  bl?: number; // bold (1 = true)
  it?: number; // italic (1 = true)
  ul?: { s: number }; // underline
  st?: { s: number }; // strikethrough
  cl?: { rgb?: string }; // text color
  bg?: { rgb?: string }; // background
  ht?: number; // horizontal align (1=left, 2=center, 3=right)
  vt?: number; // vertical align (1=top, 2=middle, 3=bottom)
  tb?: number; // text wrap (1 = overflow, 2 = wrap, 3 = clip)
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
      const colIndex = colLetterToIndex(colLetter);
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
    mergeData.push(parseRangeToIRange(range));
  }

  // 5. Freeze Panes
  let freeze: IFreezeData | undefined;
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

/**
 * Convert a SheetDocument or WorkbookDocument into a complete Univer IWorkbookData object.
 */
export function convertToUniverWorkbook(docOrWb: SheetDocument | WorkbookDocument): {
  id: string;
  name: string;
  appVersion: string;
  locale: string;
  sheetOrder: string[];
  sheets: Record<string, UniverSheetData>;
  styles: Record<string, unknown>;
} {
  const isMultiSheet = "sheets" in docOrWb && Array.isArray(docOrWb.sheets);
  const sheetDocs = isMultiSheet ? docOrWb.sheets : [docOrWb as SheetDocument];

  const sheets: Record<string, UniverSheetData> = {};
  const sheetOrder: string[] = [];

  for (let i = 0; i < sheetDocs.length; i++) {
    const sDoc = sheetDocs[i];
    const sId = `sheet-${i + 1}`;
    const uSheet = convertToUniver(sDoc);
    uSheet.id = sId;

    sheets[sId] = uSheet;
    sheetOrder.push(sId);
  }

  const primaryName = sheetDocs[0]?.name || "Workbook";

  return {
    id: "antsheet-workbook",
    name: primaryName,
    appVersion: "0.25.1",
    locale: "en-US",
    sheetOrder,
    sheets,
    styles: {},
  };
}

// ─────────────────────────────────────────────
// Univer -> AntSheet Converter
// ─────────────────────────────────────────────

export function convertFromUniver(univerData: UniverSheetData): SheetDocument {
  const name = univerData.name || "Untitled";
  const meta: SheetDocument["meta"] = {};
  const cells: SheetDocument["cells"] = {};
  const merges: string[] = [];

  // 1. Column Widths
  if (univerData.columnData) {
    const colWidths: Record<string, number> = {};
    for (const [colIdxStr, colInfo] of Object.entries(univerData.columnData)) {
      if (colInfo?.w) {
        colWidths[indexToColLetter(Number(colIdxStr))] = colInfo.w;
      }
    }
    if (Object.keys(colWidths).length > 0) meta.columnWidth = colWidths;
  }

  // 2. Row Heights
  if (univerData.rowData) {
    const rowHeights: Record<string, number> = {};
    for (const [rowIdxStr, rowInfo] of Object.entries(univerData.rowData)) {
      if (rowInfo?.h) {
        rowHeights[String(Number(rowIdxStr) + 1)] = rowInfo.h;
      }
    }
    if (Object.keys(rowHeights).length > 0) meta.rowHeight = rowHeights;
  }

  // 3. Freeze Panes
  if (univerData.freeze) {
    if (univerData.freeze.startRow) meta.freezeRow = univerData.freeze.startRow;
    if (univerData.freeze.startColumn) meta.freezeColumn = univerData.freeze.startColumn;
  }

  // 4. Cells
  if (univerData.cellData) {
    for (const [rowIdxStr, rowObj] of Object.entries(univerData.cellData)) {
      const row = Number(rowIdxStr);
      for (const [colIdxStr, cell] of Object.entries(rowObj)) {
        const col = Number(colIdxStr);
        const ref = `${indexToColLetter(col)}${row + 1}`;

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
      const startRef = `${indexToColLetter(m.startColumn)}${m.startRow + 1}`;
      const endRef = `${indexToColLetter(m.endColumn)}${m.endRow + 1}`;
      merges.push(startRef === endRef ? startRef : `${startRef}:${endRef}`);
    }
  }

  return { name, meta, cells, merges };
}

/**
 * Convert a complete Univer workbook object back into a WorkbookDocument.
 */
export function convertFromUniverWorkbook(univerWorkbook: {
  sheets: Record<string, UniverSheetData>;
  sheetOrder?: string[];
}): WorkbookDocument {
  const sheetOrder = univerWorkbook.sheetOrder || Object.keys(univerWorkbook.sheets);
  const sheets: SheetDocument[] = [];
  for (const sId of sheetOrder) {
    const sData = univerWorkbook.sheets[sId];
    if (sData) {
      sheets.push(convertFromUniver(sData));
    }
  }
  return { sheets };
}

// ─────────────────────────────────────────────
// Style Converters
// ─────────────────────────────────────────────

function convertStyle(style: CellStyle): IStyleData {
  const s: IStyleData = {};

  if (style.fontFamily) s.ff = style.fontFamily;
  if (style.fontSize) s.fs = style.fontSize;

  if (
    style.fontWeight === "bold" ||
    style.fontWeight === 700 ||
    style.fontWeight === "700"
  ) {
    s.bl = 1;
  }

  if (style.fontStyle === "italic") s.it = 1;
  if (style.textDecoration?.includes("underline")) s.ul = { s: 1 };
  if (style.textDecoration?.includes("lineThrough")) s.st = { s: 1 };

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
  const hAlignMap: Record<string, number> = { left: 1, center: 2, right: 3 };
  if (style.textAlign && hAlignMap[style.textAlign]) {
    s.ht = hAlignMap[style.textAlign];
  }

  const vAlignMap: Record<string, number> = { top: 1, middle: 2, bottom: 3 };
  if (style.verticalAlign && vAlignMap[style.verticalAlign]) {
    s.vt = vAlignMap[style.verticalAlign];
  }

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

  const hAlignRevMap: Record<number, "left" | "center" | "right"> = {
    1: "left",
    2: "center",
    3: "right",
  };
  if (s.ht && hAlignRevMap[s.ht]) style.textAlign = hAlignRevMap[s.ht];

  const vAlignRevMap: Record<number, "top" | "middle" | "bottom"> = {
    1: "top",
    2: "middle",
    3: "bottom",
  };
  if (s.vt && vAlignRevMap[s.vt]) style.verticalAlign = vAlignRevMap[s.vt];

  const textWrapMap: Record<number, "overflow" | "wrap" | "clip"> = {
    1: "overflow",
    2: "wrap",
    3: "clip",
  };
  if (s.tb && textWrapMap[s.tb]) style.textWrap = textWrapMap[s.tb];

  return Object.keys(style).length > 0 ? style : undefined;
}

function convertBorder(style: CellStyle): IBorderData | undefined {
  const bd: IBorderData = {};

  const applySide = (side: "t" | "r" | "b" | "l", value?: string | number) => {
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

  if (style.border === "bottom") applySide("b", style.borderWidth || 1);
  if (style.border === "top") applySide("t", style.borderWidth || 1);
  if (style.border === "left") applySide("l", style.borderWidth || 1);
  if (style.border === "right") applySide("r", style.borderWidth || 1);

  applySide("t", style.borderTop);
  applySide("r", style.borderRight);
  applySide("b", style.borderBottom);
  applySide("l", style.borderLeft);

  return Object.keys(bd).length > 0 ? bd : undefined;
}

// ─────────────────────────────────────────────
// Format Mapping Table
// ─────────────────────────────────────────────

const FORMAT_PATTERN_MAP: Record<string, string> = {
  currency: '"Rp"#,##0',
  "currency:IDR": '"Rp"#,##0',
  "currency:USD": '"$"#,##0.00',
  "currency:EUR": '"€"#,##0.00',
  "currency:JPY": '"¥"#,##0',
  percent: "0%",
  "percent:1": "0.0%",
  "percent:2": "0.00%",
  number: "#,##0.###",
  "number:2": "#,##0.00",
  "number:0": "#,##0",
  compact: "0.0,,",
  date: "dd/mm/yyyy",
  "date:long": "dd mmmm yyyy",
  "date:iso": "yyyy-mm-dd",
  datetime: "dd/mm/yyyy hh:mm",
  time: "hh:mm",
};

function mapFormatToPattern(format: string): string {
  return FORMAT_PATTERN_MAP[format] ?? format;
}

function getCellType(value: string | number | boolean | null): number {
  if (typeof value === "number") return 2;
  if (typeof value === "boolean") return 3;
  return 1; // string
}

function parseRangeToIRange(range: string): IRange {
  const coords = parseRangeCoordinates(range);
  return {
    startRow: coords.startRow,
    startColumn: coords.startCol,
    endRow: coords.endRow,
    endColumn: coords.endCol,
  };
}