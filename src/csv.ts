// csv.ts
import type { SheetDocument, CellData } from "./types";

export interface CSVParseOptions {
  /** Whether the first row is a header row (used as column keys). Default: false */
  hasHeader?: boolean;
  /** Column delimiter. Default: "," */
  delimiter?: string;
  /** Sheet name for the resulting document. Default: "Sheet1" */
  sheetName?: string;
}

export interface CSVExportOptions {
  /** Column delimiter. Default: "," */
  delimiter?: string;
  /** Whether to include a header row from column letters (A, B, C...). Default: false */
  includeHeader?: boolean;
  /** Line ending. Default: "\n" */
  lineEnding?: "\n" | "\r\n";
}

/**
 * Parse CSV text into a SheetDocument.
 */
export function parseCSV(text: string, options: CSVParseOptions = {}): SheetDocument {
  const {
    hasHeader = false,
    delimiter = ",",
    sheetName = "Sheet1",
  } = options;

  const lines = text.split(/\r?\n/).filter((l) => l.trim().length > 0);
  const cells: Record<string, CellData> = {};

  if (hasHeader && lines.length > 0) {
    // First row treated as header — data still imported from row 1 as-is
  }

  for (let rowIdx = 0; rowIdx < lines.length; rowIdx++) {
    const rowNum = rowIdx + 1;
    const cols = parseCSVLine(lines[rowIdx], delimiter);

    for (let colIdx = 0; colIdx < cols.length; colIdx++) {
      const colLetter = indexToColumnLetter(colIdx);
      const ref = `${colLetter}${rowNum}`;
      const raw = cols[colIdx].trim();

      let value: string | number | boolean | null = raw;

      if (raw === "") {
        value = null;
      } else if (raw === "true") {
        value = true;
      } else if (raw === "false") {
        value = false;
      } else if (!isNaN(Number(raw)) && raw.trim() !== "") {
        value = Number(raw);
      }

      if (value !== null) {
        cells[ref] = { value };
      }
    }
  }

  return {
    name: sheetName,
    meta: {},
    cells,
    merges: [],
  };
}

/**
 * Export a SheetDocument to CSV text.
 */
export function exportToCSV(doc: SheetDocument, options: CSVExportOptions = {}): string {
  const {
    delimiter = ",",
    lineEnding = "\n",
  } = options;

  if (Object.keys(doc.cells).length === 0) return "";

  // Find bounds
  let maxRow = 0;
  let maxCol = 0;

  for (const ref of Object.keys(doc.cells)) {
    const match = ref.match(/^([A-Z]+)(\d+)$/);
    if (!match) continue;
    const col = columnLetterToIndex(match[1]);
    const row = parseInt(match[2], 10);
    if (row > maxRow) maxRow = row;
    if (col > maxCol) maxCol = col;
  }

  const rows: string[] = [];

  for (let row = 1; row <= maxRow; row++) {
    const cols: string[] = [];
    for (let col = 0; col <= maxCol; col++) {
      const ref = `${indexToColumnLetter(col)}${row}`;
      const cell = doc.cells[ref];
      cols.push(cell ? csvEscape(String(cell.value ?? ""), delimiter) : "");
    }
    rows.push(cols.join(delimiter));
  }

  return rows.join(lineEnding) + lineEnding;
}

// ─────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────

/** Parse a single CSV line respecting quoted fields */
function parseCSVLine(line: string, delimiter: string): string[] {
  const result: string[] = [];
  let current = "";
  let insideQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];

    if (char === '"') {
      if (insideQuotes && line[i + 1] === '"') {
        // Escaped quote
        current += '"';
        i++;
      } else {
        insideQuotes = !insideQuotes;
      }
    } else if (char === delimiter && !insideQuotes) {
      result.push(current);
      current = "";
    } else {
      current += char;
    }
  }

  result.push(current);
  return result;
}

function csvEscape(value: string, delimiter: string): string {
  if (value.includes(delimiter) || value.includes('"') || value.includes("\n")) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

function indexToColumnLetter(index: number): string {
  let letter = "";
  let n = index;
  while (n >= 0) {
    letter = String.fromCharCode((n % 26) + 65) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}

function columnLetterToIndex(letter: string): number {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index - 1;
}
