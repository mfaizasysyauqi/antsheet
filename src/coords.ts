// coords.ts - Single source of truth for spreadsheet coordinates & cell references

export const CELL_REF_REGEX = /^([A-Z]+)(\d+)$/;

export interface CellCoordinates {
  col: number; // 0-indexed
  row: number; // 0-indexed
}

export interface RangeCoordinates {
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
}

/**
 * Converts a 1-based column letter (e.g. "A", "Z", "AA") to a 0-based column index.
 */
export function colLetterToIndex(letter: string): number {
  let index = 0;
  const upper = letter.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

/**
 * Converts a 0-based column index (e.g. 0 -> "A", 25 -> "Z", 26 -> "AA") to column letters.
 */
export function indexToColLetter(index: number): string {
  let temp = index + 1;
  let letter = "";
  while (temp > 0) {
    const mod = (temp - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    temp = Math.floor((temp - mod) / 26);
  }
  return letter;
}

// Aliases for backward compatibility
export const columnLetterToIndex = colLetterToIndex;
export const indexToColumnLetter = indexToColLetter;

/**
 * Parses a cell reference like "A1" or "BC42" into 0-based { col, row }.
 */
export function parseCellRef(ref: string): CellCoordinates {
  const match = ref.toUpperCase().match(CELL_REF_REGEX);
  if (!match) {
    throw new Error(`Invalid cell ref: "${ref}"`);
  }
  return {
    col: colLetterToIndex(match[1]),
    row: parseInt(match[2], 10) - 1,
  };
}

/**
 * Formats 0-based column and row coordinates into a cell reference (e.g. 0, 0 -> "A1").
 */
export function formatCellRef(col: number, row: number): string {
  return `${indexToColLetter(col)}${row + 1}`;
}

/**
 * Validates whether a string is a valid cell reference (e.g. "A1", "Z99").
 */
export function isValidCellRef(ref: string): boolean {
  return CELL_REF_REGEX.test(ref.toUpperCase());
}

/**
 * Parses a range string like "A1:C5" or a single cell "B2" into normalized boundary coordinates.
 */
export function parseRangeCoordinates(range: string): RangeCoordinates {
  const [start, end] = range.split(":");
  const startPos = parseCellRef(start.trim());
  const endPos = end ? parseCellRef(end.trim()) : startPos;

  return {
    startCol: Math.min(startPos.col, endPos.col),
    startRow: Math.min(startPos.row, endPos.row),
    endCol: Math.max(startPos.col, endPos.col),
    endRow: Math.max(startPos.row, endPos.row),
  };
}

/**
 * Resolves all cell reference strings contained within a range (e.g. "A1:B2" -> ["A1", "B1", "A2", "B2"]).
 */
export function resolveRangeCells(range: string): string[] {
  const coords = parseRangeCoordinates(range);
  const cells: string[] = [];

  for (let r = coords.startRow; r <= coords.endRow; r++) {
    for (let c = coords.startCol; c <= coords.endCol; c++) {
      cells.push(formatCellRef(c, r));
    }
  }

  return cells;
}

/**
 * Comparator for sorting cell references in reading order (top-to-bottom, left-to-right).
 */
export function compareCellRefs(a: string, b: string): number {
  const pa = parseCellRef(a);
  const pb = parseCellRef(b);
  return pa.row !== pb.row ? pa.row - pb.row : pa.col - pb.col;
}
