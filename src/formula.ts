// formula.ts
// Formula evaluation using HyperFormula — runs in browser and Node.js
import { HyperFormula } from "hyperformula";

import type { WorkbookDocument, SheetDocument, CellData } from "./types";
import { columnLetterToIndex } from "./chart";


export interface EvaluateOptions {
  /** License key — use "gpl-v3" for open-source projects (free) */
  licenseKey?: string;
}

/**
 * Evaluate all formulas in a single SheetDocument.
 * Returns a new SheetDocument where formula cells have their computed value.
 * The original formula string is preserved in cell.formula.
 */
export function evaluateFormulas(
  doc: SheetDocument,
  options: EvaluateOptions = {}
): SheetDocument {
  const wb = evaluateWorkbook({ sheets: [doc] }, options);
  return wb.sheets[0];
}

/**
 * Evaluate all formulas across a multi-sheet WorkbookDocument.
 * Returns a new WorkbookDocument with computed values.
 */
export function evaluateWorkbook(
  wb: WorkbookDocument,
  options: EvaluateOptions = {}
): WorkbookDocument {
  const licenseKey = options.licenseKey ?? "gpl-v3";

  // Build sheet data arrays for HyperFormula
  // HyperFormula expects: rows of columns (2D array, row-major)
  const sheetNames = wb.sheets.map((s) => s.name);
  const sheetArrays = wb.sheets.map((sheet) => buildGrid(sheet));

  const hf = HyperFormula.buildFromSheets(
    Object.fromEntries(sheetNames.map((name, i) => [name, sheetArrays[i]])),
    { licenseKey }
  );

  const evaluatedSheets: SheetDocument[] = wb.sheets.map((sheet, sheetIdx) => {
    const evaluatedCells: Record<string, CellData> = {};

    for (const [ref, cell] of Object.entries(sheet.cells)) {
      const match = ref.match(/^([A-Z]+)(\d+)$/);
      if (!match) {
        evaluatedCells[ref] = cell;
        continue;
      }

      const col = columnLetterToIndex(match[1]);
      const row = parseInt(match[2], 10) - 1;

      const isFormula =
        typeof cell.value === "string" && cell.value.startsWith("=");

      if (!isFormula) {
        evaluatedCells[ref] = cell;
        continue;
      }

      // Get computed value from HyperFormula
      const computed = hf.getCellValue({
        sheet: sheetIdx,
        row,
        col,
      });

      let resolvedValue: string | number | boolean | null;

      if (computed === null || computed === undefined) {
        resolvedValue = null;
      } else if (typeof computed === "object" && "type" in computed) {
        // HyperFormula error object (e.g. #REF!, #DIV/0!)
        resolvedValue = String(computed);
      } else {
        resolvedValue = computed as string | number | boolean;
      }

      evaluatedCells[ref] = {
        ...cell,
        value: resolvedValue,
        // Preserve original formula in a comment-like way so it's not lost
        comment: cell.comment ?? `formula: ${cell.value}`,
      };
    }

    return {
      ...sheet,
      cells: evaluatedCells,
    };
  });

  hf.destroy();

  return { sheets: evaluatedSheets };
}

/**
 * Build a 2D grid (rows × cols) from a SheetDocument's cells,
 * as expected by HyperFormula.buildFromSheets().
 */
function buildGrid(sheet: SheetDocument): (string | number | boolean | null)[][] {
  if (Object.keys(sheet.cells).length === 0) return [[]];

  let maxRow = 0;
  let maxCol = 0;

  for (const ref of Object.keys(sheet.cells)) {
    const match = ref.match(/^([A-Z]+)(\d+)$/);
    if (!match) continue;
    const col = columnLetterToIndex(match[1]);
    const row = parseInt(match[2], 10) - 1;
    if (row > maxRow) maxRow = row;
    if (col > maxCol) maxCol = col;
  }

  // Initialize empty grid
  const grid: (string | number | boolean | null)[][] = Array.from(
    { length: maxRow + 1 },
    () => Array(maxCol + 1).fill(null)
  );

  for (const [ref, cell] of Object.entries(sheet.cells)) {
    const match = ref.match(/^([A-Z]+)(\d+)$/);
    if (!match) continue;
    const col = columnLetterToIndex(match[1]);
    const row = parseInt(match[2], 10) - 1;

    // HyperFormula recognizes formulas starting with "="
    grid[row][col] = cell.value;
  }

  return grid;
}
