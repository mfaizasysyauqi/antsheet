// chart.ts
import type { SheetDocument, ChartConfig, ChartType } from "./types";

export interface ChartDataset {
  label: string;
  data: (number | null)[];
  color?: string;
}

export interface ResolvedChartData {
  id: string;
  type: ChartType;
  title?: string;
  position?: string;
  labels: string[];
  datasets: ChartDataset[];
  options?: {
    stacked?: boolean;
    horizontal?: boolean;
    smooth?: boolean;
    colors?: string[];
  };
}

/**
 * Extracts data from SheetDocument based on ChartConfig into ready-to-render chart data
 */
export function extractChartData(
  doc: SheetDocument,
  chart: ChartConfig
): ResolvedChartData {
  let labels: string[] = [];
  const datasets: ChartDataset[] = [];

  const defaultColors = [
    "#3b82f6", // blue
    "#10b981", // emerald
    "#f59e0b", // amber
    "#ef4444", // red
    "#8b5cf6", // violet
    "#ec4899", // pink
    "#06b6d4", // cyan
    "#84cc16", // lime
  ];

  // Helper to get cell value
  const getCellValue = (ref: string): any => {
    return doc.cells[ref]?.value ?? null;
  };

  const getNumericValue = (val: any): number | null => {
    if (val === null || val === undefined || val === "") return null;
    if (typeof val === "number") return isNaN(val) ? null : val;
    const num = Number(val);
    return isNaN(num) ? null : num;
  };

  // Case 1: Explicit labels & values (Pie / Doughnut / Single Series)
  if (chart.labels && chart.values) {
    const labelCells = resolveRangeCells(chart.labels);
    const valueCells = resolveRangeCells(chart.values);

    labels = labelCells.map((ref) => {
      const val = getCellValue(ref);
      return val !== null && val !== undefined ? String(val) : "";
    });

    const data = valueCells.map((ref) => getNumericValue(getCellValue(ref)));
    datasets.push({
      label: chart.title || "Data",
      data,
      color: chart.colors?.[0] || defaultColors[0],
    });
  }
  // Case 2: Explicit xAxis & series columns/ranges
  else if (chart.xAxis && chart.series && chart.series.length > 0) {
    // Determine row range
    let labelCells: string[] = [];
    if (chart.xAxis.includes(":")) {
      labelCells = resolveRangeCells(chart.xAxis);
    } else {
      // It's a column letter like "A" - find all populated rows in that column
      labelCells = getColumnCells(doc, chart.xAxis);
    }

    labels = labelCells.map((ref) => {
      const val = getCellValue(ref);
      return val !== null && val !== undefined ? String(val) : "";
    });

    chart.series.forEach((seriesSpec, index) => {
      let seriesLabel = `Series ${index + 1}`;
      let dataCells: string[] = [];

      if (seriesSpec.includes(":")) {
        // e.g. "B2:B5"
        dataCells = resolveRangeCells(seriesSpec);
      } else {
        // e.g. "B"
        const colLetter = seriesSpec.trim();
        // Check if row 1 is a header name
        const headerRef = `${colLetter}1`;
        const headerVal = getCellValue(headerRef);
        if (headerVal !== null && headerVal !== undefined && typeof headerVal === "string") {
          seriesLabel = headerVal;
        }

        // Map matching rows from labelCells
        dataCells = labelCells.map((lRef) => {
          const rowNum = lRef.match(/\d+$/)?.[0];
          return `${colLetter}${rowNum}`;
        });
      }

      const data = dataCells.map((ref) => getNumericValue(getCellValue(ref)));
      datasets.push({
        label: seriesLabel,
        data,
        color: chart.colors?.[index] || defaultColors[index % defaultColors.length],
      });
    });
  }
  // Case 3: Table range like "A1:C5"
  else if (chart.range) {
    const coords = parseRangeCoordinates(chart.range);
    const hasHeader = true; // By default top row is series labels, left col is X categories

    const startRow = hasHeader ? coords.startRow + 1 : coords.startRow;
    const endRow = coords.endRow;

    // Collect X labels from first column
    for (let r = startRow; r <= endRow; r++) {
      const ref = `${indexToColumnLetter(coords.startCol)}${r + 1}`;
      const val = getCellValue(ref);
      labels.push(val !== null && val !== undefined ? String(val) : "");
    }

    // Collect datasets from subsequent columns
    let colorIdx = 0;
    for (let c = coords.startCol + 1; c <= coords.endCol; c++) {
      const colLetter = indexToColumnLetter(c);
      const headerRef = `${colLetter}${coords.startRow + 1}`;
      const headerVal = getCellValue(headerRef);
      const seriesLabel =
        headerVal !== null && headerVal !== undefined ? String(headerVal) : `Series ${c - coords.startCol}`;

      const data: (number | null)[] = [];
      for (let r = startRow; r <= endRow; r++) {
        const cellRef = `${colLetter}${r + 1}`;
        data.push(getNumericValue(getCellValue(cellRef)));
      }

      datasets.push({
        label: seriesLabel,
        data,
        color: chart.colors?.[colorIdx] || defaultColors[colorIdx % defaultColors.length],
      });
      colorIdx++;
    }
  }

  return {
    id: chart.id,
    type: chart.type,
    title: chart.title,
    position: chart.position,
    labels,
    datasets,
    options: {
      stacked: chart.stacked,
      horizontal: chart.horizontal,
      smooth: chart.smooth,
      colors: chart.colors,
    },
  };
}

// ─────────────────────────────────────────────
// Coordinate & Range Helpers
// ─────────────────────────────────────────────

export function columnLetterToIndex(letter: string): number {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index - 1;
}

export function indexToColumnLetter(index: number): string {
  let temp = index + 1;
  let letter = "";
  while (temp > 0) {
    const mod = (temp - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    temp = Math.floor((temp - mod) / 26);
  }
  return letter;
}

export function parseCellRef(ref: string): { col: number; row: number } {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) throw new Error(`Invalid cell ref: ${ref}`);
  return {
    col: columnLetterToIndex(match[1]),
    row: parseInt(match[2], 10) - 1,
  };
}

export function parseRangeCoordinates(range: string): {
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
} {
  const [start, end] = range.split(":");
  const startPos = parseCellRef(start);
  const endPos = end ? parseCellRef(end) : startPos;

  return {
    startCol: Math.min(startPos.col, endPos.col),
    startRow: Math.min(startPos.row, endPos.row),
    endCol: Math.max(startPos.col, endPos.col),
    endRow: Math.max(startPos.row, endPos.row),
  };
}

export function resolveRangeCells(range: string): string[] {
  const cells: string[] = [];
  const coords = parseRangeCoordinates(range);

  for (let r = coords.startRow; r <= coords.endRow; r++) {
    for (let c = coords.startCol; c <= coords.endCol; c++) {
      cells.push(`${indexToColumnLetter(c)}${r + 1}`);
    }
  }
  return cells;
}

function getColumnCells(doc: SheetDocument, colLetter: string): string[] {
  const rowNumbers = new Set<number>();
  for (const ref of Object.keys(doc.cells)) {
    const match = ref.match(/^([A-Z]+)(\d+)$/);
    if (match && match[1] === colLetter) {
      const row = parseInt(match[2], 10);
      if (row > 1) {
        // Skip header row by default
        rowNumbers.add(row);
      }
    }
  }

  return Array.from(rowNumbers)
    .sort((a, b) => a - b)
    .map((r) => `${colLetter}${r}`);
}
