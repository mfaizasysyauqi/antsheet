// chart.ts
import type { SheetDocument, ChartConfig, ChartType } from "./types";
import {
  parseRangeCoordinates,
  resolveRangeCells,
  indexToColLetter,
} from "./coords";

// Re-export coordinate utilities for backward compatibility
export {
  colLetterToIndex as columnLetterToIndex,
  indexToColLetter as indexToColumnLetter,
  parseCellRef,
  parseRangeCoordinates,
  resolveRangeCells,
} from "./coords";

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

const DEFAULT_CHART_COLORS = [
  "#3b82f6", // blue
  "#10b981", // emerald
  "#f59e0b", // amber
  "#ef4444", // red
  "#8b5cf6", // violet
  "#ec4899", // pink
  "#06b6d4", // cyan
  "#84cc16", // lime
];

function toNumericValue(val: unknown): number | null {
  if (val === null || val === undefined || val === "") return null;
  if (typeof val === "number") return Number.isNaN(val) ? null : val;
  const num = Number(val);
  return Number.isNaN(num) ? null : num;
}

function getCellValueString(doc: SheetDocument, ref: string): string {
  const val = doc.cells[ref]?.value;
  return val !== null && val !== undefined ? String(val) : "";
}

function getColumnPopulatedCells(doc: SheetDocument, colLetter: string): string[] {
  const rowNumbers = new Set<number>();
  const upperCol = colLetter.toUpperCase();

  for (const ref of Object.keys(doc.cells)) {
    const match = ref.match(/^([A-Z]+)(\d+)$/);
    if (match && match[1] === upperCol) {
      const row = parseInt(match[2], 10);
      if (row > 1) {
        rowNumbers.add(row);
      }
    }
  }

  return Array.from(rowNumbers)
    .sort((a, b) => a - b)
    .map((r) => `${upperCol}${r}`);
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

  const getCellValue = (ref: string) => doc.cells[ref]?.value ?? null;

  // Case 1: Explicit labels & values (e.g. Pie / Doughnut / Single Series)
  if (chart.labels && chart.values) {
    const labelCells = resolveRangeCells(chart.labels);
    const valueCells = resolveRangeCells(chart.values);

    labels = labelCells.map((ref) => getCellValueString(doc, ref));
    const data = valueCells.map((ref) => toNumericValue(getCellValue(ref)));

    datasets.push({
      label: chart.title || "Data",
      data,
      color: chart.colors?.[0] || DEFAULT_CHART_COLORS[0],
    });
  }
  // Case 2: Explicit xAxis & series columns/ranges
  else if (chart.xAxis && chart.series && chart.series.length > 0) {
    const labelCells = chart.xAxis.includes(":")
      ? resolveRangeCells(chart.xAxis)
      : getColumnPopulatedCells(doc, chart.xAxis);

    labels = labelCells.map((ref) => getCellValueString(doc, ref));

    for (let index = 0; index < chart.series.length; index++) {
      const seriesSpec = chart.series[index];
      let seriesLabel = `Series ${index + 1}`;
      let dataCells: string[];

      if (seriesSpec.includes(":")) {
        dataCells = resolveRangeCells(seriesSpec);
      } else {
        const colLetter = seriesSpec.trim().toUpperCase();
        const headerRef = `${colLetter}1`;
        const headerVal = getCellValue(headerRef);

        if (typeof headerVal === "string" && headerVal.length > 0) {
          seriesLabel = headerVal;
        }

        dataCells = labelCells.map((lRef) => {
          const rowNum = lRef.match(/\d+$/)?.[0];
          return `${colLetter}${rowNum}`;
        });
      }

      const data = dataCells.map((ref) => toNumericValue(getCellValue(ref)));
      datasets.push({
        label: seriesLabel,
        data,
        color: chart.colors?.[index] || DEFAULT_CHART_COLORS[index % DEFAULT_CHART_COLORS.length],
      });
    }
  }
  // Case 3: Table range like "A1:C5"
  else if (chart.range) {
    const coords = parseRangeCoordinates(chart.range);
    const startRow = coords.startRow + 1; // By default top row is headers
    const endRow = coords.endRow;

    // Collect X labels from first column
    for (let r = startRow; r <= endRow; r++) {
      const ref = `${indexToColLetter(coords.startCol)}${r + 1}`;
      labels.push(getCellValueString(doc, ref));
    }

    // Collect datasets from subsequent columns
    let colorIdx = 0;
    for (let c = coords.startCol + 1; c <= coords.endCol; c++) {
      const colLetter = indexToColLetter(c);
      const headerRef = `${colLetter}${coords.startRow + 1}`;
      const headerVal = getCellValue(headerRef);
      const seriesLabel =
        typeof headerVal === "string" && headerVal.length > 0
          ? headerVal
          : `Series ${c - coords.startCol}`;

      const data: (number | null)[] = [];
      for (let r = startRow; r <= endRow; r++) {
        const cellRef = `${colLetter}${r + 1}`;
        data.push(toNumericValue(getCellValue(cellRef)));
      }

      datasets.push({
        label: seriesLabel,
        data,
        color: chart.colors?.[colorIdx] || DEFAULT_CHART_COLORS[colorIdx % DEFAULT_CHART_COLORS.length],
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
