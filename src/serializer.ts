// serializer.ts
import type {
  WorkbookDocument,
  SheetDocument,
  CellData,
  CellStyle,
  ChartConfig,
  DataValidationRule,
  ConditionalFormattingRule,
} from "./types";
import { compareCellRefs } from "./coords";
import { formatProperties } from "./utils";

// ─────────────────────────────────────────────
// Public API
// ─────────────────────────────────────────────

/** Serialize a single SheetDocument back to .antsheet text */
export function serializeAntsheet(doc: SheetDocument): string {
  return serializeSheet(doc);
}

/** Serialize a WorkbookDocument (multi-sheet) to .antsheet text */
export function serializeWorkbook(wb: WorkbookDocument): string {
  return wb.sheets.map(serializeSheet).join("\n\n---\n\n");
}

// ─────────────────────────────────────────────
// Sheet Serializer
// ─────────────────────────────────────────────

function serializeSheet(doc: SheetDocument): string {
  const lines: string[] = [`# ${doc.name}`, ""];

  // 1. Meta
  const metaLines = serializeMeta(doc.meta);
  if (metaLines.length > 0) {
    lines.push("## Meta", ...metaLines, "");
  }

  // 2. Cells
  lines.push("## Cells");
  const sortedRefs = Object.keys(doc.cells).sort(compareCellRefs);
  for (const ref of sortedRefs) {
    lines.push(serializeCell(ref, doc.cells[ref]));
  }

  // 3. Merges
  if (doc.merges.length > 0) {
    lines.push("", "## Merges", ...doc.merges);
  }

  // 4. Data Validation
  if (doc.dataValidations && doc.dataValidations.length > 0) {
    lines.push("", "## DataValidation");
    for (const rule of doc.dataValidations) {
      lines.push(serializeDataValidation(rule));
    }
  }

  // 5. Conditional Formatting
  if (doc.conditionalFormattings && doc.conditionalFormattings.length > 0) {
    lines.push("", "## ConditionalFormatting");
    for (const cf of doc.conditionalFormattings) {
      lines.push(serializeConditionalFormatting(cf));
    }
  }

  // 6. Charts
  if (doc.charts && doc.charts.length > 0) {
    lines.push("", "## Charts");
    for (const chart of doc.charts) {
      lines.push(serializeChart(chart));
    }
  }

  return lines.join("\n") + "\n";
}

// ─────────────────────────────────────────────
// Meta Serialization
// ─────────────────────────────────────────────

function serializeMeta(meta: SheetDocument["meta"]): string[] {
  const lines: string[] = [];

  if (meta.columnWidth && Object.keys(meta.columnWidth).length > 0) {
    const parts = Object.entries(meta.columnWidth)
      .map(([c, w]) => `${c}=${w}`)
      .join(", ");
    lines.push(`columnWidth: ${parts}`);
  }

  if (meta.rowHeight && Object.keys(meta.rowHeight).length > 0) {
    const parts = Object.entries(meta.rowHeight)
      .map(([r, h]) => `${r}=${h}`)
      .join(", ");
    lines.push(`rowHeight: ${parts}`);
  }

  if (meta.freezeRow !== undefined) lines.push(`freezeRow: ${meta.freezeRow}`);
  if (meta.freezeColumn !== undefined) lines.push(`freezeColumn: ${meta.freezeColumn}`);
  if (meta.filter !== undefined) lines.push(`filter: ${meta.filter}`);

  if (meta.namedRanges) {
    for (const [name, range] of Object.entries(meta.namedRanges)) {
      lines.push(`namedRange: ${name} = ${range}`);
    }
  }

  if (meta.hiddenRows && meta.hiddenRows.length > 0) {
    lines.push(`hiddenRows: ${meta.hiddenRows.join(", ")}`);
  }

  if (meta.hiddenColumns && meta.hiddenColumns.length > 0) {
    lines.push(`hiddenColumns: ${meta.hiddenColumns.join(", ")}`);
  }

  return lines;
}

// ─────────────────────────────────────────────
// Cell Serialization
// ─────────────────────────────────────────────

function serializeCell(ref: string, cell: CellData): string {
  const parts: string[] = [];

  // Rich text cell
  if (cell.richText && cell.richText.length > 0) {
    const segs = cell.richText.map((seg) => {
      const segProps: Record<string, unknown> = {
        text: `"${seg.text}"`,
        bold: seg.bold,
        italic: seg.italic,
        underline: seg.underline,
        strikethrough: seg.strikethrough,
        color: seg.color,
        backgroundColor: seg.backgroundColor,
        fontSize: seg.fontSize,
      };
      return `{${formatProperties(segProps)}}`;
    });
    parts.push(`segments: [${segs.join(", ")}]`);
  }

  if (cell.style && Object.keys(cell.style).length > 0) {
    const styleStr = serializeStyle(cell.style);
    if (styleStr) parts.push(styleStr);
  }

  if (cell.comment !== undefined) {
    parts.push(`comment: "${cell.comment}"`);
  }

  const valueStr = cell.richText && cell.richText.length > 0
    ? "richtext"
    : serializeCellValue(cell.value);

  return parts.length === 0 ? `${ref} = ${valueStr}` : `${ref} = ${valueStr} | ${parts.join(", ")}`;
}

function serializeCellValue(value: string | number | boolean | null): string {
  if (value === null || value === undefined) return `""`;
  if (typeof value === "string") return `"${value}"`;
  if (typeof value === "boolean") return value ? "true" : "false";
  return String(value);
}

function serializeStyle(style: CellStyle): string {
  return formatProperties(style as Record<string, unknown>);
}

// ─────────────────────────────────────────────
// Feature Serialization
// ─────────────────────────────────────────────

function serializeDataValidation(rule: DataValidationRule): string {
  const optionsFormatted = rule.options?.map((o) => (o.includes(" ") ? `"${o}"` : o));

  const propsStr = formatProperties({
    range: rule.range,
    options: optionsFormatted,
    operator: rule.operator,
    min: rule.min,
    max: rule.max,
    value: rule.value,
    allowBlank: rule.allowBlank,
    error: rule.error !== undefined ? `"${rule.error}"` : undefined,
    prompt: rule.prompt !== undefined ? `"${rule.prompt}"` : undefined,
  });

  return `${rule.id} = ${rule.type} | ${propsStr}`;
}

function serializeConditionalFormatting(cf: ConditionalFormattingRule): string {
  const propsStr = formatProperties({
    range: cf.range,
    operator: cf.operator !== undefined ? `"${cf.operator}"` : undefined,
    value: cf.value,
    value2: cf.value2,
    color: cf.color,
    backgroundColor: cf.backgroundColor,
    minColor: cf.minColor,
    maxColor: cf.maxColor,
    midColor: cf.midColor,
  });

  return `${cf.id} = ${cf.type} | ${propsStr}`;
}

function serializeChart(chart: ChartConfig): string {
  const propsStr = formatProperties({
    title: chart.title !== undefined ? `"${chart.title}"` : undefined,
    range: chart.range,
    x: chart.xAxis,
    series: chart.series,
    labels: chart.labels,
    values: chart.values,
    position: chart.position,
    colors: chart.colors,
    stacked: chart.stacked,
    horizontal: chart.horizontal,
    smooth: chart.smooth,
  });

  return propsStr ? `${chart.id} = ${chart.type} | ${propsStr}` : `${chart.id} = ${chart.type}`;
}