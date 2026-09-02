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
// Sheet serializer
// ─────────────────────────────────────────────
function serializeSheet(doc: SheetDocument): string {
  const lines: string[] = [];

  // 1. Sheet Name
  lines.push(`# ${doc.name}`);
  lines.push("");

  // 2. Meta
  const hasMeta =
    (doc.meta.columnWidth && Object.keys(doc.meta.columnWidth).length > 0) ||
    (doc.meta.rowHeight && Object.keys(doc.meta.rowHeight).length > 0) ||
    doc.meta.freezeRow !== undefined ||
    doc.meta.freezeColumn !== undefined ||
    doc.meta.filter !== undefined ||
    (doc.meta.namedRanges && Object.keys(doc.meta.namedRanges).length > 0) ||
    (doc.meta.hiddenRows && doc.meta.hiddenRows.length > 0) ||
    (doc.meta.hiddenColumns && doc.meta.hiddenColumns.length > 0);

  if (hasMeta) {
    lines.push("## Meta");

    if (doc.meta.columnWidth) {
      const parts = Object.entries(doc.meta.columnWidth).map(([c, w]) => `${c}=${w}`).join(", ");
      lines.push(`columnWidth: ${parts}`);
    }
    if (doc.meta.rowHeight) {
      const parts = Object.entries(doc.meta.rowHeight).map(([r, h]) => `${r}=${h}`).join(", ");
      lines.push(`rowHeight: ${parts}`);
    }
    if (doc.meta.freezeRow !== undefined) lines.push(`freezeRow: ${doc.meta.freezeRow}`);
    if (doc.meta.freezeColumn !== undefined) lines.push(`freezeColumn: ${doc.meta.freezeColumn}`);
    if (doc.meta.filter !== undefined) lines.push(`filter: ${doc.meta.filter}`);
    if (doc.meta.namedRanges) {
      for (const [name, range] of Object.entries(doc.meta.namedRanges)) {
        lines.push(`namedRange: ${name} = ${range}`);
      }
    }
    if (doc.meta.hiddenRows && doc.meta.hiddenRows.length > 0) {
      lines.push(`hiddenRows: ${doc.meta.hiddenRows.join(", ")}`);
    }
    if (doc.meta.hiddenColumns && doc.meta.hiddenColumns.length > 0) {
      lines.push(`hiddenColumns: ${doc.meta.hiddenColumns.join(", ")}`);
    }

    lines.push("");
  }

  // 3. Cells
  lines.push("## Cells");
  const sortedRefs = Object.keys(doc.cells).sort(compareCellRef);
  for (const ref of sortedRefs) {
    lines.push(serializeCell(ref, doc.cells[ref]));
  }

  // 4. Merges
  if (doc.merges.length > 0) {
    lines.push("");
    lines.push("## Merges");
    for (const merge of doc.merges) lines.push(merge);
  }

  // 5. Data Validation
  if (doc.dataValidations && doc.dataValidations.length > 0) {
    lines.push("");
    lines.push("## DataValidation");
    for (const rule of doc.dataValidations) lines.push(serializeDataValidation(rule));
  }

  // 6. Conditional Formatting
  if (doc.conditionalFormattings && doc.conditionalFormattings.length > 0) {
    lines.push("");
    lines.push("## ConditionalFormatting");
    for (const cf of doc.conditionalFormattings) lines.push(serializeConditionalFormatting(cf));
  }

  // 7. Charts
  if (doc.charts && doc.charts.length > 0) {
    lines.push("");
    lines.push("## Charts");
    for (const chart of doc.charts) lines.push(serializeChart(chart));
  }

  return lines.join("\n") + "\n";
}

// ─────────────────────────────────────────────
// Cell
// ─────────────────────────────────────────────
function serializeCell(ref: string, cell: CellData): string {
  const parts: string[] = [];

  // Rich text cell
  if (cell.richText && cell.richText.length > 0) {
    const segs = cell.richText.map((seg) => {
      const props: string[] = [`text: "${seg.text}"`];
      if (seg.bold) props.push("bold: true");
      if (seg.italic) props.push("italic: true");
      if (seg.underline) props.push("underline: true");
      if (seg.strikethrough) props.push("strikethrough: true");
      if (seg.color) props.push(`color: ${seg.color}`);
      if (seg.backgroundColor) props.push(`backgroundColor: ${seg.backgroundColor}`);
      if (seg.fontSize) props.push(`fontSize: ${seg.fontSize}`);
      return `{${props.join(", ")}}`;
    });
    parts.push(`segments: [${segs.join(", ")}]`);
  }

  if (cell.style && Object.keys(cell.style).length > 0) {
    parts.push(serializeStyle(cell.style));
  }

  if (cell.comment !== undefined) {
    parts.push(`comment: "${cell.comment}"`);
  }

  const valueStr = cell.richText && cell.richText.length > 0
    ? "richtext"
    : serializeValue(cell.value);

  if (parts.length === 0) return `${ref} = ${valueStr}`;
  return `${ref} = ${valueStr} | ${parts.join(", ")}`;
}

function serializeValue(value: string | number | boolean | null): string {
  if (value === null || value === undefined) return `""`;
  if (typeof value === "string") return `"${value}"`;
  if (typeof value === "boolean") return value ? "true" : "false";
  return String(value);
}

// ─────────────────────────────────────────────
// Style
// ─────────────────────────────────────────────
function serializeStyle(style: CellStyle): string {
  const parts: string[] = [];

  if (style.fontFamily !== undefined) parts.push(`fontFamily: ${style.fontFamily}`);
  if (style.fontSize !== undefined) parts.push(`fontSize: ${style.fontSize}`);
  if (style.fontWeight !== undefined) parts.push(`fontWeight: ${style.fontWeight}`);
  if (style.fontStyle !== undefined) parts.push(`fontStyle: ${style.fontStyle}`);
  if (style.textDecoration !== undefined) parts.push(`textDecoration: ${style.textDecoration}`);
  if (style.textTransform !== undefined) parts.push(`textTransform: ${style.textTransform}`);
  if (style.letterSpacing !== undefined) parts.push(`letterSpacing: ${style.letterSpacing}`);
  if (style.lineHeight !== undefined) parts.push(`lineHeight: ${style.lineHeight}`);
  if (style.textWrap !== undefined) parts.push(`textWrap: ${style.textWrap}`);
  if (style.wrapText !== undefined) parts.push(`wrapText: ${style.wrapText}`);
  if (style.color !== undefined) parts.push(`color: ${style.color}`);
  if (style.backgroundColor !== undefined) parts.push(`backgroundColor: ${style.backgroundColor}`);
  if (style.textAlign !== undefined) parts.push(`textAlign: ${style.textAlign}`);
  if (style.verticalAlign !== undefined) parts.push(`verticalAlign: ${style.verticalAlign}`);
  if (style.format !== undefined) parts.push(`format: ${style.format}`);
  if (style.border !== undefined) parts.push(`border: ${style.border}`);
  if (style.borderTop !== undefined) parts.push(`borderTop: ${style.borderTop}`);
  if (style.borderRight !== undefined) parts.push(`borderRight: ${style.borderRight}`);
  if (style.borderBottom !== undefined) parts.push(`borderBottom: ${style.borderBottom}`);
  if (style.borderLeft !== undefined) parts.push(`borderLeft: ${style.borderLeft}`);
  if (style.borderColor !== undefined) parts.push(`borderColor: ${style.borderColor}`);
  if (style.borderWidth !== undefined) parts.push(`borderWidth: ${style.borderWidth}`);
  if (style.borderStyle !== undefined) parts.push(`borderStyle: ${style.borderStyle}`);
  if (style.padding !== undefined) parts.push(`padding: ${style.padding}`);
  if (style.paddingTop !== undefined) parts.push(`paddingTop: ${style.paddingTop}`);
  if (style.paddingRight !== undefined) parts.push(`paddingRight: ${style.paddingRight}`);
  if (style.paddingBottom !== undefined) parts.push(`paddingBottom: ${style.paddingBottom}`);
  if (style.paddingLeft !== undefined) parts.push(`paddingLeft: ${style.paddingLeft}`);

  return parts.join(", ");
}

// ─────────────────────────────────────────────
// Data Validation
// ─────────────────────────────────────────────
function serializeDataValidation(rule: DataValidationRule): string {
  const props: string[] = [`range: ${rule.range}`];

  if (rule.options && rule.options.length > 0) {
    const opts = rule.options.map((o) => (o.includes(" ") ? `"${o}"` : o)).join(", ");
    props.push(`options: [${opts}]`);
  }
  if (rule.operator !== undefined) props.push(`operator: ${rule.operator}`);
  if (rule.min !== undefined) props.push(`min: ${rule.min}`);
  if (rule.max !== undefined) props.push(`max: ${rule.max}`);
  if (rule.value !== undefined) props.push(`value: ${rule.value}`);
  if (rule.allowBlank !== undefined) props.push(`allowBlank: ${rule.allowBlank}`);
  if (rule.error !== undefined) props.push(`error: "${rule.error}"`);
  if (rule.prompt !== undefined) props.push(`prompt: "${rule.prompt}"`);

  return `${rule.id} = ${rule.type} | ${props.join(", ")}`;
}

// ─────────────────────────────────────────────
// Conditional Formatting
// ─────────────────────────────────────────────
function serializeConditionalFormatting(cf: ConditionalFormattingRule): string {
  const props: string[] = [`range: ${cf.range}`];

  if (cf.operator !== undefined) props.push(`operator: "${cf.operator}"`);
  if (cf.value !== undefined) props.push(`value: ${cf.value}`);
  if (cf.value2 !== undefined) props.push(`value2: ${cf.value2}`);
  if (cf.color !== undefined) props.push(`color: ${cf.color}`);
  if (cf.backgroundColor !== undefined) props.push(`backgroundColor: ${cf.backgroundColor}`);
  if (cf.minColor !== undefined) props.push(`minColor: ${cf.minColor}`);
  if (cf.maxColor !== undefined) props.push(`maxColor: ${cf.maxColor}`);
  if (cf.midColor !== undefined) props.push(`midColor: ${cf.midColor}`);

  return `${cf.id} = ${cf.type} | ${props.join(", ")}`;
}

// ─────────────────────────────────────────────
// Chart
// ─────────────────────────────────────────────
function serializeChart(chart: ChartConfig): string {
  const props: string[] = [];

  if (chart.title !== undefined) props.push(`title: "${chart.title}"`);
  if (chart.range !== undefined) props.push(`range: ${chart.range}`);
  if (chart.xAxis !== undefined) props.push(`x: ${chart.xAxis}`);
  if (chart.series && chart.series.length > 0) props.push(`series: [${chart.series.join(", ")}]`);
  if (chart.labels !== undefined) props.push(`labels: ${chart.labels}`);
  if (chart.values !== undefined) props.push(`values: ${chart.values}`);
  if (chart.position !== undefined) props.push(`position: ${chart.position}`);
  if (chart.colors && chart.colors.length > 0) props.push(`colors: [${chart.colors.join(", ")}]`);
  if (chart.stacked !== undefined) props.push(`stacked: ${chart.stacked}`);
  if (chart.horizontal !== undefined) props.push(`horizontal: ${chart.horizontal}`);
  if (chart.smooth !== undefined) props.push(`smooth: ${chart.smooth}`);

  return props.length === 0
    ? `${chart.id} = ${chart.type}`
    : `${chart.id} = ${chart.type} | ${props.join(", ")}`;
}

// ─────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────
function compareCellRef(a: string, b: string): number {
  const parse = (ref: string) => {
    const match = ref.match(/^([A-Z]+)(\d+)$/)!;
    return { col: columnLetterToIndex(match[1]), row: parseInt(match[2], 10) };
  };
  const pa = parse(a), pb = parse(b);
  if (pa.row !== pb.row) return pa.row - pb.row;
  return pa.col - pb.col;
}

function columnLetterToIndex(letter: string): number {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index - 1;
}