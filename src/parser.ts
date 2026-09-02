// parser.ts
import type {
  WorkbookDocument,
  SheetDocument,
  SheetMeta,
  CellData,
  CellStyle,
  ChartConfig,
  ChartType,
  DataValidationRule,
  DataValidationType,
  DataValidationOperator,
  ConditionalFormattingRule,
  ConditionalFormatType,
  ConditionalFormatOperator,
  RichTextSegment,
} from "./types";
import { columnLetterToIndex } from "./chart";

type Section =
  | "meta"
  | "cells"
  | "merges"
  | "charts"
  | "dataValidation"
  | "conditionalFormatting"
  | null;

// ─────────────────────────────────────────────
// Public API
// ─────────────────────────────────────────────

/**
 * Parse a single-sheet .antsheet text into a SheetDocument.
 * For multi-sheet documents, use parseWorkbook().
 */
export function parseAntsheet(text: string): SheetDocument {
  const wb = parseWorkbook(text);
  return wb.sheets[0] ?? { name: "Untitled", meta: {}, cells: {}, merges: [] };
}

/**
 * Parse a (possibly multi-sheet) .antsheet text into a WorkbookDocument.
 * Sheets are separated by lines containing only `---`.
 */
export function parseWorkbook(text: string): WorkbookDocument {
  // Split on sheet-separator lines (--- alone on a line)
  const sheetTexts = text.split(/^---+\s*$/m).map((s) => s.trim()).filter(Boolean);
  const sheets = sheetTexts.map(parseSheet);
  return { sheets };
}

// ─────────────────────────────────────────────
// Single sheet parser
// ─────────────────────────────────────────────
function parseSheet(text: string): SheetDocument {
  const lines = text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l.length > 0 && !l.startsWith("//"));

  let name = "Untitled";
  let section: Section = null;

  const meta: SheetMeta = {};
  const cells: Record<string, CellData> = {};
  const merges: string[] = [];
  const charts: ChartConfig[] = [];
  const dataValidations: DataValidationRule[] = [];
  const conditionalFormattings: ConditionalFormattingRule[] = [];

  for (const line of lines) {
    if (line.startsWith("# ")) {
      name = line.slice(2).trim();
      continue;
    }
    if (line === "## Meta") { section = "meta"; continue; }
    if (line === "## Cells") { section = "cells"; continue; }
    if (line === "## Merges") { section = "merges"; continue; }
    if (line === "## Charts") { section = "charts"; continue; }
    if (line === "## DataValidation" || line === "## Validation") {
      section = "dataValidation"; continue;
    }
    if (
      line === "## ConditionalFormatting" ||
      line === "## ConditionalFormat" ||
      line === "## Rules"
    ) {
      section = "conditionalFormatting"; continue;
    }

    if (section === "meta") {
      parseMetaLine(line, meta);
    } else if (section === "cells") {
      parseCellLine(line, cells);
    } else if (section === "merges") {
      merges.push(line);
    } else if (section === "charts") {
      const chart = parseChartLine(line);
      if (chart) charts.push(chart);
    } else if (section === "dataValidation") {
      const rule = parseDataValidationLine(line);
      if (rule) dataValidations.push(rule);
    } else if (section === "conditionalFormatting") {
      const cf = parseConditionalFormattingLine(line);
      if (cf) conditionalFormattings.push(cf);
    }
  }

  const doc: SheetDocument = { name, meta, cells, merges };
  if (charts.length > 0) doc.charts = charts;
  if (dataValidations.length > 0) doc.dataValidations = dataValidations;
  if (conditionalFormattings.length > 0) doc.conditionalFormattings = conditionalFormattings;

  return doc;
}

// ─────────────────────────────────────────────
// Meta
// ─────────────────────────────────────────────
function parseMetaLine(line: string, meta: SheetMeta) {
  if (line.startsWith("columnWidth:")) {
    meta.columnWidth = parseKeyValueNumbers(line.slice("columnWidth:".length).trim());
  } else if (line.startsWith("rowHeight:")) {
    meta.rowHeight = parseKeyValueNumbers(line.slice("rowHeight:".length).trim());
  } else if (line.startsWith("freezeRow:")) {
    const raw = line.slice("freezeRow:".length).trim();
    if (!isNaN(Number(raw))) meta.freezeRow = Number(raw);
  } else if (line.startsWith("freezeColumn:")) {
    const raw = line.slice("freezeColumn:".length).trim();
    meta.freezeColumn = !isNaN(Number(raw))
      ? Number(raw)
      : columnLetterToIndex(raw.toUpperCase()) + 1;
  } else if (line.startsWith("freeze:")) {
    const parts = line.slice("freeze:".length).trim().split(",").map((p) => p.trim());
    for (const part of parts) {
      const [k, v] = part.split("=").map((s) => s.trim());
      if (k === "row" && !isNaN(Number(v))) meta.freezeRow = Number(v);
      if (k === "col" || k === "column") {
        meta.freezeColumn = !isNaN(Number(v)) ? Number(v) : columnLetterToIndex(v.toUpperCase()) + 1;
      }
    }
  } else if (line.startsWith("filter:") || line.startsWith("autoFilter:")) {
    meta.filter = line.slice(line.indexOf(":") + 1).trim().toUpperCase();
  } else if (line.startsWith("namedRange:")) {
    const raw = line.slice("namedRange:".length).trim();
    const eqIdx = raw.indexOf("=");
    if (eqIdx !== -1) {
      const rangeName = raw.slice(0, eqIdx).trim();
      const rangeVal = raw.slice(eqIdx + 1).trim().toUpperCase();
      if (!meta.namedRanges) meta.namedRanges = {};
      meta.namedRanges[rangeName] = rangeVal;
    }
  } else if (line.startsWith("hiddenRows:")) {
    const raw = line.slice("hiddenRows:".length).trim();
    meta.hiddenRows = raw.split(",").map((s) => parseInt(s.trim(), 10)).filter((n) => !isNaN(n));
  } else if (line.startsWith("hiddenColumns:")) {
    const raw = line.slice("hiddenColumns:".length).trim();
    meta.hiddenColumns = raw.split(",").map((s) => s.trim().toUpperCase()).filter(Boolean);
  }
}

function parseKeyValueNumbers(raw: string): Record<string, number> {
  const result: Record<string, number> = {};
  for (const part of raw.split(",").map((p) => p.trim()).filter(Boolean)) {
    const [key, value] = part.split("=").map((s) => s.trim());
    if (key && value && !isNaN(Number(value))) {
      result[key] = Number(value);
    }
  }
  return result;
}

// ─────────────────────────────────────────────
// Cells
// ─────────────────────────────────────────────
function parseCellLine(line: string, cells: Record<string, CellData>) {
  const eqIndex = line.indexOf("=");
  if (eqIndex === -1) return;

  const ref = line.slice(0, eqIndex).trim().toUpperCase();
  if (!/^[A-Z]+\d+$/.test(ref)) return;

  const rest = line.slice(eqIndex + 1).trim();
  const pipeIndex = rest.indexOf("|");

  let valueRaw: string;
  let propsRaw: string | null = null;

  if (pipeIndex === -1) {
    valueRaw = rest;
  } else {
    valueRaw = rest.slice(0, pipeIndex).trim();
    propsRaw = rest.slice(pipeIndex + 1).trim();
  }

  // Rich text cell
  if (valueRaw === "richtext" || valueRaw === "richText") {
    const cellData: CellData = { value: null };
    if (propsRaw) {
      const parsed = parseCellProps(propsRaw);
      if (parsed.richText) cellData.richText = parsed.richText;
      if (parsed.comment) cellData.comment = parsed.comment;
      if (parsed.style && Object.keys(parsed.style).length > 0) cellData.style = parsed.style;
    }
    cells[ref] = cellData;
    return;
  }

  const value = parseValue(valueRaw);
  const cellData: CellData = { value };

  if (propsRaw) {
    const parsed = parseCellProps(propsRaw);
    if (parsed.style && Object.keys(parsed.style).length > 0) cellData.style = parsed.style;
    if (parsed.comment) cellData.comment = parsed.comment;
    if (parsed.richText) cellData.richText = parsed.richText;
  }

  cells[ref] = cellData;
}

interface ParsedCellProps {
  style: CellStyle;
  comment?: string;
  richText?: RichTextSegment[];
}

function parseCellProps(raw: string): ParsedCellProps {
  const style: CellStyle = {};
  let comment: string | undefined;
  let richText: RichTextSegment[] | undefined;

  // Extract segments: [...] before splitting by commas
  const segmentsMatch = raw.match(/segments:\s*(\[.*?\])/s);
  if (segmentsMatch) {
    richText = parseRichTextSegments(segmentsMatch[1]);
    raw = raw.replace(segmentsMatch[0], "").replace(/,\s*,/, ",").replace(/^\s*,|,\s*$/, "").trim();
  }

  const parts = splitStyleParts(raw);
  for (const part of parts) {
    const colonIndex = part.indexOf(":");
    if (colonIndex === -1) continue;
    const key = part.slice(0, colonIndex).trim();
    const value = part.slice(colonIndex + 1).trim();

    if (key === "comment") {
      comment = cleanQuotedString(value);
    } else {
      assignStyle(style, key, value);
    }
  }

  return { style, comment, richText };
}

function parseRichTextSegments(raw: string): RichTextSegment[] {
  // Simple JSON-like array parser: [{text: "...", bold: true}, ...]
  const cleaned = raw.trim().slice(1, -1).trim(); // remove [ ]
  const segments: RichTextSegment[] = [];

  // Split on }, { boundary
  const segRaw = cleaned.split(/\},\s*\{/).map((s, i, arr) => {
    let str = s.trim();
    if (i > 0) str = "{" + str;
    if (i < arr.length - 1) str = str + "}";
    if (!str.startsWith("{")) str = "{" + str;
    if (!str.endsWith("}")) str = str + "}";
    return str;
  });

  for (const segStr of segRaw) {
    const seg: RichTextSegment = { text: "" };
    const textMatch = segStr.match(/text:\s*"([^"]*)"/);
    if (textMatch) seg.text = textMatch[1];
    if (/bold:\s*true/.test(segStr)) seg.bold = true;
    if (/italic:\s*true/.test(segStr)) seg.italic = true;
    if (/underline:\s*true/.test(segStr)) seg.underline = true;
    if (/strikethrough:\s*true/.test(segStr)) seg.strikethrough = true;
    const colorMatch = segStr.match(/color:\s*(#[0-9a-fA-F]{3,8}|[a-z]+)/);
    if (colorMatch) seg.color = colorMatch[1];
    const bgMatch = segStr.match(/backgroundColor:\s*(#[0-9a-fA-F]{3,8}|[a-z]+)/);
    if (bgMatch) seg.backgroundColor = bgMatch[1];
    const fsMatch = segStr.match(/fontSize:\s*(\d+)/);
    if (fsMatch) seg.fontSize = parseInt(fsMatch[1], 10);
    if (seg.text) segments.push(seg);
  }

  return segments;
}

function parseValue(raw: string): string | number | boolean | null {
  if (raw === "true") return true;
  if (raw === "false") return false;
  if (raw === "null" || raw === "") return null;
  if (
    (raw.startsWith('"') && raw.endsWith('"')) ||
    (raw.startsWith("'") && raw.endsWith("'"))
  ) {
    return raw.slice(1, -1);
  }
  if (!isNaN(Number(raw)) && raw.trim() !== "") return Number(raw);
  return raw;
}

function splitStyleParts(raw: string): string[] {
  return raw.split(",").map((p) => p.trim()).filter(Boolean);
}

function assignStyle(style: CellStyle, key: string, value: string) {
  switch (key) {
    case "fontFamily": style.fontFamily = value; break;
    case "fontSize": style.fontSize = Number(value); break;
    case "fontWeight": style.fontWeight = isNaN(Number(value)) ? value : Number(value); break;
    case "fontStyle": style.fontStyle = value as CellStyle["fontStyle"]; break;
    case "textDecoration": style.textDecoration = value; break;
    case "textTransform": style.textTransform = value; break;
    case "letterSpacing": style.letterSpacing = Number(value); break;
    case "lineHeight": style.lineHeight = isNaN(Number(value)) ? value : Number(value); break;
    case "textWrap": style.textWrap = value as CellStyle["textWrap"]; break;
    case "wrapText": style.wrapText = value.toLowerCase() === "true"; break;
    case "color": style.color = value; break;
    case "backgroundColor": style.backgroundColor = value; break;
    case "textAlign": style.textAlign = value as CellStyle["textAlign"]; break;
    case "verticalAlign": style.verticalAlign = value as CellStyle["verticalAlign"]; break;
    case "format": style.format = value; break;
    case "border": style.border = value; break;
    case "borderTop": style.borderTop = parseBorderValue(value); break;
    case "borderRight": style.borderRight = parseBorderValue(value); break;
    case "borderBottom": style.borderBottom = parseBorderValue(value); break;
    case "borderLeft": style.borderLeft = parseBorderValue(value); break;
    case "borderColor": style.borderColor = value; break;
    case "borderWidth": style.borderWidth = Number(value); break;
    case "borderStyle": style.borderStyle = value; break;
    case "padding": style.padding = isNaN(Number(value)) ? value : Number(value); break;
    case "paddingTop": style.paddingTop = Number(value); break;
    case "paddingRight": style.paddingRight = Number(value); break;
    case "paddingBottom": style.paddingBottom = Number(value); break;
    case "paddingLeft": style.paddingLeft = Number(value); break;
  }
}

function parseBorderValue(value: string): string | number {
  if (!isNaN(Number(value)) && value.trim() !== "") return Number(value);
  return value;
}

// ─────────────────────────────────────────────
// Charts
// ─────────────────────────────────────────────
function parseChartLine(line: string): ChartConfig | null {
  const eqIndex = line.indexOf("=");
  if (eqIndex === -1) return null;

  const id = line.slice(0, eqIndex).trim();
  const rest = line.slice(eqIndex + 1).trim();
  const pipeIndex = rest.indexOf("|");

  const type = (pipeIndex === -1 ? rest : rest.slice(0, pipeIndex)).trim().toLowerCase() as ChartType;
  const config: ChartConfig = { id, type };

  if (pipeIndex !== -1) {
    parseChartProperties(rest.slice(pipeIndex + 1).trim(), config);
  }

  return config;
}

function parseChartProperties(raw: string, config: ChartConfig) {
  for (const part of splitConfigProps(raw)) {
    const colonIndex = part.indexOf(":");
    if (colonIndex === -1) continue;
    const key = part.slice(0, colonIndex).trim();
    const value = part.slice(colonIndex + 1).trim();

    switch (key) {
      case "title": config.title = cleanQuotedString(value); break;
      case "range": config.range = value.toUpperCase(); break;
      case "x": case "xAxis": config.xAxis = value.toUpperCase(); break;
      case "series": config.series = parseArrayOrList(value); break;
      case "labels": config.labels = value.toUpperCase(); break;
      case "values": config.values = value.toUpperCase(); break;
      case "position": config.position = value.toUpperCase(); break;
      case "colors": config.colors = parseArrayOrList(value); break;
      case "stacked": config.stacked = value.toLowerCase() === "true"; break;
      case "horizontal": config.horizontal = value.toLowerCase() === "true"; break;
      case "smooth": config.smooth = value.toLowerCase() === "true"; break;
    }
  }
}

// ─────────────────────────────────────────────
// Data Validation
// ─────────────────────────────────────────────
function parseDataValidationLine(line: string): DataValidationRule | null {
  const eqIndex = line.indexOf("=");
  if (eqIndex === -1) return null;

  const id = line.slice(0, eqIndex).trim();
  const rest = line.slice(eqIndex + 1).trim();
  const pipeIndex = rest.indexOf("|");

  const typeRaw = (pipeIndex === -1 ? rest : rest.slice(0, pipeIndex)).trim().toLowerCase() as DataValidationType;
  const rule: DataValidationRule = { id, type: typeRaw, range: "" };

  if (pipeIndex !== -1) {
    for (const prop of splitConfigProps(rest.slice(pipeIndex + 1).trim())) {
      const colonIndex = prop.indexOf(":");
      if (colonIndex === -1) continue;
      const key = prop.slice(0, colonIndex).trim();
      const value = prop.slice(colonIndex + 1).trim();

      switch (key) {
        case "range": rule.range = value.toUpperCase(); break;
        case "options": case "list": rule.options = parseArrayOrList(value); break;
        case "operator": rule.operator = value as DataValidationOperator; break;
        case "min": rule.min = isNaN(Number(value)) ? cleanQuotedString(value) : Number(value); break;
        case "max": rule.max = isNaN(Number(value)) ? cleanQuotedString(value) : Number(value); break;
        case "value": rule.value = isNaN(Number(value)) ? cleanQuotedString(value) : Number(value); break;
        case "allowBlank": rule.allowBlank = value.toLowerCase() === "true"; break;
        case "error": rule.error = cleanQuotedString(value); break;
        case "prompt": rule.prompt = cleanQuotedString(value); break;
      }
    }
  }

  return rule.range ? rule : null;
}

// ─────────────────────────────────────────────
// Conditional Formatting
// ─────────────────────────────────────────────
function parseConditionalFormattingLine(line: string): ConditionalFormattingRule | null {
  const eqIndex = line.indexOf("=");
  if (eqIndex === -1) return null;

  const id = line.slice(0, eqIndex).trim();
  const rest = line.slice(eqIndex + 1).trim();
  const pipeIndex = rest.indexOf("|");

  const typeRaw = (pipeIndex === -1 ? rest : rest.slice(0, pipeIndex)).trim() as ConditionalFormatType;
  const cf: ConditionalFormattingRule = { id, type: typeRaw, range: "" };

  if (pipeIndex !== -1) {
    for (const prop of splitConfigProps(rest.slice(pipeIndex + 1).trim())) {
      const colonIndex = prop.indexOf(":");
      if (colonIndex === -1) continue;
      const key = prop.slice(0, colonIndex).trim();
      const value = prop.slice(colonIndex + 1).trim();

      switch (key) {
        case "range": cf.range = value.toUpperCase(); break;
        case "operator": cf.operator = cleanQuotedString(value) as ConditionalFormatOperator; break;
        case "value": cf.value = isNaN(Number(value)) ? cleanQuotedString(value) : Number(value); break;
        case "value2": cf.value2 = isNaN(Number(value)) ? cleanQuotedString(value) : Number(value); break;
        case "color": cf.color = value; break;
        case "backgroundColor": case "bg": cf.backgroundColor = value; break;
        case "minColor": cf.minColor = value; break;
        case "maxColor": cf.maxColor = value; break;
        case "midColor": cf.midColor = value; break;
      }
    }
  }

  return cf.range ? cf : null;
}

// ─────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────
function splitConfigProps(raw: string): string[] {
  const parts: string[] = [];
  let current = "";
  let insideQuotes = false;
  let insideBracket = false;

  for (const char of raw) {
    if (char === '"' || char === "'") {
      insideQuotes = !insideQuotes;
    } else if (char === "[" && !insideQuotes) {
      insideBracket = true;
    } else if (char === "]" && !insideQuotes) {
      insideBracket = false;
    }

    if (char === "," && !insideQuotes && !insideBracket) {
      if (current.trim()) parts.push(current.trim());
      current = "";
    } else {
      current += char;
    }
  }

  if (current.trim()) parts.push(current.trim());
  return parts;
}

function cleanQuotedString(val: string): string {
  if (
    (val.startsWith('"') && val.endsWith('"')) ||
    (val.startsWith("'") && val.endsWith("'"))
  ) {
    return val.slice(1, -1);
  }
  return val;
}

function parseArrayOrList(val: string): string[] {
  let cleaned = val.trim();
  if (cleaned.startsWith("[") && cleaned.endsWith("]")) {
    cleaned = cleaned.slice(1, -1);
  }
  return cleaned
    .split(",")
    .map((s) => cleanQuotedString(s.trim()))
    .filter(Boolean);
}