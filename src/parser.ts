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
import { colLetterToIndex, CELL_REF_REGEX } from "./coords";
import {
  cleanQuotedString,
  parseArrayOrList,
  parseNumericOrString,
  parseDeclarationLine,
  splitPipePayload,
  parseProperties,
} from "./utils";

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
 * Sheets can be separated either by `---` or by top-level `# SheetName` headers.
 */
export function parseWorkbook(text: string): WorkbookDocument {
  const explicitChunks = text
    .split(/^---+\s*$/m)
    .map((s) => s.trim())
    .filter(Boolean);

  const rawSheets: string[] = [];

  for (const chunk of explicitChunks) {
    const lines = chunk.split(/\r?\n/);
    let currentSheetLines: string[] = [];

    for (const line of lines) {
      const trimmed = line.trim();
      // New top-level sheet header (# SheetName) starts a new sheet if we already have content
      if (trimmed.startsWith("# ") && currentSheetLines.some((l) => l.trim().length > 0)) {
        rawSheets.push(currentSheetLines.join("\n").trim());
        currentSheetLines = [line];
      } else {
        currentSheetLines.push(line);
      }
    }

    if (currentSheetLines.length > 0 && currentSheetLines.some((l) => l.trim().length > 0)) {
      rawSheets.push(currentSheetLines.join("\n").trim());
    }
  }

  const sheets = (rawSheets.length > 0 ? rawSheets : [text.trim()])
    .map(parseSheet)
    .filter((s) => Boolean(s.name || Object.keys(s.cells).length > 0));

  return { sheets: sheets.length > 0 ? sheets : [parseSheet(text)] };
}

// ─────────────────────────────────────────────
// Single Sheet Parser
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

    const nextSection = identifySection(line);
    if (nextSection) {
      section = nextSection;
      continue;
    }

    switch (section) {
      case "meta":
        applyMetaLine(line, meta);
        break;
      case "cells": {
        const parsedCell = parseCellLine(line);
        if (parsedCell) cells[parsedCell.ref] = parsedCell.data;
        break;
      }
      case "merges":
        merges.push(line);
        break;
      case "charts": {
        const chart = parseChartLine(line);
        if (chart) charts.push(chart);
        break;
      }
      case "dataValidation": {
        const rule = parseDataValidationLine(line);
        if (rule) dataValidations.push(rule);
        break;
      }
      case "conditionalFormatting": {
        const cf = parseConditionalFormattingLine(line);
        if (cf) conditionalFormattings.push(cf);
        break;
      }
    }
  }

  const doc: SheetDocument = { name, meta, cells, merges };
  if (charts.length > 0) doc.charts = charts;
  if (dataValidations.length > 0) doc.dataValidations = dataValidations;
  if (conditionalFormattings.length > 0) doc.conditionalFormattings = conditionalFormattings;

  return doc;
}

function identifySection(line: string): Section | null {
  switch (line) {
    case "## Meta":
      return "meta";
    case "## Cells":
      return "cells";
    case "## Merges":
      return "merges";
    case "## Charts":
      return "charts";
    case "## DataValidation":
    case "## Validation":
      return "dataValidation";
    case "## ConditionalFormatting":
    case "## ConditionalFormat":
    case "## Rules":
      return "conditionalFormatting";
    default:
      return null;
  }
}

// ─────────────────────────────────────────────
// Meta Parsing
// ─────────────────────────────────────────────

function applyMetaLine(line: string, meta: SheetMeta): void {
  if (line.startsWith("columnWidth:")) {
    meta.columnWidth = parseKeyValueNumbers(line.slice("columnWidth:".length));
  } else if (line.startsWith("rowHeight:")) {
    meta.rowHeight = parseKeyValueNumbers(line.slice("rowHeight:".length));
  } else if (line.startsWith("freezeRow:")) {
    const raw = line.slice("freezeRow:".length).trim();
    if (!Number.isNaN(Number(raw))) meta.freezeRow = Number(raw);
  } else if (line.startsWith("freezeColumn:")) {
    const raw = line.slice("freezeColumn:".length).trim();
    meta.freezeColumn = !Number.isNaN(Number(raw))
      ? Number(raw)
      : colLetterToIndex(raw.toUpperCase()) + 1;
  } else if (line.startsWith("freeze:")) {
    const parts = line.slice("freeze:".length).trim().split(",").map((p) => p.trim());
    for (const part of parts) {
      const [k, v] = part.split("=").map((s) => s.trim());
      if (k === "row" && !Number.isNaN(Number(v))) meta.freezeRow = Number(v);
      if (k === "col" || k === "column") {
        meta.freezeColumn = !Number.isNaN(Number(v)) ? Number(v) : colLetterToIndex(v.toUpperCase()) + 1;
      }
    }
  } else if (line.startsWith("filter:") || line.startsWith("autoFilter:")) {
    meta.filter = line.slice(line.indexOf(":") + 1).trim().toUpperCase();
  } else if (line.startsWith("namedRange:")) {
    const decl = parseDeclarationLine(line.slice("namedRange:".length).trim());
    if (decl) {
      if (!meta.namedRanges) meta.namedRanges = {};
      meta.namedRanges[decl.key] = decl.payload.toUpperCase();
    }
  } else if (line.startsWith("hiddenRows:")) {
    meta.hiddenRows = line
      .slice("hiddenRows:".length)
      .split(",")
      .map((s) => parseInt(s.trim(), 10))
      .filter((n) => !Number.isNaN(n));
  } else if (line.startsWith("hiddenColumns:")) {
    meta.hiddenColumns = line
      .slice("hiddenColumns:".length)
      .split(",")
      .map((s) => s.trim().toUpperCase())
      .filter(Boolean);
  }
}

function parseKeyValueNumbers(raw: string): Record<string, number> {
  const result: Record<string, number> = {};
  for (const part of raw.split(",").map((p) => p.trim()).filter(Boolean)) {
    const [key, value] = part.split("=").map((s) => s.trim());
    if (key && value && !Number.isNaN(Number(value))) {
      result[key] = Number(value);
    }
  }
  return result;
}

// ─────────────────────────────────────────────
// Cell Parsing
// ─────────────────────────────────────────────

function parseCellLine(line: string): { ref: string; data: CellData } | null {
  const decl = parseDeclarationLine(line);
  if (!decl) return null;

  const ref = decl.key.toUpperCase();
  if (!CELL_REF_REGEX.test(ref)) return null;

  const { main: valueRaw, props: propsRaw } = splitPipePayload(decl.payload);

  // Rich text cell
  if (valueRaw.toLowerCase() === "richtext") {
    const cellData: CellData = { value: null };
    if (propsRaw) {
      const parsed = parseCellProps(propsRaw);
      if (parsed.richText) cellData.richText = parsed.richText;
      if (parsed.comment) cellData.comment = parsed.comment;
      if (Object.keys(parsed.style).length > 0) cellData.style = parsed.style;
    }
    return { ref, data: cellData };
  }

  const value = parseCellValue(valueRaw);
  const cellData: CellData = { value };

  if (propsRaw) {
    const parsed = parseCellProps(propsRaw);
    if (Object.keys(parsed.style).length > 0) cellData.style = parsed.style;
    if (parsed.comment) cellData.comment = parsed.comment;
    if (parsed.richText) cellData.richText = parsed.richText;
  }

  return { ref, data: cellData };
}

interface ParsedCellProps {
  style: CellStyle;
  comment?: string;
  richText?: RichTextSegment[];
}

function parseCellProps(raw: string): ParsedCellProps {
  let cleanedRaw = raw;
  let richText: RichTextSegment[] | undefined;

  // Extract segments: [...] before parsing regular properties
  const segmentsMatch = cleanedRaw.match(/segments:\s*(\[.*?\])/s);
  if (segmentsMatch) {
    richText = parseRichTextSegments(segmentsMatch[1]);
    cleanedRaw = cleanedRaw
      .replace(segmentsMatch[0], "")
      .replace(/,\s*,/, ",")
      .replace(/^\s*,|,\s*$/, "")
      .trim();
  }

  const props = parseProperties(cleanedRaw);
  const style: CellStyle = {};
  let comment: string | undefined;

  for (const [key, value] of Object.entries(props)) {
    if (key === "comment") {
      comment = cleanQuotedString(value);
    } else {
      assignStyleProperty(style, key, value);
    }
  }

  return { style, comment, richText };
}

const NUMERIC_STYLE_KEYS = new Set([
  "fontSize",
  "letterSpacing",
  "borderWidth",
  "paddingTop",
  "paddingRight",
  "paddingBottom",
  "paddingLeft",
]);

const NUMERIC_OR_STRING_STYLE_KEYS = new Set([
  "fontWeight",
  "lineHeight",
  "padding",
  "borderTop",
  "borderRight",
  "borderBottom",
  "borderLeft",
]);

function assignStyleProperty(style: CellStyle, key: string, value: string): void {
  if (NUMERIC_STYLE_KEYS.has(key)) {
    const num = Number(value);
    if (!Number.isNaN(num)) {
      (style as Record<string, unknown>)[key] = num;
    }
  } else if (NUMERIC_OR_STRING_STYLE_KEYS.has(key)) {
    (style as Record<string, unknown>)[key] = parseNumericOrString(value);
  } else if (key === "wrapText") {
    style.wrapText = value.toLowerCase() === "true";
  } else {
    // String style properties
    (style as Record<string, unknown>)[key] = cleanQuotedString(value);
  }
}

function parseCellValue(raw: string): string | number | boolean | null {
  if (raw === "true") return true;
  if (raw === "false") return false;
  if (raw === "null" || raw === "") return null;
  if (
    (raw.startsWith('"') && raw.endsWith('"')) ||
    (raw.startsWith("'") && raw.endsWith("'"))
  ) {
    return raw.slice(1, -1);
  }
  if (!Number.isNaN(Number(raw)) && raw.trim() !== "") return Number(raw);
  return raw;
}

function parseRichTextSegments(raw: string): RichTextSegment[] {
  const cleaned = raw.trim().slice(1, -1).trim(); // strip outer [ ]
  if (!cleaned) return [];

  const segRaw = cleaned.split(/\},\s*\{/).map((s, i, arr) => {
    let str = s.trim();
    if (i > 0 && !str.startsWith("{")) str = "{" + str;
    if (i < arr.length - 1 && !str.endsWith("}")) str = str + "}";
    if (!str.startsWith("{")) str = "{" + str;
    if (!str.endsWith("}")) str = str + "}";
    return str;
  });

  const segments: RichTextSegment[] = [];

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

// ─────────────────────────────────────────────
// Chart Parsing
// ─────────────────────────────────────────────

function parseChartLine(line: string): ChartConfig | null {
  const decl = parseDeclarationLine(line);
  if (!decl) return null;

  const { main: typeRaw, props: propsRaw } = splitPipePayload(decl.payload);
  const type = typeRaw.toLowerCase() as ChartType;
  const config: ChartConfig = { id: decl.key, type };

  if (propsRaw) {
    const props = parseProperties(propsRaw);
    for (const [key, value] of Object.entries(props)) {
      switch (key) {
        case "title":
          config.title = cleanQuotedString(value);
          break;
        case "range":
          config.range = value.toUpperCase();
          break;
        case "x":
        case "xAxis":
          config.xAxis = value.toUpperCase();
          break;
        case "series":
          config.series = parseArrayOrList(value);
          break;
        case "labels":
          config.labels = value.toUpperCase();
          break;
        case "values":
          config.values = value.toUpperCase();
          break;
        case "position":
          config.position = value.toUpperCase();
          break;
        case "colors":
          config.colors = parseArrayOrList(value);
          break;
        case "stacked":
          config.stacked = value.toLowerCase() === "true";
          break;
        case "horizontal":
          config.horizontal = value.toLowerCase() === "true";
          break;
        case "smooth":
          config.smooth = value.toLowerCase() === "true";
          break;
      }
    }
  }

  return config;
}

// ─────────────────────────────────────────────
// Data Validation Parsing
// ─────────────────────────────────────────────

function parseDataValidationLine(line: string): DataValidationRule | null {
  const decl = parseDeclarationLine(line);
  if (!decl) return null;

  const { main: typeRaw, props: propsRaw } = splitPipePayload(decl.payload);
  const type = typeRaw.toLowerCase() as DataValidationType;
  const rule: DataValidationRule = { id: decl.key, type, range: "" };

  if (propsRaw) {
    const props = parseProperties(propsRaw);
    for (const [key, value] of Object.entries(props)) {
      switch (key) {
        case "range":
          rule.range = value.toUpperCase();
          break;
        case "options":
        case "list":
          rule.options = parseArrayOrList(value);
          break;
        case "operator":
          rule.operator = value as DataValidationOperator;
          break;
        case "min":
          rule.min = parseNumericOrString(value);
          break;
        case "max":
          rule.max = parseNumericOrString(value);
          break;
        case "value":
          rule.value = parseNumericOrString(value);
          break;
        case "allowBlank":
          rule.allowBlank = value.toLowerCase() === "true";
          break;
        case "error":
          rule.error = cleanQuotedString(value);
          break;
        case "prompt":
          rule.prompt = cleanQuotedString(value);
          break;
      }
    }
  }

  return rule.range ? rule : null;
}

// ─────────────────────────────────────────────
// Conditional Formatting Parsing
// ─────────────────────────────────────────────

function parseConditionalFormattingLine(line: string): ConditionalFormattingRule | null {
  const decl = parseDeclarationLine(line);
  if (!decl) return null;

  const { main: typeRaw, props: propsRaw } = splitPipePayload(decl.payload);
  const type = typeRaw as ConditionalFormatType;
  const cf: ConditionalFormattingRule = { id: decl.key, type, range: "" };

  if (propsRaw) {
    const props = parseProperties(propsRaw);
    for (const [key, value] of Object.entries(props)) {
      switch (key) {
        case "range":
          cf.range = value.toUpperCase();
          break;
        case "operator":
          cf.operator = cleanQuotedString(value) as ConditionalFormatOperator;
          break;
        case "value":
          cf.value = parseNumericOrString(value);
          break;
        case "value2":
          cf.value2 = parseNumericOrString(value);
          break;
        case "color":
          cf.color = value;
          break;
        case "backgroundColor":
        case "bg":
          cf.backgroundColor = value;
          break;
        case "minColor":
          cf.minColor = value;
          break;
        case "maxColor":
          cf.maxColor = value;
          break;
        case "midColor":
          cf.midColor = value;
          break;
      }
    }
  }

  return cf.range ? cf : null;
}