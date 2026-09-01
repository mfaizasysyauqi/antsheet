// parser.ts
import type { SheetDocument, SheetMeta, CellData, CellStyle } from "../src/types";

type Section = "meta" | "cells" | "merges" | null;

export function parseAntsheet(text: string): SheetDocument {
  const lines = text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l.length > 0 && !l.startsWith("//"));

  let name = "Untitled";
  let section: Section = null;

  const meta: SheetMeta = {};
  const cells: Record<string, CellData> = {};
  const merges: string[] = [];

  for (const line of lines) {
    // Sheet name
    if (line.startsWith("# ")) {
      name = line.slice(2).trim();
      continue;
    }

    // Section headers
    if (line === "## Meta") {
      section = "meta";
      continue;
    }
    if (line === "## Cells") {
      section = "cells";
      continue;
    }
    if (line === "## Merges") {
      section = "merges";
      continue;
    }

    // Content
    if (section === "meta") {
      parseMetaLine(line, meta);
    } else if (section === "cells") {
      parseCellLine(line, cells);
    } else if (section === "merges") {
      merges.push(line);
    }
  }

  return { name, meta, cells, merges };
}

// ─────────────────────────────────────────────
// Meta
// ─────────────────────────────────────────────
function parseMetaLine(line: string, meta: SheetMeta) {
  if (line.startsWith("columnWidth:")) {
    const raw = line.slice("columnWidth:".length).trim();
    meta.columnWidth = parseKeyValueNumbers(raw);
  } else if (line.startsWith("rowHeight:")) {
    const raw = line.slice("rowHeight:".length).trim();
    meta.rowHeight = parseKeyValueNumbers(raw);
  }
}

function parseKeyValueNumbers(raw: string): Record<string, number> {
  const result: Record<string, number> = {};
  const parts = raw.split(",").map((p) => p.trim());

  for (const part of parts) {
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
  // Format: A1 = "value" | style1: val1, style2: val2
  const eqIndex = line.indexOf("=");
  if (eqIndex === -1) return;

  const ref = line.slice(0, eqIndex).trim().toUpperCase();
  if (!/^[A-Z]+\d+$/.test(ref)) return;

  const rest = line.slice(eqIndex + 1).trim();

  // Pisahkan value dan style
  const pipeIndex = rest.indexOf("|");
  let valueRaw: string;
  let styleRaw: string | null = null;

  if (pipeIndex === -1) {
    valueRaw = rest;
  } else {
    valueRaw = rest.slice(0, pipeIndex).trim();
    styleRaw = rest.slice(pipeIndex + 1).trim();
  }

  const value = parseValue(valueRaw);
  const style = styleRaw ? parseStyle(styleRaw) : undefined;

  cells[ref] = { value, style };
}

function parseValue(raw: string): string | number | boolean | null {
  if (raw === "true") return true;
  if (raw === "false") return false;
  if (raw === "null" || raw === "") return null;

  // String dengan quotes
  if (
    (raw.startsWith('"') && raw.endsWith('"')) ||
    (raw.startsWith("'") && raw.endsWith("'"))
  ) {
    return raw.slice(1, -1);
  }

  // Number
  if (!isNaN(Number(raw)) && raw.trim() !== "") {
    return Number(raw);
  }

  // Fallback: treat as string
  return raw;
}

function parseStyle(raw: string): CellStyle {
  const style: CellStyle = {};
  const parts = splitStyleParts(raw);

  for (const part of parts) {
    const colonIndex = part.indexOf(":");
    if (colonIndex === -1) continue;

    const key = part.slice(0, colonIndex).trim();
    const value = part.slice(colonIndex + 1).trim();

    assignStyle(style, key, value);
  }

  return style;
}

/**
 * Split style string by comma, but ignore commas inside values if needed.
 * Simple version: split by ", " (comma + space) is usually enough.
 */
function splitStyleParts(raw: string): string[] {
  return raw.split(",").map((p) => p.trim()).filter(Boolean);
}

function assignStyle(style: CellStyle, key: string, value: string) {
  switch (key) {
    // Font
    case "fontFamily":
      style.fontFamily = value;
      break;
    case "fontSize":
      style.fontSize = Number(value);
      break;
    case "fontWeight":
      style.fontWeight = isNaN(Number(value)) ? value : Number(value);
      break;
    case "fontStyle":
      style.fontStyle = value as CellStyle["fontStyle"];
      break;
    case "textDecoration":
      style.textDecoration = value;
      break;
    case "textTransform":
      style.textTransform = value;
      break;
    case "letterSpacing":
      style.letterSpacing = Number(value);
      break;
    case "lineHeight":
      style.lineHeight = isNaN(Number(value)) ? value : Number(value);
      break;

    // Color & Align
    case "color":
      style.color = value;
      break;
    case "backgroundColor":
      style.backgroundColor = value;
      break;
    case "textAlign":
      style.textAlign = value as CellStyle["textAlign"];
      break;
    case "verticalAlign":
      style.verticalAlign = value as CellStyle["verticalAlign"];
      break;

    // Format
    case "format":
      style.format = value;
      break;

    // Border
    case "border":
      style.border = value;
      break;
    case "borderTop":
      style.borderTop = parseBorderValue(value);
      break;
    case "borderRight":
      style.borderRight = parseBorderValue(value);
      break;
    case "borderBottom":
      style.borderBottom = parseBorderValue(value);
      break;
    case "borderLeft":
      style.borderLeft = parseBorderValue(value);
      break;
    case "borderColor":
      style.borderColor = value;
      break;
    case "borderWidth":
      style.borderWidth = Number(value);
      break;
    case "borderStyle":
      style.borderStyle = value;
      break;

    // Padding
    case "padding":
      style.padding = isNaN(Number(value)) ? value : Number(value);
      break;
    case "paddingTop":
      style.paddingTop = Number(value);
      break;
    case "paddingRight":
      style.paddingRight = Number(value);
      break;
    case "paddingBottom":
      style.paddingBottom = Number(value);
      break;
    case "paddingLeft":
      style.paddingLeft = Number(value);
      break;
  }
}

function parseBorderValue(value: string): string | number {
  // "2" → number, "2 solid #000" → string
  if (!isNaN(Number(value)) && value.trim() !== "") {
    return Number(value);
  }
  return value;
}