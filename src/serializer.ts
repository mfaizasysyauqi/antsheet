// serializer.ts
import type { SheetDocument, CellData, CellStyle } from "./types";

export function serializeAntsheet(doc: SheetDocument): string {
  const lines: string[] = [];

  // 1. Nama Sheet
  lines.push(`# ${doc.name}`);
  lines.push("");

  // 2. Meta
  const hasColumnWidth =
    doc.meta.columnWidth && Object.keys(doc.meta.columnWidth).length > 0;
  const hasRowHeight =
    doc.meta.rowHeight && Object.keys(doc.meta.rowHeight).length > 0;

  if (hasColumnWidth || hasRowHeight) {
    lines.push("## Meta");

    if (hasColumnWidth) {
      const parts = Object.entries(doc.meta.columnWidth!)
        .map(([col, width]) => `${col}=${width}`)
        .join(", ");
      lines.push(`columnWidth: ${parts}`);
    }

    if (hasRowHeight) {
      const parts = Object.entries(doc.meta.rowHeight!)
        .map(([row, height]) => `${row}=${height}`)
        .join(", ");
      lines.push(`rowHeight: ${parts}`);
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
    for (const merge of doc.merges) {
      lines.push(merge);
    }
  }

  return lines.join("\n") + "\n";
}

// ─────────────────────────────────────────────
// Cell
// ─────────────────────────────────────────────
function serializeCell(ref: string, cell: CellData): string {
  const valueStr = serializeValue(cell.value);

  if (!cell.style || Object.keys(cell.style).length === 0) {
    return `${ref} = ${valueStr}`;
  }

  const styleStr = serializeStyle(cell.style);
  return `${ref} = ${valueStr} | ${styleStr}`;
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

  // Font
  if (style.fontFamily !== undefined) parts.push(`fontFamily: ${style.fontFamily}`);
  if (style.fontSize !== undefined) parts.push(`fontSize: ${style.fontSize}`);
  if (style.fontWeight !== undefined) parts.push(`fontWeight: ${style.fontWeight}`);
  if (style.fontStyle !== undefined) parts.push(`fontStyle: ${style.fontStyle}`);
  if (style.textDecoration !== undefined) parts.push(`textDecoration: ${style.textDecoration}`);
  if (style.textTransform !== undefined) parts.push(`textTransform: ${style.textTransform}`);
  if (style.letterSpacing !== undefined) parts.push(`letterSpacing: ${style.letterSpacing}`);
  if (style.lineHeight !== undefined) parts.push(`lineHeight: ${style.lineHeight}`);

  // Color & Align
  if (style.color !== undefined) parts.push(`color: ${style.color}`);
  if (style.backgroundColor !== undefined) parts.push(`backgroundColor: ${style.backgroundColor}`);
  if (style.textAlign !== undefined) parts.push(`textAlign: ${style.textAlign}`);
  if (style.verticalAlign !== undefined) parts.push(`verticalAlign: ${style.verticalAlign}`);

  // Format
  if (style.format !== undefined) parts.push(`format: ${style.format}`);

  // Border
  if (style.border !== undefined) parts.push(`border: ${style.border}`);
  if (style.borderTop !== undefined) parts.push(`borderTop: ${style.borderTop}`);
  if (style.borderRight !== undefined) parts.push(`borderRight: ${style.borderRight}`);
  if (style.borderBottom !== undefined) parts.push(`borderBottom: ${style.borderBottom}`);
  if (style.borderLeft !== undefined) parts.push(`borderLeft: ${style.borderLeft}`);
  if (style.borderColor !== undefined) parts.push(`borderColor: ${style.borderColor}`);
  if (style.borderWidth !== undefined) parts.push(`borderWidth: ${style.borderWidth}`);
  if (style.borderStyle !== undefined) parts.push(`borderStyle: ${style.borderStyle}`);

  // Padding
  if (style.padding !== undefined) parts.push(`padding: ${style.padding}`);
  if (style.paddingTop !== undefined) parts.push(`paddingTop: ${style.paddingTop}`);
  if (style.paddingRight !== undefined) parts.push(`paddingRight: ${style.paddingRight}`);
  if (style.paddingBottom !== undefined) parts.push(`paddingBottom: ${style.paddingBottom}`);
  if (style.paddingLeft !== undefined) parts.push(`paddingLeft: ${style.paddingLeft}`);

  return parts.join(", ");
}

// ─────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────
function compareCellRef(a: string, b: string): number {
  const parse = (ref: string) => {
    const match = ref.match(/^([A-Z]+)(\d+)$/)!;
    return {
      col: columnLetterToIndex(match[1]),
      row: parseInt(match[2], 10),
    };
  };

  const pa = parse(a);
  const pb = parse(b);

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