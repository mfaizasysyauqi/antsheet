// utils.ts - Common string, property, and formatting helpers

/**
 * Removes wrapping single or double quotes from a string if present.
 */
export function cleanQuotedString(val: string): string {
  const trimmed = val.trim();
  if (
    (trimmed.startsWith('"') && trimmed.endsWith('"')) ||
    (trimmed.startsWith("'") && trimmed.endsWith("'"))
  ) {
    return trimmed.slice(1, -1);
  }
  return trimmed;
}

/**
 * Parses a comma-delimited list or JSON-like array `[a, b, c]` into clean string items.
 */
export function parseArrayOrList(val: string): string[] {
  let cleaned = val.trim();
  if (cleaned.startsWith("[") && cleaned.endsWith("]")) {
    cleaned = cleaned.slice(1, -1);
  }
  return cleaned
    .split(",")
    .map((s) => cleanQuotedString(s.trim()))
    .filter(Boolean);
}

/**
 * Converts a string to number if numeric, otherwise strips surrounding quotes.
 */
export function parseNumericOrString(val: string): number | string {
  const trimmed = val.trim();
  if (trimmed !== "" && !Number.isNaN(Number(trimmed))) {
    return Number(trimmed);
  }
  return cleanQuotedString(trimmed);
}

/**
 * Parses "true" or "false" string to a boolean.
 */
export function parseBoolean(val: string): boolean {
  return val.trim().toLowerCase() === "true";
}

/**
 * Splits a declaration line like `A1 = 42 | style` into key and payload.
 */
export function parseDeclarationLine(
  line: string,
  separator = "="
): { key: string; payload: string } | null {
  const eqIdx = line.indexOf(separator);
  if (eqIdx === -1) return null;
  return {
    key: line.slice(0, eqIdx).trim(),
    payload: line.slice(eqIdx + 1).trim(),
  };
}

/**
 * Splits a payload string into primary value and trailing pipe properties.
 * E.g. `"Hello" | color: red` -> `{ main: '"Hello"', props: 'color: red' }`
 */
export function splitPipePayload(payload: string): { main: string; props: string | null } {
  const pipeIdx = payload.indexOf("|");
  if (pipeIdx === -1) {
    return { main: payload.trim(), props: null };
  }
  return {
    main: payload.slice(0, pipeIdx).trim(),
    props: payload.slice(pipeIdx + 1).trim(),
  };
}

/**
 * Splits comma-separated properties while respecting quotes ("...") and brackets ([...]).
 */
export function splitPropertyParts(raw: string): string[] {
  const parts: string[] = [];
  let current = "";
  let inQuotes = false;
  let quoteChar = "";
  let bracketDepth = 0;

  for (let i = 0; i < raw.length; i++) {
    const char = raw[i];

    if (!inQuotes && (char === '"' || char === "'")) {
      inQuotes = true;
      quoteChar = char;
    } else if (inQuotes && char === quoteChar) {
      inQuotes = false;
    } else if (!inQuotes && char === "[") {
      bracketDepth++;
    } else if (!inQuotes && char === "]") {
      bracketDepth = Math.max(0, bracketDepth - 1);
    }

    if (char === "," && !inQuotes && bracketDepth === 0) {
      if (current.trim()) parts.push(current.trim());
      current = "";
    } else {
      current += char;
    }
  }

  if (current.trim()) {
    parts.push(current.trim());
  }

  return parts;
}

/**
 * Parses a comma-delimited `key: value` property string into a typed record.
 */
export function parseProperties(raw: string): Record<string, string> {
  const result: Record<string, string> = {};
  const parts = splitPropertyParts(raw);

  for (const part of parts) {
    const colonIdx = part.indexOf(":");
    if (colonIdx === -1) continue;
    const key = part.slice(0, colonIdx).trim();
    const val = part.slice(colonIdx + 1).trim();
    if (key) {
      result[key] = val;
    }
  }

  return result;
}

/**
 * Formats a key-value record into an AntSheet property string (e.g. `title: "Sales", stacked: true`).
 */
export function formatProperties(props: Record<string, unknown>): string {
  const entries: string[] = [];

  for (const [key, value] of Object.entries(props)) {
    if (value === undefined || value === null) continue;

    if (Array.isArray(value)) {
      if (value.length > 0) {
        entries.push(`${key}: [${value.join(", ")}]`);
      }
    } else if (typeof value === "string") {
      entries.push(`${key}: ${value}`);
    } else {
      entries.push(`${key}: ${String(value)}`);
    }
  }

  return entries.join(", ");
}
