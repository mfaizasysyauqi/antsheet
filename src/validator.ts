// validator.ts
import { isValidCellRef, CELL_REF_REGEX } from "./coords";
import { parseDeclarationLine, splitPipePayload } from "./utils";

export interface ValidationResult {
  valid: boolean;
  errors: string[];
  warnings: string[];
}

const VALID_SECTIONS = new Set([
  "Meta",
  "Cells",
  "Merges",
  "Charts",
  "DataValidation",
  "Validation",
  "ConditionalFormatting",
  "ConditionalFormat",
  "Rules",
]);

const VALID_CHART_TYPES = new Set([
  "bar",
  "line",
  "pie",
  "doughnut",
  "area",
  "scatter",
  "radar",
]);

const VALID_VALIDATION_TYPES = new Set([
  "list",
  "number",
  "textLength",
  "date",
  "checkbox",
  "custom",
]);

const VALID_CONDITIONAL_TYPES = new Set([
  "highlight",
  "colorScale",
  "dataBar",
]);

const VALID_STYLE_KEYS = new Set([
  "fontFamily",
  "fontSize",
  "fontWeight",
  "fontStyle",
  "textDecoration",
  "textTransform",
  "letterSpacing",
  "lineHeight",
  "color",
  "backgroundColor",
  "textAlign",
  "verticalAlign",
  "format",
  "border",
  "borderTop",
  "borderRight",
  "borderBottom",
  "borderLeft",
  "borderColor",
  "borderWidth",
  "borderStyle",
  "textWrap",
  "wrapText",
  "comment",
  "segments",
  "padding",
  "paddingTop",
  "paddingRight",
  "paddingBottom",
  "paddingLeft",
]);

/**
 * Validates AntSheet plain-text syntax and checks for structural issues
 */
export function validateAntsheet(text: string): ValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];

  const lines = text.split(/\r?\n/);
  let currentSection: string | null = null;
  let hasSheetName = false;
  let hasCellsSection = false;

  for (let i = 0; i < lines.length; i++) {
    const rawLine = lines[i];
    const line = rawLine.trim();
    const lineNum = i + 1;

    if (!line || line.startsWith("//")) continue;

    // Sheet Name
    if (line.startsWith("# ")) {
      hasSheetName = true;
      continue;
    }

    // Section Header
    if (line.startsWith("## ")) {
      const sectionName = line.slice(3).trim();
      if (!VALID_SECTIONS.has(sectionName)) {
        warnings.push(`Line ${lineNum}: Unknown section '## ${sectionName}'`);
      }
      currentSection = sectionName;
      if (sectionName === "Cells") hasCellsSection = true;
      continue;
    }

    // Inside Meta
    if (currentSection === "Meta") {
      if (line.startsWith("columnWidth:")) {
        validateKeyValueLine(line.slice("columnWidth:".length), lineNum, "columnWidth", errors);
      } else if (line.startsWith("rowHeight:")) {
        validateKeyValueLine(line.slice("rowHeight:".length), lineNum, "rowHeight", errors);
      } else if (line.startsWith("freezeRow:")) {
        const val = line.slice("freezeRow:".length).trim();
        if (Number.isNaN(Number(val))) {
          errors.push(`Line ${lineNum}: Invalid freezeRow '${val}'. Expected number.`);
        }
      } else if (line.startsWith("namedRange:")) {
        const raw = line.slice("namedRange:".length).trim();
        if (!raw.includes("=")) {
          errors.push(`Line ${lineNum}: Named range must use '=' syntax: 'namedRange: MyName = A1:B10'`);
        }
      } else if (
        !line.startsWith("freezeColumn:") &&
        !line.startsWith("freeze:") &&
        !line.startsWith("filter:") &&
        !line.startsWith("autoFilter:") &&
        !line.startsWith("hiddenRows:") &&
        !line.startsWith("hiddenColumns:")
      ) {
        warnings.push(`Line ${lineNum}: Unrecognized Meta directive: '${line}'`);
      }
    }
    // Inside Cells
    else if (currentSection === "Cells") {
      validateCellLine(line, lineNum, errors, warnings);
    }
    // Inside Merges
    else if (currentSection === "Merges") {
      validateMergeLine(line, lineNum, errors);
    }
    // Inside Charts
    else if (currentSection === "Charts") {
      validateChartLine(line, lineNum, errors, warnings);
    }
    // Inside DataValidation
    else if (currentSection === "DataValidation" || currentSection === "Validation") {
      validateDataValidationLine(line, lineNum, errors, warnings);
    }
    // Inside ConditionalFormatting
    else if (
      currentSection === "ConditionalFormatting" ||
      currentSection === "ConditionalFormat" ||
      currentSection === "Rules"
    ) {
      validateConditionalFormattingLine(line, lineNum, errors, warnings);
    }
    // Outside any section
    else {
      warnings.push(`Line ${lineNum}: Content outside of any recognized section: '${line}'`);
    }
  }

  if (!hasSheetName) {
    warnings.push("Document is missing a sheet name header ('# Sheet Name')");
  }
  if (!hasCellsSection) {
    warnings.push("Document is missing a '## Cells' section");
  }

  return {
    valid: errors.length === 0,
    errors,
    warnings,
  };
}

function validateKeyValueLine(
  raw: string,
  lineNum: number,
  type: "columnWidth" | "rowHeight",
  errors: string[]
) {
  const parts = raw.split(",").map((p) => p.trim()).filter(Boolean);
  for (const part of parts) {
    const [k, v] = part.split("=").map((s) => s.trim());
    if (!k || !v || Number.isNaN(Number(v))) {
      errors.push(`Line ${lineNum}: Invalid ${type} definition '${part}'. Expected format 'A=100' or '1=30'.`);
    }
  }
}

function validateCellLine(
  line: string,
  lineNum: number,
  errors: string[],
  warnings: string[]
) {
  const decl = parseDeclarationLine(line);
  if (!decl) {
    errors.push(`Line ${lineNum}: Cell definition must contain '=' (e.g. A1 = "Value")`);
    return;
  }

  const ref = decl.key.toUpperCase();
  if (!CELL_REF_REGEX.test(ref)) {
    errors.push(`Line ${lineNum}: Invalid cell reference '${ref}'. Expected format like A1, B2, AA10.`);
    return;
  }

  const { main: valueRaw, props: styleRaw } = splitPipePayload(decl.payload);

  // Validate value quotes
  if (
    (valueRaw.startsWith('"') && !valueRaw.endsWith('"')) ||
    (!valueRaw.startsWith('"') && valueRaw.endsWith('"')) ||
    (valueRaw.startsWith("'") && !valueRaw.endsWith("'")) ||
    (!valueRaw.startsWith("'") && valueRaw.endsWith("'"))
  ) {
    errors.push(`Line ${lineNum}: Mismatched quotes in cell value: ${valueRaw}`);
  }

  // Validate style properties
  if (styleRaw) {
    const styleParts = styleRaw.split(",").map((p) => p.trim()).filter(Boolean);
    for (const part of styleParts) {
      const colIndex = part.indexOf(":");
      if (colIndex === -1) {
        warnings.push(`Line ${lineNum}: Malformed style property '${part}'`);
        continue;
      }
      const propKey = part.slice(0, colIndex).trim();
      if (!VALID_STYLE_KEYS.has(propKey)) {
        warnings.push(`Line ${lineNum}: Unknown style property '${propKey}'`);
      }
    }
  }
}

function validateMergeLine(line: string, lineNum: number, errors: string[]) {
  const parts = line.split(":");
  if (parts.length > 2) {
    errors.push(`Line ${lineNum}: Invalid merge range '${line}'. Expected format 'A1:C1'.`);
    return;
  }

  for (const part of parts) {
    if (!isValidCellRef(part.trim())) {
      errors.push(`Line ${lineNum}: Invalid cell reference in merge range '${part}'.`);
    }
  }
}

function validateChartLine(
  line: string,
  lineNum: number,
  errors: string[],
  warnings: string[]
) {
  const decl = parseDeclarationLine(line);
  if (!decl) {
    errors.push(`Line ${lineNum}: Chart definition must contain '=' (e.g. chart1 = bar | title: "...")`);
    return;
  }

  const { main: typeRaw } = splitPipePayload(decl.payload);
  const type = typeRaw.toLowerCase();

  if (!VALID_CHART_TYPES.has(type)) {
    warnings.push(
      `Line ${lineNum}: Unsupported chart type '${type}'. Supported types: ${Array.from(VALID_CHART_TYPES).join(", ")}`
    );
  }
}

function validateDataValidationLine(
  line: string,
  lineNum: number,
  errors: string[],
  warnings: string[]
) {
  const decl = parseDeclarationLine(line);
  if (!decl) {
    errors.push(`Line ${lineNum}: DataValidation definition must contain '=' (e.g. rule1 = list | range: D2:D50)`);
    return;
  }

  const { main: typeRaw, props } = splitPipePayload(decl.payload);
  const type = typeRaw.toLowerCase();

  if (!VALID_VALIDATION_TYPES.has(type)) {
    warnings.push(
      `Line ${lineNum}: Unsupported validation type '${type}'. Supported: ${Array.from(VALID_VALIDATION_TYPES).join(", ")}`
    );
  }

  if (!props || !props.includes("range:")) {
    errors.push(`Line ${lineNum}: DataValidation rule must define a 'range:'`);
  }
}

function validateConditionalFormattingLine(
  line: string,
  lineNum: number,
  errors: string[],
  warnings: string[]
) {
  const decl = parseDeclarationLine(line);
  if (!decl) {
    errors.push(`Line ${lineNum}: ConditionalFormatting definition must contain '=' (e.g. cf1 = highlight | range: C2:C20)`);
    return;
  }

  const { main: typeRaw, props } = splitPipePayload(decl.payload);
  const type = typeRaw.toLowerCase();

  if (!VALID_CONDITIONAL_TYPES.has(type)) {
    warnings.push(
      `Line ${lineNum}: Unsupported conditional format type '${type}'. Supported: ${Array.from(VALID_CONDITIONAL_TYPES).join(", ")}`
    );
  }

  if (!props || !props.includes("range:")) {
    errors.push(`Line ${lineNum}: ConditionalFormatting rule must define a 'range:'`);
  }
}
