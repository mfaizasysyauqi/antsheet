// types.ts

// ─────────────────────────────────────────────
// Rich Text
// ─────────────────────────────────────────────
export interface RichTextSegment {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  color?: string;
  backgroundColor?: string;
  fontSize?: number;
  fontFamily?: string;
}

// ─────────────────────────────────────────────
// Workbook (multi-sheet)
// ─────────────────────────────────────────────
export interface WorkbookDocument {
  sheets: SheetDocument[];
}

export interface SheetDocument {
  name: string;
  meta: SheetMeta;
  cells: Record<string, CellData>;
  merges: string[];
  charts?: ChartConfig[];
  dataValidations?: DataValidationRule[];
  conditionalFormattings?: ConditionalFormattingRule[];
}

// ─────────────────────────────────────────────
// Chart
// ─────────────────────────────────────────────
export type ChartType =
  | "bar"
  | "line"
  | "pie"
  | "doughnut"
  | "area"
  | "scatter"
  | "radar";

export interface ChartConfig {
  id: string;
  type: ChartType;
  title?: string;
  range?: string;
  xAxis?: string;
  series?: string[];
  labels?: string;
  values?: string;
  position?: string;
  colors?: string[];
  stacked?: boolean;
  horizontal?: boolean;
  smooth?: boolean;
}

// ─────────────────────────────────────────────
// Data Validation
// ─────────────────────────────────────────────
export type DataValidationType =
  | "list"
  | "number"
  | "textLength"
  | "date"
  | "checkbox"
  | "custom";

export type DataValidationOperator =
  | "between"
  | "notBetween"
  | "equal"
  | "notEqual"
  | "greaterThan"
  | "lessThan"
  | "greaterThanOrEqual"
  | "lessThanOrEqual";

export interface DataValidationRule {
  id: string;
  type: DataValidationType;
  range: string;
  options?: string[];
  operator?: DataValidationOperator;
  min?: number | string;
  max?: number | string;
  value?: number | string;
  allowBlank?: boolean;
  error?: string;
  prompt?: string;
}

// ─────────────────────────────────────────────
// Conditional Formatting
// ─────────────────────────────────────────────
export type ConditionalFormatType =
  | "highlight"
  | "colorScale"
  | "dataBar";

export type ConditionalFormatOperator =
  | ">"
  | "<"
  | ">="
  | "<="
  | "="
  | "!="
  | "between"
  | "notBetween"
  | "containsText"
  | "notContainsText"
  | "empty"
  | "notEmpty";

export interface ConditionalFormattingRule {
  id: string;
  type: ConditionalFormatType;
  range: string;
  operator?: ConditionalFormatOperator;
  value?: string | number;
  value2?: string | number;
  color?: string;
  backgroundColor?: string;
  minColor?: string;
  maxColor?: string;
  midColor?: string;
}

// ─────────────────────────────────────────────
// Sheet Meta
// ─────────────────────────────────────────────
export interface SheetMeta {
  columnWidth?: Record<string, number>;
  rowHeight?: Record<string, number>;
  freezeRow?: number;
  freezeColumn?: number;
  filter?: string;
  namedRanges?: Record<string, string>;  // e.g. { Sales_Total: "C2:C20" }
  hiddenRows?: number[];
  hiddenColumns?: string[];
}

// ─────────────────────────────────────────────
// Cell Data
// ─────────────────────────────────────────────
export interface CellData {
  value: string | number | boolean | null;
  style?: CellStyle;
  comment?: string;
  richText?: RichTextSegment[];
}

// ─────────────────────────────────────────────
// Cell Style
// ─────────────────────────────────────────────
export interface CellStyle {
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: string | number;
  fontStyle?: "normal" | "italic" | "oblique";
  textDecoration?: string;
  textTransform?: string;
  letterSpacing?: number;
  lineHeight?: number | string;

  textWrap?: "overflow" | "wrap" | "clip";
  wrapText?: boolean;

  color?: string;
  backgroundColor?: string;
  textAlign?: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";

  format?: string;

  border?: string;
  borderTop?: string | number;
  borderRight?: string | number;
  borderBottom?: string | number;
  borderLeft?: string | number;
  borderColor?: string;
  borderWidth?: number;
  borderStyle?: string;

  padding?: string | number;
  paddingTop?: number;
  paddingRight?: number;
  paddingBottom?: number;
  paddingLeft?: number;
}