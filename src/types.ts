// types.ts
export interface SheetDocument {
  name: string;
  meta: SheetMeta;
  cells: Record<string, CellData>;
  merges: string[];
}

export interface SheetMeta {
  columnWidth?: Record<string, number>;
  rowHeight?: Record<string, number>;
}

export interface CellData {
  value: string | number | boolean | null;
  style?: CellStyle;
}

export interface CellStyle {
  // Font
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: string | number;
  fontStyle?: "normal" | "italic" | "oblique";
  textDecoration?: string;
  textTransform?: string;
  letterSpacing?: number;
  lineHeight?: number | string;

  // Color & Align
  color?: string;
  backgroundColor?: string;
  textAlign?: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";

  // Format
  format?: string;

  // Border
  border?: string;
  borderTop?: string | number;
  borderRight?: string | number;
  borderBottom?: string | number;
  borderLeft?: string | number;
  borderColor?: string;
  borderWidth?: number;
  borderStyle?: string;

  // Padding
  padding?: string | number;
  paddingTop?: number;
  paddingRight?: number;
  paddingBottom?: number;
  paddingLeft?: number;
}