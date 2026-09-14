import { describe, it, expect } from "vitest";
import { parseXLSX, exportToXLSX } from "../src/xlsx";
import { parseAntsheet } from "../src/parser";

describe("AntSheet XLSX Import/Export", () => {
  it("should export a SheetDocument to XLSX bytes", () => {
    const doc = parseAntsheet(`# Sales Report

## Meta
columnWidth: A=180, B=100, C=120

## Cells
A1 = "Product" | fontWeight: bold
B1 = "Qty" | fontWeight: bold
C1 = "Revenue" | fontWeight: bold

A2 = "Monitor"
B2 = 42
C2 = 16800000

A3 = "Keyboard"
B3 = 28
C3 = 4200000

A4 = "Total" | fontWeight: bold
B4 = "=SUM(B2:B3)"
C4 = "=SUM(C2:C3)"

## Merges
A1:C1
`);

    const bytes = exportToXLSX(doc);
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(bytes.length).toBeGreaterThan(1000); // non-trivial xlsx
  });

  it("should round-trip export -> import", () => {
    const original = parseAntsheet(`# Round Trip Test

## Cells
A1 = "Hello"
B1 = 42
C1 = true
A2 = "=SUM(B1:B1)"
`);

    const bytes = exportToXLSX(original);
    const reimported = parseXLSX(bytes);

    expect(reimported.cells["A1"]?.value).toBe("Hello");
    expect(reimported.cells["B1"]?.value).toBe(42);
    expect(reimported.cells["C1"]?.value).toBe(true);
  });
});
