import { describe, it, expect } from "vitest";
import { parseCSV, exportToCSV } from "../src/csv";
import { parseAntsheet } from "../src/parser";

describe("CSV Import", () => {
  it("should parse basic CSV", () => {
    const csv = `Product,Qty,Price\nMonitor,5,2500000\nKeyboard,10,450000`;
    const doc = parseCSV(csv);
    expect(doc.cells.A1.value).toBe("Product");
    expect(doc.cells.B2.value).toBe(5);
    expect(doc.cells.C3.value).toBe(450000);
  });

  it("should parse quoted fields with commas", () => {
    const csv = `"Smith, John",30,"New York, NY"`;
    const doc = parseCSV(csv);
    expect(doc.cells.A1.value).toBe("Smith, John");
    expect(doc.cells.B1.value).toBe(30);
    expect(doc.cells.C1.value).toBe("New York, NY");
  });

  it("should parse boolean and null values", () => {
    const csv = `true,false,`;
    const doc = parseCSV(csv);
    expect(doc.cells.A1.value).toBe(true);
    expect(doc.cells.B1.value).toBe(false);
    expect(doc.cells.C1).toBeUndefined();
  });

  it("should support custom delimiter", () => {
    const tsv = `Name\tAge\tCity\nAlice\t25\tBandung`;
    const doc = parseCSV(tsv, { delimiter: "\t", sheetName: "TSV Data" });
    expect(doc.name).toBe("TSV Data");
    expect(doc.cells.A2.value).toBe("Alice");
    expect(doc.cells.B2.value).toBe(25);
  });
});

describe("CSV Export", () => {
  it("should export cells to CSV", () => {
    const doc = parseAntsheet(`# Sales

## Cells
A1 = "Product"
B1 = "Revenue"
A2 = "Monitor"
B2 = 16800000
`);
    const csv = exportToCSV(doc);
    expect(csv).toContain("Product");
    expect(csv).toContain("16800000");
  });

  it("should escape values containing commas", () => {
    const doc = parseAntsheet(`# Test

## Cells
A1 = "Hello, World"
B1 = 42
`);
    const csv = exportToCSV(doc);
    expect(csv).toContain('"Hello, World"');
    expect(csv).toContain("42");
  });

  it("should round-trip CSV export -> import", () => {
    const original = parseAntsheet(`# Data

## Cells
A1 = "Name"
B1 = "Score"
A2 = "Alice"
B2 = 95
A3 = "Bob"
B3 = 88
`);
    const csv = exportToCSV(original);
    const reimported = parseCSV(csv);
    expect(reimported.cells.A1.value).toBe("Name");
    expect(reimported.cells.B2.value).toBe(95);
    expect(reimported.cells.A3.value).toBe("Bob");
  });
});
