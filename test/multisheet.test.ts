import { describe, it, expect } from "vitest";
import { parseWorkbook, parseAntsheet } from "../src/parser";
import { serializeWorkbook, serializeAntsheet } from "../src/serializer";

const multiSheetText = `# Sales

## Cells
A1 = "Product"
B1 = "Revenue"
A2 = "Monitor"
B2 = 16800000

---

# Summary

## Cells
A1 = "Total Revenue"
B1 = "=Sales!B2"
`;

describe("Multi-Sheet (Workbook)", () => {
  it("should parse multiple sheets separated by ---", () => {
    const wb = parseWorkbook(multiSheetText);
    expect(wb.sheets).toHaveLength(2);
    expect(wb.sheets[0].name).toBe("Sales");
    expect(wb.sheets[1].name).toBe("Summary");
  });

  it("sheet 1 should have correct cells", () => {
    const wb = parseWorkbook(multiSheetText);
    expect(wb.sheets[0].cells.A1.value).toBe("Product");
    expect(wb.sheets[0].cells.B2.value).toBe(16800000);
  });

  it("sheet 2 should have cross-sheet formula stored as string", () => {
    const wb = parseWorkbook(multiSheetText);
    expect(wb.sheets[1].cells.B1.value).toBe("=Sales!B2");
  });

  it("should serialize workbook back to multi-sheet text with ---", () => {
    const wb = parseWorkbook(multiSheetText);
    const serialized = serializeWorkbook(wb);
    expect(serialized).toContain("---");
    expect(serialized).toContain("# Sales");
    expect(serialized).toContain("# Summary");
  });

  it("should parse multiple sheets defined with # SheetName without ---", () => {
    const textWithoutDashes = `# Sheet1
## Cells
A1 = "Data 1"
B1 = 100

# Sheet2
## Cells
A1 = "Data 2"
B1 = 200
`;
    const wb = parseWorkbook(textWithoutDashes);
    expect(wb.sheets).toHaveLength(2);
    expect(wb.sheets[0].name).toBe("Sheet1");
    expect(wb.sheets[0].cells.A1.value).toBe("Data 1");
    expect(wb.sheets[1].name).toBe("Sheet2");
    expect(wb.sheets[1].cells.A1.value).toBe("Data 2");
  });

  it("parseAntsheet should still work for single-sheet (backwards compat)", () => {
    const doc = parseAntsheet(`# Inventory

## Cells
A1 = "Item"
B1 = 100
`);
    expect(doc.name).toBe("Inventory");
    expect(doc.cells.A1.value).toBe("Item");
  });
});

describe("Named Ranges & Hidden Rows/Cols", () => {
  const text = `# Report

## Meta
namedRange: Revenue_Total = C2:C20
namedRange: Tax_Rate = E1
hiddenRows: 5, 6
hiddenColumns: D, E

## Cells
A1 = "Title"
`;

  it("should parse named ranges", () => {
    const doc = parseAntsheet(text);
    expect(doc.meta.namedRanges?.Revenue_Total).toBe("C2:C20");
    expect(doc.meta.namedRanges?.Tax_Rate).toBe("E1");
  });

  it("should parse hidden rows and columns", () => {
    const doc = parseAntsheet(text);
    expect(doc.meta.hiddenRows).toEqual([5, 6]);
    expect(doc.meta.hiddenColumns).toEqual(["D", "E"]);
  });

  it("should serialize named ranges and hidden rows/cols back", () => {
    const doc = parseAntsheet(text);
    const out = serializeAntsheet(doc);
    expect(out).toContain("namedRange: Revenue_Total = C2:C20");
    expect(out).toContain("hiddenRows: 5, 6");
    expect(out).toContain("hiddenColumns: D, E");
  });
});

describe("Cell Comments", () => {
  const text = `# Notes

## Cells
A1 = "Revenue" | fontWeight: bold, comment: "Includes VAT for EU regions"
B1 = 42 | comment: "Auto-calculated"
`;

  it("should parse cell comments", () => {
    const doc = parseAntsheet(text);
    expect(doc.cells.A1.comment).toBe("Includes VAT for EU regions");
    expect(doc.cells.B1.comment).toBe("Auto-calculated");
  });

  it("should serialize cell comments back", () => {
    const doc = parseAntsheet(text);
    const out = serializeAntsheet(doc);
    expect(out).toContain(`comment: "Includes VAT for EU regions"`);
  });
});

describe("Rich Text Cells", () => {
  const text = `# Rich Text Test

## Cells
A1 = richtext | segments: [{text: "Bold", bold: true}, {text: " and normal"}]
B1 = richtext | segments: [{text: "Red text", color: #ef4444}, {text: " blue text", color: #3b82f6}]
`;

  it("should parse rich text segments", () => {
    const doc = parseAntsheet(text);
    expect(doc.cells.A1.richText).toHaveLength(2);
    expect(doc.cells.A1.richText![0].text).toBe("Bold");
    expect(doc.cells.A1.richText![0].bold).toBe(true);
    expect(doc.cells.A1.richText![1].text).toBe(" and normal");
  });

  it("should parse rich text colors", () => {
    const doc = parseAntsheet(text);
    expect(doc.cells.B1.richText![0].color).toBe("#ef4444");
    expect(doc.cells.B1.richText![1].color).toBe("#3b82f6");
  });

  it("should serialize rich text back", () => {
    const doc = parseAntsheet(text);
    const out = serializeAntsheet(doc);
    expect(out).toContain("richtext");
    expect(out).toContain('text: "Bold"');
    expect(out).toContain("bold: true");
  });
});
