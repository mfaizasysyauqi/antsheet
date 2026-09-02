import { describe, it, expect } from "vitest";
import { parseAntsheet, parseWorkbook } from "../src/parser";
import { evaluateFormulas, evaluateWorkbook } from "../src/formula";

describe("Formula Evaluation with HyperFormula", () => {
  it("should evaluate basic arithmetic and SUM formula", () => {
    const text = `# Sales Report

## Cells
A1 = "Product"
B1 = "Qty"
C1 = "Price"
D1 = "Total"

A2 = "Laptop"
B2 = 2
C2 = 10000000
D2 = "=B2*C2"

A3 = "Mouse"
B3 = 5
C3 = 200000
D3 = "=B3*C3"

A4 = "Grand Total"
D4 = "=SUM(D2:D3)"
`;

    const doc = parseAntsheet(text);
    const evaluated = evaluateFormulas(doc);

    expect(evaluated.cells.D2.value).toBe(20000000);
    expect(evaluated.cells.D3.value).toBe(1000000);
    expect(evaluated.cells.D4.value).toBe(21000000);
  });

  it("should evaluate logical formulas (IF, AND, OR)", () => {
    const text = `# Logic Test

## Cells
A1 = 85
B1 = "=IF(A1>=80, \"Pass\", \"Fail\")"
A2 = 65
B2 = "=IF(A2>=80, \"Pass\", \"Fail\")"
`;

    const doc = parseAntsheet(text);
    const evaluated = evaluateFormulas(doc);

    expect(evaluated.cells.B1.value).toBe("Pass");
    expect(evaluated.cells.B2.value).toBe("Fail");
  });

  it("should evaluate cross-sheet formulas in a workbook", () => {
    const text = `# Details

## Cells
A1 = "Sales"
B1 = 5000000

---

# Summary

## Cells
A1 = "Total from Details"
B1 = "=Details!B1*1.1"
`;

    const wb = parseWorkbook(text);
    const evaluated = evaluateWorkbook(wb);

    expect(evaluated.sheets[1].cells.B1.value).toBe(5500000);
  });
});
