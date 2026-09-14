import { describe, it, expect } from "vitest";
import { validateAntsheet } from "../src/validator";

describe("AntSheet Validator", () => {
  it("should validate a correct AntSheet document", () => {
    const text = `# Valid Document

## Meta
columnWidth: A=120, B=100
rowHeight: 1=32

## Cells
A1 = "Product" | fontWeight: bold, color: #ffffff
B1 = 100 | textAlign: right
A2 = "Sum"
B2 = "=SUM(B1:B1)" | format: currency:USD

## Merges
A3:B3

## Charts
chart1 = bar | title: "Overview", range: A1:B2, x: A, series: [B]
`;

    const result = validateAntsheet(text);
    expect(result.valid).toBe(true);
    expect(result.errors).toHaveLength(0);
  });

  it("should detect mismatched quotes and invalid cell refs", () => {
    const text = `# Broken Document

## Cells
INVALID_REF = "Hello"
A1 = "Unclosed string
A2 = 123
`;

    const result = validateAntsheet(text);
    expect(result.valid).toBe(false);
    expect(result.errors.some((e) => e.includes("Invalid cell reference"))).toBe(true);
    expect(result.errors.some((e) => e.includes("Mismatched quotes"))).toBe(true);
  });

  it("should detect invalid meta definitions", () => {
    const text = `# Bad Meta

## Meta
columnWidth: A=notanumber
`;

    const result = validateAntsheet(text);
    expect(result.valid).toBe(false);
    expect(result.errors.some((e) => e.includes("Invalid columnWidth definition"))).toBe(true);
  });
});
