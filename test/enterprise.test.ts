import { describe, it, expect } from "vitest";
import { parseAntsheet } from "../src/parser";
import { serializeAntsheet } from "../src/serializer";
import { validateAntsheet } from "../src/validator";

describe("AntSheet Enterprise Features", () => {
  const sampleText = `# Project Tracker

## Meta
columnWidth: A=180, B=120, C=120, D=120
rowHeight: 1=36
freezeRow: 1
filter: A1:D20

## Cells
A1 = "Task" | fontWeight: bold, backgroundColor: #1e293b, color: #ffffff
B1 = "Status" | fontWeight: bold, backgroundColor: #1e293b, color: #ffffff
C1 = "Score" | fontWeight: bold, backgroundColor: #1e293b, color: #ffffff
D1 = "Hours" | fontWeight: bold, backgroundColor: #1e293b, color: #ffffff

A2 = "Implement Auth"
B2 = "In Progress"
C2 = 85
D2 = 12

A3 = "Design Database"
B3 = "Done"
C3 = 98
D3 = 20

## DataValidation
rule1 = list | range: B2:B20, options: ["Pending", "In Progress", "Done", "Cancelled"]
rule2 = number | range: C2:C20, operator: between, min: 0, max: 100

## ConditionalFormatting
cf1 = highlight | range: C2:C20, operator: ">=", value: 90, backgroundColor: #dcfce7, color: #15803d
cf2 = colorScale | range: D2:D20, minColor: #fee2e2, maxColor: #dcfce7
`;

  it("should parse DataValidation, ConditionalFormatting, and Filter correctly", () => {
    const doc = parseAntsheet(sampleText);

    expect(doc.meta.filter).toBe("A1:D20");
    expect(doc.meta.freezeRow).toBe(1);

    // Data Validation
    expect(doc.dataValidations).toHaveLength(2);
    expect(doc.dataValidations![0].id).toBe("rule1");
    expect(doc.dataValidations![0].type).toBe("list");
    expect(doc.dataValidations![0].range).toBe("B2:B20");
    expect(doc.dataValidations![0].options).toEqual(["Pending", "In Progress", "Done", "Cancelled"]);

    expect(doc.dataValidations![1].id).toBe("rule2");
    expect(doc.dataValidations![1].type).toBe("number");
    expect(doc.dataValidations![1].min).toBe(0);
    expect(doc.dataValidations![1].max).toBe(100);

    // Conditional Formatting
    expect(doc.conditionalFormattings).toHaveLength(2);
    expect(doc.conditionalFormattings![0].id).toBe("cf1");
    expect(doc.conditionalFormattings![0].type).toBe("highlight");
    expect(doc.conditionalFormattings![0].operator).toBe(">=");
    expect(doc.conditionalFormattings![0].value).toBe(90);
    expect(doc.conditionalFormattings![0].backgroundColor).toBe("#dcfce7");

    expect(doc.conditionalFormattings![1].id).toBe("cf2");
    expect(doc.conditionalFormattings![1].type).toBe("colorScale");
    expect(doc.conditionalFormattings![1].minColor).toBe("#fee2e2");
  });

  it("should serialize DataValidation and ConditionalFormatting back to text", () => {
    const doc = parseAntsheet(sampleText);
    const serialized = serializeAntsheet(doc);

    expect(serialized).toContain("filter: A1:D20");
    expect(serialized).toContain("## DataValidation");
    expect(serialized).toContain('rule1 = list | range: B2:B20, options: [Pending, "In Progress", Done, Cancelled]');
    expect(serialized).toContain("## ConditionalFormatting");
    expect(serialized).toContain('cf1 = highlight | range: C2:C20, operator: ">=", value: 90');
  });

  it("should pass validation for enterprise documents", () => {
    const result = validateAntsheet(sampleText);
    expect(result.valid).toBe(true);
    expect(result.errors).toHaveLength(0);
  });
});
