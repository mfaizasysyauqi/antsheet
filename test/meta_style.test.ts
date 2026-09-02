import { describe, it, expect } from "vitest";
import { parseAntsheet } from "../src/parser";
import { serializeAntsheet } from "../src/serializer";
import { convertToUniver, convertFromUniver } from "../src/converter";

describe("AntSheet Freeze Panes & Text Wrapping", () => {
  const sampleText = `# Freeze & Wrap Test

## Meta
columnWidth: A=200, B=150
rowHeight: 1=40
freezeRow: 1
freezeColumn: 1

## Cells
A1 = "Header Long Title" | fontWeight: bold, textWrap: wrap
B1 = "Normal" | textWrap: clip
A2 = "Wrapped paragraph text inside a cell" | textWrap: wrap
B2 = 12345
`;

  it("should parse freeze panes and textWrap correctly", () => {
    const doc = parseAntsheet(sampleText);

    expect(doc.meta.freezeRow).toBe(1);
    expect(doc.meta.freezeColumn).toBe(1);

    expect(doc.cells.A1.style?.textWrap).toBe("wrap");
    expect(doc.cells.B1.style?.textWrap).toBe("clip");
    expect(doc.cells.A2.style?.textWrap).toBe("wrap");
  });

  it("should serialize freeze panes and textWrap back to text", () => {
    const doc = parseAntsheet(sampleText);
    const serialized = serializeAntsheet(doc);

    expect(serialized).toContain("freezeRow: 1");
    expect(serialized).toContain("freezeColumn: 1");
    expect(serialized).toContain("textWrap: wrap");
    expect(serialized).toContain("textWrap: clip");
  });

  it("should convert freeze panes and textWrap to Univer and back", () => {
    const doc = parseAntsheet(sampleText);
    const univer = convertToUniver(doc);

    expect(univer.freeze).toBeDefined();
    expect(univer.freeze?.startRow).toBe(1);
    expect(univer.freeze?.startColumn).toBe(1);

    // Cell A1 (row 0, col 0) has textWrap: wrap -> tb: 2
    expect(univer.cellData[0][0].s?.tb).toBe(2);
    // Cell B1 (row 0, col 1) has textWrap: clip -> tb: 3
    expect(univer.cellData[0][1].s?.tb).toBe(3);

    // Reverse conversion
    const backToDoc = convertFromUniver(univer);
    expect(backToDoc.meta.freezeRow).toBe(1);
    expect(backToDoc.meta.freezeColumn).toBe(1);
    expect(backToDoc.cells.A1.style?.textWrap).toBe("wrap");
    expect(backToDoc.cells.B1.style?.textWrap).toBe("clip");
  });
});
