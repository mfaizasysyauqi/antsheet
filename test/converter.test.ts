import { describe, it, expect } from "vitest";
import { convertToUniver, convertFromUniver } from "../src/converter";
import { parseAntsheet } from "../src/parser";
import { serializeAntsheet } from "../src/serializer";
import type { SheetDocument } from "../src/types";

describe("AntSheet Converter", () => {
  const doc: SheetDocument = {
    name: "Financials",
    meta: {
      columnWidth: { A: 150, B: 120 },
      rowHeight: { "1": 35, "2": 28 },
    },
    cells: {
      A1: { value: "Revenue", style: { fontWeight: "bold", textAlign: "center", color: "#ffffff" } },
      B1: { value: 1000000, style: { format: "currency:USD", textAlign: "right" } },
      A2: { value: "Expenses" },
      B2: { value: 400000 },
      A3: { value: "Net Profit", style: { fontWeight: "bold" } },
      B3: { value: "=B1-B2", style: { fontWeight: "bold", format: "currency:USD" } },
    },
    merges: ["A4:B4"],
  };

  it("should convert AntSheet to Univer sheet data including formulas", () => {
    const univer = convertToUniver(doc);

    expect(univer.name).toBe("Financials");
    expect(univer.columnData[0]).toEqual({ w: 150 });
    expect(univer.columnData[1]).toEqual({ w: 120 });
    expect(univer.rowData[0]).toEqual({ h: 35 });
    expect(univer.rowData[1]).toEqual({ h: 28 });

    // Cell A1 (row 0, col 0)
    expect(univer.cellData[0][0].v).toBe("Revenue");
    expect(univer.cellData[0][0].s?.bl).toBe(1);
    expect(univer.cellData[0][0].s?.ht).toBe(2);
    expect(univer.cellData[0][0].s?.cl?.rgb).toBe("#ffffff");

    // Cell B3 formula (row 2, col 1)
    expect(univer.cellData[2][1].f).toBe("=B1-B2");
    expect(univer.cellData[2][1].v).toBeNull();
  });

  it("should convert Univer sheet data back to AntSheet document", () => {
    const univer = convertToUniver(doc);
    const convertedDoc = convertFromUniver(univer);

    expect(convertedDoc.name).toBe("Financials");
    expect(convertedDoc.meta.columnWidth?.A).toBe(150);
    expect(convertedDoc.meta.columnWidth?.B).toBe(120);
    expect(convertedDoc.meta.rowHeight?.["1"]).toBe(35);

    expect(convertedDoc.cells.A1.value).toBe("Revenue");
    expect(convertedDoc.cells.A1.style?.fontWeight).toBe("bold");
    expect(convertedDoc.cells.A1.style?.textAlign).toBe("center");
    expect(convertedDoc.cells.A1.style?.color).toBe("#ffffff");

    expect(convertedDoc.cells.B3.value).toBe("=B1-B2");
    expect(convertedDoc.cells.B3.style?.fontWeight).toBe("bold");

    expect(convertedDoc.merges).toEqual(["A4:B4"]);
  });
});
