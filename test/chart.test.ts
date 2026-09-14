import { describe, it, expect } from "vitest";
import { parseAntsheet } from "../src/parser";
import { serializeAntsheet } from "../src/serializer";
import { extractChartData } from "../src/chart";
import type { SheetDocument } from "../src/types";

describe("AntSheet Charts", () => {
  const sampleDocText = `# Sales Report Q1

## Meta
columnWidth: A=140, B=100, C=120

## Cells
A1 = "Month" | fontWeight: bold, backgroundColor: #1e293b, color: #ffffff
B1 = "Target" | fontWeight: bold, backgroundColor: #1e293b, color: #ffffff
C1 = "Actual" | fontWeight: bold, backgroundColor: #1e293b, color: #ffffff
A2 = "Jan"
B2 = 50000000
C2 = 55000000
A3 = "Feb"
B3 = 60000000
C3 = 58000000
A4 = "Mar"
B4 = 70000000
C4 = 75000000

## Merges
A1:C1

## Charts
chart1 = bar | title: "Target vs Actual", range: A1:C4, x: A, series: [B, C], position: E2:L14, stacked: false
pie1 = pie | title: "Actual Distribution", labels: A2:A4, values: C2:C4, position: E16:I26
`;

  it("should parse ## Charts section correctly", () => {
    const doc = parseAntsheet(sampleDocText);

    expect(doc.charts).toBeDefined();
    expect(doc.charts?.length).toBe(2);

    const barChart = doc.charts?.[0];
    expect(barChart?.id).toBe("chart1");
    expect(barChart?.type).toBe("bar");
    expect(barChart?.title).toBe("Target vs Actual");
    expect(barChart?.range).toBe("A1:C4");
    expect(barChart?.xAxis).toBe("A");
    expect(barChart?.series).toEqual(["B", "C"]);
    expect(barChart?.position).toBe("E2:L14");
    expect(barChart?.stacked).toBe(false);

    const pieChart = doc.charts?.[1];
    expect(pieChart?.id).toBe("pie1");
    expect(pieChart?.type).toBe("pie");
    expect(pieChart?.title).toBe("Actual Distribution");
    expect(pieChart?.labels).toBe("A2:A4");
    expect(pieChart?.values).toBe("C2:C4");
    expect(pieChart?.position).toBe("E16:I26");
  });

  it("should extract chart data for bar chart with xAxis and series", () => {
    const doc = parseAntsheet(sampleDocText);
    const barChart = doc.charts![0];

    const resolved = extractChartData(doc, barChart);
    expect(resolved.id).toBe("chart1");
    expect(resolved.type).toBe("bar");
    expect(resolved.title).toBe("Target vs Actual");
    expect(resolved.labels).toEqual(["Jan", "Feb", "Mar"]);
    expect(resolved.datasets.length).toBe(2);

    // Series 1: Target
    expect(resolved.datasets[0].label).toBe("Target");
    expect(resolved.datasets[0].data).toEqual([50000000, 60000000, 70000000]);

    // Series 2: Actual
    expect(resolved.datasets[1].label).toBe("Actual");
    expect(resolved.datasets[1].data).toEqual([55000000, 58000000, 75000000]);
  });

  it("should extract chart data for pie chart with labels and values", () => {
    const doc = parseAntsheet(sampleDocText);
    const pieChart = doc.charts![1];

    const resolved = extractChartData(doc, pieChart);
    expect(resolved.id).toBe("pie1");
    expect(resolved.type).toBe("pie");
    expect(resolved.title).toBe("Actual Distribution");
    expect(resolved.labels).toEqual(["Jan", "Feb", "Mar"]);
    expect(resolved.datasets.length).toBe(1);
    expect(resolved.datasets[0].data).toEqual([55000000, 58000000, 75000000]);
  });

  it("should serialize ## Charts back to antsheet text", () => {
    const doc = parseAntsheet(sampleDocText);
    const serialized = serializeAntsheet(doc);

    expect(serialized).toContain("## Charts");
    expect(serialized).toContain('chart1 = bar | title: "Target vs Actual", range: A1:C4, x: A, series: [B, C], position: E2:L14');
    expect(serialized).toContain('pie1 = pie | title: "Actual Distribution", labels: A2:A4, values: C2:C4, position: E16:I26');
  });

  it("should handle roundtrip parsing and serializing cleanly", () => {
    const doc1 = parseAntsheet(sampleDocText);
    const serialized = serializeAntsheet(doc1);
    const doc2 = parseAntsheet(serialized);

    expect(doc2.charts).toEqual(doc1.charts);
  });
});
