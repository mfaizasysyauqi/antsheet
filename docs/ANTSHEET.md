# AntSheet Specification

AntSheet (`.antsheet`) is a lightweight, human-readable, and agent-friendly plain-text spreadsheet format.  
It is designed to be effortlessly parsed, generated, and edited by AI models while supporting multi-sheet workbooks, styling, formulas, merges, charts, data validation, conditional formatting, rich text, named ranges, freeze panes, AutoFilter, CSV and XLSX import/export, and bi-directional conversion with modern spreadsheet engines like Univer and Google Sheets.

**File Extension:** `.antsheet`

---

## Architecture & Flow

```text
  .antsheet Text (single or multi-sheet)
              |
    parseWorkbook() / parseAntsheet()
              |
              v
  WorkbookDocument { sheets: SheetDocument[] }
              |
     +--------+--------+------------+
     |                 |            |
     v                 v            v
convertToUniver()  extractChartData()  exportToXLSX()
     |                 |            |
     v                 v            v
Univer Sheets      Chart Dataset  Excel File
(UI Engine)    (Recharts/Chart.js) (.xlsx)
     |
 convertFromUniver()
     |
 serializeWorkbook() / serializeAntsheet()
     |
  .antsheet Text
```

---

## 1. Syntax Overview

```text
# Q1 Sales Report

## Meta
columnWidth: A=140, B=100, C=140
rowHeight: 1=36

## Cells
A1 = "Product" | fontWeight: bold, fontSize: 14, backgroundColor: #1e293b, color: #ffffff, textAlign: center
B1 = "Units"   | fontWeight: bold, fontSize: 14, backgroundColor: #1e293b, color: #ffffff, textAlign: center
C1 = "Total"   | fontWeight: bold, fontSize: 14, backgroundColor: #1e293b, color: #ffffff, textAlign: center

A2 = "Laptop"
B2 = 15 | textAlign: right, format: number
C2 = 22500000 | textAlign: right, format: currency:IDR, fontWeight: bold

A3 = "Mouse"
B3 = 42 | textAlign: right, format: number
C3 = 2100000 | textAlign: right, format: currency:IDR

A4 = "Keyboard"
B4 = 28 | textAlign: right, format: number
C4 = 4200000 | textAlign: right, format: currency:IDR

A5 = "Total" | fontWeight: bold, backgroundColor: #f1f5f9, borderTop: 2 solid #94a3b8
B5 = "=SUM(B2:B4)" | fontWeight: bold, backgroundColor: #f1f5f9, textAlign: right, borderTop: 2 solid #94a3b8
C5 = "=SUM(C2:C4)" | fontWeight: bold, backgroundColor: #f1f5f9, format: currency:IDR, textAlign: right, borderTop: 2 solid #94a3b8

## Merges
A1:C1

## Charts
chart1 = bar | title: "Units Sold", x: A, series: [B], position: E2:K12
chart2 = pie | title: "Revenue Share", labels: A2:A4, values: C2:C4, position: E14:I24
```

### Core Rules

1. **One cell per line**: Each line under `## Cells` declares exactly one cell.
2. **Quoted strings**: Text strings must be wrapped in quotes (`"..."`).
3. **Numbers & booleans**: Numbers, booleans (`true`/`false`), and null values are written without quotes.
4. **Formulas**: Formulas start with an equal sign inside quotes, e.g. `"=SUM(B2:B4)"` or `"=Sheet2!A1"`.
5. **Styles are optional**: Styles follow a pipe delimiter `|` and use camelCase properties.
6. **Section order**: `# Name`, `## Meta`, `## Cells`, `## Merges`, `## DataValidation`, `## ConditionalFormatting`, `## Charts`.
7. **Multi-sheet**: Separate sheets with a line containing only `---`.

---

## 2. Meta Section

Configures global sheet properties.

```text
## Meta
columnWidth: A=140, B=100, C=120
rowHeight: 1=36, 5=40
freezeRow: 1
freezeColumn: 1
filter: A1:G50
namedRange: Revenue_Total = C2:C20
namedRange: Tax_Rate = E1
hiddenRows: 5, 6
hiddenColumns: D, E
```

| Property        | Description                                         |
|-----------------|-----------------------------------------------------|
| `columnWidth`   | Column widths (e.g. `A=140, B=100`) (px)            |
| `rowHeight`     | Row heights (e.g. `1=36, 5=40`) (px)               |
| `freezeRow`     | Number of frozen top rows (e.g. `1`)               |
| `freezeColumn`  | Number of frozen left columns (e.g. `1`, `A`)       |
| `filter`        | AutoFilter range (e.g. `A1:G50`)                   |
| `namedRange`    | Named range alias (e.g. `Revenue_Total = C2:C20`)  |
| `hiddenRows`    | Comma-separated row numbers to hide (e.g. `5, 6`)  |
| `hiddenColumns` | Comma-separated column letters to hide (e.g. `D, E`)|

---

## 3. Cells Section

### Format
```text
CellRef = Value | style1: value1, style2: value2
```

### Examples
```text
A1 = "Title"
B2 = 15 | textAlign: right, format: number
C5 = 28800000 | format: currency:IDR, fontWeight: bold, backgroundColor: #f1f5f9
D5 = "=SUM(D2:D4)" | format: currency:USD, fontWeight: bold

// Cell comment
A6 = "Revenue" | fontWeight: bold, comment: "Includes VAT for EU regions"

// Rich text (multiple segments with different styles)
A7 = richtext | segments: [{text: "Total", bold: true}, {text: " (net)", color: #64748b}]
```

---

## 3a. Multi-Sheet Workbooks

Separate multiple sheets within one `.antsheet` file using a line containing only `---`.

```text
# Sheet 1: Sales

## Cells
A1 = "Product"
B1 = 16800000

---

# Sheet 2: Summary

## Cells
A1 = "Total Revenue"
B1 = "=Sales!B1"
```

**API:**
```typescript
import { parseWorkbook, serializeWorkbook } from "antsheet";

const wb = parseWorkbook(text);  // WorkbookDocument { sheets: SheetDocument[] }
const text = serializeWorkbook(wb);
```

---


## 4. Merges Section

Combines cell ranges into merged cells.

```text
## Merges
A1:C1
A5:B5
```

---

## 5. Style System

### Typography
| Property         | Supported Values                  |
|------------------|-----------------------------------|
| `fontFamily`     | `Inter`, `Arial`, `Roboto`, etc.  |
| `fontSize`       | `12`, `14`, `16`, `18`, `20`      |
| `fontWeight`     | `normal`, `bold`, `100`–`900`     |
| `fontStyle`      | `normal`, `italic`, `oblique`     |
| `textDecoration` | `none`, `underline`, `lineThrough`|
| `textTransform`  | `none`, `uppercase`, `lowercase`  |
| `letterSpacing`  | `0`, `0.5`, `1`                   |
| `lineHeight`     | `1.2`, `1.5`, `24`                |

### Colors & Alignment
| Property          | Supported Values          |
|-------------------|---------------------------|
| `color`           | `#ffffff`, `red`, etc.    |
| `backgroundColor` | `#1e293b`, `transparent`  |
| `textAlign`       | `left`, `center`, `right` |
| `verticalAlign`   | `top`, `middle`, `bottom` |

### Text Wrapping
| Property   | Supported Values              | Description                               |
|------------|-------------------------------|-------------------------------------------|
| `textWrap` | `wrap`, `clip`, `overflow`    | Controls how multi-line/overflow text fits |
| `wrapText` | `true`, `false`               | Boolean shortcut for wrapping text        |

### Value Formatting

| Format Pattern | Formatted Output Example |
|----------------|--------------------------|
| `number`       | `1,234.5`                |
| `number:2`     | `1,234.57`               |
| `number:0`     | `1,235`                  |
| `compact`      | `1.5M`                   |
| `currency`     | `Rp 22.500.000`          |
| `currency:IDR` | `Rp 22.500.000`          |
| `currency:USD` | `$1,500.00`              |
| `currency:EUR` | `€1,500.00`              |
| `percent`      | `85%`                    |
| `percent:1`    | `85.6%`                  |
| `percent:2`    | `85.67%`                 |
| `date`         | `15/03/2026`             |
| `date:long`    | `15 March 2026`          |
| `date:iso`     | `2026-03-15`             |
| `datetime`     | `15/03/2026 14:30`       |
| `time`         | `14:30`                  |

### Borders & Padding
- **Shorthand:** `border: all`, `border: bottom`, `borderTop: 2 solid #0f172a`
- **Padding:** `padding: 8`, `padding: 4 12`, `paddingTop: 8`

---

## 6. Charts Section

### Format
```text
## Charts
ChartID = ChartType | prop1: value1, prop2: value2
```

### Supported Chart Types
- `bar` - Column / Bar chart
- `line` - Line trend chart
- `pie` - Circular proportional pie
- `doughnut` - Ring chart
- `area` - Filled area chart
- `scatter` - Scatter plot
- `radar` - Radar / spider chart

### Properties
| Property     | Example              | Description                                                    |
|--------------|----------------------|----------------------------------------------------------------|
| `title`      | `"Q1 Sales"`         | Chart title                                                    |
| `range`      | `A1:C5`              | Full table range (first row = headers, first column = X labels)|
| `x` / `xAxis`| `A` or `A2:A5`       | Category / X-axis column or range                              |
| `series`     | `[B, C]`             | Series column letters or ranges                                |
| `labels`     | `A2:A5`              | Category labels (for pie / doughnut)                           |
| `values`     | `B2:B5`              | Numerical values (for pie / doughnut)                          |
| `position`   | `E2:L14`             | Grid position overlay range                                    |
| `stacked`    | `true`, `false`      | Stacked bars/areas                                             |
| `horizontal` | `true`, `false`      | Horizontal orientation                                         |
| `smooth`     | `true`, `false`      | Curved spline lines                                            |
| `colors`     | `[#3b82f6, #10b981]` | Custom series palette                                          |

---

## 7. Data Validation Section

Defines input constraints and dropdown lists for cell ranges.

### Format
```text
## DataValidation
rule1 = list | range: D2:D50, options: ["Pending", "In Progress", "Done", "Cancelled"]
rule2 = number | range: E2:E50, operator: between, min: 0, max: 100
rule3 = checkbox | range: F2:F50
```

### Validation Types
| Type         | Description                            |
|--------------|----------------------------------------|
| `list`       | Dropdown of predefined string options  |
| `number`     | Numeric constraint (min, max, operator)|
| `textLength` | Character length constraint            |
| `date`       | Date range constraint                  |
| `checkbox`   | Boolean true/false checkbox            |
| `custom`     | Custom formula validation              |

### Properties
| Property     | Description                                        |
|--------------|----------------------------------------------------|
| `range`      | Cell range to apply the rule (required)            |
| `options`    | Array of allowed values (for `list` type)          |
| `operator`   | `between`, `greaterThan`, `lessThan`, `equal`, etc |
| `min` / `max`| Numeric bounds                                     |
| `allowBlank` | Allow empty cells (`true` / `false`)               |
| `error`      | Custom error message on invalid input              |
| `prompt`     | Helper tooltip shown when cell is selected         |

---

## 8. Conditional Formatting Section

Auto-applies visual formatting to cells based on their values.

### Format
```text
## ConditionalFormatting
cf1 = highlight | range: C2:C20, operator: ">=", value: 90, backgroundColor: #dcfce7, color: #166534
cf2 = colorScale | range: D2:D20, minColor: #fee2e2, maxColor: #dcfce7
cf3 = dataBar | range: E2:E20, minColor: #bfdbfe, maxColor: #3b82f6
```

### Rule Types
| Type          | Description                                       |
|---------------|---------------------------------------------------|
| `highlight`   | Highlight cells matching a condition              |
| `colorScale`  | Gradient color scale across a value range         |
| `dataBar`     | In-cell proportional bar indicator                |

### Properties
| Property          | Description                                        |
|-------------------|----------------------------------------------------|
| `range`           | Cell range to apply the rule (required)            |
| `operator`        | `>`, `<`, `>=`, `<=`, `=`, `!=`, `between`, etc   |
| `value`           | Threshold value for comparison                     |
| `value2`          | Second value for `between` operator                |
| `color`           | Text color to apply                                |
| `backgroundColor` | Background cell highlight color                    |
| `minColor`        | Low end color for `colorScale` / `dataBar`         |
| `maxColor`        | High end color for `colorScale` / `dataBar`        |
| `midColor`        | Mid-point color for 3-color scale                  |

---

## 9. TypeScript Data Model (`SheetDocument`)

```typescript
export interface SheetDocument {
  name: string;
  meta: SheetMeta;
  cells: Record<string, CellData>;
  merges: string[];
  charts?: ChartConfig[];
  dataValidations?: DataValidationRule[];
  conditionalFormattings?: ConditionalFormattingRule[];
}

export interface SheetMeta {
  columnWidth?: Record<string, number>;
  rowHeight?: Record<string, number>;
  freezeRow?: number;
  freezeColumn?: number;
  filter?: string;              // e.g. "A1:G50"
}
```

---

## 10. API & Usage

```typescript
import {
  parseAntsheet,       // Single sheet
  parseWorkbook,       // Multi-sheet workbook
  serializeAntsheet,   // Single sheet -> text
  serializeWorkbook,   // Workbook -> text
  convertToUniver,
  convertFromUniver,
  extractChartData,
  validateAntsheet,
} from "antsheet";

// Excel files (.xlsx)
import { parseXLSX, exportToXLSX } from "antsheet/xlsx";

// CSV files
import { parseCSV, exportToCSV } from "antsheet/csv";

// Formula evaluation (HyperFormula runtime)
import { evaluateFormulas, evaluateWorkbook } from "antsheet/formula";


// 1. Parse single sheet
const doc = parseAntsheet(antsheetText);

// 2. Parse multi-sheet workbook
const wb = parseWorkbook(antsheetText);
wb.sheets[0].cells; // first sheet

// 3. Validate syntax (errors + warnings)
const { valid, errors, warnings } = validateAntsheet(antsheetText);

// 4. Extract chart dataset for rendering
if (doc.charts?.length) {
  const chart = extractChartData(doc, doc.charts[0]);
  console.log(chart.labels, chart.datasets);
}

// 5. Convert to/from Univer sheets engine
const univerData = convertToUniver(doc);
const backToDoc = convertFromUniver(univerData);

// 6. Import Excel -> AntSheet
const importedDoc = parseXLSX(fs.readFileSync("report.xlsx"));
const antsheetOut = serializeAntsheet(importedDoc);

// 7. Export AntSheet -> Excel
const xlsxBytes = exportToXLSX(doc);
fs.writeFileSync("output.xlsx", xlsxBytes);

// 8. Import CSV -> AntSheet
const csvDoc = parseCSV(csvText, { hasHeader: true, sheetName: "Data" });

// 9. Export AntSheet -> CSV
const csvOut = exportToCSV(doc);

// 10. Evaluate formulas in-memory (computes all =SUM, =IF, etc.)
const evaluatedDoc = evaluateFormulas(doc);
console.log(evaluatedDoc.cells["D4"].value); // e.g. 21000000

// 11. Serialize back to .antsheet text
const outputText = serializeWorkbook(wb);
```

---

## 11. AI Agent Instructions

When prompting an AI assistant to generate or modify `.antsheet` documents, use these guidelines:

```text
You generate and edit spreadsheets in the AntSheet (.antsheet) format.

Strict rules:
1.  Section order: # Sheet Name, ## Meta, ## Cells, ## Merges,
    ## DataValidation, ## ConditionalFormatting, ## Charts
2.  One cell per line: CellRef = Value | styles
3.  Text values MUST be quoted: A1 = "Product Name"
4.  Numbers/booleans NOT quoted: B1 = 15000, C1 = true
5.  Formulas start with = inside quotes: D1 = "=SUM(B2:B10)"
6.  Cross-sheet formulas: B1 = "=Sheet2!C5"
7.  Cell comments: A1 = "Value" | comment: "Note here"
8.  Rich text: A1 = richtext | segments: [{text: "Bold", bold: true}, {text: " normal"}]
9.  Named ranges: namedRange: Sales_Total = C2:C20 (in ## Meta)
10. Hidden rows/cols: hiddenRows: 5, 6 / hiddenColumns: D, E (in ## Meta)
11. Multi-sheet: separate sheets with a line containing only ---
12. For dropdowns: rule1 = list | range: D2:D50, options: ["A","B"]
13. For color rules: cf1 = highlight | range: C2:C20, operator: ">", value: 90, backgroundColor: #dcfce7
14. For charts: chart1 = bar | title: "...", x: A, series: [B, C], position: E2:L14
15. NEVER output Markdown tables or CSV when editing .antsheet files.
```