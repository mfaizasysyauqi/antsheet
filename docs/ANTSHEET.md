```markdown
# Sheetdown Specification

Format teks polos untuk spreadsheet yang agent-friendly.  
Dirancang agar mudah diedit AI, tetap mendukung styling, dan bisa di-preview memakai Univer.

**File extension:** `.antsheet`

---

## Alur Utama

```text
Teks .antsheet
     ↓
  Parser
     ↓
SheetDocument
     ↓
 Converter  →  Univer (preview)
     ↓
 Serializer
     ↓
Teks .antsheet (lagi)
```

---

## 1. Syntax Dasar

```text
# Nama Sheet

## Meta
columnWidth: A=140, B=100, C=120
rowHeight: 1=32

## Cells
A1 = "Produk" | fontWeight: bold, backgroundColor: #1e293b, color: #ffffff, textAlign: center
B1 = "Terjual" | fontWeight: bold, backgroundColor: #1e293b, color: #ffffff, textAlign: center
C1 = "Total" | fontWeight: bold, backgroundColor: #1e293b, color: #ffffff, textAlign: center

A2 = "Laptop"
B2 = 15 | textAlign: right, format: number
C2 = 22500000 | textAlign: right, format: currency:IDR, fontWeight: bold

## Merges
A1:C1
```

### Aturan Umum

- Satu cell = satu baris
- String harus memakai tanda kutip `"..."`
- Angka dan boolean tidak memakai tanda kutip
- Style bersifat opsional
- Style memakai nama lengkap (Dart-style)
- File disimpan dengan extension `.antsheet`

---

## 2. Bagian Meta

```text
## Meta
columnWidth: A=140, B=100, C=120
rowHeight: 1=32, 5=40
```

| Property      | Keterangan        |
|---------------|-------------------|
| `columnWidth` | Lebar kolom (px)  |
| `rowHeight`   | Tinggi baris (px) |

---

## 3. Cells

Format:

```text
CellRef = Nilai | style1: value1, style2: value2
```

Contoh:

```text
A1 = "Judul"
B2 = 15 | textAlign: right, format: number
C5 = 28800000 | format: currency:IDR, fontWeight: bold, backgroundColor: #f1f5f9
```

---

## 4. Merges

```text
## Merges
A1:C1
A5:B5
```

Setiap baris berisi satu range yang digabung.

---

## 5. Style System

### Font

| Property         | Contoh Nilai                       |
|------------------|------------------------------------|
| `fontFamily`     | `Inter`, `Arial`, `Roboto`         |
| `fontSize`       | `12`, `14`, `16`, `18`, `20`       |
| `fontWeight`     | `normal`, `bold`, `w100`–`w900`    |
| `fontStyle`      | `normal`, `italic`, `oblique`      |
| `textDecoration` | `none`, `underline`, `lineThrough` |
| `textTransform`  | `none`, `uppercase`, `lowercase`   |
| `letterSpacing`  | `0`, `0.5`, `1`                    |
| `lineHeight`     | `1.2`, `1.5`, `24`                 |

### Warna & Alignment

| Property          | Contoh Nilai              |
|-------------------|---------------------------|
| `color`           | `#ffffff`, `red`          |
| `backgroundColor` | `#1e293b`, `transparent`  |
| `textAlign`       | `left`, `center`, `right` |
| `verticalAlign`   | `top`, `middle`, `bottom` |

### Format Nilai

| Format         | Contoh Hasil         |
|----------------|----------------------|
| `number`       | `1,234.5`            |
| `number:2`     | `1,234.57`           |
| `number:0`     | `1,235`              |
| `compact`      | `1.5M`               |
| `currency`     | `Rp 22.500.000`      |
| `currency:IDR` | `Rp 22.500.000`      |
| `currency:USD` | `$1,500.00`          |
| `currency:EUR` | `€1,500.00`          |
| `percent`      | `85%`                |
| `percent:1`    | `85.6%`              |
| `percent:2`    | `85.67%`             |
| `date`         | `15/03/2026`         |
| `date:long`    | `15 Maret 2026`      |
| `date:iso`     | `2026-03-15`         |
| `datetime`     | `15/03/2026 14:30`   |
| `time`         | `14:30`              |
| `text`         | `00123` (tetap teks) |
| `boolean`      | `Ya` / `Tidak`       |

### Border

**Cara singkat:**

```text
border: all
border: bottom
border: none
```

**Cara detail:**

```text
borderTop: 1
borderBottom: 2 solid #0f172a
borderLeft: 4 solid #3b82f6
borderRight: 1 solid #e2e8f0
```

| Property       | Keterangan                                      |
|----------------|-------------------------------------------------|
| `border`       | `all`, `bottom`, `top`, `left`, `right`, `none` |
| `borderTop`    | number atau `2 solid #000`                      |
| `borderRight`  | sama                                            |
| `borderBottom` | sama                                            |
| `borderLeft`   | sama                                            |
| `borderColor`  | warna global                                    |
| `borderWidth`  | ketebalan global                                |
| `borderStyle`  | `solid`, `dashed`, `dotted`                     |

### Padding

**Cara singkat:**

```text
padding: 8
padding: 4 12
padding: 8 12 8 12
```

**Cara detail:**

```text
paddingTop: 8
paddingRight: 12
paddingBottom: 8
paddingLeft: 12
```

---

## 6. Data Structure (`SheetDocument`)

```typescript
interface SheetDocument {
  name: string;
  meta: SheetMeta;
  cells: Record<string, CellData>;
  merges: string[];
}

interface SheetMeta {
  columnWidth?: Record<string, number>;
  rowHeight?: Record<string, number>;
}

interface CellData {
  value: string | number | boolean | null;
  style?: CellStyle;
}

interface CellStyle {
  // Font
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: string | number;
  fontStyle?: "normal" | "italic" | "oblique";
  textDecoration?: string;
  textTransform?: string;
  letterSpacing?: number;
  lineHeight?: number | string;

  // Color & Align
  color?: string;
  backgroundColor?: string;
  textAlign?: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";

  // Format
  format?: string;

  // Border
  border?: string;
  borderTop?: string | number;
  borderRight?: string | number;
  borderBottom?: string | number;
  borderLeft?: string | number;
  borderColor?: string;
  borderWidth?: number;
  borderStyle?: string;

  // Padding
  padding?: string | number;
  paddingTop?: number;
  paddingRight?: number;
  paddingBottom?: number;
  paddingLeft?: number;
}
```

---

## 7. Parser

Tugas: Teks `.antsheet` → `SheetDocument`

Langkah utama:

1. Deteksi nama sheet dari baris `# ...`
2. Deteksi section (`## Meta`, `## Cells`, `## Merges`)
3. Parse Meta → `columnWidth` & `rowHeight`
4. Parse setiap baris Cells → `cells[ref]`
5. Parse Merges → array range

---

## 8. Converter (ke Univer)

Tugas: `SheetDocument` → data / command Univer

Yang dikonversi:

- `columnWidth` → `columnData`
- `rowHeight` → `rowData`
- `cells` → `cellData` + style mapping
- `merges` → `mergeData`

### Mapping Style ke Univer

| Sheetdown           | Univer                  |
|---------------------|-------------------------|
| `fontFamily`        | `ff`                    |
| `fontSize`          | `fs`                    |
| `fontWeight: bold`  | `bl: 1`                 |
| `fontStyle: italic` | `it: 1`                 |
| `color`             | `cl: { rgb: "..." }`    |
| `backgroundColor`   | `bg: { rgb: "..." }`    |
| `textAlign: center` | `ht: 2`                 |
| `textAlign: right`  | `ht: 3`                 |
| `format`            | `n: { pattern: "..." }` |
| Border              | `bd`                    |

---

## 9. Serializer

Tugas: `SheetDocument` → Teks `.antsheet`

Aturan:

- Tulis `# Nama Sheet`
- Tulis `## Meta` jika ada data
- Tulis `## Cells` (urutkan dari A1, B1, C1, A2 ...)
- Tulis `## Merges` jika ada
- Style ditulis lengkap (Dart-style)
- String selalu dibungkus `"..."`

---

## 10. Aturan untuk Agent

```text
Kamu mengedit spreadsheet dalam format Sheetdown (.antsheet).

Aturan wajib:
1. Pertahankan struktur #, ## Meta, ## Cells, ## Merges
2. Satu cell = satu baris
3. Format cell: A1 = "nilai" | fontWeight: bold, color: #ffffff
4. String pakai "...", angka tidak
5. Style memakai nama lengkap (fontWeight, backgroundColor, textAlign, dll)
6. Urutkan cell dari kiri ke kanan, atas ke bawah
7. Kalau hanya ubah nilai, pertahankan style yang sudah ada
8. Jangan memakai Markdown table atau CSV
```

---

## 11. Contoh Lengkap

```text
# Penjualan Q1

## Meta
columnWidth: A=140, B=100, C=140
rowHeight: 1=36

## Cells
A1 = "Produk" | fontWeight: bold, fontSize: 14, backgroundColor: #1e293b, color: #ffffff, textAlign: center
B1 = "Terjual" | fontWeight: bold, fontSize: 14, backgroundColor: #1e293b, color: #ffffff, textAlign: center
C1 = "Total" | fontWeight: bold, fontSize: 14, backgroundColor: #1e293b, color: #ffffff, textAlign: center

A2 = "Laptop"
B2 = 15 | textAlign: right, format: number
C2 = 22500000 | textAlign: right, format: currency:IDR, fontWeight: bold

A3 = "Mouse"
B3 = 42 | textAlign: right, format: number
C3 = 2100000 | textAlign: right, format: currency:IDR

A4 = "Keyboard"
B4 = 28 | textAlign: right, format: number
C4 = 4200000 | textAlign: right, format: currency:IDR

A6 = "Total" | fontWeight: bold, backgroundColor: #f1f5f9, borderTop: 2 solid #94a3b8
C6 = 28800000 | fontWeight: bold, backgroundColor: #f1f5f9, format: currency:IDR, textAlign: right, borderTop: 2 solid #94a3b8

## Merges
A1:C1
```

---

## Catatan Desain

- File extension resmi: **`.antsheet`**
- Versi ini **belum mendukung formula hidup**
- Chart tidak disimpan di dalam format ini (bisa ditambahkan nanti lewat section terpisah atau lewat Univer)
- Fokus utama: agent-friendly + styling dasar yang cukup kaya
```