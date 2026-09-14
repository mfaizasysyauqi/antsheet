# AntSheet

**AntSheet** is a lightweight, type-safe TypeScript library for working with spreadsheet data. It provides utilities for parsing, converting, and serializing spreadsheet formats.

## Features

- ✅ **Type-safe** - Full TypeScript support with comprehensive type definitions
- 📈 **Chart generation** - Declare and extract chart datasets (Bar, Line, Pie, Area, etc.)
- 🔄 **Format conversion** - Convert between different spreadsheet formats
- 📊 **Data parsing** - Parse spreadsheet data with ease
- 💾 **Serialization** - Serialize spreadsheet data efficiently
- 🎨 **Theme support** - Built-in dark/light mode theming

## Installation

```bash
npm install antsheet
# or
pnpm add antsheet
# or
yarn add antsheet
```

## Usage

```typescript
import { parseAntsheet, serializeAntsheet, extractChartData } from 'antsheet';

// Parse .antsheet file
const doc = parseAntsheet(antsheetContent);

// Extract chart data (ready for Recharts, Chart.js, etc.)
if (doc.charts && doc.charts.length > 0) {
  const chartData = extractChartData(doc, doc.charts[0]);
  console.log(chartData.labels, chartData.datasets);
}

// Serialize back to .antsheet
const serialized = serializeAntsheet(doc);
```

## Documentation

See [ANTSHEET.md](./docs/ANTSHEET.md) for detailed documentation.

## API Reference

### Types

See [types.ts](./src/types.ts) for complete type definitions.

### Parser

```typescript
import { parseAntsheet } from 'antsheet/parser';
```

### Serializer

```typescript
import { serializeAntsheet } from 'antsheet/serializer';
```

### Chart Extractor

```typescript
import { extractChartData } from 'antsheet/chart';
```

### Converter

```typescript
import { convertToUniver } from 'antsheet/converter';
```


## Development

```bash
# Install dependencies
pnpm install

# Build
pnpm build

# Test
pnpm test
```

## License

MIT © Anticeil

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Related Projects

- [Anticeil](https://github.com/mfaizasysyauqi/anticeil) - The main Anticeil application
- [Univer](https://github.com/dream-num/univer) - Office suite engine

---

Made with ❤️ by the Anticeil team
