# AntSheet

**AntSheet** is a lightweight, type-safe TypeScript library for working with spreadsheet data. It provides utilities for parsing, converting, and serializing spreadsheet formats.

## Features

- ✅ **Type-safe** - Full TypeScript support with comprehensive type definitions
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
import { parseSpreadsheet, convertFormat, serialize } from 'antsheet';

// Parse spreadsheet data
const data = parseSpreadsheet(rawData);

// Convert format
const converted = convertFormat(data, 'xlsx');

// Serialize
const serialized = serialize(data);
```

## Documentation

See [ANTSHEET.md](./docs/ANTSHEET.md) for detailed documentation.

## API Reference

### Types

See [types.ts](./src/types.ts) for complete type definitions.

### Parser

```typescript
import { parseSpreadsheet } from 'antsheet/parser';
```

### Converter

```typescript
import { convertFormat } from 'antsheet/converter';
```

### Serializer

```typescript
import { serialize } from 'antsheet/serializer';
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
