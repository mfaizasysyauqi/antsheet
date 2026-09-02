import { defineConfig } from 'tsup';

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    parser: 'src/parser.ts',
    converter: 'src/converter.ts',
    serializer: 'src/serializer.ts',
    types: 'src/types.ts',
    chart: 'src/chart.ts',
    validator: 'src/validator.ts',
    xlsx: 'src/xlsx.ts',
    csv: 'src/csv.ts',
    formula: 'src/formula.ts',
  },



  format: ['cjs', 'esm'],
  dts: true,
  splitting: false,
  sourcemap: true,
  clean: true,
  minify: false,
  treeshake: true,
});
