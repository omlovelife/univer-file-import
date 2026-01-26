import { defineConfig } from 'tsup';

export default defineConfig({
  entry: ['src/index.ts'],
  format: ['cjs', 'esm'],
  dts: true,
  splitting: false,
  sourcemap: true,
  clean: true,
  treeshake: true,
  minify: true,
  external: [
    '@univerjs/core',
    '@univerjs/presets',
  ],
  esbuildOptions(options) {
    options.define = {
      UNIVER_VERSION: '"0.15.0"',
    };
  },
});
