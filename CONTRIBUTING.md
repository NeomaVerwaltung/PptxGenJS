# Contributing

Thanks for helping improve PptxGenJS.

## Development setup

Requires Node.js **20 or newer** (matches CI and the `engines` field).

```bash
npm ci        # install exact locked dependencies
npm run check # lint + typecheck + strict typecheck + tests — run before every PR
```

## Before opening a PR

`npm run check` must pass. It runs:

- `npm run lint` — ESLint
- `npm run typecheck` — full-project TypeScript
- `npm run typecheck:strict` — `strictNullChecks` on migrated files (see below)
- `npm test` — unit, e2e, and XML snapshot tests

Build the distributables locally with `npm run build` (or `npm run dist` for
the full minified/bundled `dist/` artifacts).

## Tests

Tests live in `test/`. When you change XML generation, update or add a
golden-file snapshot (`test/__snapshots__/`) and include the regenerated
snapshot in your PR. Coverage can be inspected with `npm run test:coverage`.

## `strictNullChecks` migration

The codebase is migrating to `strictNullChecks` incrementally (see
`tsconfig.strict.json`). When you make a source file null-safe, add it to the
`include` list in `tsconfig.strict.json` so it stays that way. When every
`src/**` file is listed, `strictNullChecks` will be enabled in the base
`tsconfig.json` and `tsconfig.strict.json` deleted.

## Coding style

Match the surrounding code. ESLint and the stylistic plugin enforce
formatting — run `npm run lint` and fix reported issues before pushing.
