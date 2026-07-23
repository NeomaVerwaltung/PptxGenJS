# Release Procedure

This document describes the release process for the NEOMA distribution of PptxGenJS. Maintainers must complete this checklist before publishing to npm or GitHub.

## Beta releases

1. Update the version in `package.json` (for example, `4.1.0-beta.0`).
2. Update the version constant in `src/pptxgen.ts`.
3. Build the library: npm scripts > `ship`.
4. Publish: `npm publish --tag beta`.

## Build and update files

1. Update the version in `package.json`.
2. Update the version constant in `src/pptxgen.ts` (for example, `const VERSION = '4.0.1'`).
3. Update `CHANGELOG.md` with the release date.
4. Build the library: npm scripts > `ship`.
5. Consolidate new type changes from `src/bld/*.ts` into `types/index.d.ts` and update the version in the header comment.
6. Inspect the headers of the generated `dist/*.js` files.
7. Update the version in `demos/node/package.json` (its `@neo-ma/pptxgenjs` dependency is `file:../..`; no dependency bump is required).

## Pre-release testing

Run the standard test suite as documented in [TESTING.md](./TESTING.md), then record the results:

| Dist File      | Test      | Tested Via             | Result |
| -------------- | --------- | ---------------------- | ------ |
| pptxgen.es.js  | Webpack 4 | SPFx (v1.16.1) project |        |
| pptxgen.es.js  | Webpack 5 | SPFx (v1.19.1) project |        |
| pptxgen.es.js  | Rollup 4  | Vite (v6) scaffold     |        |
| pptxgen.cjs.js | Node/CJS  | Node demo              |        |

## Pre-release verification

Confirm the following before publishing:

1. The `version` field in `package.json` is updated.
2. The version constant in `src/pptxgen.ts` is updated.
3. The version in the `types/index.d.ts` header is updated.

## Release: GitHub

1. Commit all changes.
2. Merge the working branch into `main`.
3. Copy the CHANGELOG entry and draft a new release on the [Releases page](https://github.com/NeomaVerwaltung/PptxGenJS/releases).
4. Use "Version x.x.x" as the title and "vX.X.X" as the tag.
5. Verify the title and tag on the Releases page, then publish.

## Release: npm

```bash
cd ~/GitHub/PptxGenJS
npm publish
```

## Post-release tasks

1. Verify the CDN links in `README.md`.
2. Update `docs/installation.md` if the CDN version changed.
3. Update the documentation under `docs/` for any API changes.

The documentation site (GitHub Pages) redeploys automatically from `docs/` on push to `master` via `.github/workflows/docs.yml`; no manual gh-pages step is required.
