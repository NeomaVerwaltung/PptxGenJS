# Release Procedure

This document describes the release process for the NEOMA distribution of PptxGenJS. A published GitHub Release triggers the npm publication and bundle upload.

## Beta releases

Follow the standard preparation below, then create and publish a **pre-release** in GitHub for a tag such as `v4.1.0-beta.0`. The release workflow publishes the corresponding npm package under the `beta` tag and attaches the bundles.

## Build and update files

1. Update the version in `package.json`.
2. Update the version constant in `src/pptxgen.ts` (for example, `const VERSION = '4.0.1'`).
3. Update `CHANGELOG.md` with the release date.
4. Build the library: npm scripts > `ship`.
5. Consolidate new type changes from `src/bld/*.ts` into `types/index.d.ts` and update the version in the header comment.
6. Inspect the headers of the generated `dist/*.js` files.

> `dist/` is not committed. When a GitHub Release is published, the release workflow rebuilds it (`npm ci` runs `prepare`), publishes the package to npm, and attaches `pptxgen.bundle.js`, `pptxgen.min.js`, `pptxgen.cjs.js` and `pptxgen.es.js` to that release.

## Pre-release testing

Run the standard test suite as documented in [TESTING.md](./TESTING.md), then record the results:

| Dist File      | Test      | Tested Via             | Result |
| -------------- | --------- | ---------------------- | ------ |
| pptxgen.es.js  | Webpack 4 | SPFx (v1.16.1) project |        |
| pptxgen.es.js  | Webpack 5 | SPFx (v1.19.1) project |        |
| pptxgen.es.js  | Rollup 4  | Vite (v6) scaffold     |        |
| pptxgen.cjs.js | Node/CJS  | Node scaffold          |        |

## Pre-release verification

Confirm the following before publishing:

1. The `version` field in `package.json` is updated.
2. The version constant in `src/pptxgen.ts` is updated.
3. The version in the `types/index.d.ts` header is updated.

## Release: GitHub

1. Commit all changes.
2. Merge the working branch into `main`.
3. Draft a new release on the [Releases page](https://github.com/NeomaVerwaltung/PptxGenJS/releases), using the release commit as the target.
4. Use "Version x.x.x" as the title and "vX.X.X" as the tag; generate or paste the release notes from `CHANGELOG.md`.
5. Verify the title, tag, and target commit, then publish the release. This starts the release workflow; it publishes npm and uploads the generated bundles to the same GitHub Release.

## Post-release tasks

1. Verify the CDN links in `README.md`.
2. Update `docs/installation.md` if the CDN version changed.
3. Update the documentation under `docs/` for any API changes.

The documentation site (GitHub Pages) redeploys automatically from `docs/` on push to `master` via `.github/workflows/docs.yml`; no manual gh-pages step is required.
