# Release Procedure

This document describes the release process for the NEOMA distribution of PptxGenJS. There are two paths, both running the same `Release` workflow: the **automatic release** below (recommended), and a manually published GitHub Release.

## Two packages, independent versions

| Package | Directory | Tag scheme |
| --- | --- | --- |
| `@neo-ma/pptxgenjs` | `packages/core` | `vX.Y.Z` |
| `@neo-ma/pptxgenjs-std` | `packages/std` | `std-vX.Y.Z` |

One workflow releases either. On the dispatch path the `package` input selects it; on the published-release path the tag prefix does (`std-v*` → std, anything else → core). The versions are unrelated — never bump them together.

## Automatic release

Run the [Release workflow](https://github.com/NeomaVerwaltung/PptxGenJS/actions/workflows/release.yml) via **Run workflow** on `master`, pick the package (`core`, `std`) and the increment (`patch`, `minor`, `major`). After an admin approves the `release` environment, the job:

1. Bumps that package's `package.json` — for core also `src/pptxgen.ts` and the `types/index.d.ts` header via `scripts/sync-version.mjs` — prepends the `CHANGELOG.md` section via `scripts/changelog.mjs`, commits `chore: release <tag>` and pushes the tag.
2. Rebuilds the package's `dist/`, verifies every published artifact, and publishes to npm under the `latest` tag.
3. Creates the GitHub Release with generated notes; core also gets its browser bundles attached.

The `release` environment's deployment branch policy must list `master` (dispatch path) and both tag patterns `v*` and `std-v*` (published-release path); a missing entry makes the run fail before the approval prompt appears.

`CHANGELOG.md` is written by the run itself: `scripts/changelog.mjs` collects the commits that touched the package being released since its previous tag, buckets them by conventional-commit type (`feat` → Added, `fix` → Fixed, `perf`/`refactor`/`revert`/breaking → Changed) and prepends the section. Plumbing types (`docs`, `chore`, `ci`, `test`, `build`, `style`) and commits outside the package directory are left out, so write commit subjects as the changelog line you want. Update the docs before starting the run.

For a beta, use the manual path below; the dispatch inputs cannot express a pre-release tag. That path does not generate a changelog entry — the tag already exists when the workflow runs — so write it by hand.

## Beta releases

Follow the standard preparation below, then create and publish a **pre-release** in GitHub for a tag such as `v4.1.0-beta.0` (or `std-v0.2.0-beta.0`). The release workflow publishes the corresponding npm package under the `beta` tag and attaches the bundles.

## Build and update files

Only needed for the manual path; the automatic release handles steps 1, 2 and 5.

Paths below are relative to the package being released (`packages/core` unless noted).

1. Update the version in `package.json` (or run `npm version <level> --no-git-tag-version -w <package>` from the root, which syncs steps 2 and 5 for you).
2. Update the version constant in `src/pptxgen.ts` (for example, `const VERSION = '4.0.1'`).
3. Update `CHANGELOG.md` with the release date (the automatic path does this for you).
4. Build the library: npm scripts > `ship`.
5. Consolidate new type changes from `src/bld/*.ts` into `types/index.d.ts` and update the version in the header comment.
6. Inspect the headers of the generated `dist/*.js` files.

> `dist/` is not committed. When a GitHub Release is published, the release workflow rebuilds it (`npm ci` runs each workspace's `prepare`), publishes the selected package to npm, and — for core — attaches every generated static asset in `dist/` to that release.

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

1. The `version` field in the package's `package.json` is updated.
2. For core: the version constant in `packages/core/src/pptxgen.ts` is updated.
3. For core: the version in the `packages/core/types/index.d.ts` header is updated.

## Release: GitHub (manual path)

1. Commit all changes.
2. Merge the working branch into `main`.
3. Draft a new release on the [Releases page](https://github.com/NeomaVerwaltung/PptxGenJS/releases), using the release commit as the target.
4. Use "Version x.x.x" as the title and "vX.X.X" (core) or "std-vX.Y.Z" (std) as the tag; generate or paste the release notes from `CHANGELOG.md`. The tag prefix is what selects the package, so a typo publishes the wrong one.
5. Verify the title, tag, and target commit, then publish the release. This starts the release workflow; it publishes npm and uploads the generated bundles to the same GitHub Release.

## Post-release tasks

1. Verify the CDN links in `README.md`.
2. Update `docs/installation.md` if the CDN version changed.
3. Update the documentation under `docs/` for any API changes.

The documentation site (GitHub Pages) redeploys automatically from `docs/` on push to `master` via `.github/workflows/docs.yml`; no manual gh-pages step is required.
