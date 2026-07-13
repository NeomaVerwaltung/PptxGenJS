# PptxGenJS Release Checklist

> This guide documents how to perform a PptxGenJS release.
> Maintainers should follow this checklist before pushing to npm or GitHub.

## 📋 Beta Releases

1. Update `package.json` version (ex: `4.1.0-beta.0`)
2. Update `src/pptxgen.ts` version
3. Build library: npm scripts > `ship`
4. `npm publish --tag beta`

## 🚀 Build Library, Update Files

1. Update `package.json` version
2. Update `src/pptxgen.ts` version (eg: `const VERSION = '4.0.1'`)
3. Update `CHANGELOG.md` with new date
4. Build library: npm scripts > `ship`
5. Consolidate new changes from `src/bld/*.ts` into `types/index.d.ts` and update version in head comment
6. Open `dist/*.js` and check headers
7. Update version in: `demos/node/package.json` (its `@neoma/pptxgenjs` dep is `file:../..` — no dep bump needed)

## 🧪 Run Tests Before Release

### ⚠️ Run Standard Test Suite

See [TESTING.md](./TESTING.md) for complete test instructions.

### ⚠️ Capture Testing Results

| Dist File         | Test       | Tested Via             | Result |
| ----------------- | ---------- | ---------------------- | ------ |
| pptxgen.es.js     | Webpack 4  | SPFx (v1.16.1) project | ✅?🟡    |
| pptxgen.es.js     | Webpack 5  | SPFx (v1.19.1) project | ✅?🟡    |
| pptxgen.es.js     | Rollup 4   | Vite (v6) scaffold     | ✅?🟡    |
| pptxgen.cjs.js    | Node/CJS   | Node demo              | ✅?🟡    |

## 🚌 Release New Version

### 🟡 Pre-Release Checklist

1. Check: Is `version` updated in package.json?
2. Check: Is `version` updated in src/pptxgen.ts?
3. Check: Is `types/index.d.ts` version in header updated?

### 🟢 Release: GitHub

1. Checkin all changes via GitHub Desktop
2. Merge working branch into `main`
3. Copy CHANGELOG entry and draft new release: [Releases](https://github.com/NeomaVerwaltung/PptxGenJS/releases)
4. Use "Version x.x.x" as title and "vX.X.X" as tag
5. Go back to Releases page, double-check title/tag, release when ready

### 🟢 Release: NPM

```bash
cd ~/GitHub/PptxGenJS
npm publish
```

## 🏁 Post-Release Tasks

1. Test CDN links on README.md
2. Update `docs/installation.md` with the latest CDN version if it changed
3. Update the docs under `docs/` for any API changes

> The docs site (GitHub Pages) redeploys automatically from `docs/` on push to `master`
> via `.github/workflows/docs.yml` — no manual gh-pages branch step.
