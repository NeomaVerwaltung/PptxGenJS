# Testing Guide

This document describes how to manually verify PptxGenJS across supported platforms and environments prior to a release.

The automated suite (`npm run check`) covers linting, type checks, and unit/e2e/snapshot tests. The manual steps below validate real runtimes and bundlers.

Procedure:

1. Run `npm run ship`.
2. Execute the tests in each section below.

## Test suite overview

| Platform        | Tooling              |
| --------------- | -------------------- |
| Node.js         | Native CLI           |
| Vite/TypeScript | Modern front-end SPA |
| Webpack         | SharePoint Framework |

## Node.js tests

Purpose: validate the CommonJS build in a pure Node environment.

### CLI tests

```bash
cd demos/node
npm run demo
npm run demo-all
```

1. Confirm the console output and the exported PPTX files are correct.

### Stream test

```bash
npm run demo-stream
```

1. Confirm the streamed PPTX download is correct.
2. Open the [stream URL](http://192.168.254.x:3000/) on a mobile device and verify the download.

## Vite + TypeScript tests

Purpose: validate integration with modern front-end toolchains (Vite, TypeScript, React-compatible).

No Vite application is checked in, as a pinned SPA becomes outdated between releases. Scaffold a fresh application and link this repository:

```bash
npm create vite@latest pptxgenjs-vite-test -- --template react-ts
cd pptxgenjs-vite-test
npm install
npm install /path/to/this/repo   # installs @neo-ma/pptxgenjs from the local checkout
npm run dev
```

1. In a component, `import pptxgen from "@neo-ma/pptxgenjs"` and export a test slide.
2. Verify that IntelliSense autocompletes, for example, `pptxgen.ChartType.` (types resolve correctly).
3. Export a PowerPoint file and confirm it renders correctly.
4. Delete the scaffold when finished; nothing is retained.

## Completion checklist

Record the result of each test before release:

| Dist File      | Test      | Tested Via             | Result |
| -------------- | --------- | ---------------------- | ------ |
| pptxgen.es.js  | Webpack 4 | SPFx (v1.16.1) project |        |
| pptxgen.es.js  | Webpack 5 | SPFx (v1.19.1) project |        |
| pptxgen.es.js  | Rollup 4  | Vite (v6) scaffold     |        |
| pptxgen.cjs.js | Node/CJS  | Node demo              |        |
