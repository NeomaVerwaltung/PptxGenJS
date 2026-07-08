# PptxGenJS Testing Guide

This document outlines how to manually test PptxGenJS across supported platforms and environments prior to release.

> ✅ The automated suite (`npm run check`) covers lint, typechecks, and unit/e2e/snapshot tests. The manual steps below validate real runtimes and bundlers.

Testing Steps

1. Run `npm run ship`
2. Execute tests from each section below

## 🧪 Test Suites Overview

| Platform        | Tooling              | Status |
| --------------- | -------------------- | ------ |
| Node.js         | Native CLI           | ✅      |
| Vite/TypeScript | Modern front-end SPA | ✅      |
| Webpack         | SharePoint Framework | ✅      |

---

## 📦 Node.js Tests

**Purpose:** Validate functionality of CommonJS module in pure Node environments.

### CLI Tests

Run the following test commands:

```bash
cd demos/node
npm run demo
npm run demo-all
```

1. Confirm console output and exported PPTX files are correct.

### Stream Test

```bash
npm run demo-stream
```

1. Confirm stream download PPTX file is correct.
2. Open the [Stream URL](http://192.168.254.x:3000/) on iPhone & test.

---

## ⚛️ Vite + TypeScript Tests

**Purpose:** Validate integration in modern front-end SPA toolchains (Vite, TypeScript, React-compatible).

There is no checked-in Vite app (a pinned SPA rots between releases). Scaffold a fresh one and link this repo:

```bash
npm create vite@latest pptxgenjs-vite-test -- --template react-ts
cd pptxgenjs-vite-test
npm install
npm install /path/to/this/repo   # installs @neoma/pptxgenjs from the local checkout
npm run dev
```

1. In a component, `import pptxgen from "@neoma/pptxgenjs"` and export a test slide.
2. Verify IntelliSense autocompletes e.g. `pptxgen.ChartType.` (types resolve).
3. Export a PowerPoint file and open it to verify it renders correctly.
4. Delete the scaffold when done — nothing to maintain.

---

## 🏁 Test Completion Checklist

| Dist File         | Test       | Tested Via             | Result |
| ----------------- | ---------- | ---------------------- | ------ |
| pptxgen.es.js     | Webpack 4  | SPFx (v1.16.1) project | ✅?🟡    |
| pptxgen.es.js     | Webpack 5  | SPFx (v1.19.1) project | ✅?🟡    |
| pptxgen.es.js     | Rollup 4   | Vite (v6) scaffold     | ✅?🟡    |
| pptxgen.cjs.js    | Node/CJS   | Node demo              | ✅?🟡    |
