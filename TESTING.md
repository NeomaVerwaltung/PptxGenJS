# PptxGenJS Testing Guide

This document outlines how to manually test PptxGenJS across supported platforms and environments prior to release.

> ✅ Run these tests to ensure compatibility with major bundlers, runtimes, and front-end frameworks.

Config Notes

> ⚠️ Disable VPN on the server machine, otherwise, clients using the local IP address cannot connect.

Testing Steps

1. Run `npm run ship`
2. Execute tests from each section below

## 🧪 Test Suites Overview

| Platform        | Tooling              | Status |
| --------------- | -------------------- | ------ |
| Browser         | Standalone HTML demo | ✅      |
| Node.js         | Native CLI           | ✅      |
| Web Worker      | JS Worker demo       | ✅      |
| Vite/TypeScript | Modern front-end SPA | ✅      |
| Webpack         | SharePoint Framework | ✅      |

---

## 🌐 Browser Tests

**Purpose:** Validate browser compatibility using the standalone bundle as script.

### Desktop & Mobile Browsers

Run local test server:

```bash
cd demos
npm install   # installs express (demo-only; no longer in the library's deps)
node browser_server.mjs
```

1. Open the [Demo Page](http://localhost:8000/browser/index.html).
2. In DevTools, confirm the latest `pptxgen.bundle.js` is loaded (`Sources` tab).
3. Run all UI-driven demos and verify demo presentation render correctly.
4. Open the [Demo Page](http://192.168.254.x:8000/browser/index.html) on iPhone & test.

### Web Worker API

1. Open the [Web Worker Demo Page](localhost:8000/browser/worker_test.html).
2. Note: Use Chrome (Safari *will not work*)
3. Run the test; verify result & library version

### Microsoft 365 Check

1. Upload the full demo output from above to M365/Office/OneDrive.
2. Use web viewer to validate file

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
| pptxgen.es.js     | Rollup 4   | Vite (v6) demo         | ✅?🟡    |
| pptxgen.es.js     | Webworkers | worker_test demo       | ✅?🟡    |
| pptxgen.cjs.js    | Node/CJS   | Node demo              | ✅?🟡    |
| pptxgen.bundle.js | Script     | Browser demo (desktop) | ✅?🟡    |
| pptxgen.bundle.js | Script     | Browser demo (iOS)     | ✅?🟡    |
