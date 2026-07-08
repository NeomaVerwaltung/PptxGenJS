---
id: integration
title: Integration by Environment
---

PptxGenJS can be used in various JavaScript environments. Choose the integration method below that best suits your project setup.

## Available Distributions

- ES6 Module `dist/pptxgen.es.js`
- CommonJS `dist/pptxgen.cjs.js`
- Browser `dist/pptxgen.min.js`

## Environment Guide

| Environment(s)                                                                                                | Import / Usage                                                                                                         | Notes / Details                                                                                                                                                                                                          |
| ------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **Node.js (Version 18 and higher)**| `import pptxgen from "@neoma/pptxgenjs"`| Automatically uses the appropriate Node.js build based on your project's module type (`package.json#type`). Both ESM and CommonJS formats are fully supported.|
| **Browser Bundlers** (Webpack, Vite, Rollup, Parcel, Browserify, Create React App, Next.js, Angular, Vue CLI) | `import pptxgen from '@neoma/pptxgenjs'`| Your bundler will automatically select the optimized ES Module build (`dist/pptxgen.es.js`). This enables effective tree-shaking to minimize your final bundle size. No extra bundler configuration is typically needed. |
| **Plain Browser (`<script>` tag, no bundler)**| Include the bundled script directly in your HTML: `<script src=".../pptxgen.bundle.js"></script>`| This provides a self-contained build (`dist/pptxgen.bundle.js`) that adds the `PptxGenJS` object to the global `window` scope. Useful for simple scripts or environments without a module bundler.|
| **Web Worker / Service Worker**| `import pptxgen from '@neoma/pptxgenjs'` (Requires a module worker (`type: "module"`) or the use of import maps)| Utilize the ES Module build (`dist/pptxgen.es.js`). Remember that data (like the final presentation `ArrayBuffer`) will need to be transferred back to the main thread using `postMessage`.|
| **Serverless Functions** (AWS Lambda, Cloudflare Workers, etc.)| `import pptxgen from '@neoma/pptxgenjs'` (for ESM runtimes) OR `const pptxgen = require('@neoma/pptxgenjs')` (for CommonJS runtimes) | Bundle your function code using a tool like esbuild or Vite SSR; Be mindful of function size limits and potential cold start impacts from larger dependencies.|
| **Electron (Main Process)**| Same as **Node.js**| In the main Electron process, you have full access to Node.js APIs, including the filesystem, which is useful for directly saving presentation files using the `writeFile()` method.|
| **Electron (Renderer Process)**| Same as **Browser Bundlers**| The renderer process is similar to a browser environment. If `nodeIntegration` is enabled and securely configured, you may also be able to use Node.js filesystem access from the renderer.|

## Integration Demos

Many of the common integration methods have working demos and code available.

### Node.js

A runnable CLI demo covering every feature ships in the repo:

- Source Code: [demos/node](https://github.com/NeomaVerwaltung/PptxGenJS/tree/master/demos/node)
- Run it: `cd demos/node && npm install && npm run demo-all`

### React + Vite / other bundlers

There is no pinned SPA demo (a checked-in app rots between releases). Scaffold a fresh
one and install the library — see the Vite steps in
[TESTING.md](https://github.com/NeomaVerwaltung/PptxGenJS/blob/master/TESTING.md).
The [Quick Start](./quick-start) shows the browser and Node code side by side.

## Troubleshooting

### Webpack

The pptxgenjs library has been tested with several different framework and bundler combinations
including Vite and Webpack. While most projects can simply install @neoma/pptxgenjs and go, there are times
when errors occur.

Here's an example from the latest version of Docusaurus (v3.7) where Webpack (v5) fails during the build process:

```text
[ERROR] Client bundle compiled with errors therefore further build is impossible.
Module build failed: UnhandledSchemeError: Reading from "node:fs" is not handled by plugins (Unhandled scheme).
Webpack supports "data:" and "file:" URIs by default.
You may need an additional plugin to handle "node:" URIs.
```

The error is being caused by the use of the "node:" prefix within "browser" field in pptxgenjs' `package.json` file.

```json
// @neoma/pptxgenjs package.json
{
  "name": "@neoma/pptxgenjs",
  "browser": {
    "fs": false,
    "image-size": false,
    "node:fs": false,
    "node:https": false,
    "os": false,
    "path": false
  }
}
```

Starting in version 4.0.0, node modules are excluded using this format in the library source, so most
bundlers resolve them to empty modules automatically.

### Excluding "node:" from Webpack Builds

Older Webpack setups (e.g. Docusaurus v3 on Webpack v5) may still fail on the `node:` scheme:

```text
UnhandledSchemeError: Reading from "node:fs" is not handled by plugins (Unhandled scheme).
```

Tell Webpack to ignore the `node:` prefix by adding a `NormalModuleReplacementPlugin` in your
`webpack.config.js` (or the equivalent hook in your framework's build config):

```js
const webpack = require("webpack")

module.exports = {
  plugins: [
    new webpack.NormalModuleReplacementPlugin(/^node:/, (resource) => {
      resource.request = resource.request.replace(/^node:/, "")
    }),
  ],
}
```

Webpack now resolves the `node:` items via the `browser` field and builds successfully.
