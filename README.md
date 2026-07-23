# PptxGenJS

![PptxGenJS Sample Slides](https://raw.githubusercontent.com/gitbrent/PptxGenJS/gh-pages/img/readme_banner.png)

![GitHub Repo stars](https://img.shields.io/github/stars/NeomaVerwaltung/PptxGenJS?style=flat-square)
![GitHub License](https://img.shields.io/github/license/NeomaVerwaltung/PptxGenJS?style=flat-square)

PptxGenJS is a JavaScript library for generating PowerPoint presentations programmatically. It runs in Node.js, browsers, and bundler-based environments (React, Angular, Vite, Webpack, Electron), and requires no PowerPoint installation or license.

This repository contains the NEOMA-maintained distribution, published as `@neoma/pptxgenjs` and maintained by [NEOMA GmbH](https://github.com/NeomaVerwaltung).

## Overview

The library produces standards-compliant Open Office XML (OOXML) files compatible with:

- Microsoft PowerPoint
- Apple Keynote
- LibreOffice Impress
- Google Slides (via import)

## Features

### Platform support

- All major modern browsers, desktop and mobile
- Node.js, React, Angular, Vite, and Electron
- Serverless and edge runtimes (AWS Lambda, Vercel, Cloudflare Workers)

### Slide content

- All major slide objects: text, tables, shapes, images, charts, and media
- Custom Slide Masters for consistent corporate branding
- SVG images, animated GIFs, YouTube embeds, right-to-left text, and Asian fonts

### Developer experience

- Minimal API: a complete presentation can be created in four statements
- Full TypeScript definitions for autocomplete and inline documentation
- More than 75 demo slides covering every feature and usage pattern

### Output options

- Direct `.pptx` download from the browser with correct MIME handling
- Export as base64, Blob, Buffer, or Node stream
- Compression and additional output options for production use

### Production readiness

The NEOMA distribution is hardened for production use:

- **Strict null-safety** — the entire source compiles under TypeScript `strictNullChecks` with zero non-null assertions (`!`) and zero unchecked `as` casts; null-safety is enforced through explicit types, guard clauses, and validated defaults
- **Continuous integration** — unit, end-to-end, and golden-file OOXML snapshot tests run on Node.js 20 and 22
- **Security** — no known vulnerabilities in the published package, [CodeQL](https://github.com/NeomaVerwaltung/PptxGenJS/security/code-scanning) scanning, and a published [security policy](SECURITY.md)
- **Requires Node.js 20 or newer** — see [CONTRIBUTING.md](CONTRIBUTING.md) for build and test instructions

## Installation

```bash
npm install @neoma/pptxgenjs
```

```bash
yarn add @neoma/pptxgenjs
```

## Compatibility

PptxGenJS ships dual ESM and CJS builds with zero runtime dependencies. Bundlers select the correct build automatically via the `exports` field in `package.json`.

Supported environments:

- **Node.js** — backend scripts, APIs, and CLI tools
- **React / Angular / Vite / Webpack** — import directly; no additional configuration required
- **Electron** — native applications with full filesystem access
- **Browser (vanilla JavaScript)** — web applications with direct download support
- **Serverless / edge functions** — AWS Lambda, Vercel, Cloudflare Workers, and similar platforms

Builds provided:

- **CommonJS**: [`dist/pptxgen.cjs.js`](./dist/pptxgen.cjs.js)
- **ES Module**: [`dist/pptxgen.es.js`](./dist/pptxgen.es.js)

## Quick start

A presentation is created in four steps: instantiate, add a slide, add content, save.

### Angular/React, ES6, TypeScript

```typescript
import pptxgen from "@neoma/pptxgenjs";

// 1. Create a new Presentation
let pres = new pptxgen();

// 2. Add a Slide
let slide = pres.addSlide();

// 3. Add one or more objects (Tables, Shapes, Images, Text and Media) to the Slide
let textboxText = "Hello World from PptxGenJS!";
let textboxOpts = { x: 1, y: 1, color: "363636" };
slide.addText(textboxText, textboxOpts);

// 4. Save the Presentation
pres.writeFile();
```

### Script/Web browser

```javascript
// 1. Create a new Presentation
let pres = new PptxGenJS();

// 2. Add a Slide
let slide = pres.addSlide();

// 3. Add one or more objects (Tables, Shapes, Images, Text and Media) to the Slide
let textboxText = "Hello World from PptxGenJS!";
let textboxOpts = { x: 1, y: 1, color: "363636" };
slide.addText(textboxText, textboxOpts);

// 4. Save the Presentation
pres.writeFile();
```

## HTML table conversion

The `tableToSlides` method converts an HTML `<table>` element into one or more formatted slides:

```javascript
let pptx = new pptxgen();
pptx.tableToSlides("tableElementId");
pptx.writeFile({ fileName: "html2pptx-demo.pptx" });
```

Typical use cases include exporting dashboards, data reports, and tabular content from web applications or BI tools.

See the [HTML-to-PowerPoint guide](https://neomaverwaltung.github.io/PptxGenJS/html-to-powerpoint) for details.

## Documentation

The complete API reference, tutorials, and integration guides are available at [https://neomaverwaltung.github.io/PptxGenJS](https://neomaverwaltung.github.io/PptxGenJS).

The documentation source is plain Markdown under [`docs/`](docs) and builds with VitePress (`npm run docs:dev` to preview locally). It deploys to GitHub Pages automatically on push to `master`.

## Issues and support

Report defects and feature requests on the [issue tracker](https://github.com/NeomaVerwaltung/PptxGenJS/issues/new), or [submit a pull request](https://github.com/NeomaVerwaltung/PptxGenJS/pulls). When reporting an issue, include a code snippet or a link that demonstrates the problem — this [jsFiddle](https://jsfiddle.net/gitbrent/L1uctxm0/) is preconfigured with the latest PptxGenJS build.

Additional resources:

- The Node demo (`demos/node`) contains working examples of every library feature
- [Questions tagged `PptxGenJS` on Stack Overflow](https://stackoverflow.com/questions/tagged/pptxgenjs?sort=votes&pageSize=50); tag new questions `pptxgenjs`

## Contributors

NEOMA thanks all contributors to the original project and this distribution.

Notable contributions:

- [Dzmitry Dulko](https://github.com/DzmitryDulko) — initial NPM publication
- [Michal Kacerovský](https://github.com/kajda90) — Master Slide layouts and chart expertise
- [Connor Bowman](https://github.com/conbow) — placeholder support
- [Reima Frgos](https://github.com/ReimaFrgos) — chart and general functionality patches
- [Matt King](https://github.com/kyrrigle) — chart expertise
- [Mike Wilcox](https://github.com/clubajax) — chart expertise
- [Joonas](https://github.com/wyozi) — [react-pptx](https://github.com/wyozi/react-pptx)

PowerPoint shape definitions and portions of the XML generation are derived from the [Officegen Project](https://github.com/Ziv-Barber/officegen).

## License

Copyright &copy; 2015-present [Brent Ely](https://github.com/gitbrent/), &copy; 2026-present [NEOMA GmbH](https://github.com/NeomaVerwaltung)

Licensed under the [MIT License](https://github.com/NeomaVerwaltung/PptxGenJS/blob/master/LICENSE).
