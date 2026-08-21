# @neo-ma/pptxgenjs-std

Helpers that compose the public `@neo-ma/pptxgenjs` API. Plain functions, no runtime dependency on
the core - every helper takes the object it acts on and calls the same `addX` methods you would.

They live outside the core on purpose: the core is an OOXML emitter bound by upstream API
compatibility, while these are opinions about layout and diagram shapes that should be free to change.

## Install

```sh
npm install @neo-ma/pptxgenjs-std
```

> **Beta.** This package is `0.x` and its API may change between minor versions. It is versioned
> independently of the core - the two version numbers are unrelated.

Requires `@neo-ma/pptxgenjs` 4.2.0 or later - `waterfall` relies on per-series `'transparent'`
colours (4.1.0) and per-point data labels (4.2.0), both silently ignored by older cores.

`@neo-ma/pptxgenjs` is a **peer** dependency, not a regular one: these helpers act on the slide
objects your own presentation created, so a second copy of the core at another version would be
exactly wrong. It is used for types only - nothing here imports it at runtime.
The package ships `dist/index.mjs` and `dist/index.cjs`, so plain Node resolves both without a
bundler. It versions independently from the core; its release tags are `std-vX.Y.Z`.

## `grid` - placement without arithmetic

`addText`/`addShape`/`addChart` take absolute inches, which is where overlapping and off-slide
content comes from. `grid` turns that into cell coordinates and throws on anything out of range.

```ts
import { grid } from '@neo-ma/pptxgenjs-std'

const at = grid() // 12 x 6 cells over a 10 x 5.625in slide, 0.5in margin, 0.2in gutter
slide.addText('Revenue', { ...at(0, 0, 12, 1), fontSize: 32 })
slide.addChart('bar', data, at(0, 1, 6, 5))
slide.addTable(rows, at(6, 1, 6, 5))

at(12, 0) // throws: col 12 span 1 exceeds 12 columns
```

`gridFor(pres, opts)` reads the size from `pres.presLayout`, so it follows `layout` and `defineLayout`.

## `row` / `column` - divide an area

`grid` needs integer spans over a fixed cell count; these take any area and split it, so nested and
uneven layouts do not need a second grid. A number means equal slots, an array means weights.

```ts
import { grid, row, column } from '@neo-ma/pptxgenjs-std'

const [header, body] = column(grid()(0, 0, 12, 6), [1, 4])
const [left, right] = row(body, [1, 2])
```

Output is the same `{ x, y, w, h }` shape as the input, so slots nest.

## `measureText` / `fitText` - text size, computed

Nothing in a `.pptx` records where text wraps. `measureText` works it out, using a browser canvas
when there is one, bundled advance widths otherwise (Calibri, via the metric-compatible OFL font
Carlito), and a flat estimate as a last resort - `source` on the result says which. `fitText` binary
searches it for the largest whole point size that fits a box.

```ts
import { measureText, fitText } from '@neo-ma/pptxgenjs-std'

const { h } = measureText(paragraph, { w: 4, fontSize: 14 })
const { fontSize } = fitText({ w: 4, h: 2 }, headline)
```

`registerFontMetrics(face, { widths })` adds a font; `scripts/extract-font-metrics.mjs` generates the
table from a `.ttf`.

## `paginateTable` / `tableFromHtml` - tables across slides

The core's `autoPage` guesses where rows break from a per-character constant, which is why it ships
`autoPageCharWeight` to tune by hand. These measure each cell instead.

```ts
import { paginateTable } from '@neo-ma/pptxgenjs-std'

paginateTable(pres, rows, { x: 0.5, y: 0.5, w: 9, fontSize: 11, repeatHeaderRows: 1 })
```

They take the presentation, not a slide, because they create the slides. `tableFromHtml` is the same
thing over a table-shaped object - a live `<table>` or anything with `rows`/`cells`.

## `waterfall` - bridge chart

PowerPoint has no waterfall type reachable through ECMA-376 `c:barChart`. This builds the standard
construction: a stacked bar chart with a transparent riser carrying each bar to its starting value,
plus separate increase/decrease series so each side keeps its own color and signed data label.

```ts
import { waterfall } from '@neo-ma/pptxgenjs-std'

waterfall(slide, {
	labels: ['Opening', 'New sales', 'Churn', 'Upsell'],
	values: [1200, 340, -180, 95],
	total: 'Closing',
}, { x: 0.5, y: 0.5, w: 9, h: 4.5, showValue: true })
```

`values` are signed deltas. Everything after `props` is passed through to `addChart`, so titles, axes
and label formatting work as usual - only `barDir`, `barGrouping` and the series colors are fixed.
