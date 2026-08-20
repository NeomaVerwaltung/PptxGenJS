# @neo-ma/pptxgenjs-std

Helpers that compose the public `@neo-ma/pptxgenjs` API. Plain functions, no runtime dependency on
the core - every helper takes the object it acts on and calls the same `addX` methods you would.

They live outside the core on purpose: the core is an OOXML emitter bound by upstream API
compatibility, while these are opinions about layout and diagram shapes that should be free to change.

## Install

```sh
npm install @neo-ma/pptxgenjs-std
```

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
