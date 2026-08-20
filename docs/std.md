---
title: Std Library
---

`@neo-ma/pptxgenjs-std` is a separate package of helpers that compose the public API. Nothing in it
reaches into the presentation internals - each helper takes the object it acts on and calls the same
`addText`/`addShape`/`addChart` methods you would.

It ships separately from the core because it is opinionated: the core is an OOXML emitter bound by
upstream API compatibility, while these helpers are free to change.

## Install

```bash
npm install @neo-ma/pptxgenjs @neo-ma/pptxgenjs-std
```

`@neo-ma/pptxgenjs` is a peer dependency used for types only - the helpers never import it at
runtime. The package ships `.mjs` and `.cjs` builds, so plain Node resolves both without a bundler.
It versions independently from the core.

Helpers are grouped into categories, each reachable as its own subpath. Import from the root barrel
or from a category - the same functions either way:

```typescript
import { grid, waterfall } from "@neo-ma/pptxgenjs-std";
import { grid } from "@neo-ma/pptxgenjs-std/layout";
import { waterfall } from "@neo-ma/pptxgenjs-std/charts";
```

| Subpath | Contents |
| --- | --- |
| `@neo-ma/pptxgenjs-std` | every helper, re-exported |
| `@neo-ma/pptxgenjs-std/layout` | `grid`, `gridFor` |
| `@neo-ma/pptxgenjs-std/charts` | `waterfall` |

## `grid` - placement without arithmetic

`addText`, `addShape` and `addChart` take absolute inches, which is where overlapping and off-slide
content comes from. `grid` returns a placement function: give it a cell, get an `{ x, y, w, h }`
object ready to spread into any options object.

```typescript
import pptxgen from "@neo-ma/pptxgenjs";
import { grid } from "@neo-ma/pptxgenjs-std";

const pres = new pptxgen();
const slide = pres.addSlide();
const at = grid(); // 12 x 6 cells over a 10 x 5.625in slide

slide.addText("Revenue by region", { ...at(0, 0, 12, 1), fontSize: 32, bold: true });
slide.addChart("bar", data, at(0, 1, 6, 5));
slide.addTable(rows, at(6, 1, 6, 5));
```

The call signature is `at(col, row, colSpan?, rowSpan?)`, zero-indexed, spans defaulting to `1`.
Anything out of range **throws** rather than placing content past the slide edge:

```typescript
at(12, 0); // Error: grid: col 12 span 1 exceeds 12 columns
at(10, 0, 3); // Error: grid: col 10 span 3 exceeds 12 columns
at(0, 0, 0); // Error: grid: spans must be >= 1
```

### Options

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `w` | number | `10` | Slide width in inches (`LAYOUT_16x9`) |
| `h` | number | `5.625` | Slide height in inches (`LAYOUT_16x9`) |
| `cols` | number | `12` | Column count |
| `rows` | number | `6` | Row count |
| `gutter` | number | `0.2` | Space between cells, inches |
| `margin` | number | `0.5` | Space outside the grid, inches |

```typescript
const at = grid({ cols: 3, rows: 3, gutter: 0.15, margin: 0.4 });
```

### `gridFor` - follow the presentation layout

`gridFor(pres, opts)` reads the size from `pres.presLayout`, so it tracks `layout` and
`defineLayout` instead of hard-coding inches:

```typescript
import { gridFor } from "@neo-ma/pptxgenjs-std";

const pres = new pptxgen();
pres.layout = "LAYOUT_4x3";
const at = gridFor(pres, { cols: 2, rows: 2 });
```

## `waterfall` - bridge chart

PowerPoint has no waterfall chart type reachable through ECMA-376 `c:barChart`. `waterfall` builds
the standard construction instead: a stacked bar chart whose first series is transparent and carries
each bar up to its starting value, plus separate increase and decrease series so each side keeps its
own colour and a signed data label.

```typescript
import { waterfall } from "@neo-ma/pptxgenjs-std";

waterfall(
    slide,
    {
        labels: ["Opening", "New sales", "Churn", "Upsell"],
        values: [1200, 340, -180, 95],
        total: "Closing",
    },
    { x: 0.5, y: 0.5, w: 9, h: 4.5, showValue: true }
);
```

`values` are signed deltas: each bar spans from the running total to running + value. A negative
value hangs below its starting point. `total` appends a final bar spanning from zero to the closing
total, labelled with the string you pass.

### Properties

| Property | Type | Default | Description |
| --- | --- | --- | --- |
| `labels` | string[] | required | Category label per delta |
| `values` | number[] | required | Signed change per category; must match `labels` in length |
| `total` | string | - | Label for a final bar spanning zero to the closing total; omit for no total |
| `increaseColor` | string | `'2E7D32'` | Bar colour for positive deltas |
| `decreaseColor` | string | `'C62828'` | Bar colour for negative deltas |
| `increaseName` | string | `'Increase'` | Legend entry for positive deltas |
| `decreaseName` | string | `'Decrease'` | Legend entry for negative deltas |

The third argument is passed straight through to `addChart`, so titles, axes, legends and label
formatting work as usual. Only `barDir`, `barGrouping` and the series colours are fixed by the
construction.

Malformed input throws rather than emitting a chart PowerPoint cannot read:

```typescript
waterfall(slide, { labels: ["A"], values: [1, 2] }); // Error: waterfall: labels (1) and values (2) must be the same length
waterfall(slide, { labels: [], values: [] }); // Error: waterfall: at least one value is required
waterfall(slide, { labels: ["A"], values: [NaN] }); // Error: waterfall: values[0] is not a finite number
```

## Combining them

`grid` returns exactly the shape `waterfall`'s options argument expects, so they compose directly:

```typescript
const at = grid();
slide.addText("Q3 bridge", { ...at(0, 0, 12, 1), fontSize: 28 });
waterfall(slide, { labels, values, total: "Net" }, { ...at(0, 1, 12, 5), showValue: true });
```

## Types

The helpers do not import the `Slide` class - they declare the slice of it they need. `waterfall`
accepts any object with a conforming `addChart`, and `gridFor` any object with a `presLayout`. Both
are generic over the value you pass, so `waterfall` returns your slide unchanged in type.

```typescript
import type { GridProps, GridArea, WaterfallProps, ChartSlide } from "@neo-ma/pptxgenjs-std";
```

Types are exported from their category subpath as well, so `GridProps` is also available from
`@neo-ma/pptxgenjs-std/layout`.
