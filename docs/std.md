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

::: warning Beta
`@neo-ma/pptxgenjs-std` is `0.x` and its API may still change between minor versions. It is
versioned independently of the core: a core release never forces a std release, and the two version
numbers are unrelated.
:::

**Requires `@neo-ma/pptxgenjs` 4.2.0 or later.** `waterfall` builds its invisible riser from a
per-series `'transparent'` colour (4.1.0) and its signed labels from per-point data labels (4.2.0);
on an older core both are ignored and the chart renders wrong rather than failing.

`@neo-ma/pptxgenjs` is a **peer** dependency rather than a regular one: the helpers act on the slide
objects your own presentation created, so a second copy of the core resolved at a different version
would be exactly wrong. It is used for types only - nothing is imported from it at runtime. The package ships `.mjs` and `.cjs` builds, so plain Node resolves both without a bundler.
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
| `@neo-ma/pptxgenjs-std/layout` | `grid`, `gridFor`, `row`, `column` |
| `@neo-ma/pptxgenjs-std/charts` | `waterfall` |
| `@neo-ma/pptxgenjs-std/text` | `measureText`, `fitText`, `registerFontMetrics` |
| `@neo-ma/pptxgenjs-std/tables` | `paginateTable`, `tableFromHtml`, `cssColorToHex` |

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

### `row` and `column` - divide an area

`grid` needs integer spans over a fixed cell count. `row` and `column` take any area and divide it,
so nesting and uneven splits do not need a second grid:

```typescript
import { grid, row, column } from "@neo-ma/pptxgenjs-std";

const at = grid();
const [header, body] = column(at(0, 0, 12, 6), [1, 4]);
const [left, right] = row(body, [1, 2]);

slide.addText("Q3 review", { ...header, fontSize: 28 });
slide.addChart("bar", data, left);
slide.addTable(rows, right);
```

A number means equal slots, an array means weights - `row(area, 3)` and `row(area, [1, 1, 1])` are
the same call. The third argument is the gap in inches, defaulting to `0.2`. Output is the same
`{ x, y, w, h }` shape as the input, so slots nest without special handling, and the cross axis is
left untouched. Out-of-range input throws:

```typescript
row(area, 0); // Error: row: slot count must be an integer >= 1 (got 0)
row(area, [1, 0]); // Error: row: every weight must be > 0 (got [1, 0])
row({ x: 0, y: 0, w: 4, h: 2 }, 5, 1); // Error: row: gap 1 leaves no room across 4
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

## `measureText` - how big text actually is

PowerPoint decides where text wraps; nothing in a `.pptx` records it. `measureText` computes it up
front, so a caller can size a box to its content instead of guessing.

```typescript
import { measureText } from "@neo-ma/pptxgenjs-std";

const { w, h, lines, source } = measureText(paragraph, { w: 4, fontSize: 14 });
slide.addText(paragraph, { x: 1, y: 1, w: 4, h });
```

Every value is in inches, except `fontSize`, which is in points like the rest of the API. Omit `w`
to measure the text as one unbroken line; `\n` always splits.

### Where the numbers come from

`source` on the result says which of three tiers answered, best first:

| `source` | When | Accuracy |
| --- | --- | --- |
| `'canvas'` | a browser DOM is present | exact for the font the browser has |
| `'metrics'` | the font's advance widths are bundled or registered | exact advance widths, no kerning |
| `'estimate'` | neither | a flat per-character guess |

Bundled metrics cover **Calibri** in all four styles, by way of
[Carlito](https://github.com/googlefonts/carlito) - an OFL font that is metric-compatible with it,
so the widths are Calibri's without shipping Calibri. Check `source === 'estimate'` when accuracy
matters; that is the tier that behaves like the core's own auto-paging guess.

::: warning
PowerPoint substitutes a different font when the viewer's machine lacks yours, and re-wraps with the
substitute's metrics. Treat any measurement as accurate for *this* font, not guaranteed on every
machine. Pair it with `fit: 'shrink'` when overflow must not happen.
:::

### `registerFontMetrics` - teach it another font

```typescript
import { registerFontMetrics } from "@neo-ma/pptxgenjs-std";

registerFontMetrics("Aptos", { widths: { " ": 0.22, a: 0.5 }, ascent: 0.94, descent: -0.27 });
registerFontMetrics("Aptos Bold", { widths: { " ": 0.22, a: 0.54 } });
```

Widths are in ems, keyed by single character or codepoint. Register a style under its full name plus
the suffix (`'Aptos Bold'`, `'Aptos Italic'`, `'Aptos Bold Italic'`) so `bold: true` finds it; a
missing style falls back to the regular weight rather than to the estimate.

`scripts/extract-font-metrics.mjs` in the repository generates a table from a `.ttf`, if hand-listing
widths is not practical.

### Options

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `fontFace` | string | `'Calibri'` | Typeface name, matched case-insensitively |
| `fontSize` | number | `12` | Font size in points |
| `bold` / `italic` | boolean | `false` | Style, used to pick the metrics |
| `w` | number | - | Wrap width in inches; omit for a single line |
| `lineSpacingMultiple` | number | `1` | Line spacing as a multiple of single |

## `fitText` - largest size that fits

`fit: 'shrink'` makes PowerPoint shrink text at render time, but the chosen size is not readable
back, so nothing else on the slide can react to it. `fitText` computes it instead:

```typescript
import { fitText } from "@neo-ma/pptxgenjs-std";

const { fontSize, h, overflows } = fitText({ w: 4, h: 2 }, headline, { max: 40 });
slide.addText(headline, { x: 1, y: 1, w: 4, h: 2, fontSize });
```

Whole points only, which is what the PowerPoint UI offers. `overflows` is `true` when even `min`
does not fit - the text is too long for the box at any size in range, and you get `min` back rather
than a silent overflow. `margin` (a number or a TRBL tuple, inches) is subtracted from the area
first, for text boxes with inset.

## `paginateTable` - a table across slides, measured

The core's `autoPage` decides where rows break from a per-character constant, which is why it ships
`autoPageCharWeight` for callers to tune by hand. `paginateTable` measures each cell with
`measureText` instead, so the row heights it adds up are the ones PowerPoint will lay out.

```typescript
import { paginateTable } from "@neo-ma/pptxgenjs-std";

const { slides, rowsPerSlide, estimated } = paginateTable(pres, rows, {
    x: 0.5,
    y: 0.5,
    w: 9,
    fontSize: 11,
    repeatHeaderRows: 1,
});
```

It creates the slides itself, so it takes the presentation rather than a slide. Everything other
than the paging options is passed straight through to `addTable`, once per slide.

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `repeatHeaderRows` | number | `0` | Leading rows repeated at the top of every slide |
| `continueY` | number | the initial `y` | `y` for tables on slides after the first |
| `masterName` | string | - | Master applied to every slide it creates |
| `bottomMargin` | number | `0.5` | Usable-area bottom, from the slide bottom, inches |
| `slideHeight` | number | from `presLayout` | Slide height override, inches |

`estimated` on the result is `true` when any cell fell back to the guessing tier - the signal that
the page breaks are approximate for the same reason the core's are.

## `tableFromHtml` - an HTML table, measured

The measured counterpart to the core's `tableToSlides`. It takes the table as an object rather than
an element id, so it is not tied to a live document and can be driven by anything table-shaped.

```typescript
import { tableFromHtml } from "@neo-ma/pptxgenjs-std";

tableFromHtml(pres, document.getElementById("report"), { x: 0.5, y: 0.5, w: 9, fontSize: 10 });
```

Cell fill, colour, weight, size, face and alignment come from `window.getComputedStyle` when a DOM
is present. Pass `styleOf` to supply them from anything else - it receives each cell and returns the
formatting to apply. `<th>` cells are bold, and leading rows made entirely of `<th>` become repeated
headers unless `detectHeaderRows: false` or an explicit `repeatHeaderRows` says otherwise. `colspan`
and `rowspan` carry over.

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
import type {
    GridProps,
    GridArea,
    Slots,
    WaterfallProps,
    ChartSlide,
    MeasureProps,
    Measurement,
    FitTextProps,
    PaginateTableProps,
    TableFromHtmlProps,
} from "@neo-ma/pptxgenjs-std";
```

Types are exported from their category subpath as well, so `GridProps` is also available from
`@neo-ma/pptxgenjs-std/layout`.
