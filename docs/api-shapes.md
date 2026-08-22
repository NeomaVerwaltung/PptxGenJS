---
title: Shapes
---

Almost 200 shape types can be added to Slides (see [`ShapeType`](https://github.com/NeomaVerwaltung/PptxGenJS/blob/master/types/index.d.ts) enum).

## Usage

```typescript
// Shapes without text
slide.addShape(pres.ShapeType.rect, { fill: { color: "FF0000" } });
slide.addShape(pres.ShapeType.ellipse, {
  fill: { type: "solid", color: "0088CC" },
});
slide.addShape(pres.ShapeType.line, { line: { color: "FF0000", width: 1 } });

// Shapes with text
slide.addText("ShapeType.rect", {
  shape: pres.ShapeType.rect,
  fill: { color: "FF0000" },
});
slide.addText("ShapeType.ellipse", {
  shape: pres.ShapeType.ellipse,
  fill: { color: "FF0000" },
});
slide.addText("ShapeType.line", {
  shape: pres.ShapeType.line,
  line: { color: "FF0000", width: 1, dashType: "lgDash" },
});
```

## Properties

### Position/Size Props (`PositionProps`)

| Name | Type   | Default | Description            | Possible Values                              |
| :--- | :----- | :------ | :--------------------- | :------------------------------------------- |
| `x`  | number | `1.0`   | hor location (inches)  | 0-n                                          |
| `x`  | string |         | hor location (percent) | 'n%'. (Ex: `{x:'50%'}` middle of the Slide)  |
| `y`  | number | `1.0`   | ver location (inches)  | 0-n                                          |
| `y`  | string |         | ver location (percent) | 'n%'. (Ex: `{y:'50%'}` middle of the Slide)  |
| `w`  | number | `1.0`   | width (inches)         | 0-n                                          |
| `w`  | string |         | width (percent)        | 'n%'. (Ex: `{w:'50%'}` 50% the Slide width)  |
| `h`  | number | `1.0`   | height (inches)        | 0-n                                          |
| `h`  | string |         | height (percent)       | 'n%'. (Ex: `{h:'50%'}` 50% the Slide height) |

### Shape Props (`ShapeProps`)

| Name         | Type                                                                    | Description         | Possible Values                                             |
| :----------- | :---------------------------------------------------------------------- | :------------------ | :---------------------------------------------------------- |
| `align`      | string                                                                  | alignment           | `left` or `center` or `right`. Default: `left`              |
| `fill`       | `ShapeFillProps`       | fill props          | Solid color, gradient, pattern, or picture fill (see below) |
| `flipH`      | boolean                                                                 | flip Horizontal     | `true` or `false`                                           |
| `flipV`      | boolean                                                                 | flip Vertical       | `true` or `false`                                           |
| `hyperlink`  | `HyperlinkProps`  | hyperlink props     | (see type link)                                             |
| `line`       | `ShapeLineProps` | border line props   | (see type link)                                             |
| `rectRadius` | number                                                                  | rounding radius     | 0 to 1. (Ex: 0.5. Only for `pptx.shapes.ROUNDED_RECTANGLE`) |
| `rotate`     | number                                                                  | rotation (degrees)  | -360 to 360. Default: `0`                                   |
| `shadow`     | `ShadowProps`           | shadow props        | (see type link)                                             |
| `shapeName`  | string                                                                  | optional shape name | Ex: "Customer Network Diagram 99"                           |

## Examples

![Shapes with Text Demo](./assets/ex-shape-slide.png)

## Gradient Fills

Set `fill.type` (or `line.type`) to `gradient` and supply at least two color stops. Gradients work anywhere a
`ShapeFillProps` is accepted: shape fills, shape/line outlines, table cell fills, and slide backgrounds.

### Gradient Props (`ShapeGradientProps`)

| Name              | Type                        | Default  | Description                                                          |
| :---------------- | :-------------------------- | :------- | :------------------------------------------------------------------- |
| `stops`           | `ShapeGradientStopProps[]`  |          | **required** - 2 or more `{ color, position, transparency? }` stops   |
| `type`            | string                      | `linear` | `linear` or `radial`                                                 |
| `angle`           | number                      | `90`     | linear angle (degrees, clockwise): 0 = left-to-right, 90 = top-to-bottom |
| `scaled`          | boolean                     | `false`  | whether the linear angle scales with the fill region                 |
| `rotateWithShape` | boolean                     | `true`   | whether the gradient rotates with its shape                          |

Stop `position` and `transparency` are percents (0-100) and are clamped to that range; stops are sorted by
`position` before they are written. Fewer than two stops falls back to a solid fill with a console warning.

```javascript
const slide = pptx.addSlide()

slide.background = { type: 'gradient', gradient: { angle: 45, stops: [{ color: 'FFFFFF', position: 0 }, { color: 'E7E6E6', position: 100 }] } }

slide.addShape(pptx.ShapeType.rect, {
    x: 1, y: 1, w: 4, h: 2,
    fill: { type: 'gradient', gradient: { stops: [{ color: 'FF0000', position: 0 }, { color: '0000FF', position: 100, transparency: 20 }] } },
    line: { type: 'gradient', width: 2, gradient: { type: 'radial', stops: [{ color: '00FF00', position: 0 }, { color: '000000', position: 100 }] } },
})
```

## Pattern Fills

Set `fill.type` to `pattern` with one of the 54 ECMA-376 preset patterns.

```javascript
slide.addShape(pptx.ShapeType.rect, {
    x: 1, y: 1, w: 4, h: 2,
    fill: { type: 'pattern', pattern: { preset: 'diagCross', color: '0000FF', backColor: 'FFFF00' } },
})
```

| Option      | Type   | Default    | Description                        |
| :---------- | :----- | :--------- | :--------------------------------- |
| `preset`    | string |            | **required** - preset pattern name |
| `color`     | Color  | `000000`   | pattern (foreground) color         |
| `backColor` | Color  | `FFFFFF`   | color behind the pattern           |

Presets are the `ST_PresetPatternVal` names: percentage shades (`pct5`…`pct90`), lines (`horz`, `vert`,
`ltHorz`, `dkVert`, `narHorz`, `dashVert`, …), diagonals (`dnDiag`, `ltUpDiag`, `wdDnDiag`, `diagCross`, …),
grids and checks (`smGrid`, `lgCheck`, `dotGrid`, `smConfetti`, …), and textures (`plaid`, `sphere`,
`weave`, `divot`, `shingle`, `wave`, `trellis`, `zigZag`, `horzBrick`, `diagBrick`, `solidDmnd`,
`openDmnd`, `dotDmnd`). An unknown preset falls back to `pct50` with a warning.

## Picture Fills

Set `fill.type` to `image` to fill a shape or table cell with a picture. The image becomes a normal
image relationship, so the same file used twice by path is stored once.

```javascript
slide.addShape(pptx.ShapeType.rect, {
    x: 1, y: 1, w: 4, h: 2,
    fill: { type: 'image', image: { path: '/img/texture.png', sizing: 'tile', scale: 50 } },
})
```

| Option            | Type    | Default   | Description                                                |
| :---------------- | :------ | :-------- | :--------------------------------------------------------- |
| `data`            | string  |           | base64 image data with a mime header (`data` or `path`)     |
| `path`            | string  |           | image path or URL (`data` or `path`)                        |
| `sizing`          | string  | `stretch` | `stretch` scales to the shape, `tile` repeats the image     |
| `scale`           | number  | `100`     | `tile` only: per-tile scale (percent)                       |
| `alignment`       | string  | `tl`      | `tile` only: `tl`,`t`,`tr`,`l`,`ctr`,`r`,`bl`,`b`,`br`     |
| `rotateWithShape` | boolean | `true`    | whether the fill rotates with the shape                     |

A picture fill without usable image data is dropped with a warning rather than written, because a
`a:blip` with no relationship makes PowerPoint report the file as damaged.

Both fill types work on shapes and table cells.

## Group Fill

`fill: { type: 'group' }` emits `a:grpFill` — the shape inherits its parent group's fill instead of
defining its own. It is legal anywhere a fill is, but only has a visible effect on a shape inside a
group shape.

## Effects

`glow`, `shadow`, `reflection` and `softEdge` are joined by two more DrawingML effects. All of them go
into one `a:effectLst`, in the order the schema fixes (blur, fill overlay, glow, shadow, reflection,
soft edge) regardless of the order you set them in.

```typescript
slide.addShape(pptx.ShapeType.rect, {
    x: 1, y: 1, w: 3, h: 2, fill: { color: 'CCCCCC' },
    blur: { radius: 4, grow: false },
    fillOverlay: { blend: 'mult', fill: { color: 'FF0000', transparency: 40 } },
});
```

| Option        | Type   | Description                                                                    |
| :------------ | :----- | :----------------------------------------------------------------------------- |
| `blur`        | object | `radius` (points, required) and `grow` (default `true`) — whether the blur may extend past the shape's bounds |
| `fillOverlay` | object | `blend` (`over`, `mult`, `screen`, `darken`, `lighten`) and `fill` — a second fill blended over the shape's own |

Both `blend` and `fill` are required by the schema, so a `fillOverlay` missing either is dropped
rather than written as an element PowerPoint would refuse to open.

### Preset Shadows

Besides `outer` and `inner`, `shadow.type` accepts `preset` — PowerPoint's twenty built-in shadows.
`preset` is required in that case; without it the shadow is dropped.

```typescript
slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 3, h: 2, shadow: { type: 'preset', preset: 'shdw7', color: '333333' } });
```

Preset names are `shdw1` through `shdw20`. `offset` and `angle` still apply; `blur` and `opacity` do
not (the preset defines them).

### Composed Effect Graphs (`effectDag`)

`a:effectLst` and `a:effectDag` are alternatives in the OOXML schema. Setting `effectDag` emits the
same effects inside an `a:effectDag` instead of an `a:effectLst`:

```typescript
slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 3, h: 2, glow: { size: 6, color: 'FFFF00', opacity: 0.6 }, effectDag: { type: 'sib' } });
```

| Option | Type   | Default | Description                                                    |
| :----- | :----- | :------ | :------------------------------------------------------------- |
| `type` | string | `sib`   | `sib` (applies to siblings) or `tree` (applies to the subtree)  |

The graph is flat: nested `a:cont` containers and named `a:effect` references are not emitted.
PowerPoint's UI never produces them, so there is nothing to round-trip against.

## Line Properties

`line` accepts the full `CT_LineProperties` model (ECMA-376 §20.1.2.1). Everything below is optional
and omitted when unset.

```javascript
slide.addShape(pptx.ShapeType.line, {
    x: 1, y: 1, w: 4, h: 0,
    line: {
        color: '0000FF', width: 3,
        compound: 'thickThin',
        join: 'miter', miterLimit: 400,
        beginArrowType: 'arrow', beginArrowSize: { width: 'lg', length: 'lg' },
        endArrowType: 'triangle', endArrowSize: { width: 'sm' },
    },
})
```

| Option           | Type   | Default | Description                                                     |
| :--------------- | :----- | :------ | :-------------------------------------------------------------- |
| `compound`       | string | `sng`   | `sng`, `dbl`, `thickThin`, `thinThick`, `tri`                   |
| `join`           | string | `round` | how segments meet at a corner: `round`, `bevel`, `miter`        |
| `miterLimit`     | number | `800`   | `join: 'miter'` only — percent of line width                     |
| `customDash`     | array  |         | `[{ dash, space }]` percentages of line width; replaces `dashType` |
| `beginArrowSize` | object |         | `{ width, length }`, each `sm`/`med`/`lg`                        |
| `endArrowSize`   | object |         | `{ width, length }`                                             |

`customDash` and `dashType` are the same choice in the schema, so a custom pattern replaces the preset
one rather than appearing alongside it. An invalid `compound`, `join`, or dash stop is dropped with a
warning; usable stops in a partly-invalid pattern still apply.

These options also drive `a:uLn` via the text `underlineLine` option — see the text docs.

## Locks and Non-Visual Properties

Every object accepts editing locks and the non-visual properties PowerPoint shows in the selection and
alt-text panes. All are optional; the locks the library has always written (`noChangeAspect` on
pictures, `noGrp` on table frames) are unchanged and caller locks are added to them.

```javascript
slide.addShape(pptx.ShapeType.rect, {
    x: 1, y: 1, w: 3, h: 1,
    objectName: 'Logo frame', title: 'Company logo', hidden: false,
    lock: { noMove: true, noResize: true, noTextEdit: true },
})
```

| Option   | Type   | Description                                                    |
| :------- | :----- | :------------------------------------------------------------- |
| `title`  | string | alt-text **title**, distinct from `altText` (the description)   |
| `hidden` | boolean | hide the object; it stays in the file and in the selection pane |
| `lock`   | object | editing locks — see below                                      |

Locks: `noGroup`, `noSelect`, `noRotate`, `noChangeAspect`, `noMove`, `noResize`, `noEditPoints`,
`noAdjustHandles`, `noChangeArrowheads`, `noChangeShapeType`, `noTextEdit`, plus `noCrop` and
`preferRelativeResize` for pictures.

Not every lock applies to every object — `a:spLocks`, `a:picLocks`, and `a:graphicFrameLocks` permit
different sets. A lock the object's element does not accept is dropped with a warning rather than
written, since an unexpected attribute is a schema violation.
