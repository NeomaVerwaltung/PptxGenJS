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
