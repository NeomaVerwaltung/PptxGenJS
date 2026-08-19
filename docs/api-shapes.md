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
| `fill`       | `ShapeFillProps`       | fill props          | Fill color/transparency props, or a gradient (see below)    |
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
