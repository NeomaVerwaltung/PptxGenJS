---
title: Slide Properties and Methods
---

## Slide Properties

| Option               | Type                                                                 | Default  | Description                      | Possible Values                                                                                                                                                    |
| :------------------- | :------------------------------------------------------------------- | :------- | :------------------------------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `background`         | `BackgroundProps`                                                    | `FFFFFF` | background color/images          | add background color or image `DataOrPathProps` and/or `ShapeFillProps` |
| `color`              | string                                                               | `000000` | default text color               | hex color or [scheme color](./shapes-and-schemes.md).                                                                                                              |
| `hidden`             | boolean                                                              | `false`  | whether slide is hidden          | Ex: `slide.hidden = true`                                                                                                                                          |
| `newAutoPagedSlides` | PresSlide[]                                                          |          | all slides created by autopaging | Contains slides automatically created when content (e.g. a table) overflows the current slide using autoPage:true                                                  |
| `slideNumber`        | `SlideNumberProps` |          | slide number props               | (see examples below)                                                                                                                                               |

## Full Examples

### Example: Background/Foreground

```typescript
// EX: Use several methods to set a background
slide.background = { color: "F1F1F1" }; // Solid color
slide.background = { color: "FF3399", transparency: 50 }; // hex fill color with transparency of 50%
slide.background = { data: "image/png;base64,ABC[...]123" }; // image: base64 data
slide.background = { path: "https://some.url/image.jpg" }; // image: url
```

```typescript
// EX: Set slide default font color
slide.color = "696969";
```

### Example: Slide Number

```typescript
// EX: Add a Slide Number at a given location
slide.slideNumber = { x: 1.0, y: "90%" };

// EX: Styled Slide Numbers
slide.slideNumber = { x: 1.0, y: "95%", fontFace: "Courier", fontSize: 32, color: "CF0101" };
```

## Slide Transitions

Set `transition` when adding the slide, or assign `slide.transition` afterwards.

```typescript
pptx.addSlide({ transition: { type: 'wipe', direction: 'left', speed: 'slow' } });

const slide = pptx.addSlide();
slide.transition = { type: 'morph', duration: 1200, advClick: false, advTm: 5000 };
```

### Transition Props (`SlideTransitionProps`)

| Option      | Type                | Default | Description                                                        |
| :---------- | :------------------ | :------ | :----------------------------------------------------------------- |
| `type`      | string              |         | **required** - transition effect (see below)                       |
| `direction` | string              |         | direction, if the effect takes one - OOXML tokens or friendly names |
| `orient`    | `horz` \| `vert`    | `horz`  | `split` only: which axis splits                                    |
| `speed`     | `slow`\|`med`\|`fast` |       | coarse speed - ignored when `duration` is set                      |
| `duration`  | number              |         | duration in milliseconds (PowerPoint 2010+)                        |
| `advClick`  | boolean             | `true`  | advance the slide on mouse click                                   |
| `advTm`     | number              |         | advance automatically after this many milliseconds                 |
| `spokes`    | 1\|2\|3\|4\|8       | `4`     | `wheel` only: number of spokes                                     |
| `thruBlk`   | boolean             | `false` | `fade` and `cut` only: transition through black                    |

**Base effects** (every consumer): `blinds`, `checker`, `circle`, `comb`, `cover`, `cut`, `diamond`,
`dissolve`, `fade`, `newsflash`, `none`, `plus`, `pull`, `push`, `random`, `randomBar`, `split`,
`strips`, `wedge`, `wheel`, `wipe`, `zoom`.

**PowerPoint 2010+ effects**: `conveyor`, `doors`, `ferris`, `flash`, `flip`, `flythrough`, `gallery`,
`glitter`, `honeycomb`, `morph`, `pan`, `prism`, `reveal`, `ripple`, `shred`, `switch`, `vortex`,
`warp`, `wheelReverse`, `window`. These are written with a base-effect fallback, so other tools show a
similar transition rather than failing to open the file.

Directions accept friendly aliases (`left`, `right`, `up`, `down`, `horizontal`, `vertical`,
`topLeft`, `bottomRight`, ...). A direction the effect does not support is ignored with a warning, as
are an unknown `type`, an out-of-range `spokes`, and a non-numeric `duration`/`advTm`.
