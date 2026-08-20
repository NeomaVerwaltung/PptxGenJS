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
## Animations

Set `animation` on any object you add to a slide.

```typescript
slide.addText('Headline', { x: 1, y: 1, w: 6, h: 1, animation: { type: 'fadeIn' } });
slide.addText('Subtitle', { x: 1, y: 2, w: 6, h: 1, animation: { type: 'wipeIn', direction: 'left', trigger: 'withPrevious' } });
slide.addShape(pptx.ShapeType.rect, { x: 1, y: 3, w: 3, h: 1, animation: { type: 'zoomOut', trigger: 'afterPrevious', delay: 250 } });
```

### Animation Props (`AnimationProps`)

| Option      | Type   | Default   | Description                                             |
| :---------- | :----- | :-------- | :------------------------------------------------------ |
| `type`      | string |           | **required** - preset (see below)                       |
| `trigger`   | string | `onClick` | `onClick`, `withPrevious`, or `afterPrevious`            |
| `direction` | string |           | for presets that take one (see below)                   |
| `delay`     | number | `0`       | delay before the effect starts (milliseconds)            |
| `duration`  | number | `500`     | effect length (milliseconds)                            |

**Entrance**: `appear`, `fadeIn`, `wipeIn`, `zoomIn` — **Exit**: `disappear`, `fadeOut`, `wipeOut`, `zoomOut`.

`wipeIn`/`wipeOut` take `direction: 'up' | 'right' | 'down' | 'left'` (default `up`);
`zoomIn`/`zoomOut` take `direction: 'in' | 'out'`. `appear`/`disappear` are instant and take none.

`withPrevious` and `afterPrevious` join the click group before them, so a run of them plays together
with (or after) the preceding `onClick` effect.

Motion-path and emphasis effects are not supported yet. An unknown preset or trigger is skipped with a
warning; an unsupported `direction`, a negative `delay`, and a non-numeric `duration` fall back to
their defaults.
## Slide Creation Id

PowerPoint tags each slide with a stable id so it can recognise the same slide across saves
(MS-PPTX §2.2.9 `p14:creationId`). It is opt-in per slide:

```typescript
const slide = pptx.addSlide();
slide.creationId = true;   // PptxGenJS assigns one
slide.creationId = 918273; // or supply your own unsigned 32-bit value
```

Values must be integers in the range 0–4294967295; anything else is ignored with a warning. Generated
ids are derived from the slide number, so repeated exports of the same presentation produce the same
value.
