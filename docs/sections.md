---
title: Slide Sections
---

Sections group slides in PowerPoint's slide navigator, giving long decks a collapsible outline (for
example "Intro", "Financials", "Appendix"). They are organizational only — they do not change how slides
render. Use them when a presentation is large enough that named groups help reviewers navigate it.

## Syntax

Define a section on the presentation, then assign slides to it by title when you add them:

```typescript
pptx.addSection({ title: "Tables" })
pptx.addSection({ title: "Charts", order: 3 })
```

## Section Options

| Option  | Type    | Description   | Notes                                                                              |
| :------ | :------ | :------------ | :--------------------------------------------------------------------------------- |
| `title` | string  | section title | Required. Shown in the PowerPoint navigator; must be unique.                       |
| `order` | integer | section order | Optional, `1`-based. Inserts the section at a specific index instead of appending. |

## Section Example

```typescript
import pptxgen from "@neo-ma/pptxgenjs";
let pptx = new pptxgen();

// STEP 1: Create a section
pptx.addSection({ title: "Tables" });

// STEP 2: Provide section title to a slide that you want in corresponding section
let slide = pptx.addSlide({ sectionTitle: "Tables" });

slide.addText("This slide is in the Tables section!", { x: 1.5, y: 1.5, fontSize: 18, color: "363636" });
pptx.writeFile({ fileName: "Section Sample.pptx" });
```

## Zoom Objects

A zoom is a shape that jumps to a slide or a section when clicked. All three kinds are opt-in and add
nothing to a presentation that does not use them.

```typescript
pptx.addSection({ title: 'Intro' });
pptx.addSection({ title: 'Results' });

const hub = pptx.addSlide({ sectionTitle: 'Intro' });
hub.addZoom({ slideNumber: 4, x: 1, y: 1, w: 3, h: 2 });
hub.addSectionZoom({ sectionTitle: 'Results', x: 5, y: 1, w: 3, h: 2 });
hub.addSummaryZoom({ sectionTitles: ['Intro', 'Results'], x: 1, y: 4, w: 8, h: 3 });
```

| Method             | Target                                      | Required prop     |
| :----------------- | :------------------------------------------ | :---------------- |
| `addZoom()`        | a single slide                              | `slideNumber`     |
| `addSectionZoom()` | a section                                   | `sectionTitle`    |
| `addSummaryZoom()` | several sections, laid out in a grid        | `sectionTitles`   |

### Shared Props

| Option           | Type    | Default | Description                                                    |
| :--------------- | :------ | :------ | :------------------------------------------------------------- |
| `x`/`y`/`w`/`h`  | number  |         | position and size (inches)                                     |
| `cover`          | string  |         | base64 thumbnail; PowerPoint replaces it with a live one       |
| `returnToParent` | boolean | `true`  | return to this slide when the zoom finishes                    |
| `showBg`         | boolean | `true`  | keep the parent slide's background while zooming               |
| `transitionDur`  | number  |         | zoom transition duration (milliseconds)                        |
| `altText`        | string  |         | alt text for the zoom object                                   |

Targets are resolved during export, so a zoom may point at a slide or section added **after** it. A
target that never appears is dropped with a warning rather than written as a broken reference. Section
zooms reuse the exact GUID written into the presentation's section list.

Because the zoom elements are PowerPoint-2016 markup, each is written inside `mc:AlternateContent`:
other consumers render the fallback — a picture for slide and section zooms, an empty group for a
summary zoom — instead of failing to open the file.
