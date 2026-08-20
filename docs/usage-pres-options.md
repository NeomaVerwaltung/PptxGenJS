---
title: Presentation Options
---

## Metadata

### Metadata Properties

These optional metadata properties correspond to built-in PowerPoint document properties (visible under File > Info). They help describe the presentation’s content and ownership.

| Name       | Description                  |
| :--------- | :--------------------------- |
| `title`    | title shown in PowerPoint UI |
| `author`   | presentation author          |
| `subject`  | presentation subject         |
| `company`  | company name                 |
| `revision` | revision number (as string)  |

## Library Version

> You can also check the current PptxGenJS library version using the read-only `version` property

```typescript
console.log(pptx.version); // e.g. "4.0.0"
```

### Metadata Properties Examples

PptxGenJS uses ES6-style getters/setters.

```typescript
pptx.title = 'My Awesome Presentation';
pptx.author = 'Brent Ely';
pptx.subject = 'Annual Report';
pptx.company = 'Computer Science Chair';
pptx.revision = '15';
```

## Slide Layouts (Sizes)

Layout option applies to all slides in the current Presentation.

### Slide Layout Syntax

```typescript
pptx.layout = 'LAYOUT_NAME';
```

### Standard Slide Layouts

| Layout Name    | Default | Layout Slide Size |
| :------------- | :------ | :---------------- |
| `LAYOUT_16x9`  | Yes     | 10 x 5.625 inches |
| `LAYOUT_16x10` | No      | 10 x 6.25 inches  |
| `LAYOUT_4x3`   | No      | 10 x 7.5 inches   |
| `LAYOUT_WIDE`  | No      | 13.3 x 7.5 inches |

### Custom Slide Layouts

Custom layouts of any size are supported.

* Use the `defineLayout()` method to create a custom layout of any size
* Multiple layouts are supported. For example, define both an 'A3' and an 'A4' layout, then apply whichever is required

### Custom Slide Layout Example

```typescript
// Define new layout for the Presentation
pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });

// Set presentation to use new layout
pptx.layout = 'A3';
```

> To inspect the current layout size:

```typescript
console.log(pptx.presLayout); // { width: 10, height: 5.625 }
```

## Text Direction

### Text Direction Options

Right-to-Left (RTL) text is supported. Set the RTL mode presentation property to enable it.

### Text Direction Examples

```typescript
pptx.rtlMode = true; // set RTL text mode to true
pptx.theme = { lang: "he" }; // set RTL language to use (default is 'EN-US')
```

Notes:

* You may also need to set an RTL lang value such as `lang='he'` as the default lang is 'EN-US'
* See [Issue#600](https://github.com/gitbrent/PptxGenJS/issues/600) for more

## Default Font

### Default Font Options

Use the `headFontFace` and `bodyFontFace` properties to set the default font used in the presentation.

### Default Font Examples

```typescript
pptx.theme = { headFontFace: "Arial Light" };
pptx.theme = { bodyFontFace: "Arial" };
```

## Chart Tracking

PowerPoint marks the presentations it creates so charts track their data *references* rather than cell
positions (MS-PPTX §2.2.12 `p15:chartTrackingRefBased`). PptxGenJS does the same, so generated decks
match PowerPoint-authored ones.

The flag affects no rendering — it only decides whether editing a chart's data follows cell references
or cell positions. Turn it off to write `ppt/presProps.xml` exactly as versions before v4.2.0 did:

```typescript
pptx.chartTrackingRefBased = false;
```

## Slide Show Options

```typescript
pptx.slideShow = { mode: 'kiosk', loop: true, useTimings: false };
pptx.slideShow = { mode: 'browse', browseMode: true, laserColor: 'FF0000' };
```

| Option          | Type    | Default   | Description                                                       |
| :-------------- | :------ | :-------- | :---------------------------------------------------------------- |
| `mode`          | string  | `present` | `present` (full screen), `browse` (windowed), `kiosk` (self-running) |
| `loop`          | boolean | `false`   | restart after the last slide                                      |
| `showNarration` | boolean | `true`    | play recorded narration                                           |
| `showAnimation` | boolean | `true`    | play animations                                                   |
| `useTimings`    | boolean | `true`    | use the recorded slide timings                                    |
| `browseMode`    | boolean |           | show the browse-mode UI (PowerPoint 2010+)                        |
| `laserColor`    | Color   |           | laser-pointer color (PowerPoint 2010+)                            |

## Image and View Preferences

```typescript
pptx.defaultImageDpi = 220;        // PowerPoint's own default; 0 = do not compress
pptx.discardImageEditData = true;  // smaller file, picture edits no longer undoable
pptx.readonlyRecommended = true;   // PowerPoint suggests opening read-only
```

All of the above are opt-in: while unset, `ppt/presProps.xml` is written exactly as in prior versions.
An unknown `slideShow.mode`, a negative `defaultImageDpi`, and a non-boolean flag are ignored with a
warning.

Recorded laser traces (`p14:laserTraceLst`) and show-event lists (`p14:showEvtLst`) are recordings
PowerPoint captures while presenting, not authoring options, so PptxGenJS does not generate them.
## Embedded Fonts

Embedding a font makes text render with that typeface on machines that do not have it installed. This
is opt-in: without an `addFont()` call the package has no font parts and is byte-for-byte unchanged.

```typescript
pptx.addFont({ fontFace: 'Custom Sans', data: eotBase64 });
pptx.addFont({ fontFace: 'Custom Sans', data: boldEotBase64, style: 'bold' });

pptx.addSlide().addText('Branded headline', { x: 1, y: 1, w: 8, h: 1, fontFace: 'Custom Sans' });
```

| Option     | Type                                  | Default   | Description                                              |
| :--------- | :------------------------------------ | :-------- | :------------------------------------------------------- |
| `fontFace` | string                                |           | **required** - typeface name, as used in `fontFace`      |
| `data`     | string \| ArrayBuffer \| Uint8Array   |           | **required** - EOT font data (base64 or binary)          |
| `style`    | string                                | `regular` | `regular`, `bold`, `italic`, or `boldItalic`             |

Call `addFont()` once per style of a typeface; they are grouped into one `<p:embeddedFont>` entry.

### Font data must be EOT

PowerPoint stores embedded fonts as **Embedded OpenType**. PptxGenJS does not convert font files and
ships none: convert your TTF/OTF/WOFF first (for example with `ttf2eot` or `fonteditor-core`) and pass
the resulting bytes. Data that lacks the EOT magic number is rejected with a warning rather than
written, because a `.fntdata` part holding raw TTF produces a deck PowerPoint cannot open.

### Licensing

**You are responsible for holding a licence that permits embedding the font.** Many commercial and
some open fonts restrict or forbid embedding. PptxGenJS performs no licence checks and bundles no
font files.
