---
title: Content Parts & Ink
---

A content part embeds markup PresentationML does not define — most commonly InkML for pen
annotations — as its own package part, referenced from the slide (MS-PPTX §2.2.3).

```typescript
slide.addContentPart({
    data: inkmlMarkup,
    contentType: 'application/inkml+xml',
    relationshipType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
    ink: true,
    cover: previewPngBase64,
    x: 1, y: 1, w: 4, h: 3,
});
```

| Option             | Type    | Default              | Description                                                  |
| :----------------- | :------ | :------------------- | :----------------------------------------------------------- |
| `data`             | string  |                      | **required** - the payload markup                            |
| `contentType`      | string  |                      | **required** - content type of the payload                   |
| `relationshipType` | string  |                      | **required** - relationship type linking slide to payload    |
| `ink`              | boolean | `false`              | payload is InkML; makes `cover` required                     |
| `cover`            | string  |                      | base64 raster preview for the fallback                       |
| `fileName`         | string  | `contentPart<n>.xml` | payload file name inside the package                         |
| `x`/`y`/`w`/`h`    | number  |                      | position and size (inches)                                   |
| `altText`          | string  |                      | alt text on the fallback shape                               |

## Why `contentType` and `relationshipType` are required

They belong to the format being embedded, not to PptxGenJS. Writing the wrong value produces a package
PowerPoint reports as damaged, and no amount of validation here can detect it — so rather than guess,
both are required and used verbatim. Take them from the file or specification you are embedding.

## Fallbacks are not optional

`p14:contentPart` replaces a shape, so it is written inside `mc:AlternateContent`. The specification
mandates the fallback shape:

- **ink** falls back to `p:pic` showing your `cover` image (§2.2.3.1) — which is why `cover` is
  required for ink; without it the fallback would render nothing
- **any other payload** falls back to an empty `p:sp` of the same size (§2.2.3)

A content part missing `data`, `contentType`, `relationshipType`, or (for ink) `cover` is refused with a
warning rather than written.

## Not supported

Office App / web-extension references (§2.2.13) are not written: they need the web-extension part graph
from [MS-OWEXML] §2.1.3, which is not verifiable from the specification text available here. See
`docs/ms-pptx-profile.md`.
