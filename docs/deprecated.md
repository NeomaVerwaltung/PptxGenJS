---
title: Deprecated
---

## Version 3.0 Breaking Changes

Please see the [Version 3.0 Migration Guide](https://github.com/gitbrent/PptxGenJS/wiki/Version-3.0-Migration-Guide)

- `pptx.colors` is deprecated - use `pptx.SchemeColor`
- `pptx.charts` is deprecated - use `pptx.ChartType`
- `pptx.shapes` is deprecated - use `pptx.ShapeType`

## Version 2.0 Breaking Changes

Version 2.0.0 included a cleanup of the API surface and may break existing code. In most cases a
search-and-replace of the renamed option names resolves the breakage.

Although the changes primarily affect cosmetic properties, test your solutions thoroughly before upgrading PptxGenJS to the 2.0 version.

### All Users

The library `getVersion()` method is now a property: `version`

Option names are now camelCase across all methods:

- `font_face` renamed to `fontFace`
- `font_size` renamed to `fontSize`
- `line_dash` renamed to `lineDash`
- `line_head` renamed to `lineHead`
- `line_size` renamed to `lineSize`
- `line_tail` renamed to `lineTail`

Options deprecated in early 1.0 versions:

- `marginPt` renamed to `margin`

### Node Users

- `require('@neoma/pptxgenjs')` no longer returns a singleton instance
- `pptx = new PptxGenJS()` will create a single, unique instance
- Advantage: This simplifies creating [multiple presentations](#saving-multiple-presentations) - see [Issue #83](https://github.com/gitbrent/PptxGenJS/issues/83) for details.
