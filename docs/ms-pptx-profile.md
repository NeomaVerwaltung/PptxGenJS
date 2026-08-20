---
title: MS-PPTX Conformance Profile
---

PptxGenJS generates PresentationML that ECMA-376 defines, plus a small set of Microsoft extensions
that PowerPoint needs to render or track certain features. This page records **exactly which
extensions are emitted**, where each one lives, and which specification defines it.

Two rules keep this page truthful, both enforced by `test/ms-pptx-profile.test.ts`:

1. An extension may only be written through the `OOXML_EXT` registry in `src/core-enums.ts` — an
   inline GUID literal in emitting code fails the test.
2. Every extension that reaches a generated package must sit inside an `extLst` of the same namespace
   prefix, which is what [MS-PPTX] §2.2 requires of a conforming extension. Markup that *replaces*
   standard elements rather than extending them — zoom objects, modern transitions, Office math — uses
   the other wrapper §2.2 allows, `mc:AlternateContent` with a `mc:Fallback`.

## References

- [ECMA-376 Office Open XML](https://ecma-international.org/publications-and-standards/standards/ecma-376/) — Part 1 (PresentationML/DrawingML), Part 2 (packaging), Part 3 (markup compatibility)
- [MS-PPTX](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-pptx/) v20240820 — Microsoft's PresentationML extensions
- [MS-OI29500](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/) — where Office behaviour differs from the standard

## Emitted extensions

| Extension | Host | `p:ext@uri` | Namespace | Spec | API |
| :-- | :-- | :-- | :-- | :-- | :-- |
| `thm15:themeFamily` | `a:theme/a:extLst` | `{05A4C25C-…}` | thememl/2012/main | ECMA-376 + MS-ODRAWXML | none (always written; theme identity) |
| `ma14:wrappingTextBoxFlag` | `p:cNvSpPr/a:extLst` | `{C572A759-…}` | mac/drawingml/2011/main | MS-ODRAWXML | none (text boxes; Office for Mac) |
| `asvg:svgBlip` | `a:blip/a:extLst` | `{96DAC541-…}` | drawing/2016/SVG/main | MS-ODRAWXML | `addImage()` with an SVG (raster fallback in the blip) |
| `ahyp:hlinkClr` | `a:hlinkClick/a:extLst` | `{A12FA001-…}` | drawing/2018/hyperlinkcolor | MS-ODRAWXML | `hyperlink` with an explicit `color` |
| `p14:sectionLst` | `p:presentation/p:extLst` | `{521415D9-…}` | powerpoint/2010/main | §2.2.5 / §2.3.1.25 | `pptx.addSection()` |
| `p15:sldGuideLst` | `p:presentation/p:extLst` | `{EFAFB233-…}` | powerpoint/2012/main | §2.2.11 / §2.4.1.6 | `pptx.guides` |
| `p14:creationId` | `p:cSld/p:extLst` | `{BB962C8B-…}` | powerpoint/2010/main | §2.2.9 / §2.3.1.4 | `slide.creationId`; always written on notes slides |
| `p14:modId` | `p:nvPr/p:extLst` | `{D42A27DB-…}` | powerpoint/2010/main | §2.2.9 / §2.3.1.19 | none (tables; unique per shape on the slide) |
| `p14:media` | `p:nvPr/p:extLst` | `{DAA4B4D4-…}` | powerpoint/2010/main | §2.3.3.14 | `addMedia()` (embedded audio/video) |
| `c15:*` | `c:extLst` | `{CE6537A1-…}` | drawing/2012/chart | MS-ODRAWXML | `showLeaderLines`, extended data labels |
| `c16:uniqueId` | `c:extLst` | `{C3380CC4-…}` | drawing/2014/chart | MS-ODRAWXML | none (chart series identity) |
| `p16:sldZm` | `p:spTree` via `mc:AlternateContent` | n/a (`mc:Choice Requires="p16"`) | powerpoint/2016/slidezoom | §2.2.15 / §2.10 | `slide.addZoom()` |
| `p16:sectionZm` | `p:spTree` via `mc:AlternateContent` | n/a (`mc:Choice Requires="p16"`) | powerpoint/2016/sectionzoom | §2.2.15 / §2.9 | `slide.addSectionZoom()` |
| `p16:summaryZm` | `p:spTree` via `mc:AlternateContent` | n/a (`mc:Choice Requires="p16"`) | powerpoint/2016/summaryzoom | §2.2.15 / §2.11 | `slide.addSummaryZoom()` |
| `p166:zmPr` | inside a zoom object | n/a | powerpoint/2016/6/main | §2.2.15 | zoom `returnToParent`/`showBg`/`transitionDur` |

Each row's package contract is asserted in `test/contracts.test.ts` or `test/issues.test.ts`, and the
LibreOffice round-trip in `test/office-open.test.ts` (`npm run test:office`) covers them end to end.

## Not emitted

These are defined by MS-PPTX but PptxGenJS does not write them. Each has a reason, so the profile
distinguishes "unsupported" from "overlooked":

| Area | Spec | Why not |
| :-- | :-- | :-- |
| Revision / changes information parts | §2.1.2, §2.1.4 | The child cardinality of `revInfo`/`chgInfo` needs Appendix A 5.4/5.16; emitting a guessed body risks the repair dialog |
| Laser traces, show-event lists | §2.2.6 | Recordings PowerPoint captures while presenting, not authoring options |
| Comments, authors, tasks, reactions | §2.1.5, §2.1.6, §2.2.10, §2.16, §2.18–2.21 | Collaboration state; needs its own parts and an author model |
| Content parts, ink, Office App references | §2.2.3, §2.2.13 | Each needs an `mc:AlternateContent` fallback and extra package parts |
| Design elements, Designer properties/tags | §2.2.17, §2.2.19, §2.2.20 | Server-driven Designer metadata with no authoring meaning |
| Classification outcomes | §2.2.18 | Compliance metadata; must stay opt-in and is not implemented |
| Modern placeholder type extension | §2.22.1.1 | `p:ph@type` alone is currently emitted, per ECMA-376 |

## Adding an extension

1. Add an entry to `OOXML_EXT` with the URI, namespace, host, and spec section in its doc comment.
2. Emit it through that entry, inside an `extLst`. If the extension replaces standard markup rather
   than adding to it, wrap it in `mc:AlternateContent` with a `mc:Fallback`, as the transitions and
   Office-math paths do.
3. Add a semantic contract test — the emitted URI, namespace, and element, plus proof that output is
   unchanged when the feature is unused.
4. Add a row to the table above, and extend `test/office-open.test.ts` so the round-trip covers it.
