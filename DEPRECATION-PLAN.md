# Deprecation & API-Cleanup Plan

Audit date: 2026-07-23 (F1–F5) · 2026-08-20 (F6–F9) · Package version: 4.1.0

## Findings

### F1 — `compression: boolean` is a boolean trap
`WriteBaseProps.compression?: boolean` (src/core-interfaces.ts:1692) maps to a JSZip
strategy choice (`DEFLATE` vs `STORE`) in src/pptxgen.ts:563/569. A boolean can't
express compression *level* and reads as noise at call sites.

### F2 — Bug: compression silently ignored for explicit output types
`write({ outputType: 'base64', compression: true })` hits the middle branch at
src/pptxgen.ts:566, which calls `zip.generateAsync({ type })` **without** the
compression option. Only STREAM and the browser-blob default honor it.

### F3 — Deprecated union-typed method params (marked "remove in v4.0.0", still present)
- `write(props?: WriteProps | WRITE_OUTPUT_TYPE)` — string form deprecated v3.5.0
- `writeFile(props?: WriteFileProps | string)` — string form deprecated v3.5.0

### F4 — v4.0 deprecation cleanup never executed
~30 `@deprecated` aliases from v3.3–v3.11 remain in src/core-interfaces.ts plus
their runtime shims in src/gen-objects.ts (lines 332–334, 732–738, 1081):
`line` as color string, `lineSize/lineDash/lineHead/lineTail`, `bkgd`,
`autoFit/shrinkText` (→ `fit`), `inset` (→ `margin`), chart `border`/`fill`
(→ `plotArea.*`), `dashType` on old paths, bullet `code/startAt/style`, etc.

### F5 — Minor
- `strike?: boolean | 'dblStrike' | 'sngStrike'` — mixed bool/enum.
- `verbose?: boolean` on TableToSlidesProps — undocumented.
- Plain on/off booleans (`bold`, `flipH`, `showLegend`, …) are **fine**; leave them.

### F6 — Duplicate `p:cNvPr@id` values on a single slide
Three unrelated id schemes coexist in src/xml/slide.ts: most objects use their index
(`idx + 2`), media uses `mediaRid + 2` (lines 749 and 766), and tables use
`intTableNum * slide._slideNum + 1` (line 310). A slide holding text + media + a table
emits ids `2, 3, 2` — the table collides with the text shape. ECMA-376 requires
`cNvPr@id` to be unique within the slide, and duplicates are a known repair-dialog
cause. The media formula also can't change freely: the comment at line 765 records that
it must match the preview-image rId or PowerPoint errors.

Blocked on: one id allocator per slide, which changes every emitted id and therefore
any consumer diffing generated XML.

### F7 — `fill: { type: 'none' }` emits no fill element, while *omitting* `fill` emits `<a:noFill/>`
src/xml/slide.ts:648 reads `options.fill ? genXmlColorSelection(fill) : '<a:noFill/>'`,
and `genXmlColorSelection` has no `none` branch (src/gen-utils.ts:378), so an explicit
`type: 'none'` falls through to `''`. The shape then inherits the theme fill instead of
being transparent — the opposite of what was asked for, and the opposite of what
omitting `fill` does.

Blocked on: shapes that pass `type: 'none'` today are theme-filled; fixing it makes them
transparent, which is a visible change.

### F8 — `p14:modId` is emitted on tables only
The MS-PPTX 2.2.9 modification id is written for `p:graphicFrame` tables
(src/xml/slide.ts:313) and nowhere else. PowerPoint writes it on every shape it creates.
Widening it is additive per shape but changes the XML of every existing shape.

### F9 — Chart part numbering is process-global
The chart counter lives at module scope, so the first chart of the *second* presentation
created in a process is `ppt/charts/chart2.xml`. Verified: two fresh presentations in one
process yield `chart1.xml` and `chart2.xml`. Relationships are internally consistent, so
nothing breaks, but part names leak state across presentations and make output
non-reproducible across a process. Counting per presentation would fix it and renames
parts.

## Plan

### Phase 1 — bug fix (now, patch release)
Fix F2: pass the compression option in the explicit-outputType branch too, so all
three `zip.generateAsync` calls behave the same. One-line change + one test.

> **Status 2026-08-04**: Phase 1 shipped (PR #50). Phase 2 shipped: `pptx.compression`
> enum, per-call boolean deprecated with one-time warning, warn-once on all v3.x shims
> (`line` string, `lineSize/lineDash/lineHead/lineTail`, chart `border`/`fill`, `bkgd`,
> `strike: true`). Phase 3 remains for the next major. The breaking remainder of
> issue #29 (margin inches-vs-points, alignment vocab collapse) is folded into Phase 3.

### Phase 2 — deprecate, don't break (next minor)
1. Move compression to presentation-level config and kill the boolean in one move:
   `new PptxGenJS({ compression: 'none' | 'fast' | 'best' })` (maps to STORE /
   DEFLATE level 1 / DEFLATE level 9). Export methods stop growing options; the
   setting is document config like layout, not a per-call flag.
   Keep the old per-call `compression: boolean` working (`true` → `'best'`,
   overrides the constructor value) but mark it
   `@deprecated - set compression on the PptxGenJS constructor` and warn once.
2. Emit a single `console.warn` (once per process) when any deprecated alias from
   F3/F4 is used — today most shims are silent, so users have no migration signal.
3. Document `verbose` or mark it `@deprecated`.

### Phase 3 — removal (next major, v5.0)
1. Remove the per-call `compression` prop from `WriteBaseProps` entirely; the
   constructor enum is the only knob.
2. Remove string/union overloads of `write()` / `writeFile()` — options object only.
3. Delete all F4 aliases from interfaces and their shims in gen-objects.ts.
4. Normalize `strike` to enum-only (`'dblStrike' | 'sngStrike'`), map `true` removed.
5. Issue #29 remainder: standardize `margin` on inches everywhere (SlideNumberProps is
   points today), remove `BorderProps.pt`, collapse the alignment vocab to the public
   string unions and internalize `TEXT_HALIGN`/`TEXT_VALIGN`.
5. Ship a MIGRATION.md table: old prop → new prop (the `@deprecated` JSDoc tags
   already contain the mapping; generate the table from them).

### Phase 3b — output-behaviour cleanups (v5.0)
F6–F9 are all "the emitted XML is wrong or inconsistent, and fixing it changes bytes for
existing users". None can ship in a minor, and none is urgent: F6 and F7 have workarounds
(avoid mixing media and tables on one slide; omit `fill` instead of `type: 'none'`), F8
and F9 are cosmetic. Group them in the major so consumers absorb one XML change, not four.

6. F6: single per-slide id allocator; `cNvPr@id` unique per slide. Keep the media/preview
   rId invariant noted at src/xml/slide.ts:747.
7. F7: give `genXmlColorSelection` a `none` branch emitting `<a:noFill/>`, so explicit and
   omitted both mean "no fill".
8. F8: emit `p14:modId` for every shape, using the same derivation tables use.
9. F9: count chart parts per presentation.

### The rule for changing a default
`chartTrackingRefBased` moved from opt-in to on-by-default in a minor (PR #107) on this
test: **the change is not breaking if slides stay visually intact.** A presentation-level
flag that alters only editing or tracking behaviour, and provably leaves every rendered
part byte-identical, can flip. Anything that changes pixels — on open *or* on re-save —
stays opt-in. That is why `defaultImageDpi` did not flip: an absent extension means "use
the user's own Office compression preference", not "220", so writing it overrides that
preference per file (PR #123).

### Non-goals
- Renaming the flat chart-axis namespace (`catAxisLabelFontBold`, …). It's ugly but
  pervasive, documented, and a rewrite would churn every chart user for zero
  functional gain. Revisit only if a chart-options overhaul is planned anyway.
