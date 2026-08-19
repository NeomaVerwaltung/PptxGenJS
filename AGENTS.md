# PptxGenJS agent guidance

## Repository targeting

This checkout is the NEOMA-maintained fork: open pull requests against `NeomaVerwaltung/PptxGenJS` (`origin`), not the original `gitbrent/PptxGenJS` (`upstream`). Verify the remotes before GitHub issue or PR work.

After creating a pull request, run `gh pr view <number> --repo NeomaVerwaltung/PptxGenJS --json url,state,baseRefName,headRefName,title` and verify the target is `master` and the head is the intended branch before reporting it.

## OOXML specification

For OOXML generation or package changes, work from the official [ECMA-376 Office Open XML specification](https://ecma-international.org/publications-and-standards/standards/ecma-376/). It provides the current downloadable parts:

- Part 1: Fundamentals and Markup Language Reference (DrawingML and PresentationML)
- Part 2: Open Packaging Conventions
- Part 3: Markup Compatibility and Extensibility
- Part 4: Transitional Migration Features

For PowerPoint compatibility, also consult Microsoft's [MS-OI29500 Office implementation notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/1fd4a662-8623-49c0-82f0-18fa91b413b8), which documents Office behavior that varies from or extends the standard. Use the relevant source when a generated package or element is in question; keep package-contract tests semantic rather than snapshotting generated XML.

## Verification

`npm run check` (lint + typecheck + all tests) is the gate; it must pass before any PR. `npm run build` produces `dist/`.

| Command | Purpose |
| --- | --- |
| `npm run check` | lint + typecheck + tests — run this before reporting done |
| `npm run lint` | ESLint over `src` and `test` |
| `npm run typecheck` | `tsc --noEmit` |
| `npm test` | Node test runner over the test files listed in `package.json` |
| `npm run build` | Rollup ESM + CJS bundles into `dist/` |
| `npm run test:office` | Opens a generated `.pptx` in LibreOffice; needs `PPTXGENJS_OFFICE_BIN` and a local LibreOffice install. CI runs it — skip locally unless installed. |

**The `test` script hardcodes its file list.** A new `test/*.test.ts` will not run until it is added to both the `test` and `test:coverage` scripts in `package.json`. Prefer adding cases to an existing file.

## Conventions

- The source compiles under `strictNullChecks`. Do not introduce non-null assertions (`!`) or unchecked `as` casts — use explicit types, guard clauses, and validated defaults instead.
- The public API stays compatible with upstream PptxGenJS. Additive changes only; deprecations go through `DEPRECATION-PLAN.md`.
- Keep package-contract tests semantic (assert on parsed structure), not snapshots of generated XML.
- Hand-written public types live in `types/index.d.ts` and must be updated alongside API changes.

## Source map

| Path | Contents |
| --- | --- |
| `src/pptxgen.ts` | Entry point; the `PptxGenJS` presentation class and save/write pipeline |
| `src/slide.ts` | `Slide` class — per-slide state and the `addX` surface |
| `src/gen-objects.ts` | Builds slide object models (text, shapes, images, charts, tables) |
| `src/gen-tables.ts` | Table layout, row splitting, and auto-paging |
| `src/gen-media.ts` | Image/media fetching and encoding |
| `src/gen-utils.ts` | Units, colors, escaping, and other shared helpers |
| `src/core-enums.ts` | Frozen constant objects (shape types, chart types, schemes) |
| `src/core-interfaces.ts` | Internal + public TypeScript interfaces |
| `src/xml/` | OOXML emitters: `package.ts` (package parts), `slide.ts`, `text.ts`, `relationships.ts` |
| `src/charts/` | Chart XML: `xml.ts`, `axes.ts`, `title.ts`, `workbook.ts`, `utils.ts` |
| `test/pptx-contracts.ts` | Shared helpers that parse a generated `.pptx` and assert package semantics |
| `docs/` | VitePress site (`npm run docs:dev`) |

Deeper testing guidance is in `TESTING.md`; release mechanics are in `RELEASING.md`.
