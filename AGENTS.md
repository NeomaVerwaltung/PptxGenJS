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

`npm run check` (lint + typecheck + all tests) is the gate; it must pass before any PR. Root scripts fan out over the workspaces, so run them from the repo root. `npm run build` produces each package's `dist/`.

| Command | Purpose |
| --- | --- |
| `npm run check` | lint + typecheck + tests — run this before reporting done |
| `npm run lint` | ESLint over every package's `src` and `test` |
| `npm run typecheck` | `tsc --noEmit` per package |
| `npm test` | Node test runner over the test files listed in `package.json` |
| `npm run build` | Rollup ESM + CJS bundles into each package's `dist/` |
| `npm run test:office` | Opens a generated `.pptx` in LibreOffice; needs `PPTXGENJS_OFFICE_BIN` and a local LibreOffice install. CI runs it — skip locally unless installed. |

**The `test` script discovers its files.** `scripts/run-tests.mjs` finds every `test/**/*.test.ts` recursively, so a new test file runs as soon as it exists — no `package.json` change, and subdirectories work. (It is a script rather than a shell glob because `cmd.exe` does not expand globs and Node only globs them itself from v21, while this repo supports v20.)

Prefer a new test file over appending to a large one — the shared ones are where feature branches collide. `contracts-<area>.test.ts` is split by feature area, and the LibreOffice round-trip builds its deck from one module per feature under `test/office-fixtures/`, listed in that directory's `index.ts`. Adding a round-trip fixture is a new file plus one line; `office-fixtures.test.ts` fails if you forget the line, and runs every fixture so a broken one surfaces without LibreOffice installed.

## Conventions

- The source compiles under `strictNullChecks`. Do not introduce non-null assertions (`!`) or unchecked `as` casts — use explicit types, guard clauses, and validated defaults instead.
- The public API stays compatible with upstream PptxGenJS. Additive changes only; deprecations go through `DEPRECATION-PLAN.md`.
- Keep package-contract tests semantic (assert on parsed structure), not snapshots of generated XML.
- Hand-written public types live in `packages/core/types/index.d.ts` and must be updated alongside API changes.
- Anything that only produces or consumes plain option objects belongs in `packages/std/`, not the core; the core is for code that needs the internal slide object model. Std helpers type slides structurally (an object with the `addX` method they call) rather than importing the `Slide` class.
- A new std helper is a file plus one line in its category's `index.ts`. A new std category also needs an entry in `packages/std/rollup.config.mjs` and in `exports` in `packages/std/package.json`; `src/index.ts` is a barrel and defines nothing.
- The two packages version and release independently: core tags are `vX.Y.Z`, std tags are `std-vX.Y.Z`. One `Release` workflow serves both. std is in beta (`0.x`) and stays independent; its only tie to the core is its `peerDependencies` floor, which must be raised whenever a helper starts using a newer core feature. See the coupling policy in `RELEASING.md`.

## Source map

| Path | Contents |
| --- | --- |
| `packages/core/src/pptxgen.ts` | Entry point; the `PptxGenJS` presentation class and save/write pipeline |
| `packages/core/src/slide.ts` | `Slide` class — per-slide state and the `addX` surface |
| `packages/core/src/gen-objects.ts` | Builds slide object models (text, shapes, images, charts, tables) |
| `packages/core/src/gen-tables.ts` | Table layout, row splitting, and auto-paging |
| `packages/core/src/gen-media.ts` | Image/media fetching and encoding |
| `packages/core/src/gen-utils.ts` | Units, colors, escaping, and other shared helpers |
| `packages/core/src/core-enums.ts` | Frozen constant objects (shape types, chart types, schemes) |
| `packages/core/src/core-interfaces.ts` | Internal + public TypeScript interfaces |
| `packages/core/src/xml/` | OOXML emitters: `package.ts` (package parts), `slide.ts`, `text.ts`, `relationships.ts` |
| `packages/core/src/charts/` | Chart XML: `xml.ts`, `axes.ts`, `title.ts`, `workbook.ts`, `utils.ts` |
| `packages/core/test/pptx-contracts.ts` | Shared helpers that parse a generated `.pptx` and assert package semantics; typechecked via `packages/std`, not by core's own `tsc` |
| `packages/core/test/fixtures.ts` | Sample inputs shared by the contract tests (PNG/WAV/EOT blobs, the small presentation the package contracts inspect) |
| `packages/core/test/office-fixtures/` | One module per feature area, each adding its slides to the LibreOffice round-trip deck; `index.ts` lists them in order |
| `packages/std/test/` | One file per category (`layout`, `charts`) plus `exports.test.ts`, which fails when a category is added without wiring its subpath |
| `docs/` | VitePress site (`npm run docs:dev`), shared by both packages |
| `tsconfig.base.json` | Shared compiler options; each package's `tsconfig.json` extends it |
| `scripts/` | Repo-level tooling, not shipped in any package (`sync-version.mjs`, run by a package's `version` hook; `run-tests.mjs`, the test-file discovery behind `npm test`) |
| `packages/std/` | `@neo-ma/pptxgenjs-std` — helpers composing the public API; published separately, no runtime dep on the core |
| `packages/std/src/<category>/` | One directory per category (`layout`, `charts`), each with an `index.ts` that is its whole public surface and a matching subpath export |

Deeper testing guidance is in `TESTING.md`; release mechanics are in `RELEASING.md`.
