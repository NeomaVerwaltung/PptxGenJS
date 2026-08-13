# PptxGenJS agent guidance

## OOXML specification

For OOXML generation or package changes, work from the official [ECMA-376 Office Open XML specification](https://ecma-international.org/publications-and-standards/standards/ecma-376/). It provides the current downloadable parts:

- Part 1: Fundamentals and Markup Language Reference (DrawingML and PresentationML)
- Part 2: Open Packaging Conventions
- Part 3: Markup Compatibility and Extensibility
- Part 4: Transitional Migration Features

Use the relevant part as the authority when a generated package or element is in question; keep package-contract tests semantic rather than snapshotting generated XML.
