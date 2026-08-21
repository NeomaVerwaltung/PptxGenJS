---
title: Comments
---

PptxGenJS writes PowerPoint's modern threaded comments (MS-PPTX §2.16). Comments are opt-in: a
presentation that never calls `addComment()` gains no parts, content types, or relationships.

```typescript
const slide = pptx.addSlide();
slide.addText('Q3 revenue', { x: 1, y: 1, w: 4, h: 1 });

slide.addComment({
    text: 'Check this figure against the ledger',
    author: 'Ada Lovelace',
    x: 4,
    y: 2,
    replies: [{ text: 'Fixed in the latest export', author: 'Grace Hopper' }],
});
```

## Comment Props (`CommentProps`)

| Option     | Type                  | Description                                                        |
| :--------- | :-------------------- | :----------------------------------------------------------------- |
| `text`     | string                | **required** - comment body                                        |
| `author`   | string                | **required** - author name; added to the author table automatically |
| `x`/`y`    | number                | anchor position (inches) - pass **both** or neither                |
| `created`  | string                | ISO 8601 timestamp; defaults to the time of export                 |
| `resolved` | boolean               | mark the thread resolved                                           |
| `replies`  | `CommentReplyProps[]` | replies in the thread (`text`, `author`, `created`)                |
| `id`       | string                | comment GUID; derived from its position when omitted               |

## Authors

Any author named by `addComment()` is added to `ppt/authors.xml` automatically, with initials derived
from the name. Set `pptx.commentAuthors` to supply the identity metadata PowerPoint shows:

```typescript
pptx.commentAuthors = [
    { name: 'Ada Lovelace', initials: 'AL', userId: 'ada@example.com', providerId: 'AD' },
];
```

| Option       | Type   | Default  | Description                                    |
| :----------- | :----- | :------- | :--------------------------------------------- |
| `name`       | string |          | **required** - matched against comment authors |
| `initials`   | string | derived  | shown in the comment avatar                    |
| `userId`     | string | `''`     | identity-provider user id                      |
| `providerId` | string | `'None'` | identity provider                              |
| `id`         | string | derived  | author GUID                                    |

## Reproducible output

Comment, reply, and author ids are derived from position rather than randomised, so exporting the same
presentation twice produces the same package — **provided you pass `created`**. A timestamp cannot be
derived, so when it is omitted every comment on that export gets the export time instead.

## Not written

Comment change records (§2.18–2.19) need the changes-information part that is not yet emitted, and the
extension URIs for tasks (§2.20) and reactions (§2.21) have no documented source — writing a guessed
URI risks the repair dialog. See `docs/ms-pptx-profile.md`.
