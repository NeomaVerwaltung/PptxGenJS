/**
 * PptxGenJS: Modern threaded comments (MS-PPTX 2.16)
 *
 * Two part families:
 * - `ppt/authors.xml` — the author table (`p188:authorLst`, 2.1.6), reached by an implicit
 *   relationship from the presentation part
 * - `ppt/comments/commentSlideN.xml` — one comment list per slide (`p188:cmLst`, 2.1.5), reached by
 *   an explicit relationship from that slide
 *
 * Each comment anchors to its slide through a `pc` moniker chain (2.12) and carries its text in a
 * DrawingML `a:txBody`. Replies nest in `p188:replyLst`.
 *
 * Ids are derived from position rather than randomised, so exporting the same presentation twice
 * produces the same package. Timestamps cannot be derived - a comment without one is meaningless -
 * so `created` defaults to export time and callers who need byte-identical output pass it in.
 */

import { COMMENT, EMU } from './core-enums'
import { CommentAuthorProps, CommentReplyProps, PresSlide } from './core-interfaces'
import { encodeXmlEntities } from './gen-utils'

/**
 * Hex discriminators for derived ids - a GUID may only contain hex digits, so these cannot be
 * mnemonic letters like `r` for "reply"
 */
const ID_KIND = { author: 'a', comment: 'c', reply: 'e' } as const

/** Format an index as a GUID, so derived ids are valid without being random */
function derivedGuid (kind: typeof ID_KIND[keyof typeof ID_KIND], index: number): string {
	return `{${kind.padStart(8, '0')}-0000-0000-0000-${String(index).padStart(12, '0')}}`
}

/**
 * Build the author table from the declared authors plus any author named only by a comment
 * @param {PresSlide[]} slides - presentation slides
 * @param {CommentAuthorProps[]} declared - authors set on the presentation
 * @returns {Required<CommentAuthorProps>[]} authors with every field resolved
 */
export function collectCommentAuthors (slides: PresSlide[], declared?: CommentAuthorProps[]): Array<Required<CommentAuthorProps>> {
	const authors: Array<Required<CommentAuthorProps>> = []
	const byName = new Map<string, number>()

	const add = (props: CommentAuthorProps): void => {
		const name = String(props.name ?? '').trim()
		if (!name) {
			console.warn('[pptxgenjs] comment author `name` is required - author ignored')
			return
		}
		const key = name.toLowerCase()
		if (byName.has(key)) return
		byName.set(key, authors.length)
		authors.push({
			name,
			// PowerPoint shows initials in the comment avatar; derive them when not supplied
			initials: props.initials ?? name.split(/\s+/).map(part => part[0] ?? '').join('').slice(0, 3).toUpperCase(),
			id: props.id ?? derivedGuid(ID_KIND.author, authors.length + 1),
			userId: props.userId ?? '',
			providerId: props.providerId ?? 'None',
		})
	}

	;(declared ?? []).forEach(add)
	slides.forEach(slide => {
		;(slide.comments ?? []).forEach(comment => {
			add({ name: comment.author })
			;(comment.replies ?? []).forEach(reply => add({ name: reply.author }))
		})
	})

	return authors
}

/** Resolve an author name to its id, falling back to the first author */
function authorId (name: string, authors: Array<Required<CommentAuthorProps>>): string {
	const match = authors.find(author => author.name.toLowerCase() === String(name ?? '').trim().toLowerCase())
	return match?.id ?? authors[0]?.id ?? derivedGuid(ID_KIND.author, 1)
}

/** Comment text as a DrawingML text body */
function commentText (text: string): string {
	return (
		'<p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/>' +
		`<a:t>${encodeXmlEntities(String(text ?? ''))}</a:t></a:r></a:p></p188:txBody>`
	)
}

/**
 * Create `ppt/authors.xml`
 * @param {Required<CommentAuthorProps>[]} authors - resolved author table
 * @returns {string} XML string
 */
export function makeXmlCommentAuthors (authors: Array<Required<CommentAuthorProps>>): string {
	const entries = authors
		.map(author =>
			`<p188:author id="${author.id}" name="${encodeXmlEntities(author.name)}"` +
			`${author.initials ? ` initials="${encodeXmlEntities(author.initials)}"` : ''}` +
			` userId="${encodeXmlEntities(author.userId)}" providerId="${encodeXmlEntities(author.providerId)}"/>`
		)
		.join('')

	return (
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
		`<p188:authorLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p188="${COMMENT.p188}">` +
		entries +
		'</p188:authorLst>'
	)
}

/**
 * Create one slide's `ppt/comments/commentSlideN.xml`
 * @param {PresSlide} slide - slide whose comments to write
 * @param {Required<CommentAuthorProps>[]} authors - resolved author table
 * @param {string} exportedAt - ISO timestamp used when a comment has no `created`
 * @returns {string} XML string
 */
export function makeXmlSlideComments (slide: PresSlide, authors: Array<Required<CommentAuthorProps>>, exportedAt: string): string {
	const comments = slide.comments ?? []

	const body = comments
		.map((comment, idx) => {
			const created = comment.created ?? exportedAt
			// both coordinates are needed for an anchor; a partial one would place the comment at 0
			const pos = typeof comment.x === 'number' && typeof comment.y === 'number' && isFinite(comment.x) && isFinite(comment.y)
				? `<p188:pos x="${Math.round(comment.x * EMU)}" y="${Math.round(comment.y * EMU)}"/>`
				: ''
			if ((comment.x !== undefined) !== (comment.y !== undefined)) {
				console.warn('[pptxgenjs] addComment: an anchored comment needs both `x` and `y` - anchor ignored')
			}

			const replies = (comment.replies ?? [])
				.map((reply, replyIdx) => replyXml(reply, authors, created, slide._slideNum ?? 0, idx, replyIdx))
				.join('')

			return (
				`<p188:cm id="${comment.id ?? derivedGuid(ID_KIND.comment, (slide._slideNum ?? 0) * 1000 + idx + 1)}" ` +
				`authorId="${authorId(comment.author, authors)}" created="${encodeXmlEntities(created)}"` +
				`${comment.resolved ? ' status="resolved"' : ''}>` +
				// EG_CommentAnchor: a document moniker plus the slide moniker (MS-PPTX 2.12)
				`<pc:sldMkLst><pc:docMkLst><pc:docMk/></pc:docMkLst><pc:sldMk sldId="${slide._slideId}"/></pc:sldMkLst>` +
				pos +
				(replies ? `<p188:replyLst>${replies}</p188:replyLst>` : '') +
				commentText(comment.text) +
				'</p188:cm>'
			)
		})
		.join('')

	return (
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
		'<p188:cmLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
		'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
		`xmlns:p188="${COMMENT.p188}" xmlns:pc="${COMMENT.pc}">` +
		body +
		'</p188:cmLst>'
	)
}

/** One reply inside a comment's `replyLst` */
function replyXml (
	reply: CommentReplyProps,
	authors: Array<Required<CommentAuthorProps>>,
	fallbackCreated: string,
	slideNum: number,
	commentIdx: number,
	replyIdx: number
): string {
	return (
		`<p188:reply id="${reply.id ?? derivedGuid(ID_KIND.reply, slideNum * 1000000 + commentIdx * 1000 + replyIdx + 1)}" ` +
		`authorId="${authorId(reply.author, authors)}" created="${encodeXmlEntities(reply.created ?? fallbackCreated)}">` +
		commentText(reply.text) +
		'</p188:reply>'
	)
}

/** Package path of a slide's comment part */
export function commentPartName (slideNumber: number): string {
	return `ppt/comments/commentSlide${slideNumber}.xml`
}
