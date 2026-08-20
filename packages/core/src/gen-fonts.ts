/**
 * PptxGenJS: Opt-in embedded fonts (`ppt/fonts/*.fntdata`)
 *
 * PowerPoint stores an embedded font as Embedded OpenType (EOT) bytes in a `.fntdata` part, listed
 * in `p:embeddedFontLst` (ECMA-376 19.2.1.13 `CT_EmbeddedFontList`) and referenced through a
 * `relationships/font` relationship on the presentation part.
 *
 * This library ships no font files and performs no font conversion: the caller supplies EOT bytes
 * and is responsible for holding a licence that permits embedding them. Nothing here runs unless
 * `addFont()` was called, so a presentation without embedded fonts is byte-for-byte unchanged.
 */

import { base64ToBytes } from './gen-utils'
import { AddFontProps, EmbeddedFont, FontEmbedStyle } from './core-interfaces'

/** `p:embeddedFont` child element for each style (ECMA-376 19.2.1.13) */
const STYLE_ELEMENTS: Record<FontEmbedStyle, string> = {
	regular: 'regular',
	bold: 'bold',
	italic: 'italic',
	boldItalic: 'boldItalic',
}

/** EOT header: a magic number of 0x504C ("LP") at byte offset 34 (MS-EOT 2.1) */
const EOT_MAGIC_OFFSET = 34

/**
 * Coerce caller-supplied font data to bytes
 * @param {AddFontProps['data']} data - base64 string (with or without a data-URI header) or binary data
 * @returns {Uint8Array | undefined} font bytes, or `undefined` when the value is unusable
 */
function toBytes (data: AddFontProps['data']): Uint8Array | undefined {
	if (typeof data === 'string') {
		const base64 = data.includes('base64,') ? data.slice(data.indexOf('base64,') + 'base64,'.length) : data
		if (!base64.trim()) return undefined
		try {
			return base64ToBytes(base64)
		} catch {
			return undefined
		}
	}
	if (data instanceof Uint8Array) return data
	if (data instanceof ArrayBuffer) return new Uint8Array(data)
	return undefined
}

/**
 * Whether the bytes carry the EOT magic number
 * - a raw TTF/OTF/WOFF file will not, and embedding one produces a deck PowerPoint cannot open
 * @param {Uint8Array} bytes - font bytes
 * @returns {boolean} whether the data looks like EOT
 */
function isEot (bytes: Uint8Array): boolean {
	if (bytes.length < EOT_MAGIC_OFFSET + 2) return false
	return bytes[EOT_MAGIC_OFFSET] === 0x4c && bytes[EOT_MAGIC_OFFSET + 1] === 0x50
}

/**
 * Validate a font registration and turn it into an entry for the export
 * @param {AddFontProps} props - font props
 * @param {EmbeddedFont[]} existing - fonts registered so far
 * @returns {EmbeddedFont | undefined} the entry, or `undefined` when the input was rejected
 */
export function createEmbeddedFont (props: AddFontProps, existing: EmbeddedFont[]): EmbeddedFont | undefined {
	if (!props || typeof props !== 'object') {
		console.warn('[pptxgenjs] addFont: an object is required, ex: `addFont({ fontFace:"Custom", data:"<base64 eot>" })` - call ignored')
		return undefined
	}

	const fontFace = typeof props.fontFace === 'string' ? props.fontFace.trim() : ''
	if (!fontFace) {
		console.warn('[pptxgenjs] addFont: `fontFace` is required and must name the typeface as used in text options - call ignored')
		return undefined
	}

	const style: FontEmbedStyle = props.style ?? 'regular'
	if (!STYLE_ELEMENTS[style]) {
		console.warn(`[pptxgenjs] addFont: \`style\` must be 'regular' | 'bold' | 'italic' | 'boldItalic' - "${String(props.style)}" ignored, "${fontFace}" not embedded`)
		return undefined
	}

	const bytes = toBytes(props.data)
	if (!bytes?.length) {
		console.warn(`[pptxgenjs] addFont: \`data\` must be base64 or binary font bytes - "${fontFace}" not embedded`)
		return undefined
	}
	if (!isEot(bytes)) {
		console.warn(
			`[pptxgenjs] addFont: "${fontFace}" does not look like EOT data. PowerPoint stores embedded fonts as ` +
			'Embedded OpenType - convert the TTF/OTF/WOFF first (ex: `ttf2eot`, `fonteditor-core`) - font not embedded'
		)
		return undefined
	}

	if (existing.some(font => font.fontFace === fontFace && font.style === style)) {
		console.warn(`[pptxgenjs] addFont: "${fontFace}" (${style}) is already embedded - the previous registration is replaced`)
	}

	return { fontFace, style, data: bytes }
}

/**
 * Create `p:embeddedFontLst`, or `''` when no fonts are embedded
 * - one `p:embeddedFont` per typeface, with a child element per registered style
 * @note belongs after `p:notesSz` and before `p:defaultTextStyle` in the CT_Presentation sequence
 * @param {EmbeddedFont[]} fonts - registered fonts
 * @param {number} firstRelId - relationship id of the first font part
 * @param {(value: string) => string} encodeAttr - XML attribute encoder
 * @returns {string} XML string
 */
export function makeXmlEmbeddedFontLst (fonts: EmbeddedFont[], firstRelId: number, encodeAttr: (value: string) => string): string {
	if (fonts.length === 0) return ''

	const typefaces = [...new Set(fonts.map(font => font.fontFace))]
	const parts = typefaces.map(typeface => {
		const styles = fonts
			.map((font, idx) => ({ font, relId: firstRelId + idx }))
			.filter(entry => entry.font.fontFace === typeface)
			.map(entry => `<p:${STYLE_ELEMENTS[entry.font.style]} r:id="rId${entry.relId}"/>`)
			.join('')
		return `<p:embeddedFont><p:font typeface="${encodeAttr(typeface)}"/>${styles}</p:embeddedFont>`
	})

	return `<p:embeddedFontLst>${parts.join('')}</p:embeddedFontLst>`
}

/** Package path of the font part for a registered font */
export function fontPartName (index: number): string {
	return `ppt/fonts/font${index + 1}.fntdata`
}
