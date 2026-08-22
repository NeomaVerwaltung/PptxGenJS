/**
 * PptxGenJS: Line properties (`a:ln` and its aliases)
 *
 * The same `CT_LineProperties` content model is used by shape outlines (`a:ln`), underlines
 * (`a:uLn`), and table cell borders, so the element name is a parameter rather than baked in.
 */

import { ARROW_SIZES, COMPOUND_TYPES } from '../core-enums'
import { ShapeLineProps } from '../core-interfaces'
import { genXmlColorSelection, valToPts } from '../gen-utils'

/**
 * Create a line element
 * @param {ShapeLineProps} line - line options
 * @param {string} tag - element name; `a:ln` for a shape outline, `a:uLn` for an underline
 * @return {string} XML
 */
export function genXmlLine (line: ShapeLineProps, tag = 'a:ln'): string {
	// `cmpd` was hardcoded to "sng" on table borders; it is written only when asked for elsewhere,
	// so existing output is unchanged
	const compound = COMPOUND_TYPES.has(String(line.compound)) ? ` cmpd="${String(line.compound)}"` : ''
	if (line.compound && !COMPOUND_TYPES.has(String(line.compound))) {
		console.warn(`[pptxgenjs] line \`compound\` must be one of ${[...COMPOUND_TYPES].join(', ')} - "${String(line.compound)}" ignored`)
	}

	let xml = line.width ? `<${tag} w="${valToPts(line.width)}"${compound}>` : `<${tag}${compound}>`
	if (line.color || line.type === 'gradient') xml += genXmlColorSelection(line)

	// CT_LineProperties sequence: fill, dash, join, headEnd, tailEnd. A custom pattern replaces
	// the preset one - emitting both would be ambiguous, since `a:custDash` is the same choice.
	xml += genXmlDash(line)
	xml += genXmlJoin(line)
	xml += genXmlArrow('a:headEnd', line.beginArrowType, line.beginArrowSize)
	xml += genXmlArrow('a:tailEnd', line.endArrowType, line.endArrowSize)

	xml += `</${tag}>`
	return xml
}

/** The dash choice: a custom pattern when given, otherwise the preset */
function genXmlDash (line: ShapeLineProps): string {
	const stops = (line.customDash ?? []).filter(stop => {
		const ok = typeof stop?.dash === 'number' && isFinite(stop.dash) && stop.dash > 0 &&
			typeof stop.space === 'number' && isFinite(stop.space) && stop.space >= 0
		if (!ok) console.warn('[pptxgenjs] each `customDash` stop needs a positive `dash` and a non-negative `space` (percent of line width) - stop ignored')
		return ok
	})
	if (stops.length > 0) {
		// `d`/`sp` are percentages of the line width, in 1000ths of a percent
		const items = stops.map(stop => `<a:ds d="${Math.round(stop.dash * 1000)}" sp="${Math.round(stop.space * 1000)}"/>`).join('')
		return `<a:custDash>${items}</a:custDash>`
	}
	return line.dashType ? `<a:prstDash val="${line.dashType}"/>` : ''
}

/** The join choice: `a:round`, `a:bevel`, or `a:miter` with its limit */
function genXmlJoin (line: ShapeLineProps): string {
	if (!line.join) return ''
	if (line.join === 'round') return '<a:round/>'
	if (line.join === 'bevel') return '<a:bevel/>'
	if (line.join === 'miter') {
		const lim = typeof line.miterLimit === 'number' && isFinite(line.miterLimit) && line.miterLimit >= 0 ? line.miterLimit : 800
		return `<a:miter lim="${Math.round(lim * 1000)}"/>`
	}
	console.warn(`[pptxgenjs] line \`join\` must be 'round' | 'bevel' | 'miter' - "${String(line.join)}" ignored`)
	return ''
}

/** An arrow end with its optional sizing */
function genXmlArrow (tag: string, type?: string, size?: { width?: string, length?: string }): string {
	if (!type) return ''
	const w = ARROW_SIZES.has(String(size?.width)) ? ` w="${String(size?.width)}"` : ''
	const len = ARROW_SIZES.has(String(size?.length)) ? ` len="${String(size?.length)}"` : ''
	return `<${tag} type="${type}"${w}${len}/>`
}
