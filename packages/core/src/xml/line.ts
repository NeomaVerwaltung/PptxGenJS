/**
 * PptxGenJS: Line properties (`a:ln` and its aliases)
 *
 * The same `CT_LineProperties` content model is used by shape outlines (`a:ln`), underlines
 * (`a:uLn`), and table cell borders, so the element name is a parameter rather than baked in.
 */

import { ShapeLineProps } from '../core-interfaces'
import { genXmlColorSelection, valToPts } from '../gen-utils'

/**
 * Create a line element
 * @param {ShapeLineProps} line - line options
 * @param {string} tag - element name; `a:ln` for a shape outline, `a:uLn` for an underline
 * @return {string} XML
 */
export function genXmlLine (line: ShapeLineProps, tag = 'a:ln'): string {
	let xml = line.width ? `<${tag} w="${valToPts(line.width)}">` : `<${tag}>`
	if (line.color || line.type === 'gradient') xml += genXmlColorSelection(line)
	if (line.dashType) xml += `<a:prstDash val="${line.dashType}"/>`
	if (line.beginArrowType) xml += `<a:headEnd type="${line.beginArrowType}"/>`
	if (line.endArrowType) xml += `<a:tailEnd type="${line.endArrowType}"/>`
	// FUTURE: `endArrowSize` < a: headEnd type = "arrow" w = "lg" len = "lg" /> 'sm' | 'med' | 'lg'(values are 1 - 9, making a 3x3 grid of w / len possibilities)
	xml += `</${tag}>`
	return xml
}
