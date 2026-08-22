/** Shared chart XML helpers. */

import { DEF_CHART_GRIDLINE, DEF_SHAPE_SHADOW, LETTERS } from '../core-enums'
import { ChartLineCap, OptsChartGridLine, ShadowProps } from '../core-interfaces'
import { valToPts } from '../gen-utils'

export function getExcelColName (colIndex: number): string {
	let colStr = ''
	const colIdx = colIndex - 1 // Subtract 1 so `LETTERS[columnIndex]` returns "A" etc

	if (colIdx <= 25) {
		// A-Z
		colStr = LETTERS[colIdx]
	} else {
		// AA-ZZ (ZZ = index 702)
		colStr = `${LETTERS[Math.floor(colIdx / LETTERS.length - 1)]}${LETTERS[colIdx % LETTERS.length]}`
	}

	return colStr
}

/**
 * Creates `a:innerShdw` or `a:outerShdw` depending on pass options `opts`.
 * @param {Object} opts optional shadow properties
 * @param {Object} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 * @example { type: 'outer', blur: 3, offset: (23000 / 12700), angle: 90, color: '000000', opacity: 0.35, rotateWithShape: true };
 * @return {string} XML
 */
/**
 * Nominal ("coloured") brand for resolved shadow options. The symbol is module-private and unexported,
 * so `ResolvedShadowProps` values can ONLY be produced by `resolveShadowOptions` below - a hand-built
 * `Required<ShadowProps>` is not assignable. This statically guarantees anything reaching
 * `createShadowElement` has passed through the defaults-merge boundary.
 */
const shadowBrand: unique symbol = Symbol('resolvedShadow')
// `preset` stays optional: chart shadows are outer/inner only, so a resolved chart shadow never
// carries a `a:prstShdw` preset name
type ResolvedShadowProps = Required<Omit<ShadowProps, 'preset'>> & Pick<ShadowProps, 'preset'> & { readonly [shadowBrand]: boolean }

/**
 * Resolve boundary: merge user shadow options over the documented defaults and brand the result.
 * Returns undefined for absent/invalid input. The only constructor of `ResolvedShadowProps` (no cast -
 * the brand is added by this factory).
 */
export function resolveShadowOptions (options: ShadowProps | undefined): ResolvedShadowProps | undefined {
	if (!options) return undefined
	if (typeof options !== 'object') {
		console.warn('`shadow` options must be an object. Ex: `{shadow: {type:\'none\'}}`')
		return undefined
	}
	return { ...DEF_SHAPE_SHADOW, ...options, [shadowBrand]: true }
}

export function createShadowElement (shadow: ResolvedShadowProps | undefined): string {
	if (!shadow) {
		return '<a:effectLst/>'
	}

	let strXml = '<a:effectLst>'
	const type = shadow.type
	const blur = valToPts(shadow.blur)
	const offset = valToPts(shadow.offset)
	const angle = Math.round(shadow.angle * 60000)
	const color = shadow.color
	const opacity = Math.round(shadow.opacity * 100000)
	const rotShape = shadow.rotateWithShape ? 1 : 0

	strXml += `<a:${type}Shdw sx="100000" sy="100000" kx="0" ky="0"  algn="bl" blurRad="${blur}" rotWithShape="${rotShape}" dist="${offset}" dir="${angle}">`
	strXml += `<a:srgbClr val="${color}">`
	strXml += `<a:alpha val="${opacity}"/></a:srgbClr>`
	strXml += `</a:${type}Shdw>`
	strXml += '</a:effectLst>'

	return strXml
}

/**
 * Create Grid Line Element
 * @param {OptsChartGridLine} glOpts {size, color, style}
 * @return {string} XML
 */
export function createGridLineElement (glOpts: OptsChartGridLine): string {
	let strXml = '<c:majorGridlines>'
	strXml += ' <c:spPr>'
	strXml += `  <a:ln w="${valToPts(glOpts.size || DEF_CHART_GRIDLINE.size)}" cap="${createLineCap(glOpts.cap || DEF_CHART_GRIDLINE.cap)}">`
	strXml += '  <a:solidFill><a:srgbClr val="' + (glOpts.color || DEF_CHART_GRIDLINE.color) + '"/></a:solidFill>' // should accept scheme colors as implemented in [Pull #135]
	strXml += '   <a:prstDash val="' + (glOpts.style || DEF_CHART_GRIDLINE.style) + '"/><a:round/>'
	strXml += '  </a:ln>'
	strXml += ' </c:spPr>'
	strXml += '</c:majorGridlines>'

	return strXml
}

export function createLineCap (lineCap: ChartLineCap | undefined): string {
	if (!lineCap || lineCap === 'flat') {
		return 'flat'
	} else if (lineCap === 'square') {
		return 'sq'
	} else if (lineCap === 'round') {
		return 'rnd'
	} else {
		const neverLineCap: never = lineCap
		throw new Error(`Invalid chart line cap: ${neverLineCap}`)
	}
}

