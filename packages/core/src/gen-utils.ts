/**
 * PptxGenJS: Utility Methods
 */

import { DEF_FONT_COLOR, DEF_TEXT_GLOW, EMU, ONEPT, PATTERN_TYPES, PRESET_COLOR_VALUES, REGEX_HEX_COLOR, SCHEME_COLORS, SCHEME_COLOR_VALUES, SYSTEM_COLOR_VALUES, SchemeColor, TILE_ALIGNMENTS } from './core-enums'
import { Color, ColorProps, ColorTransformProps, Coord, PresLayout, PresSlide, ShadowProps, ShapeFillProps, ShapeGradientProps, ShapeGradientStopProps, ShapeImageFillProps, ShapeLineProps, ShapePatternProps, SlideLayout, TextGlowProps } from './core-interfaces'

/** debug namespace, used for both the log prefix and the `NODE_DEBUG` section name */
const DEBUG_NS = 'pptxgenjs'

/**
 * Whether verbose diagnostics are enabled
 * - set `PPTXGENJS_DEBUG=1`, or include `pptxgenjs` in Node's `NODE_DEBUG`
 * @returns {boolean} debug enabled
 */
export function isDebugEnabled (): boolean {
	if (typeof process === 'undefined' || !process.env) return false
	return Boolean(process.env.PPTXGENJS_DEBUG) || (process.env.NODE_DEBUG ?? '').split(/[\s,]+/).includes(DEBUG_NS)
}

/**
 * Log a diagnostic message (no-op unless debug is enabled)
 * @param {unknown[]} args - console.debug arguments
 */
export function debugLog (...args: unknown[]): void {
	if (isDebugEnabled()) console.debug(`[${DEBUG_NS}]`, ...args)
}

/** Encode bytes as base64 without referencing Node's Buffer global. */
export function bytesToBase64 (bytes: Uint8Array): string {
	let binary = ''
	const chunkSize = 0x8000
	for (let offset = 0; offset < bytes.length; offset += chunkSize) binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize))
	return btoa(binary)
}

/** Decode a base64 string or data URL without referencing Node's Buffer global. */
export function base64ToBytes (data: string): Uint8Array {
	const marker = 'base64,'
	const base64 = data.includes(marker) ? data.slice(data.indexOf(marker) + marker.length) : data
	const binary = atob(base64)
	return Uint8Array.from(binary, character => character.charCodeAt(0))
}

/** Encode UTF-8 text as base64. */
export function utf8ToBase64 (text: string): string {
	return bytesToBase64(new TextEncoder().encode(text))
}

/** Encode a Node binary string as base64. */
export function binaryStringToBase64 (text: string): string {
	return bytesToBase64(Uint8Array.from(text, character => character.charCodeAt(0) & 0xff))
}

/** friendly `dataLabelPosition` names mapped to their OOXML `c:dLblPos` codes (the codes stay accepted too) */
const DATA_LABEL_POS_CODES: Record<string, string> = {
	bottom: 'b',
	center: 'ctr',
	left: 'l',
	right: 'r',
	top: 't',
	insideEnd: 'inEnd',
	insideBase: 'inBase',
	outsideEnd: 'outEnd',
	bestFit: 'bestFit',
}

/**
 * Resolve a `dataLabelPosition` to the OOXML code valid for this chart type
 * - a value the chart type does not accept makes PowerPoint declare the file corrupt, so it is dropped
 * @param {string} position - user value: a friendly name (`'outsideEnd'`) or an OOXML code (`'outEnd'`)
 * @param {string} chartType - chart type being rendered, or undefined for a multi-type chart (translate only)
 * @param {string} barGrouping - bar grouping (stacked bars accept fewer positions than clustered)
 * @returns {string | undefined} OOXML code, or undefined when not valid for this chart type
 */
export function resolveDataLabelPosition (position: string, chartType?: string, barGrouping?: string): string | undefined {
	const code = DATA_LABEL_POS_CODES[position] ?? position

	// a multi-type chart has no single type to validate against - each sub-chart validates its own options
	if (!chartType) return code
	// REFERENCE: https://docs.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e2b1697c-7adc-463d-9081-3daef72f656f
	let valid: string[]
	switch (chartType) {
		case 'pie':
			valid = ['bestFit', 'ctr', 'inEnd', 'outEnd']
			break
		case 'bubble':
		case 'bubble3D':
		case 'line':
		case 'scatter':
			valid = ['b', 'ctr', 'l', 'r', 't']
			break
		case 'bar':
			// stacked bars have no "outside end" to sit against
			valid = (barGrouping ?? '').includes('tacked') ? ['ctr', 'inBase', 'inEnd'] : ['ctr', 'inBase', 'inEnd', 'outEnd']
			break
		default:
			// area, bar3D, doughnut, radar: PowerPoint takes no `c:dLblPos` at all
			valid = []
	}

	if (!valid.includes(code)) {
		console.warn(
			`[pptxgenjs] dataLabelPosition '${position}' is not valid for a '${chartType}' chart - ignoring it (valid: ${valid.length > 0 ? valid.join(', ') : 'none'})`
		)
		return undefined
	}

	return code
}

/** Units per inch, for the length suffixes accepted on `x`/`y`/`w`/`h` and the other inch-valued options */
const UNITS_PER_INCH: Record<string, number> = { in: 1, cm: 2.54, mm: 25.4, pt: 72 }

/**
 * A number with a length suffix - `'2.5cm'`, `'-4mm'`, `'18pt'`, `'5in'`
 * - the mantissa alternates rather than using `\d*\.?\d+`: that form lets the two digit runs
 *   split a plain integer many ways, so a long non-matching digit string backtracks quadratically
 */
const REGEX_UNIT_LENGTH = /^\s*(-?(?:\d+(?:\.\d+)?|\.\d+))\s*(in|cm|mm|pt)\s*$/i

/**
 * Parse a unit-suffixed length into inches
 * - a suffixed value carries its own unit, so it bypasses the inches-vs-EMU magnitude heuristic entirely
 * @param {string} value - ex: `'2.5cm'`, `'18pt'`
 * @returns {number|undefined} inches, or `undefined` when `value` is not a suffixed length
 */
export function parseUnitLength (value: string): number | undefined {
	const match = REGEX_UNIT_LENGTH.exec(value)
	if (!match) return undefined
	const num = Number(match[1])
	if (!Number.isFinite(num)) return undefined
	return num / UNITS_PER_INCH[match[2].toLowerCase()]
}

/**
 * Translates any type of `x`/`y`/`w`/`h` prop to EMU
 * - guaranteed to return a result regardless of undefined, null, etc. (0)
 * - {number} - 12800 (EMU)
 * - {number} - 0.5 (inches)
 * - {string} - "75%"
 * - {string} - "2.5cm" (also in/mm/pt)
 * @param {number|string} size - numeric ("5.5"), unit-suffixed ("2.5cm") or percentage ("90%")
 * @param {'X' | 'Y'} xyDir - direction
 * @param {PresLayout} layout - presentation layout
 * @returns {number} calculated size
 */
export function getSmartParseNumber (size: Coord | undefined, xyDir: 'X' | 'Y', layout: PresLayout): number {
	// FIRST: Convert string numeric value if reqd
	if (typeof size === 'string' && !isNaN(Number(size))) size = Number(size)

	// CASE 1: Number in inches
	// Assume any number less than 100 is inches
	if (typeof size === 'number' && size < 100) return inch2Emu(size)

	// CASE 2: Number is already converted to something other than inches
	// Assume any number greater than 100 sure isnt inches! Just return it (assume value is EMU already).
	if (typeof size === 'number' && size >= 100) return size

	// CASE 3: Unit-suffixed length (ex: '2.5cm') - unambiguous, so no magnitude heuristic
	if (typeof size === 'string') {
		const inches = parseUnitLength(size)
		if (inches !== undefined) return Math.round(EMU * inches)
	}

	// CASE 4: Percentage (ex: '50%')
	if (typeof size === 'string' && size.includes('%')) {
		if (xyDir && xyDir === 'X') return Math.round((parseFloat(size) / 100) * layout.width)
		if (xyDir && xyDir === 'Y') return Math.round((parseFloat(size) / 100) * layout.height)

		// Default: Assume width (x/cx)
		return Math.round((parseFloat(size) / 100) * layout.width)
	}

	// LAST: Default value
	return 0
}

/**
 * Basic UUID Generator Adapted
 * @link https://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript#answer-2117523
 * @param {string} uuidFormat - UUID format
 * @returns {string} UUID
 */
export function getUuid (uuidFormat: string): string {
	return uuidFormat.replace(/[xy]/g, function (c) {
		// Web Crypto API - a global in browsers and in Node >=20 (this package's engines floor).
		// Mask the low nibble (0-15) - an unbiased reduction (no modulo) for a hex digit.
		const r = globalThis.crypto.getRandomValues(new Uint8Array(1))[0] & 0x0f
		const v = c === 'x' ? r : (r & 0x3) | 0x8
		return v.toString(16)
	})
}

/** deprecation keys already warned about - each fires once per process */
const _warnedDeprecations = new Set<string>()

/**
 * Warn about a deprecated option/usage - once per key, so migration guidance appears without flooding the console
 * @param {string} key - unique key for this deprecation
 * @param {string} message - migration guidance
 */
export function warnDeprecatedOnce (key: string, message: string): void {
	if (_warnedDeprecations.has(key)) return
	_warnedDeprecations.add(key)
	console.warn(`[pptxgenjs] DEPRECATED: ${message}`)
}

/**
 * Replace special XML characters with HTML-encoded strings
 * @param {string} xml - XML string to encode
 * @returns {string} escaped XML
 */
export function encodeXmlEntities (xml: string | undefined): string {
	// NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
	if (typeof xml === 'undefined' || xml == null) return ''
	return xml.toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
}

/**
 * Convert inches into EMU
 * @param {number|string} inches - as string or number
 * @returns {number} EMU value
 */
export function inch2Emu (inches: number | string): number {
	// A unit-suffixed string states its own unit - convert it before any inches assumption
	if (typeof inches === 'string') {
		const parsed = parseUnitLength(inches)
		if (parsed !== undefined) return Math.round(EMU * parsed)
	}
	// NOTE: Provide Caller Safety: Numbers may get conv<->conv during flight, so be kind and do some simple checks to ensure inches were passed
	// Any value over 100 damn sure isnt inches, so lets assume its in EMU already, therefore, just return the same value
	if (typeof inches === 'number' && inches > 100) return inches
	if (typeof inches === 'string') inches = Number(inches.replace(/in*/gi, ''))
	return Math.round(EMU * inches)
}

/**
 * Convert `pt` into points (using `ONEPT`)
 * @param {number|string} pt
 * @returns {number} value in points (`ONEPT`)
 */
export function valToPts (pt: number | string | undefined): number {
	const points = Number(pt) || 0
	return isNaN(points) ? 0 : Math.round(points * ONEPT)
}

/**
 * Convert degrees (0..360) to PowerPoint `rot` value
 * @param {number} d degrees
 * @returns {number} calculated `rot` value
 */
export function convertRotationDegrees (d: number | undefined): number {
	d = d || 0
	return Math.round((d > 360 ? d - 360 : d) * 60000)
}

/**
 * Converts component value to hex value
 * @param {number} c - component color
 * @returns {string} hex string
 */
export function componentToHex (c: number): string {
	const hex = c.toString(16)
	return hex.length === 1 ? '0' + hex : hex
}

/**
 * Converts RGB colors from css selectors to Hex for Presentation colors
 * @param {number} r - red value
 * @param {number} g - green value
 * @param {number} b - blue value
 * @returns {string} XML string
 */
export function rgbToHex (r: number, g: number, b: number): string {
	return (componentToHex(r) + componentToHex(g) + componentToHex(b)).toUpperCase()
}

/**  TODO: FUTURE: TODO-4.0:
 * @date 2022-04-10
 * @tldr this s/b a private method with all current calls switched to `genXmlColorSelection()`
 * @desc lots of code calls this method
 * @example [gen-charts.tx] `strXml += '<a:solidFill>' + createColorElement(seriesColor, `<a:alpha val="${Math.round(opts.chartColorsOpacity * 1000)}"/>`) + '</a:solidFill>'`
 * Thi sis wrong. We s/b calling `genXmlColorSelection()` instead as it returns `<a:solidfill>BLAH</a:solidFill>`!!
 */
/**
 * Create either a `a:schemeClr` - (scheme color) or `a:srgbClr` (hexa representation).
 * @param {string|SCHEME_COLORS} colorStr - hexa representation (eg. "FFFF00") or a scheme color constant (eg. pptx.SchemeColor.ACCENT1)
 * @param {string} innerElements - additional elements that adjust the color and are enclosed by the color element
 * @returns {string} XML string
 */
/**
 * DrawingML stores percentages in 1000ths of a percent
 * - `CT_PositiveFixedPercentage` (tint, shade, alpha) is 0-100
 * - `CT_FixedPercentage` (the `*Off` transforms) is -100 to 100
 * - `CT_PositivePercentage` (the `*Mod` transforms) is a *scale* and is unbounded above:
 *   real themes carry values such as `satMod val="170000"`
 */
function pct (value: number, min = 0, max = 100): string {
	const clamped = Math.min(max, Math.max(min, isFinite(value) ? value : 0))
	return String(Math.round(clamped * 1000))
}

/**
 * Whether a value is a color object rather than a fill object
 * - the six specification fields are unique to `ColorProps`, so their presence is a sound
 *   discriminator against `ShapeFillProps`/`ShapeLineProps`
 * @param {unknown} value - candidate
 * @returns {boolean} whether it specifies a color
 */
export function isColorProps (value: unknown): value is ColorProps {
	if (!value || typeof value !== 'object') return false
	return ['hex', 'scheme', 'system', 'preset', 'hsl', 'scrgb'].some(key => key in value)
}

/**
 * Create the color transform children shared by every color element (ECMA-376 20.1.2.3)
 * - transforms may be combined; order is not significant, so a stable order is used
 * @param {ColorTransformProps} props - transform props
 * @returns {string} XML string
 */
function createColorTransforms (props: ColorTransformProps): string {
	let xml = ''

	// scaling percentages are positive; offsets are signed
	if (typeof props.tint === 'number') xml += `<a:tint val="${pct(props.tint)}"/>`
	if (typeof props.shade === 'number') xml += `<a:shade val="${pct(props.shade)}"/>`
	if (props.inverse === true) xml += '<a:inv/>'
	if (props.grayscale === true) xml += '<a:gray/>'
	if (typeof props.alpha === 'number') xml += `<a:alpha val="${pct(props.alpha)}"/>`
	if (typeof props.alphaOff === 'number') xml += `<a:alphaOff val="${pct(props.alphaOff, -100, 100)}"/>`
	if (typeof props.alphaMod === 'number') xml += `<a:alphaMod val="${pct(props.alphaMod, 0, Number.MAX_SAFE_INTEGER)}"/>`
	if (typeof props.hueMod === 'number') xml += `<a:hueMod val="${pct(props.hueMod, 0, Number.MAX_SAFE_INTEGER)}"/>`
	// ST_Angle: 60000ths of a degree
	if (typeof props.hueOff === 'number') xml += `<a:hueOff val="${Math.round(Math.min(360, Math.max(-360, isFinite(props.hueOff) ? props.hueOff : 0)) * 60000)}"/>`
	if (typeof props.satMod === 'number') xml += `<a:satMod val="${pct(props.satMod, 0, Number.MAX_SAFE_INTEGER)}"/>`
	if (typeof props.satOff === 'number') xml += `<a:satOff val="${pct(props.satOff, -100, 100)}"/>`
	if (typeof props.lumMod === 'number') xml += `<a:lumMod val="${pct(props.lumMod, 0, Number.MAX_SAFE_INTEGER)}"/>`
	if (typeof props.lumOff === 'number') xml += `<a:lumOff val="${pct(props.lumOff, -100, 100)}"/>`
	if (props.complement === true) xml += '<a:comp/>'
	if (props.gamma === true) xml += '<a:gamma/>'
	if (props.inverseGamma === true) xml += '<a:invGamma/>'

	return xml
}

/**
 * Create the color element for a `ColorProps` object
 * - an unknown scheme/system/preset name would make the element unparseable, so it falls back
 * @param {ColorProps} props - color props
 * @param {string} extra - additional transform XML from internal callers
 * @returns {string} XML string
 */
function createColorPropsElement (props: ColorProps, extra: string): string {
	const inner = createColorTransforms(props) + extra
	const wrap = (tag: string, attrs: string): string => (inner ? `<a:${tag} ${attrs}>${inner}</a:${tag}>` : `<a:${tag} ${attrs}/>`)

	if ('hex' in props) {
		const hex = String(props.hex ?? '').replace('#', '')
		if (!REGEX_HEX_COLOR.test(hex)) {
			console.warn(`[pptxgenjs] "${hex}" is not a 6-digit hex color - "${DEF_FONT_COLOR}" used instead`)
			return wrap('srgbClr', `val="${DEF_FONT_COLOR}"`)
		}
		return wrap('srgbClr', `val="${hex.toUpperCase()}"`)
	}
	if ('scheme' in props) {
		if (!SCHEME_COLOR_VALUES.has(props.scheme)) {
			console.warn(`[pptxgenjs] "${String(props.scheme)}" is not a theme color slot - "${DEF_FONT_COLOR}" used instead`)
			return wrap('srgbClr', `val="${DEF_FONT_COLOR}"`)
		}
		return wrap('schemeClr', `val="${props.scheme}"`)
	}
	if ('system' in props) {
		if (!SYSTEM_COLOR_VALUES.has(props.system)) {
			console.warn(`[pptxgenjs] "${String(props.system)}" is not a system color - "${DEF_FONT_COLOR}" used instead`)
			return wrap('srgbClr', `val="${DEF_FONT_COLOR}"`)
		}
		const last = props.lastColor && REGEX_HEX_COLOR.test(props.lastColor.replace('#', '')) ? ` lastClr="${props.lastColor.replace('#', '').toUpperCase()}"` : ''
		return wrap('sysClr', `val="${props.system}"${last}`)
	}
	if ('preset' in props) {
		if (!PRESET_COLOR_VALUES.has(props.preset)) {
			console.warn(`[pptxgenjs] "${String(props.preset)}" is not a preset color name - "${DEF_FONT_COLOR}" used instead`)
			return wrap('srgbClr', `val="${DEF_FONT_COLOR}"`)
		}
		return wrap('prstClr', `val="${props.preset}"`)
	}
	if ('hsl' in props) {
		// `hue` is ST_PositiveFixedAngle (0-21599999); sat/lum are percentages
		const hue = Math.round(((((props.hsl.hue % 360) + 360) % 360) || 0) * 60000)
		return wrap('hslClr', `hue="${isFinite(hue) ? hue : 0}" sat="${pct(props.hsl.sat)}" lum="${pct(props.hsl.lum)}"`)
	}
	// scrgb: linear-gamma percentages
	return wrap('scrgbClr', `r="${pct(props.scrgb.r)}" g="${pct(props.scrgb.g)}" b="${pct(props.scrgb.b)}"`)
}

export function createColorElement (colorStr: Color | SCHEME_COLORS | undefined, innerElements?: string): string {
	// The object form covers the whole DrawingML color model; the string forms are unchanged
	if (isColorProps(colorStr)) return createColorPropsElement(colorStr, innerElements ?? '')

	let colorVal = (colorStr ?? '').replace('#', '')

	if (
		!REGEX_HEX_COLOR.test(colorVal) &&
		colorVal !== SchemeColor.background1 &&
		colorVal !== SchemeColor.background2 &&
		colorVal !== SchemeColor.text1 &&
		colorVal !== SchemeColor.text2 &&
		colorVal !== SchemeColor.accent1 &&
		colorVal !== SchemeColor.accent2 &&
		colorVal !== SchemeColor.accent3 &&
		colorVal !== SchemeColor.accent4 &&
		colorVal !== SchemeColor.accent5 &&
		colorVal !== SchemeColor.accent6
	) {
		console.warn(`"${colorVal}" is not a valid scheme color or hex RGB! "${DEF_FONT_COLOR}" used instead. Only provide 6-digit RGB or 'pptx.SchemeColor' values!`)
		colorVal = DEF_FONT_COLOR
	}

	const tagName = REGEX_HEX_COLOR.test(colorVal) ? 'srgbClr' : 'schemeClr'
	const colorAttr = 'val="' + (REGEX_HEX_COLOR.test(colorVal) ? colorVal.toUpperCase() : colorVal) + '"'

	return innerElements ? `<a:${tagName} ${colorAttr}>${innerElements}</a:${tagName}>` : `<a:${tagName} ${colorAttr}/>`
}

/**
 * Creates `a:glow` element
 * @param {TextGlowProps} options glow properties
 * @param {TextGlowProps} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 * { size: 8, color: 'FFFFFF', opacity: 0.75 };
 */
/**
 * Nominal ("coloured") brand for resolved glow options. The symbol is module-private and unexported,
 * so `ResolvedGlowProps` values can ONLY be produced by `resolveGlowOptions` below - a hand-built
 * object (even a full `Required<TextGlowProps>`) is not assignable. This statically guarantees that
 * anything reaching `createGlowElement` has passed through the defaults-merge boundary.
 */
const glowBrand: unique symbol = Symbol('resolvedGlow')
export type ResolvedGlowProps = Required<TextGlowProps> & { readonly [glowBrand]: boolean }

/**
 * Resolve boundary: merge user glow options over the documented defaults and brand the result. The
 * only constructor of `ResolvedGlowProps` (no cast - the brand is added by this factory).
 */
export function resolveGlowOptions (options: TextGlowProps | undefined): ResolvedGlowProps | undefined {
	if (!options) return undefined
	return { ...DEF_TEXT_GLOW, ...options, [glowBrand]: true }
}

export function createGlowElement (glow: ResolvedGlowProps): string {
	let strXml = ''
	const size = Math.round(glow.size * ONEPT)
	const color = glow.color
	const opacity = Math.round(glow.opacity * 100000)

	strXml += `<a:glow rad="${size}">`
	strXml += createColorElement(color, `<a:alpha val="${opacity}"/>`)
	strXml += '</a:glow>'

	return strXml
}

/** Clamp a user-supplied percent (0-100) and drop non-finite values. */
function clampPercent (value: number | undefined, fallback: number): number {
	if (typeof value !== 'number' || !isFinite(value)) return fallback
	return Math.min(100, Math.max(0, value))
}

/**
 * Create a DrawingML `a:gradFill` element
 * - `a:gs@pos` and `a:alpha@val` are ST_Percentage in 1000ths of a percent (0-100000)
 * - `a:lin@ang` is ST_PositiveFixedAngle in 60000ths of a degree (0-21599999)
 * @param {ShapeGradientProps} gradient - gradient props (already known to have 2+ stops)
 * @returns {string} XML string
 */
function createGradientFillElement (gradient: ShapeGradientProps): string {
	const stops = [...gradient.stops]
		.map((stop, idx): ShapeGradientStopProps & { _idx: number } => ({ ...stop, _idx: idx }))
		.sort((a, b) => clampPercent(a.position, 0) - clampPercent(b.position, 0) || a._idx - b._idx)
	const gsLst = stops
		.map(stop => {
			const pos = Math.round(clampPercent(stop.position, 0) * 1000)
			const alpha = typeof stop.transparency === 'number' ? `<a:alpha val="${Math.round((100 - clampPercent(stop.transparency, 0)) * 1000)}"/>` : ''
			return `<a:gs pos="${pos}">${createColorElement(stop.color, alpha)}</a:gs>`
		})
		.join('')

	// `a:lin@ang` is ST_PositiveFixedAngle, so normalize into 0-359 before scaling
	const angleDeg = typeof gradient.angle === 'number' && isFinite(gradient.angle) ? (((gradient.angle % 360) + 360) % 360) : 90
	const geometry =
		gradient.type === 'radial'
			? '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>'
			: `<a:lin ang="${Math.round(angleDeg * 60000)}" scaled="${gradient.scaled === true ? 1 : 0}"/>`

	return `<a:gradFill rotWithShape="${gradient.rotateWithShape === false ? 0 : 1}"><a:gsLst>${gsLst}</a:gsLst>${geometry}</a:gradFill>`
}

/**
 * Create a DrawingML `a:pattFill` element (ECMA-376 20.1.8.47)
 * @param {ShapePatternProps} pattern - pattern props
 * @returns {string} XML string
 */
function createPatternFillElement (pattern: ShapePatternProps): string {
	// `prst` is an enum: an unknown value makes the element unparseable, so fall back to a safe preset
	const preset = PATTERN_TYPES.has(pattern.preset) ? pattern.preset : 'pct50'
	if (!PATTERN_TYPES.has(pattern.preset)) {
		console.warn(`[pptxgenjs] unknown fill pattern "${String(pattern.preset)}" - "pct50" used instead`)
	}

	return (
		`<a:pattFill prst="${preset}">` +
		`<a:fgClr>${createColorElement(pattern.color ?? '000000')}</a:fgClr>` +
		`<a:bgClr>${createColorElement(pattern.backColor ?? 'FFFFFF')}</a:bgClr>` +
		'</a:pattFill>'
	)
}

/**
 * Create a DrawingML `a:blipFill` element for a picture fill (ECMA-376 20.1.8.14)
 * - the image relationship is resolved when the object is created, so `_rId` is set by then
 * @param {ShapeImageFillProps} image - picture fill props
 * @returns {string} XML string
 */
function createImageFillElement (image: ShapeImageFillProps): string {
	// `stretch` scales the image to the shape (20.1.8.56); `tile` repeats it (20.1.8.58)
	let mode = '<a:stretch><a:fillRect/></a:stretch>'
	if (image.sizing === 'tile') {
		const scale = typeof image.scale === 'number' && isFinite(image.scale) && image.scale > 0 ? Math.round(image.scale * 1000) : 100000
		const algn = TILE_ALIGNMENTS.has(image.alignment ?? 'tl') ? (image.alignment ?? 'tl') : 'tl'
		if (image.alignment && !TILE_ALIGNMENTS.has(image.alignment)) {
			console.warn(`[pptxgenjs] unknown tile alignment "${String(image.alignment)}" - "tl" used instead`)
		}
		mode = `<a:tile tx="0" ty="0" sx="${scale}" sy="${scale}" flip="none" algn="${algn}"/>`
	}

	return `<a:blipFill rotWithShape="${image.rotateWithShape === false ? '0' : '1'}"><a:blip r:embed="rId${image._rId ?? 0}"/>${mode}</a:blipFill>`
}

/**
 * Create color selection
 * @param {Color | ShapeFillProps | ShapeLineProps} props fill props
 * @returns XML string
 */
export function genXmlColorSelection (props: Color | ShapeFillProps | ShapeLineProps | undefined): string {
	let fillType = 'solid'
	let colorVal: Color = ''
	let internalElements = ''
	let outText = ''
	let gradient: ShapeGradientProps | undefined
	let pattern: ShapePatternProps | undefined
	let image: ShapeImageFillProps | undefined

	if (props) {
		// A bare color - string or object - is a solid fill of that color; only a fill object
		// carries `type`/`gradient`/`pattern`/`image`
		if (typeof props === 'string' || isColorProps(props)) colorVal = props
		else {
			if (props.type) fillType = props.type
			if (props.gradient) gradient = props.gradient
			if (props.pattern) pattern = props.pattern
			if (props.image) image = props.image
			if (props.color) colorVal = props.color
			if (props.alpha) internalElements += `<a:alpha val="${Math.round((100 - props.alpha) * 1000)}"/>` // DEPRECATED: @deprecated v3.3.0
			if (props.transparency) internalElements += `<a:alpha val="${Math.round((100 - props.transparency) * 1000)}"/>`
		}

		switch (fillType) {
			case 'solid':
				outText += `<a:solidFill>${createColorElement(colorVal, internalElements)}</a:solidFill>`
				break
			case 'gradient':
				// MS-PPT requires 2+ stops; anything less is not a gradient, so degrade to the solid fill path
				if (gradient && Array.isArray(gradient.stops) && gradient.stops.length >= 2) {
					outText += createGradientFillElement(gradient)
				} else {
					console.warn('[pptxgenjs] `fill.type:"gradient"` requires `fill.gradient.stops` with at least 2 stops - solid fill used instead')
					outText += `<a:solidFill>${createColorElement(colorVal || gradient?.stops?.[0]?.color, internalElements)}</a:solidFill>`
				}
				break
			case 'pattern':
				if (pattern?.preset) {
					outText += createPatternFillElement(pattern)
				} else {
					console.warn('[pptxgenjs] `fill.type:"pattern"` requires `fill.pattern.preset` - solid fill used instead')
					outText += `<a:solidFill>${createColorElement(colorVal, internalElements)}</a:solidFill>`
				}
				break
			case 'group':
				// inherit the group shape's fill; only meaningful on a shape inside a group
				outText += '<a:grpFill/>'
				break
			case 'image':
				// without a relationship the `a:blip` would dangle, which PowerPoint reports as damage
				if (image?._rId) {
					outText += createImageFillElement(image)
				} else {
					console.warn('[pptxgenjs] `fill.type:"image"` requires `fill.image.data` or `fill.image.path` - fill omitted')
				}
				break
			default: // @note need a statement as having only "break" is removed by rollup, then tiggers "no-default" js-linter
				outText += ''
				break
		}
	}

	return outText
}

/**
 * Get a new rel ID (rId) for charts, media, etc.
 * @param {PresSlide} target - the slide to use
 * @returns {number} count of all current rels plus 1 for the caller to use as its "rId"
 */
export function getNewRelId (target: PresSlide | SlideLayout): number {
	// every relationship store on the slide has to be counted, or two of them collide on one rId
	return target._rels.length + target._relsChart.length + target._relsMedia.length + (target._contentParts?.length ?? 0) + 1
}

/**
 * Checks shadow options passed by user and performs corrections if needed.
 * @param {ShadowProps} ShadowProps - shadow options
 */
export function correctShadowOptions (ShadowProps: ShadowProps): ShadowProps | undefined {
	if (!ShadowProps || typeof ShadowProps !== 'object') {
		if (ShadowProps) console.warn('[pptxgenjs] `shadow` must be an object (ex: `{shadow: {type:\'outer\'}}`) - value ignored')
		return
	}

	// Work on a copy - never mutate the caller's options object
	ShadowProps = { ...ShadowProps }

	// OPT: `type`
	if (ShadowProps.type !== 'outer' && ShadowProps.type !== 'inner' && ShadowProps.type !== 'none') {
		console.warn('Warning: shadow.type options are `outer`, `inner` or `none`.')
		ShadowProps.type = 'outer'
	}

	// OPT: `angle`
	if (ShadowProps.angle) {
		// A: REALITY-CHECK
		if (isNaN(Number(ShadowProps.angle)) || ShadowProps.angle < 0 || ShadowProps.angle > 359) {
			console.warn('Warning: shadow.angle can only be 0-359')
			ShadowProps.angle = 270
		}

		// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
		ShadowProps.angle = Math.round(Number(ShadowProps.angle))
	}

	// OPT: `opacity`
	if (ShadowProps.opacity) {
		// A: REALITY-CHECK
		if (isNaN(Number(ShadowProps.opacity)) || ShadowProps.opacity < 0 || ShadowProps.opacity > 1) {
			console.warn('Warning: shadow.opacity can only be 0-1')
			ShadowProps.opacity = 0.75
		}

		// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
		ShadowProps.opacity = Number(ShadowProps.opacity)
	}

	// OPT: `color`
	if (ShadowProps.color) {
		// INCORRECT FORMAT
		if (ShadowProps.color.startsWith('#')) {
			console.warn('Warning: shadow.color should not include hash (#) character, , e.g. "FF0000"')
			ShadowProps.color = ShadowProps.color.replace('#', '')
		}
	}

	return ShadowProps
}
