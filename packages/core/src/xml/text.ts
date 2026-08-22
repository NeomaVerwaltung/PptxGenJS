/**
 * OOXML text and placeholder rendering.
 */

import { BULLET_TYPES, CRLF, DEF_BULLET_MARGIN, OOXML_EXT, PLACEHOLDER_TYPES, SLIDE_OBJECT_TYPES } from '../core-enums'
import { ISlideObject, ObjectOptions, ParagraphProps, TableCell, TextBodyProps, TextProps, TextPropsOptions, TextRunProps } from '../core-interfaces'
import { genXmlHyperlink } from './hyperlink'
import { genXmlLine } from './line'
import { alternateContent } from './markup-compat'
import { convertRotationDegrees, createColorElement, createGlowElement, encodeXmlEntities, genXmlColorSelection, inch2Emu, resolveGlowOptions, valToPts, warnDeprecatedOnce } from '../gen-utils'

function genXmlParagraphProperties (textObj: ISlideObject | TextProps, isDefault: boolean): string {
	let strXmlBullet = ''
	let strXmlLnSpc = ''
	let strXmlParaSpc = ''
	let strXmlTabStops = ''
	const tag = isDefault ? 'a:lvl1pPr' : 'a:pPr'
	let bulletMarL = valToPts(DEF_BULLET_MARGIN)
	const options: TextPropsOptions = textObj.options ?? {}

	let paragraphPropXml = `<${tag}${options.rtlMode ? ' rtl="1" ' : ''}`

	// A: Build paragraphProperties
	{
		// OPTION: align
		if (options.align) {
			switch (options.align) {
				case 'left':
					paragraphPropXml += ' algn="l"'
					break
				case 'right':
					paragraphPropXml += ' algn="r"'
					break
				case 'center':
					paragraphPropXml += ' algn="ctr"'
					break
				case 'justify':
					paragraphPropXml += ' algn="just"'
					break
				default:
					paragraphPropXml += ''
					break
			}
		}

		if (options.lineSpacing) {
			strXmlLnSpc = `<a:lnSpc><a:spcPts val="${Math.round(options.lineSpacing * 100)}"/></a:lnSpc>`
		} else if (options.lineSpacingMultiple) {
			strXmlLnSpc = `<a:lnSpc><a:spcPct val="${Math.round(options.lineSpacingMultiple * 100000)}"/></a:lnSpc>`
		}

		// OPTION: indent
		if (options.indentLevel && !isNaN(Number(options.indentLevel)) && options.indentLevel > 0) {
			paragraphPropXml += ` lvl="${options.indentLevel}"`
		}

		// Remaining CT_TextParagraphProperties attributes - omitted unless set, so output is unchanged
		const para = options as ParagraphProps
		if (typeof para.marginRight === 'number' && isFinite(para.marginRight) && para.marginRight >= 0) paragraphPropXml += ` marR="${inch2Emu(para.marginRight)}"`
		if (typeof para.defaultTabSize === 'number' && isFinite(para.defaultTabSize) && para.defaultTabSize > 0) paragraphPropXml += ` defTabSz="${inch2Emu(para.defaultTabSize)}"`
		if (['auto', 't', 'ctr', 'base', 'b'].includes(String(para.fontAlign))) paragraphPropXml += ` fontAlgn="${String(para.fontAlign)}"`
		// these three default to true in the schema, so only write them when turned off
		if (para.eastAsianLineBreak === false) paragraphPropXml += ' eaLnBrk="0"'
		if (para.latinLineBreak === false) paragraphPropXml += ' latinLnBrk="0"'
		if (para.hangingPunctuation === false) paragraphPropXml += ' hangingPunct="0"'

		// OPTION: Paragraph Spacing: Before/After
		if (options.paraSpaceBefore && !isNaN(Number(options.paraSpaceBefore)) && options.paraSpaceBefore > 0) {
			strXmlParaSpc += `<a:spcBef><a:spcPts val="${Math.round(options.paraSpaceBefore * 100)}"/></a:spcBef>`
		}
		if (options.paraSpaceAfter && !isNaN(Number(options.paraSpaceAfter)) && options.paraSpaceAfter > 0) {
			strXmlParaSpc += `<a:spcAft><a:spcPts val="${Math.round(options.paraSpaceAfter * 100)}"/></a:spcAft>`
		}

		// OPTION: bullet
		// NOTE: OOXML uses the unicode character set for Bullets
		// EX: Unicode Character 'BULLET' (U+2022) ==> '<a:buChar char="&#x2022;"/>'
		if (typeof options.bullet === 'object') {
			if (options.bullet?.indent) bulletMarL = valToPts(options.bullet.indent)

			if (options.bullet.type && options.bullet.type.toString().toLowerCase() === 'number') {
				// NOTE: only `type: 'number'` is a distinct branch; any other `type` (e.g. 'bullet') falls through to the char-bullet cases below (issue #1432)
				paragraphPropXml += ` marL="${options.indentLevel && options.indentLevel > 0 ? bulletMarL + bulletMarL * options.indentLevel : bulletMarL
				}" indent="-${bulletMarL}"`
				strXmlBullet = `<a:buSzPct val="100000"/><a:buFont typeface="+mj-lt"/><a:buAutoNum type="${options.bullet.style || 'arabicPeriod'}" startAt="${options.bullet.numberStartAt || options.bullet.startAt || '1'
				}"/>`
			} else if (options.bullet.characterCode) {
				let bulletCode = `&#x${options.bullet.characterCode};`

				// Check value for hex-ness (s/b 4 char hex)
				if (!/^[0-9A-Fa-f]{4}$/.test(options.bullet.characterCode)) {
					console.warn('Warning: `bullet.characterCode should be a 4-digit unicode charatcer (ex: 22AB)`!')
					bulletCode = BULLET_TYPES.DEFAULT
				}

				paragraphPropXml += ` marL="${options.indentLevel && options.indentLevel > 0 ? bulletMarL + bulletMarL * options.indentLevel : bulletMarL
				}" indent="-${bulletMarL}"`
				strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="' + bulletCode + '"/>'
			} else if (options.bullet.code) {
				// @deprecated `bullet.code` v3.3.0
				let bulletCode = `&#x${options.bullet.code};`

				// Check value for hex-ness (s/b 4 char hex)
				if (!/^[0-9A-Fa-f]{4}$/.test(options.bullet.code)) {
					console.warn('Warning: `bullet.code should be a 4-digit hex code (ex: 22AB)`!')
					bulletCode = BULLET_TYPES.DEFAULT
				}

				paragraphPropXml += ` marL="${options.indentLevel && options.indentLevel > 0 ? bulletMarL + bulletMarL * options.indentLevel : bulletMarL
				}" indent="-${bulletMarL}"`
				strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="' + bulletCode + '"/>'
			} else {
				paragraphPropXml += ` marL="${options.indentLevel && options.indentLevel > 0 ? bulletMarL + bulletMarL * options.indentLevel : bulletMarL
				}" indent="-${bulletMarL}"`
				strXmlBullet = `<a:buSzPct val="100000"/><a:buChar char="${BULLET_TYPES.DEFAULT}"/>`
			}
		} else if (options.bullet) {
			paragraphPropXml += ` marL="${options.indentLevel && options.indentLevel > 0 ? bulletMarL + bulletMarL * options.indentLevel : bulletMarL
			}" indent="-${bulletMarL}"`
			strXmlBullet = `<a:buSzPct val="100000"/><a:buChar char="${BULLET_TYPES.DEFAULT}"/>`
		} else if (!options.bullet) {
			// We only add this when the user explicitely asks for no bullet, otherwise, it can override the master defaults!
			paragraphPropXml += ' indent="0" marL="0"' // FIX: ISSUE#589 - specify zero indent and marL or default will be hanging paragraph
			strXmlBullet = '<a:buNone/>'
		}

		// OPTION: tabStops
		if (options.tabStops && Array.isArray(options.tabStops)) {
			const tabStopsXml = options.tabStops.map(stop => `<a:tab pos="${inch2Emu(stop.position || 1)}" algn="${stop.alignment || 'l'}"/>`).join('')
			strXmlTabStops = `<a:tabLst>${tabStopsXml}</a:tabLst>`
		}

		// B: Close Paragraph-Properties
		// IMPORTANT: strXmlLnSpc, strXmlParaSpc, and strXmlBullet require strict ordering - anything out of order is ignored. (PPT-Online, PPT for Mac)
		paragraphPropXml += '>' + strXmlLnSpc + strXmlParaSpc + strXmlBullet + strXmlTabStops
		if (isDefault) paragraphPropXml += genXmlTextRunProperties(options, true)
		paragraphPropXml += '</' + tag + '>'
	}

	return paragraphPropXml
}

/**
 * Generate XML Text Run Properties (`a:rPr`)
 * @param {ObjectOptions|TextPropsOptions} opts - text options
 * @param {boolean} isDefault - whether these are the default text run properties
 * @return {string} XML
 */
function genXmlTextRunProperties (opts: ObjectOptions | TextPropsOptions, isDefault: boolean): string {
	let runProps = ''
	const runPropsTag = isDefault ? 'a:defRPr' : 'a:rPr'

	// BEGIN runProperties (ex: `<a:rPr lang="en-US" sz="1600" b="1" dirty="0">`)
	runProps += '<' + runPropsTag + ' lang="' + (opts.lang ? opts.lang : 'en-US') + '"' + (opts.lang ? ' altLang="en-US"' : '')
	runProps += opts.fontSize ? ` sz="${Math.round(opts.fontSize * 100)}"` : '' // NOTE: Use round so sizes like '7.5' wont cause corrupt presentations
	runProps += opts?.bold ? ` b="${opts.bold ? '1' : '0'}"` : ''
	runProps += opts?.italic ? ` i="${opts.italic ? '1' : '0'}"` : ''

	if (opts?.strike === true) warnDeprecatedOnce('strike-boolean', '`strike: true` is deprecated - use `strike: "sngStrike"` (or `"dblStrike"`)')
	runProps += opts?.strike ? ` strike="${typeof opts.strike === 'string' ? opts.strike : 'sngStrike'}"` : ''
	if (typeof opts.underline === 'object' && opts.underline?.style) {
		runProps += ` u="${opts.underline.style}"`
	} else if (typeof opts.underline === 'string') {
		// DEPRECATED: opts.underline is an object as of v3.5.0
		runProps += ` u="${String(opts.underline)}"`
	} else if (opts.hyperlink) {
		runProps += ' u="sng"'
	}
	if (opts.baseline) {
		runProps += ` baseline="${Math.round(opts.baseline * 50)}"`
	} else if (opts.subscript) {
		runProps += ' baseline="-40000"'
	} else if (opts.superscript) {
		runProps += ' baseline="30000"'
	}
	runProps += opts.charSpacing ? ` spc="${Math.round(opts.charSpacing * 100)}" kern="0"` : '' // IMPORTANT: Also disable kerning; otherwise text won't actually expand
	// Remaining CT_TextCharacterProperties attributes - omitted unless set, so output is unchanged
	const run = opts as TextRunProps
	if (['none', 'small', 'all'].includes(String(run.capitalization))) runProps += ` cap="${String(run.capitalization)}"`
	if (run.normalizeHeight === true) runProps += ' normalizeH="1"'
	if (run.noProof === true) runProps += ' noProof="1"'
	// `dirty` was hardcoded to 0; it stays the default so existing output does not change
	runProps += ` dirty="${run.dirty === true ? '1' : '0'}">`
	// Color / Font / Highlight / Outline are children of <a:rPr>, so add them now before closing the runProperties tag
	const perScript = run.latinFontFace ?? run.eastAsianFontFace ?? run.complexScriptFontFace
	if (opts.color || opts.fontFace || opts.outline || perScript || run.underlineLine || run.symbolFontFace || (typeof opts.underline === 'object' && opts.underline.color)) {
		if (opts.outline && typeof opts.outline === 'object') {
			runProps += `<a:ln w="${valToPts(opts.outline.size || 0.75)}">${genXmlColorSelection(opts.outline.color || 'FFFFFF')}</a:ln>`
		}
		if (opts.color) runProps += genXmlColorSelection({ color: opts.color, transparency: opts.transparency })
		/* CT_TextCharacterProperties fixes the child order: ln, fill, effect, highlight, uLn, uFill,
		 * latin, ea, cs, sym. The glow effect was emitted last, after `a:uFill`, which is out of
		 * sequence - PowerPoint tolerated it, but it is invalid against the schema. */
		const resolvedGlow = resolveGlowOptions(opts.glow)
		if (resolvedGlow) runProps += `<a:effectLst>${createGlowElement(resolvedGlow)}</a:effectLst>`
		if (opts.highlight) runProps += `<a:highlight>${createColorElement(opts.highlight)}</a:highlight>`
		// `a:uLn` describes the underline's line; `a:uLnTx` follows the run's own line instead
		if (run.underlineLine === 'text') runProps += '<a:uLnTx/>'
		else if (run.underlineLine && typeof run.underlineLine === 'object') runProps += genXmlLine(run.underlineLine, 'a:uLn')
		if (typeof opts.underline === 'object' && opts.underline.color) runProps += `<a:uFill>${genXmlColorSelection(opts.underline.color)}</a:uFill>`
		if (opts.fontFace || perScript) {
			// NOTE: 'cs' = Complex Script, 'ea' = East Asian (use "-120" instead of "0" - per Issue #174); ea must come first (Issue #174)
			// `fontFace` sets all three scripts; the per-script options override individual ones
			const latin = run.latinFontFace ?? opts.fontFace
			const eastAsian = run.eastAsianFontFace ?? opts.fontFace
			const complex = run.complexScriptFontFace ?? opts.fontFace
			if (latin) runProps += `<a:latin typeface="${latin}" pitchFamily="34" charset="0"/>`
			if (eastAsian) runProps += `<a:ea typeface="${eastAsian}" pitchFamily="34" charset="-122"/>`
			if (complex) runProps += `<a:cs typeface="${complex}" pitchFamily="34" charset="-120"/>`
		}
		// `a:sym` must follow the script typefaces per CT_TextCharacterProperties
		if (run.symbolFontFace) runProps += `<a:sym typeface="${encodeXmlEntities(run.symbolFontFace)}"/>`
	}

	// Hyperlink support
	if (opts.hyperlink) {
		if (typeof opts.hyperlink !== 'object') throw new Error('ERROR: text `hyperlink` option should be an object. Ex: `hyperlink:{url:\'https://github.com\'}` ')
		else if (!opts.hyperlink.url && !opts.hyperlink.slide) throw new Error('ERROR: \'hyperlink requires either `url` or `slide`\'')
		// A colored link keeps its run color instead of the theme's hyperlink color (issue #74: an
		// `a:uFill` here breaks PPT2010, so the extension is used rather than an underline fill)
		const linkColorExt = opts.color
			? `<a:extLst><a:ext uri="${OOXML_EXT.hyperlinkColor.uri}"><ahyp:hlinkClr xmlns:ahyp="${OOXML_EXT.hyperlinkColor.ns}" val="tx"/></a:ext></a:extLst>`
			: ''
		runProps += genXmlHyperlink(opts.hyperlink, 'click', 'run', linkColorExt)
	}
	// `a:hlinkMouseOver` follows `a:hlinkClick` in CT_TextCharacterProperties
	if (opts.hyperlinkHover?._rId) runProps += genXmlHyperlink(opts.hyperlinkHover, 'hover', 'run')

	// END runProperties
	runProps += `</${runPropsTag}>`

	return runProps
}

/**
 * Build textBody text runs [`<a:r></a:r>`] for paragraphs [`<a:p>`]
 * @param {TextProps} textObj - Text object
 * @return {string} XML string
 */
/**
 * Namespaces for Office Math runs
 * - `a14` is the DrawingML 2010 extension that carries a math zone inside a text body
 * - `w` is declared defensively: LaTeX/MathML converters commonly emit `w:`-prefixed nodes
 *   inside OMML, and an undeclared prefix makes the whole part unparseable
 */
const OMML_NS = {
	a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
	m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
	mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
	w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
} as const

/**
 * Normalize caller-supplied OMML into a DrawingML math zone (`a14:m`)
 * - an inner fragment (ex: `<m:f>...</m:f>`) is given an `m:oMath` root
 * - namespace declarations are added when the caller omitted them
 * - an already-wrapped `a14:m` value is passed through untouched
 * @note the caller owns well-formedness: PptxGenJS does not parse the fragment
 * @param {string} omml - caller-supplied OMML
 * @returns {string} XML string
 */
function normalizeOmml (omml: string): string {
	let math = omml.trim()
	if (!math) return ''
	if (/^<a14:m[\s/>]/i.test(math)) return math

	if (!/^<m:oMath(Para)?[\s/>]/i.test(math)) {
		math = `<m:oMath xmlns:m="${OMML_NS.m}" xmlns:w="${OMML_NS.w}">${math}</m:oMath>`
	} else if (!/xmlns:m=/i.test(math)) {
		math = math.replace(/^<m:oMath(Para)?/i, match => `${match} xmlns:m="${OMML_NS.m}" xmlns:w="${OMML_NS.w}"`)
	}

	return `<a14:m>${math}</a14:m>`
}

function genXmlTextRun (textObj: TextProps): string {
	// NOTE: Dont create full rPr runProps for empty [lineBreak] runs
	// Why? The size of the lineBreak wont match (eg: below it will be 18px instead of the correct 36px)
	// Do this:
	/*
		<a:p>
			<a:pPr algn="r"/>
			<a:endParaRPr lang="en-US" sz="3600" dirty="0"/>
		</a:p>
	*/
	// NOT this:
	/*
		<a:p>
			<a:pPr algn="r"/>
			<a:r>
				<a:rPr lang="en-US" sz="3600" dirty="0">
					<a:solidFill>
						<a:schemeClr val="accent5"/>
					</a:solidFill>
					<a:latin typeface="Times" pitchFamily="34" charset="0"/>
					<a:ea typeface="Times" pitchFamily="34" charset="-122"/>
					<a:cs typeface="Times" pitchFamily="34" charset="-120"/>
				</a:rPr>
				<a:t></a:t>
			</a:r>
			<a:endParaRPr lang="en-US" dirty="0"/>
		</a:p>
	*/

	// Return paragraph with text run
	const plainRun = textObj.text ? `<a:r>${genXmlTextRunProperties(textObj.options ?? {}, false)}<a:t>${encodeXmlEntities(textObj.text)}</a:t></a:r>` : ''

	/* Office Math run: PowerPoint stores math in a DrawingML 2010 `a14:m` zone, which is an extension
	 * to the text-body content model. ECMA-376 Part 3 (Markup Compatibility) requires an extension to
	 * be offered through `mc:AlternateContent`, so consumers that do not understand `a14` render the
	 * `mc:Fallback` plain run instead of silently dropping the math. */
	const math = normalizeOmml(typeof textObj.options?.omml === 'string' ? textObj.options.omml : '')
	if (math) {
		return alternateContent({ namespaces: { a14: OMML_NS.a14 }, choice: math, fallback: plainRun })
	}

	return plainRun
}

/**
 * Builds `<a:bodyPr></a:bodyPr>` tag for "genXmlTextBody()"
 * @param {ISlideObject | TableCell} slideObject - various options
 * @return {string} XML string
 */
function genXmlBodyProperties (slideObject: ISlideObject | TableCell): string {
	let bodyProperties = '<a:bodyPr'

	if (slideObject && slideObject._type === SLIDE_OBJECT_TYPES.text && slideObject.options?._bodyProp) {
		// PPT-2019 EX: <a:bodyPr wrap="square" lIns="1270" tIns="1270" rIns="1270" bIns="1270" rtlCol="0" anchor="ctr"/>

		// A: Enable or disable textwrapping none or square
		bodyProperties += slideObject.options._bodyProp.wrap ? ' wrap="square"' : ' wrap="none"'

		// B: Textbox margins [padding]
		if (slideObject.options._bodyProp.lIns || slideObject.options._bodyProp.lIns === 0) bodyProperties += ` lIns="${slideObject.options._bodyProp.lIns}"`
		if (slideObject.options._bodyProp.tIns || slideObject.options._bodyProp.tIns === 0) bodyProperties += ` tIns="${slideObject.options._bodyProp.tIns}"`
		if (slideObject.options._bodyProp.rIns || slideObject.options._bodyProp.rIns === 0) bodyProperties += ` rIns="${slideObject.options._bodyProp.rIns}"`
		if (slideObject.options._bodyProp.bIns || slideObject.options._bodyProp.bIns === 0) bodyProperties += ` bIns="${slideObject.options._bodyProp.bIns}"`

		// B.2: Remaining CT_TextBodyProperties attributes - each omitted unless asked for, so
		// existing output is unchanged
		const body = slideObject.options as TextBodyProps
		if (body.upright === true) bodyProperties += ' upright="1"'
		if (typeof body.textRotate === 'number' && isFinite(body.textRotate)) bodyProperties += ` rot="${convertRotationDegrees(body.textRotate)}"`
		if (body.anchorCenter === true) bodyProperties += ' anchorCtr="1"'
		if (body.spaceFirstLastPara === true) bodyProperties += ' spcFirstLastPara="1"'
		if (body.compatLineSpacing === true) bodyProperties += ' compatLnSpc="1"'
		if (body.forceAntiAlias === true) bodyProperties += ' forceAA="1"'
		if (body.horizontalOverflow === 'overflow' || body.horizontalOverflow === 'clip') bodyProperties += ` horzOverflow="${body.horizontalOverflow}"`
		if (['overflow', 'ellipsis', 'clip'].includes(String(body.verticalOverflow))) bodyProperties += ` vertOverflow="${String(body.verticalOverflow)}"`

		// C: Columns, then rtl, after margins
		if (slideObject.options._bodyProp.numCol) bodyProperties += ` numCol="${slideObject.options._bodyProp.numCol}"`
		if (slideObject.options._bodyProp.spcCol) bodyProperties += ` spcCol="${slideObject.options._bodyProp.spcCol}"`
		bodyProperties += ' rtlCol="0"'

		// D: Add anchorPoints
		if (slideObject.options._bodyProp.anchor) bodyProperties += ' anchor="' + slideObject.options._bodyProp.anchor + '"' // VALS: [t,ctr,b]
		if (slideObject.options._bodyProp.vert) bodyProperties += ' vert="' + slideObject.options._bodyProp.vert + '"' // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]

		// E: Close <a:bodyPr element
		bodyProperties += '>'

		/**
		 * F: Text Fit/AutoFit/Shrink option
		 * @see: http://officeopenxml.com/drwSp-text-bodyPr-fit.php
		 * @see: http://www.datypic.com/sc/ooxml/g-a_EG_TextAutofit.html
		 */
		if (slideObject.options.fit) {
			// NOTE: Use of '<a:noAutofit/>' instead of '' causes issues in PPT-2013!
			if (slideObject.options.fit === 'none') bodyProperties += ''
			// NOTE: Shrink does not work automatically - PowerPoint calculates the `fontScale` value dynamically upon resize
			// else if (slideObject.options.fit === 'shrink') bodyProperties += '<a:normAutofit fontScale="85000" lnSpcReduction="20000"/>' // MS-PPT > Format shape > Text Options: "Shrink text on overflow"
			else if (slideObject.options.fit === 'shrink') bodyProperties += '<a:normAutofit/>'
			else if (slideObject.options.fit === 'resize') bodyProperties += '<a:spAutoFit/>'
		}
		//
		// DEPRECATED: below (@deprecated v3.3.0)
		if (slideObject.options.shrinkText) bodyProperties += '<a:normAutofit/>' // MS-PPT > Format shape > Text Options: "Shrink text on overflow"
		/* DEPRECATED: below (@deprecated v3.3.0)
		 * MS-PPT > Format shape > Text Options: "Resize shape to fit text" [spAutoFit]
		 * NOTE: Use of '<a:noAutofit/>' in lieu of '' below causes issues in PPT-2013
		 */
		bodyProperties += slideObject.options._bodyProp.autoFit ? '<a:spAutoFit/>' : ''

		// LAST: Close _bodyProp
		bodyProperties += '</a:bodyPr>'
	} else {
		// DEFAULT:
		bodyProperties += ' wrap="square" rtlCol="0">'
		bodyProperties += '</a:bodyPr>'
	}

	// LAST: Return Close _bodyProp
	return slideObject._type === SLIDE_OBJECT_TYPES.tablecell ? '<a:bodyPr/>' : bodyProperties
}

/**
 * Generate the XML for text and its options (bold, bullet, etc) including text runs (word-level formatting)
 * @param {ISlideObject|TableCell} slideObj - slideObj or tableCell
 * @note PPT text lines [lines followed by line-breaks] are created using <p>-aragraph's
 * @note Bullets are a paragragh-level formatting device
 * @template
 *    <p:txBody>
 *        <a:bodyPr wrap="square" rtlCol="0">
 *            <a:spAutoFit/>
 *        </a:bodyPr>
 *        <a:lstStyle/>
 *        <a:p>
 *            <a:pPr algn="ctr"/>
 *            <a:r>
 *                <a:rPr lang="en-US" dirty="0" err="1"/>
 *                <a:t>textbox text</a:t>
 *            </a:r>
 *            <a:endParaRPr lang="en-US" dirty="0"/>
 *        </a:p>
 *    </p:txBody>
 * @returns XML containing the param object's text and formatting
 */
export function genXmlTextBody (slideObj: ISlideObject | TableCell): string {
	const opts: ObjectOptions = slideObj.options || {}
	let tmpTextObjects: TextProps[] = []
	const arrTextObjects: TextProps[] = []

	// FIRST: Shapes without text still require a `<p:txBody>` child on `<p:sp>` per OOXML; returning '' omits it and triggers PowerPoint repair
	if (opts && slideObj._type !== SLIDE_OBJECT_TYPES.tablecell && (typeof slideObj.text === 'undefined' || slideObj.text === null)) {
		return `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="${opts.lang || 'en-US'}"/></a:p></p:txBody>`
	}

	// STEP 1: Start textBody
	let strSlideXml = slideObj._type === SLIDE_OBJECT_TYPES.tablecell ? '<a:txBody>' : '<p:txBody>'

	// STEP 2: Add bodyProperties
	{
		// A: 'bodyPr'
		strSlideXml += genXmlBodyProperties(slideObj)

		// B: 'lstStyle'
		// NOTE: shape type 'LINE' has different text align needs (a lstStyle.lvl1pPr between bodyPr and p)
		// FIXME: LINE horiz-align doesnt work (text is always to the left inside line) (FYI: the PPT code diff is substantial!)
		if (opts.h === 0 && opts.line && opts.align) strSlideXml += '<a:lstStyle><a:lvl1pPr algn="l"/></a:lstStyle>'
		else if (slideObj._type === 'placeholder') strSlideXml += `<a:lstStyle>${genXmlParagraphProperties(slideObj, true)}</a:lstStyle>`
		else strSlideXml += '<a:lstStyle/>'
	}

	/* STEP 3: Modify slideObj.text to array
		CASES:
		addText( 'string' ) // string
		addText( 'line1\n line2' ) // string with lineBreak
		addText( {text:'word1'} ) // TextProps object
		addText( ['barry','allen'] ) // array of strings
		addText( [{text:'word1'}, {text:'word2'}] ) // TextProps object array
		addText( [{text:'line1\n line2'}, {text:'end word'}] ) // TextProps object array with lineBreak
	*/
	if (typeof slideObj.text === 'string' || typeof slideObj.text === 'number') {
		// Handle cases 1,2
		tmpTextObjects.push({ text: slideObj.text.toString(), options: opts || {} })
	} else if (slideObj.text && !Array.isArray(slideObj.text) && typeof slideObj.text === 'object' && Object.keys(slideObj.text).includes('text')) {
		// } else if (!Array.isArray(slideObj.text) && slideObj.text!.hasOwnProperty('text')) { // 20210706: replaced with below as ts compiler rejected it
		// Handle case 3
		tmpTextObjects.push({ text: slideObj.text || '', options: slideObj.options || {} })
	} else if (Array.isArray(slideObj.text)) {
		// Handle cases 4,5,6
		tmpTextObjects = slideObj.text.map(item => ({ text: item.text, options: item.options }))
	}

	// STEP 4: Iterate over text objects, set text/options, break into pieces if '\n'/breakLine found
	tmpTextObjects.forEach((itext, idx) => {
		if (!itext.text) itext.text = ''

		// A: Set options
		itext.options = itext.options || opts || {}
		if (idx === 0 && itext.options && !itext.options.bullet && opts.bullet) itext.options.bullet = opts.bullet

		// B: Cast to text-object and fix line-breaks (if needed)
		if (typeof itext.text === 'string' || typeof itext.text === 'number') {
			// 1: Convert "\n" or any variation into CRLF
			itext.text = itext.text.toString().replace(/\r*\n/g, CRLF)
		}

		// C: If text string has line-breaks, then create a separate text-object for each (much easier than dealing with split inside a loop below)
		// NOTE: Filter for trailing lineBreak prevents the creation of an empty textObj as the last item
		if (itext.text.includes(CRLF) && itext.text.match(/\n$/g) === null) {
			itext.text.split(CRLF).forEach(line => {
				if (itext.options) itext.options.breakLine = true
				arrTextObjects.push({ text: line, options: itext.options })
			})
		} else {
			arrTextObjects.push(itext)
		}
	})

	// STEP 5: Group textObj into lines by checking for lineBreak, bullets, alignment change, etc.
	const arrLines: TextProps[][] = []
	let arrTexts: TextProps[] = []
	arrTextObjects.forEach((textObj, idx) => {
		// A: Align or Bullet trigger new line
		if (arrTexts.length > 0 && (textObj.options?.align || opts.align)) {
			// Only start a new paragraph when align *changes*
			if (textObj.options?.align !== arrTextObjects[idx - 1].options?.align) {
				arrLines.push(arrTexts)
				arrTexts = []
			}
		} else if (arrTexts.length > 0 && textObj.options?.bullet && arrTexts.length > 0) {
			arrLines.push(arrTexts)
			arrTexts = []
			if (textObj.options) textObj.options.breakLine = false // For cases with both `bullet` and `brekaLine` - prevent double lineBreak
		}

		// B: Add this text to current line
		arrTexts.push(textObj)

		// C: BreakLine begins new line **after** adding current text
		if (arrTexts.length > 0 && textObj.options?.breakLine) {
			// Avoid starting a para right as loop is exhausted
			if (idx + 1 < arrTextObjects.length) {
				arrLines.push(arrTexts)
				arrTexts = []
			}
		}

		// D: Flush buffer
		if (idx + 1 === arrTextObjects.length) arrLines.push(arrTexts)
	})

	// STEP 6: Loop over each line and create paragraph props, text run, etc.
	arrLines.forEach(line => {
		let reqsClosingFontSize = false

		// A: Start paragraph, add paraProps
		strSlideXml += '<a:p>'
		// NOTE: `rtlMode` is like other opts, its propagated up to each text:options, so just check the 1st one
		let paragraphPropXml = `<a:pPr ${line[0].options?.rtlMode ? ' rtl="1" ' : ''}`

		// B: Start paragraph, loop over lines and add text runs
		line.forEach((textObj, idx) => {
			// NOTE: `options` is always populated by the time text is serialized; bind the real object so mutations below persist
			const textOpts = textObj.options ?? (textObj.options = {})
			// A: Set line index
			textOpts._lineIdx = idx

			// A.1: Add soft break if not the first run of the line.
			if (idx > 0 && textOpts.softBreakBefore) {
				strSlideXml += '<a:br/>'
			}

			// B: Inherit pPr-type options from parent shape's `options`
			textOpts.align = textOpts.align || opts.align
			textOpts.lineSpacing = textOpts.lineSpacing || opts.lineSpacing
			textOpts.lineSpacingMultiple = textOpts.lineSpacingMultiple || opts.lineSpacingMultiple
			textOpts.indentLevel = textOpts.indentLevel || opts.indentLevel
			textOpts.paraSpaceBefore = textOpts.paraSpaceBefore || opts.paraSpaceBefore
			textOpts.paraSpaceAfter = textOpts.paraSpaceAfter || opts.paraSpaceAfter
			paragraphPropXml = genXmlParagraphProperties(textObj, false)

			if (idx === 0) strSlideXml += paragraphPropXml.replace('<a:pPr></a:pPr>', '') // A paragraph permits one pPr, before its runs
			// C: Inherit any main options (color, fontSize, etc.)
			// NOTE: We only pass the text.options to genXmlTextRun (not the Slide.options),
			// so the run building function cant just fallback to Slide.color, therefore, we need to do that here before passing options below.
			// FILTER RULE: Hyperlinks should not inherit `color` from main options (let PPT default to local color, eg: blue on MacOS)
			Object.entries(opts).filter(([key]) => !(textOpts.hyperlink && key === 'color')).forEach(([key, val]) => {
				// if (textOpts.hyperlink && key === 'color') null
				// NOTE: This loop will pick up unecessary keys (`x`, etc.), but it doesnt hurt anything
				// `omml` is a per-run math payload - inheriting it would turn every sibling run into the same equation
				if (key !== 'bullet' && key !== 'omml' && !textOpts[key]) textOpts[key] = val
			})

			// D: Add formatted textrun
			strSlideXml += genXmlTextRun(textObj)

			// E: Flag close fontSize for empty [lineBreak] elements
			if ((!textObj.text && opts.fontSize) || textOpts.fontSize) {
				reqsClosingFontSize = true
				opts.fontSize = opts.fontSize || textOpts.fontSize
			}
		})

		/* C: Append 'endParaRPr' (when needed) and close current open paragraph
		 * NOTE: (ISSUE#20, ISSUE#193): Add 'endParaRPr' with font/size props or PPT default (Arial/18pt en-us) is used making row "too tall"/not honoring options
		 */
		if (slideObj._type === SLIDE_OBJECT_TYPES.tablecell && (opts.fontSize || opts.fontFace)) {
			if (opts.fontFace) {
				strSlideXml += `<a:endParaRPr lang="${opts.lang || 'en-US'}"` + (opts.fontSize ? ` sz="${Math.round(opts.fontSize * 100)}"` : '') + ' dirty="0">'
				strSlideXml += `<a:latin typeface="${opts.fontFace}" charset="0"/>`
				strSlideXml += `<a:ea typeface="${opts.fontFace}" charset="0"/>`
				strSlideXml += `<a:cs typeface="${opts.fontFace}" charset="0"/>`
				strSlideXml += '</a:endParaRPr>'
			} else {
				strSlideXml += `<a:endParaRPr lang="${opts.lang || 'en-US'}"` + (opts.fontSize ? ` sz="${Math.round(opts.fontSize * 100)}"` : '') + ' dirty="0"/>'
			}
		} else if (reqsClosingFontSize) {
			// Empty [lineBreak] lines should not contain runProp, however, they need to specify fontSize in `endParaRPr`
			strSlideXml += `<a:endParaRPr lang="${opts.lang || 'en-US'}"` + (opts.fontSize ? ` sz="${Math.round(opts.fontSize * 100)}"` : '') + ' dirty="0"/>'
		} else {
			strSlideXml += `<a:endParaRPr lang="${opts.lang || 'en-US'}" dirty="0"/>` // Added 20180101 to address PPT-2007 issues
		}

		// D: End paragraph
		strSlideXml += '</a:p>'
	})

	// IMPORTANT: An empty txBody will cause "needs repair" error! Add <p> content if missing.
	// [FIXED in v3.13.0]: This fixes issue with table auto-paging where some cells w/b empty on subsequent pages.
	/*
		<a:txBody>
			<a:bodyPr/>
			<a:lstStyle/>
		</a:txBody>
	*/
	if (strSlideXml.indexOf('<a:p>') === -1) {
		strSlideXml += '<a:p><a:endParaRPr/></a:p>'
	}

	// STEP 7: Close the textBody
	strSlideXml += slideObj._type === SLIDE_OBJECT_TYPES.tablecell ? '</a:txBody>' : '</p:txBody>'

	// LAST: Return XML
	return strSlideXml
}

/**
 * Generate an XML Placeholder
 * @param {ISlideObject} placeholderObj
 * @returns XML
 */
export function genXmlPlaceholder (placeholderObj: ISlideObject | undefined): string {
	if (!placeholderObj) return ''

	const placeholderIdx = placeholderObj.options?._placeholderIdx ? placeholderObj.options._placeholderIdx : ''
	const placeholderTyp = placeholderObj.options?._placeholderType ? placeholderObj.options._placeholderType : ''
	// NOTE: accept both the friendly name ('image', 'table') and the OOXML code it maps to ('pic', 'tbl').
	// The old code looked the mapped code back up in the enum, which only ever hit for the types whose
	// name and code are identical - that is why picture/chart/table placeholders never got a `type` (issue #33).
	const placeholderCodes: string[] = Object.values(PLACEHOLDER_TYPES)
	const placeholderType: string = PLACEHOLDER_TYPES[placeholderTyp]?.toString() ?? (placeholderCodes.includes(placeholderTyp) ? placeholderTyp : '')

	return `<p:ph
		${placeholderIdx ? ' idx="' + placeholderIdx.toString() + '"' : ''}
		${placeholderType ? ` type="${placeholderType}"` : ''}
		${placeholderObj.text && placeholderObj.text.length > 0 ? ' hasCustomPrompt="1"' : ''}
		/>`
}

// XML-GEN: First 6 functions create the base /ppt files

/**
 * Generate XML ContentType
 * @param {PresSlide[]} slides - slides
 * @param {SlideLayout[]} slideLayouts - slide layouts
 * @param {PresSlide} masterSlide - master slide
 * @returns XML
 */
