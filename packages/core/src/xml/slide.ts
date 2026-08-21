/**
 * PptxGenJS: XML Generation
 */

import {
	DEF_CELL_MARGIN_IN,
	DEF_PRES_LAYOUT_NAME,
	DEF_TEXT_SHADOW,
	EMU,
	MS_PPTX_ID_BASE,
	OOXML_EXT,
	SLDNUMFLDID,
	SHAPE_TYPE,
	SLIDE_OBJECT_TYPES,
} from '../core-enums'
import {
	ISlideObject,
	ObjectOptions,
	PresSlide,
	ReflectionProps,
	ShadowProps,
	SectionProps,
	ShapeLineProps,
	SlideLayout,
	SoftEdgeProps,
	TableCell,
	TableCellProps,
	TableProps,
	TextGlowProps,
} from '../core-interfaces'
import {
	convertRotationDegrees,
	createColorElement,
	createGlowElement,
	encodeXmlEntities,
	genXmlColorSelection,
	getSmartParseNumber,
	inch2Emu,
	resolveGlowOptions,
	valToPts,
} from '../gen-utils'

import { slideCommentRelId } from './relationships'
import { genXmlHyperlink } from './hyperlink'
import { genXmlPlaceholder, genXmlTextBody } from './text'
import { genXmlZoom } from './zoom'

const ImageSizingXml = {
	cover: function (imgSize: { w: number, h: number }, boxDim: { w: number, h: number, x: number, y: number }) {
		const imgRatio = imgSize.h / imgSize.w
		const boxRatio = boxDim.h / boxDim.w
		const isBoxBased = boxRatio > imgRatio
		const width = isBoxBased ? boxDim.h / imgRatio : boxDim.w
		const height = isBoxBased ? boxDim.h : boxDim.w * imgRatio
		const hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width))
		const vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height))
		return `<a:srcRect l="${hzPerc}" r="${hzPerc}" t="${vzPerc}" b="${vzPerc}"/><a:stretch/>`
	},
	contain: function (imgSize: { w: number, h: number }, boxDim: { w: number, h: number, x: number, y: number }) {
		const imgRatio = imgSize.h / imgSize.w
		const boxRatio = boxDim.h / boxDim.w
		const widthBased = boxRatio > imgRatio
		const width = widthBased ? boxDim.w : boxDim.h / imgRatio
		const height = widthBased ? boxDim.w * imgRatio : boxDim.h
		const hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width))
		const vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height))
		return `<a:srcRect l="${hzPerc}" r="${hzPerc}" t="${vzPerc}" b="${vzPerc}"/><a:stretch/>`
	},
	crop: function (imgSize: { w: number, h: number }, boxDim: { w: number, h: number, x: number, y: number }) {
		const l = boxDim.x
		const r = imgSize.w - (boxDim.x + boxDim.w)
		const t = boxDim.y
		const b = imgSize.h - (boxDim.y + boxDim.h)
		const lPerc = Math.round(1e5 * (l / imgSize.w))
		const rPerc = Math.round(1e5 * (r / imgSize.w))
		const tPerc = Math.round(1e5 * (t / imgSize.h))
		const bPerc = Math.round(1e5 * (b / imgSize.h))
		return `<a:srcRect l="${lPerc}" r="${rPerc}" t="${tPerc}" b="${bPerc}"/><a:stretch/>`
	},
}

/**
 * Create the `a:tblPr` table-style block (banded rows/cols, first/last row/col emphasis, style id)
 * @param {TableProps} opts - table options
 * @return {string} XML
 */
function genXmlTblPr (opts: TableProps): string {
	const firstRow = opts.firstRow ?? (opts.autoPageRepeatHeader ? true : undefined)
	const flags: Array<[string, boolean | undefined]> = [
		['firstRow', firstRow],
		['lastRow', opts.lastRow],
		['firstCol', opts.firstCol],
		['lastCol', opts.lastCol],
		['bandRow', opts.bandRow],
		['bandCol', opts.bandCol],
	]
	const attrs = flags
		.filter(([, val]) => typeof val === 'boolean')
		.map(([name, val]) => ` ${name}="${val ? 1 : 0}"`)
		.join('')

	if (!attrs && !opts.tableStyleId) return '<a:tblPr/>'
	// NOTE: `a:tableStyleId` must be the last child of `a:tblPr` per the schema
	return opts.tableStyleId ? `<a:tblPr${attrs}><a:tableStyleId>${encodeXmlEntities(opts.tableStyleId)}</a:tableStyleId></a:tblPr>` : `<a:tblPr${attrs}/>`
}

/**
 * Modification id for a shape, unique on its slide (MS-PPTX 2.3.1.19)
 * - derived from a base plus the shape's own index rather than randomised: MS-PPTX only requires
 *   uniqueness within the slide, and deriving it keeps exports reproducible and free of state
 * @param {number} objectIndex - index of the shape among the slide's objects
 * @returns {number} `p14:modId` value
 */
function shapeModId (objectIndex: number): number {
	return MS_PPTX_ID_BASE.modId + objectIndex
}

/**
 * Create the `a:ln` outline block for a shape/image
 * @param {ShapeLineProps} line - line options
 * @return {string} XML
 */
function genXmlLine (line: ShapeLineProps): string {
	let xml = line.width ? `<a:ln w="${valToPts(line.width)}">` : '<a:ln>'
	if (line.color || line.type === 'gradient') xml += genXmlColorSelection(line)
	if (line.dashType) xml += `<a:prstDash val="${line.dashType}"/>`
	if (line.beginArrowType) xml += `<a:headEnd type="${line.beginArrowType}"/>`
	if (line.endArrowType) xml += `<a:tailEnd type="${line.endArrowType}"/>`
	// FUTURE: `endArrowSize` < a: headEnd type = "arrow" w = "lg" len = "lg" /> 'sm' | 'med' | 'lg'(values are 1 - 9, making a 3x3 grid of w / len possibilities)
	xml += '</a:ln>'
	return xml
}

/**
 * Create one shadow child for an `a:effectLst`.
 * @note Pure: XML unit conversion does not mutate the caller's options.
 * @param {ShadowProps} shadow - shadow options
 * @return {string} shadow XML
 */
function genXmlShadowElement (shadow: ShadowProps): string {
	const type = shadow.type === 'inner' ? 'inner' : 'outer'
	const blur = valToPts(shadow.blur ?? 8)
	const offset = valToPts(shadow.offset ?? 4)
	const angle = Math.round((shadow.angle ?? 270) * 60000)
	const opacity = Math.round((shadow.opacity ?? 0.75) * 100000)
	const color = shadow.color || DEF_TEXT_SHADOW.color
	const attrs = type === 'outer' ? 'sx="100000" sy="100000" kx="0" ky="0" algn="bl" rotWithShape="0"' : ''

	return `<a:${type}Shdw ${attrs} blurRad="${blur}" dist="${offset}" dir="${angle}"><a:srgbClr val="${color}"><a:alpha val="${opacity}"/></a:srgbClr></a:${type}Shdw>`
}

/**
 * Create one soft-edge child for an `a:effectLst`.
 * @param {SoftEdgeProps} softEdge - soft-edge options
 * @return {string} soft-edge XML
 */
function genXmlSoftEdgeElement (softEdge: SoftEdgeProps): string {
	return `<a:softEdge rad="${valToPts(softEdge.radius)}"/>`
}

/**
 * Create one reflection child for an `a:effectLst`.
 * @param {ReflectionProps} reflection - reflection options
 * @return {string} reflection XML
 */
function genXmlReflectionElement (reflection: ReflectionProps): string {
	return `<a:reflection blurRad="${valToPts(reflection.blur ?? 0)}" stA="${Math.round((reflection.opacity ?? 0.5) * 100000)}" endA="0" dist="${valToPts(reflection.distance ?? 0)}" dir="${Math.round((reflection.direction ?? 0) * 60000)}" sy="${Math.round((reflection.scaleY ?? -1) * 100000)}" algn="bl" rotWithShape="0"/>`
}

/**
 * Create one ordered DrawingML effect list for a shape or image.
 * @param {object} opts - supported effect options
 * @return {string} effect-list XML, or an empty string when no effects are set
 */
function genXmlEffectLst (opts: { shadow?: ShadowProps, glow?: TextGlowProps, softEdge?: SoftEdgeProps, reflection?: ReflectionProps }): string {
	const effects: string[] = []
	const glow = resolveGlowOptions(opts.glow)
	if (glow) effects.push(createGlowElement(glow))
	if (opts.shadow && opts.shadow.type !== 'none') effects.push(genXmlShadowElement(opts.shadow))
	if (opts.reflection) effects.push(genXmlReflectionElement(opts.reflection))
	if (opts.softEdge) effects.push(genXmlSoftEdgeElement(opts.softEdge))
	return effects.length ? `<a:effectLst>${effects.join('')}</a:effectLst>` : ''
}

interface SlideObjectContext {
	cx: number
	cy: number
	imgHeight: number
	imgWidth: number
	locationAttr: string
	placeholderObj: ISlideObject | undefined
	rounding: ObjectOptions['rounding']
	sizing: ObjectOptions['sizing']
	x: number
	y: number
}

/**
 * Resolve an object's exported geometry once, before its type-specific XML is rendered.
 * Placeholder geometry deliberately overrides object geometry; image dimensions are captured first
 * because image sizing is based on the source image dimensions, not its placeholder frame.
 */
function resolveSlideObjectContext (slide: PresSlide | SlideLayout, slideItemObj: ISlideObject & { options: ObjectOptions }): SlideObjectContext {
	const options = slideItemObj.options

	const placeholderObj =
		'_slideLayout' in slide && slide._slideLayout?._slideObjects !== undefined && options.placeholder
			? slide._slideLayout._slideObjects.find(object => object.options?.placeholder === options.placeholder)
			: undefined

	let x = typeof options.x !== 'undefined' ? getSmartParseNumber(options.x, 'X', slide._presLayout) : 0
	let y = typeof options.y !== 'undefined' ? getSmartParseNumber(options.y, 'Y', slide._presLayout) : 0
	let cx = typeof options.w !== 'undefined' ? getSmartParseNumber(options.w, 'X', slide._presLayout) : getSmartParseNumber('75%', 'X', slide._presLayout)
	let cy = typeof options.h !== 'undefined' ? getSmartParseNumber(options.h, 'Y', slide._presLayout) : 0

	// Image sizing needs the object's own dimensions even when the image is positioned through a placeholder.
	const imgWidth = cx
	const imgHeight = cy

	if (placeholderObj) {
		if (placeholderObj.options?.x === 0 || placeholderObj.options?.x) x = getSmartParseNumber(placeholderObj.options.x, 'X', slide._presLayout)
		if (placeholderObj.options?.y === 0 || placeholderObj.options?.y) y = getSmartParseNumber(placeholderObj.options.y, 'Y', slide._presLayout)
		if (placeholderObj.options?.w === 0 || placeholderObj.options?.w) cx = getSmartParseNumber(placeholderObj.options.w, 'X', slide._presLayout)
		if (placeholderObj.options?.h === 0 || placeholderObj.options?.h) cy = getSmartParseNumber(placeholderObj.options.h, 'Y', slide._presLayout)
	}

	// DrawingML extents cannot be negative. Lines use their frame as endpoint deltas,
	// so move a negative delta into the offset and flip that axis to retain endpoints
	// (and therefore head/tail arrow direction).
	let flipH = options.flipH
	let flipV = options.flipV
	if (slideItemObj.shape === SHAPE_TYPE.LINE) {
		if (cx < 0) {
			x += cx
			cx = -cx
			flipH = !flipH
		}
		if (cy < 0) {
			y += cy
			cy = -cy
			flipV = !flipV
		}
	}

	let locationAttr = ''
	if (flipH) locationAttr += ' flipH="1"'
	if (flipV) locationAttr += ' flipV="1"'
	if (options.rotate) locationAttr += ` rot="${convertRotationDegrees(options.rotate)}"`

	return { cx, cy, imgHeight, imgWidth, locationAttr, placeholderObj, rounding: options.rounding, sizing: options.sizing, x, y }
}

/**
 * Render the single valid background representation for a slide-like part.
 * Image backgrounds win over colors; the default master layout receives the scheme background needed by Keynote previews.
 */
function genXmlSlideBackground (slide: PresSlide | SlideLayout): string {
	let strSlideXml = ''
	// STEP 1: Add background color/image (ensure only a single `<p:bg>` tag is created, ex: when master-baskground has both `color` and `path`)
	if (slide._bkgdImgRid) {
		strSlideXml += `<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId${slide._bkgdImgRid}"><a:lum/></a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`
	} else if (slide.background?.color || slide.background?.type === 'gradient') {
		// NOTE: `<a:effectLst/>` is required by PowerPoint (matches image-bg path above); omitting it triggers the repair dialog
		strSlideXml += `<p:bg><p:bgPr>${genXmlColorSelection(slide.background)}<a:effectLst/></p:bgPr></p:bg>`
	} else if (!slide.bkgd && slide._name && slide._name === DEF_PRES_LAYOUT_NAME) {
		// NOTE: Default [white] background is needed on slideMaster1.xml to avoid gray background in Keynote (and Finder previews)
		strSlideXml += '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
	}

	return strSlideXml
}

/** Create the required group shape tree and its non-visual properties. */
function genXmlSlideTreeStart (): string {
	let strSlideXml = ''
	// STEP 2: Continue slide by starting spTree node
	strSlideXml += '<p:spTree>'
	strSlideXml += '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
	strSlideXml += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
	strSlideXml += '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'

	return strSlideXml
}

/**
 * Serialize slide objects in insertion order and retain the rendering-time normalization of each object's options.
 *
 * The local table counter and object index determine OOXML non-visual IDs, so callers must keep this phase contiguous.
 */
function genXmlSlideObjects (slide: PresSlide | SlideLayout, sections: SectionProps[]): string {
	let strSlideXml = ''
	let intTableNum = 1
	// STEP 3: Loop over all Slide.data objects and add them to this slide
	slide._slideObjects.forEach((slideObject: ISlideObject, idx: number) => {
		const slideItemObj = { ...slideObject, options: slideObject.options ?? {} }
		// XML generation has historically filled in this internal object; retain that contract for downstream renderers.
		slideObject.options = slideItemObj.options
		let arrTabRows: TableCell[][] = []
		let objTabOpts: ObjectOptions = {}
		let intColCnt = 0
		let intColW = 0
		let cellOpts: TableCellProps | undefined
		let strXml = ''
		const context = resolveSlideObjectContext(slide, slideItemObj)
		const { cx, x, y } = context
		let { cy, imgHeight, imgWidth } = context
		const { locationAttr, placeholderObj, rounding, sizing } = context

		// B: Add OBJECT to the current Slide
		switch (slideItemObj._type) {
			case SLIDE_OBJECT_TYPES.table:
				arrTabRows = slideItemObj.arrTabRows ?? []
				objTabOpts = slideItemObj.options
				intColCnt = 0
				intColW = 0

				// Calc number of columns
				// NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
				// ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
				arrTabRows[0].forEach(cell => {
					cellOpts = cell.options
					intColCnt += cellOpts?.colspan ? Number(cellOpts.colspan) : 1
				})

				// STEP 1: Start Table XML
				// NOTE: Non-numeric cNvPr id values will trigger "presentation needs repair" type warning in MS-PPT-2013
				strXml = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${intTableNum * (slide._slideNum ?? 0) + 1}" name="${slideItemObj.options.objectName}"/>`
				strXml +=
					'<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>' +
					// MS-PPTX 2.3.1.19: each `p14:modId` must be unique on the slide, so it cannot be a constant
					`  <p:nvPr><p:extLst><p:ext uri="${OOXML_EXT.modId.uri}"><p14:modId xmlns:p14="${OOXML_EXT.modId.ns}" val="${shapeModId(idx)}"/></p:ext></p:extLst></p:nvPr>` +
					'</p:nvGraphicFramePr>'
				strXml += `<p:xfrm><a:off x="${x || (x === 0 ? 0 : EMU)}" y="${y || (y === 0 ? 0 : EMU)}"/><a:ext cx="${cx || (cx === 0 ? 0 : EMU)}" cy="${cy || EMU
				}"/></p:xfrm>`
				strXml += '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
				// NOTE: banding/emphasis only renders when a table style is applied - either `tableStyleId` here or a theme default
				strXml += `<a:tbl>${genXmlTblPr(objTabOpts as TableProps)}`

				// STEP 2: Set column widths
				// Evenly distribute cols/rows across size provided when applicable (calc them if only overall dimensions were provided)
				// A: Col widths provided?
				// B: Table Width provided without colW? Then distribute cols
				if (Array.isArray(objTabOpts.colW)) {
					strXml += '<a:tblGrid>'
					for (let col = 0; col < intColCnt; col++) {
						let w = inch2Emu(objTabOpts.colW[col])
						if (w == null || isNaN(w)) {
							w = (typeof slideItemObj.options.w === 'number' ? slideItemObj.options.w : 1) / intColCnt
						}
						strXml += `<a:gridCol w="${Math.round(w)}"/>`
					}
					strXml += '</a:tblGrid>'
				} else {
					intColW = objTabOpts.colW ? objTabOpts.colW : EMU
					if (slideItemObj.options.w && !objTabOpts.colW) intColW = Math.round((typeof slideItemObj.options.w === 'number' ? slideItemObj.options.w : 1) / intColCnt)
					strXml += '<a:tblGrid>'
					for (let colw = 0; colw < intColCnt; colw++) {
						strXml += `<a:gridCol w="${intColW}"/>`
					}
					strXml += '</a:tblGrid>'
				}

				// STEP 3: Build our row arrays into an actual grid to match the XML we will be building next (ISSUE #36)
				// Note row arrays can arrive "lopsided" as in row1:[1,2,3] row2:[3] when first two cols rowspan!,
				// so a simple loop below in XML building wont suffice to build table correctly.
				// We have to build an actual grid now
				/*
					EX: (A0:rowspan=3, B1:rowspan=2, C1:colspan=2)

					/------|------|------|------\
					|  A0  |  B0  |  C0  |  D0  |
					|      |  B1  |  C1  |      |
					|      |      |  C2  |  D2  |
					\------|------|------|------/
				*/
				// A: add _hmerge cell for colspan. should reserve rowspan
				arrTabRows.forEach(cells => {
					for (let cIdx = 0; cIdx < cells.length;) {
						const cell = cells[cIdx]
						const colspan = cell.options?.colspan
						const rowspan = cell.options?.rowspan
						if (colspan && colspan > 1) {
							const vMergeCells = new Array(colspan - 1).fill(undefined).map(() => {
								return { _type: SLIDE_OBJECT_TYPES.tablecell, options: { rowspan }, _hmerge: true } as const
							})
							cells.splice(cIdx + 1, 0, ...vMergeCells)
							cIdx += colspan
						} else {
							cIdx += 1
						}
					}
				})
				// B: add _vmerge cell for rowspan. should reserve colspan/_hmerge
				arrTabRows.forEach((cells, rIdx) => {
					const nextRow = arrTabRows[rIdx + 1]
					if (!nextRow) return
					cells.forEach((cell, cIdx) => {
						const rowspan = cell._rowContinue || cell.options?.rowspan
						const colspan = cell.options?.colspan
						const _hmerge = cell._hmerge
						if (rowspan && rowspan > 1) {
							const hMergeCell = { _type: SLIDE_OBJECT_TYPES.tablecell, options: { colspan }, _rowContinue: rowspan - 1, _vmerge: true, _hmerge } as const
							nextRow.splice(cIdx, 0, hMergeCell)
						}
					})
				})

				// STEP 4: Build table rows/cells
				arrTabRows.forEach((cells, rIdx) => {
					const itemOpts: ObjectOptions = slideItemObj.options ?? {}
					// A: Table Height provided without rowH? Then distribute rows
					let intRowH = 0 // IMPORTANT: Default must be zero for auto-sizing to work
					// `rowH` is either one value for every row or a per-row array; `inch2Emu` resolves inches, EMU and unit-suffixed strings alike
					const rowHOpt = Array.isArray(objTabOpts.rowH) ? objTabOpts.rowH[rIdx] : objTabOpts.rowH
					const rowHEmu = rowHOpt ? inch2Emu(rowHOpt) : NaN
					if (!isNaN(rowHEmu)) intRowH = rowHEmu
					else if (itemOpts.cy || itemOpts.h) {
						intRowH = Math.round(
							(itemOpts.h ? inch2Emu(itemOpts.h) : typeof itemOpts.cy === 'number' ? itemOpts.cy : 1) /
							arrTabRows.length
						)
					}

					// B: Start row
					strXml += `<a:tr h="${intRowH}">`

					// C: Loop over each CELL
					cells.forEach(cellObj => {
						const cell: TableCell = cellObj

						const cellSpanAttrs = {
							rowSpan: (cell.options?.rowspan ?? 0) > 1 ? cell.options?.rowspan : undefined,
							gridSpan: (cell.options?.colspan ?? 0) > 1 ? cell.options?.colspan : undefined,
							vMerge: cell._vmerge ? 1 : undefined,
							hMerge: cell._hmerge ? 1 : undefined,
						}
						let cellSpanAttrStr = Object.keys(cellSpanAttrs)
							.map(k => [k, cellSpanAttrs[k]])
							.filter(([, v]) => !!v)
							.map(([k, v]) => `${String(k)}="${String(v)}"`)
							.join(' ')
						if (cellSpanAttrStr) cellSpanAttrStr = ' ' + cellSpanAttrStr

						// 1: COLSPAN/ROWSPAN: Add dummy cells for any active colspan/rowspan
						if (cell._hmerge || cell._vmerge) {
							strXml += `<a:tc${cellSpanAttrStr}><a:tcPr/></a:tc>`
							return
						}

						// 2: OPTIONS: Build/set cell options
						const cellOpts = cell.options || {}
						cell.options = cellOpts

						// B: Inherit some options from table when cell options dont exist
						// @see: http://officeopenxml.com/drwTableCellProperties-alignment.php
						;['align', 'bold', 'border', 'color', 'fill', 'fontFace', 'fontSize', 'margin', 'textDirection', 'underline', 'valign'].forEach(name => {
							if (objTabOpts[name] && !cellOpts[name] && cellOpts[name] !== 0) cellOpts[name] = objTabOpts[name]
						})

						const cellValign = cellOpts.valign
							? ` anchor="${cellOpts.valign.replace(/^c$/i, 'ctr').replace(/^m$/i, 'ctr').replace('center', 'ctr').replace('middle', 'ctr').replace('top', 't').replace('btm', 'b').replace('bottom', 'b')}"`
							: ''
						const cellTextDir = (cellOpts.textDirection && cellOpts.textDirection !== 'horz') ? ` vert="${cellOpts.textDirection}"` : ''

						let fillColor =
							cell._optImp?.fill?.color
								? cell._optImp.fill.color
								: cell._optImp?.fill && typeof cell._optImp.fill === 'string'
									? cell._optImp.fill
									: ''
						fillColor = fillColor || cellOpts.fill ? cellOpts.fill : ''
						// A gradient needs the whole fill object; the color-only path above cannot carry stops
						// gradient, pattern, and picture fills carry more than a color, so the whole
						// fill object has to reach the emitter - the color-only path cannot express them
						const cellFillProps = cell._optImp?.fill ?? cellOpts.fill
						const cellFill =
							typeof cellFillProps === 'object' && (cellFillProps?.type === 'gradient' || cellFillProps?.type === 'pattern' || cellFillProps?.type === 'image')
								? genXmlColorSelection(cellFillProps)
								: fillColor ? genXmlColorSelection(fillColor) : ''

						let cellMargin = cellOpts.margin === 0 || cellOpts.margin ? cellOpts.margin : DEF_CELL_MARGIN_IN
						if (!Array.isArray(cellMargin) && typeof cellMargin === 'number') cellMargin = [cellMargin, cellMargin, cellMargin, cellMargin]
						// Guard against non-number/non-array margins (e.g. object/string) which otherwise yield marL="NaN" and trigger PowerPoint repair
						if (!Array.isArray(cellMargin)) cellMargin = DEF_CELL_MARGIN_IN as [number, number, number, number]
						cellMargin = (cellMargin as number[]).map(v => (typeof v === 'number' && !isNaN(v) ? v : 0)) as [number, number, number, number]
						/** FUTURE: DEPRECATED:
						 * - Backwards-Compat: Oops! Discovered we were still using points for cell margin before v3.8.0 (UGH!)
						 * - We cant introduce a breaking change before v4.0, so...
						 */
						let cellMarginXml = ''
						if (cellMargin[0] >= 1) {
							cellMarginXml = ` marL="${valToPts(cellMargin[3])}" marR="${valToPts(cellMargin[1])}" marT="${valToPts(cellMargin[0])}" marB="${valToPts(
								cellMargin[2]
							)}"`
						} else {
							cellMarginXml = ` marL="${inch2Emu(cellMargin[3])}" marR="${inch2Emu(cellMargin[1])}" marT="${inch2Emu(cellMargin[0])}" marB="${inch2Emu(
								cellMargin[2]
							)}"`
						}

						// FUTURE: Cell NOWRAP property (textwrap: add to a:tcPr (horzOverflow="overflow" or whatever options exist)

						// 4: Set CELL content and properties ==================================
						strXml += `<a:tc${cellSpanAttrStr}>${genXmlTextBody(cell)}<a:tcPr${cellMarginXml}${cellValign}${cellTextDir}>`
						// strXml += `<a:tc${cellColspan}${cellRowspan}>${genXmlTextBody(cell)}<a:tcPr${cellMarginXml}${cellValign}${cellTextDir}>`
						// FIXME: 20200525: ^^^
						// <a:tcPr marL="38100" marR="38100" marT="38100" marB="38100" vert="vert270">

						// 5: Borders: Add any borders
						if (cellOpts.border && Array.isArray(cellOpts.border)) {
							const border = cellOpts.border
							// NOTE: *** IMPORTANT! *** LRTB order matters! (Reorder a line below to watch the borders go wonky in MS-PPT-2013!!)
							;[
								{ idx: 3, name: 'lnL' },
								{ idx: 1, name: 'lnR' },
								{ idx: 0, name: 'lnT' },
								{ idx: 2, name: 'lnB' },
							].forEach(obj => {
								if (border[obj.idx].type !== 'none') {
									strXml += `<a:${obj.name} w="${valToPts(border[obj.idx].pt)}" cap="flat" cmpd="sng" algn="ctr">`
									strXml += `<a:solidFill>${createColorElement(border[obj.idx].color)}</a:solidFill>`
									strXml += `<a:prstDash val="${border[obj.idx].type === 'dash' ? 'sysDash' : 'solid'
									}"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/>`
									strXml += `</a:${obj.name}>`
								} else {
									strXml += `<a:${obj.name} w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:${obj.name}>`
								}
							})
						}

						// 6: Close cell Properties & Cell
						strXml += cellFill
						strXml += '  </a:tcPr>'
						strXml += ' </a:tc>'
					})

					// D: Complete row
					strXml += '</a:tr>'
				})

				// STEP 5: Complete table
				strXml += '      </a:tbl>'
				strXml += '    </a:graphicData>'
				strXml += '  </a:graphic>'
				strXml += '</p:graphicFrame>'

				// STEP 6: Set table XML
				strSlideXml += strXml

				// LAST: Increment counter
				intTableNum++
				break

			case SLIDE_OBJECT_TYPES.text:
			case SLIDE_OBJECT_TYPES.placeholder:
				// Lines can have zero cy, but text should not
				if (!slideItemObj.options.line && cy === 0) cy = EMU * 0.3

				// Margin/Padding/Inset for textboxes
				if (!slideItemObj.options._bodyProp) slideItemObj.options._bodyProp = {}
				if (slideItemObj.options.margin && Array.isArray(slideItemObj.options.margin)) {
					slideItemObj.options._bodyProp.lIns = valToPts(slideItemObj.options.margin[0] || 0)
					slideItemObj.options._bodyProp.rIns = valToPts(slideItemObj.options.margin[1] || 0)
					slideItemObj.options._bodyProp.bIns = valToPts(slideItemObj.options.margin[2] || 0)
					slideItemObj.options._bodyProp.tIns = valToPts(slideItemObj.options.margin[3] || 0)
				} else if (typeof slideItemObj.options.margin === 'number') {
					slideItemObj.options._bodyProp.lIns = valToPts(slideItemObj.options.margin)
					slideItemObj.options._bodyProp.rIns = valToPts(slideItemObj.options.margin)
					slideItemObj.options._bodyProp.bIns = valToPts(slideItemObj.options.margin)
					slideItemObj.options._bodyProp.tIns = valToPts(slideItemObj.options.margin)
				}

				// A: Start SHAPE =======================================================
				strSlideXml += '<p:sp>'

				// B: The addition of the "txBox" attribute is the sole determiner of if an object is a shape or textbox
				strSlideXml += `<p:nvSpPr><p:cNvPr id="${idx + 2}" name="${slideItemObj.options.objectName}">`
				// <Hyperlink> - `a:hlinkHover` follows `a:hlinkClick` in CT_NonVisualDrawingProps
				if (slideItemObj.options.hyperlink?._rId) strSlideXml += genXmlHyperlink(slideItemObj.options.hyperlink, 'click', 'shape')
				if (slideItemObj.options.hyperlinkHover?._rId) strSlideXml += genXmlHyperlink(slideItemObj.options.hyperlinkHover, 'hover', 'shape')
				// </Hyperlink>
				strSlideXml += '</p:cNvPr>'
				strSlideXml += '<p:cNvSpPr' + (slideItemObj.options?.isTextBox ? ' txBox="1"/>' : '/>')
				strSlideXml += `<p:nvPr>${slideItemObj._type === 'placeholder' ? genXmlPlaceholder(slideItemObj) : genXmlPlaceholder(placeholderObj)}</p:nvPr>`
				strSlideXml += '</p:nvSpPr><p:spPr>'
				strSlideXml += `<a:xfrm${locationAttr}>`
				strSlideXml += `<a:off x="${x}" y="${y}"/>`
				strSlideXml += `<a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`

				if (slideItemObj.shape === 'custGeom') {
					strSlideXml += '<a:custGeom><a:avLst />'
					strSlideXml += '<a:gdLst>'
					strSlideXml += '</a:gdLst>'
					strSlideXml += '<a:ahLst />'
					strSlideXml += '<a:cxnLst>'
					strSlideXml += '</a:cxnLst>'
					strSlideXml += '<a:rect l="l" t="t" r="r" b="b" />'

					strSlideXml += '<a:pathLst>'
					strSlideXml += `<a:path w="${cx}" h="${cy}">`

					slideItemObj.options.points?.forEach((point, i) => {
						if ('curve' in point) {
							switch (point.curve.type) {
								case 'arc':
									strSlideXml += `<a:arcTo hR="${getSmartParseNumber(point.curve.hR, 'Y', slide._presLayout)}" wR="${getSmartParseNumber(
										point.curve.wR,
										'X',
										slide._presLayout
									)}" stAng="${convertRotationDegrees(point.curve.stAng)}" swAng="${convertRotationDegrees(point.curve.swAng)}" />`
									break
								case 'cubic':
									strSlideXml += `<a:cubicBezTo>
									<a:pt x="${getSmartParseNumber(point.curve.x1, 'X', slide._presLayout)}" y="${getSmartParseNumber(point.curve.y1, 'Y', slide._presLayout)}" />
									<a:pt x="${getSmartParseNumber(point.curve.x2, 'X', slide._presLayout)}" y="${getSmartParseNumber(point.curve.y2, 'Y', slide._presLayout)}" />
									<a:pt x="${getSmartParseNumber(point.x, 'X', slide._presLayout)}" y="${getSmartParseNumber(point.y, 'Y', slide._presLayout)}" />
									</a:cubicBezTo>`
									break
								case 'quadratic':
									strSlideXml += `<a:quadBezTo>
									<a:pt x="${getSmartParseNumber(point.curve.x1, 'X', slide._presLayout)}" y="${getSmartParseNumber(point.curve.y1, 'Y', slide._presLayout)}" />
									<a:pt x="${getSmartParseNumber(point.x, 'X', slide._presLayout)}" y="${getSmartParseNumber(point.y, 'Y', slide._presLayout)}" />
									</a:quadBezTo>`
									break
								default:
									break
							}
						} else if ('close' in point) {
							strSlideXml += '<a:close />'
						} else if (point.moveTo || i === 0) {
							strSlideXml += `<a:moveTo><a:pt x="${getSmartParseNumber(point.x, 'X', slide._presLayout)}" y="${getSmartParseNumber(
								point.y,
								'Y',
								slide._presLayout
							)}" /></a:moveTo>`
						} else {
							strSlideXml += `<a:lnTo><a:pt x="${getSmartParseNumber(point.x, 'X', slide._presLayout)}" y="${getSmartParseNumber(
								point.y,
								'Y',
								slide._presLayout
							)}" /></a:lnTo>`
						}
					})

					strSlideXml += '</a:path>'
					strSlideXml += '</a:pathLst>'
					strSlideXml += '</a:custGeom>'
				} else {
					strSlideXml += '<a:prstGeom prst="' + slideItemObj.shape + '"><a:avLst>'
					if (slideItemObj.options.rectRadius) {
						strSlideXml += `<a:gd name="adj" fmla="val ${Math.round((slideItemObj.options.rectRadius * EMU * 100000) / Math.min(cx, cy))}"/>`
					} else if (slideItemObj.options.angleRange) {
						for (let i = 0; i < 2; i++) {
							const angle = slideItemObj.options.angleRange[i]
							strSlideXml += `<a:gd name="adj${i + 1}" fmla="val ${convertRotationDegrees(angle)}" />`
						}

						if (slideItemObj.options.arcThicknessRatio) {
							strSlideXml += `<a:gd name="adj3" fmla="val ${Math.round(slideItemObj.options.arcThicknessRatio * 50000)}" />`
						}
					}
					strSlideXml += '</a:avLst></a:prstGeom>'
				}

				// Option: FILL
				strSlideXml += slideItemObj.options.fill ? genXmlColorSelection(slideItemObj.options.fill) : '<a:noFill/>'

				// shape Type: LINE: line color
				if (slideItemObj.options.line) strSlideXml += genXmlLine(slideItemObj.options.line)

				strSlideXml += genXmlEffectLst(slideItemObj.options)

				/* TODO: FUTURE: Text wrapping (copied from MS-PPTX export)
					// Commented out b/c i'm not even sure this works - current code produces text that wraps in shapes and textboxes, so...
					if ( slideItemObj.options.textWrap ) {
						strSlideXml += '<a:extLst>'
									+ `<a:ext uri="${OOXML_EXT.macWrappingTextBox.uri}">`
									+ '<ma14:wrappingTextBoxFlag xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main" val="1"/>'
									+ '</a:ext>'
									+ '</a:extLst>';
					}
				*/

				// B: Close shape Properties
				strSlideXml += '</p:spPr>'

				// C: Add formatted text (text body "bodyPr")
				strSlideXml += genXmlTextBody(slideItemObj)

				// LAST: Close SHAPE =======================================================
				strSlideXml += '</p:sp>'
				break

			case SLIDE_OBJECT_TYPES.image:
				strSlideXml += '<p:pic>'
				strSlideXml += '  <p:nvPicPr>'
				strSlideXml += `<p:cNvPr id="${idx + 2}" name="${slideItemObj.options.objectName}" descr="${encodeXmlEntities(
					slideItemObj.options.altText || slideItemObj.image
				)}">`
				if (slideItemObj.hyperlink?._rId) strSlideXml += genXmlHyperlink(slideItemObj.hyperlink, 'click', 'shape')
				if (slideItemObj.hyperlinkHover?._rId) strSlideXml += genXmlHyperlink(slideItemObj.hyperlinkHover, 'hover', 'shape')
				strSlideXml += '    </p:cNvPr>'
				strSlideXml += '    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
				strSlideXml += '    <p:nvPr>' + genXmlPlaceholder(placeholderObj) + '</p:nvPr>'
				strSlideXml += '  </p:nvPicPr>'
				strSlideXml += '<p:blipFill>'
				// NOTE: This works for both cases: either `path` or `data` contains the SVG
				if (
					(slide._relsMedia || []).filter(rel => rel.rId === slideItemObj.imageRid)[0] &&
					(slide._relsMedia || []).filter(rel => rel.rId === slideItemObj.imageRid)[0].extn === 'svg'
				) {
					strSlideXml += `<a:blip r:embed="rId${(slideItemObj.imageRid ?? 0) - 1}">`
					strSlideXml += slideItemObj.options.transparency ? ` <a:alphaModFix amt="${Math.round((100 - slideItemObj.options.transparency) * 1000)}"/>` : ''
					strSlideXml += ' <a:extLst>'
					strSlideXml += `  <a:ext uri="${OOXML_EXT.svgBlip.uri}">`
					strSlideXml += `   <asvg:svgBlip xmlns:asvg="${OOXML_EXT.svgBlip.ns}" r:embed="rId${slideItemObj.imageRid}"/>`
					strSlideXml += '  </a:ext>'
					strSlideXml += ' </a:extLst>'
					strSlideXml += '</a:blip>'
				} else {
					strSlideXml += `<a:blip r:embed="rId${slideItemObj.imageRid}">`
					strSlideXml += slideItemObj.options.transparency ? `<a:alphaModFix amt="${Math.round((100 - slideItemObj.options.transparency) * 1000)}"/>` : ''
					strSlideXml += '</a:blip>'
				}
				if (sizing?.type) {
					const boxW = sizing.w ? getSmartParseNumber(sizing.w, 'X', slide._presLayout) : cx
					const boxH = sizing.h ? getSmartParseNumber(sizing.h, 'Y', slide._presLayout) : cy
					const boxX = getSmartParseNumber(sizing.x || 0, 'X', slide._presLayout)
					const boxY = getSmartParseNumber(sizing.y || 0, 'Y', slide._presLayout)
					const sourceSize = typeof slideItemObj.options.w === 'number' && typeof slideItemObj.options.h === 'number'
						? { w: slideItemObj.options.w, h: slideItemObj.options.h }
						: { w: imgWidth, h: imgHeight }

					strSlideXml += ImageSizingXml[sizing.type](sourceSize, { w: boxW, h: boxH, x: boxX, y: boxY })
					imgWidth = boxW
					imgHeight = boxH
				} else {
					strSlideXml += '  <a:stretch><a:fillRect/></a:stretch>'
				}
				strSlideXml += '</p:blipFill>'
				strSlideXml += '<p:spPr>'
				strSlideXml += ' <a:xfrm' + locationAttr + '>'
				strSlideXml += `  <a:off x="${x}" y="${y}"/>`
				strSlideXml += `  <a:ext cx="${imgWidth}" cy="${imgHeight}"/>`
				strSlideXml += ' </a:xfrm>'
				strSlideXml += ` <a:prstGeom prst="${rounding ? 'ellipse' : 'rect'}"><a:avLst/></a:prstGeom>`

				// OUTLINE: picture border/frame (issue #35)
				if (slideItemObj.options.line) strSlideXml += genXmlLine(slideItemObj.options.line)

				strSlideXml += genXmlEffectLst(slideItemObj.options)
				strSlideXml += '</p:spPr>'
				strSlideXml += '</p:pic>'
				break

			case SLIDE_OBJECT_TYPES.media:
				if (slideItemObj.mtype === 'online') {
					strSlideXml += '<p:pic>'
					strSlideXml += ' <p:nvPicPr>'
					// IMPORTANT: <p:cNvPr id="" value is critical - if its not the same number as preview image `rId`, PowerPoint throws error!
					strSlideXml += `<p:cNvPr id="${(slideItemObj.mediaRid ?? 0) + 2}" name="${slideItemObj.options.objectName}"/>`
					strSlideXml += ' <p:cNvPicPr/>'
					strSlideXml += ' <p:nvPr>'
					strSlideXml += `  <a:videoFile r:link="rId${slideItemObj.mediaRid}"/>`
					strSlideXml += ' </p:nvPr>'
					strSlideXml += ' </p:nvPicPr>'
					// NOTE: `blip` is diferent than videos; also there's no preview "p:extLst" above but exists in videos
					strSlideXml += ` <p:blipFill><a:blip r:embed="rId${(slideItemObj.mediaRid ?? 0) + 1}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` // NOTE: Preview image is required!
					strSlideXml += ' <p:spPr>'
					strSlideXml += `  <a:xfrm${locationAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
					strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
					strSlideXml += ' </p:spPr>'
					strSlideXml += '</p:pic>'
				} else {
					strSlideXml += '<p:pic>'
					strSlideXml += ' <p:nvPicPr>'
					// IMPORTANT: <p:cNvPr id="" value is critical - if not the same number as preiew image rId, PowerPoint throws error!
					strSlideXml += `<p:cNvPr id="${(slideItemObj.mediaRid ?? 0) + 2}" name="${slideItemObj.options.objectName
					}"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>`
					strSlideXml += ' <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
					strSlideXml += ' <p:nvPr>'
					// ECMA-376: audio references `a:audioFile`, video `a:videoFile` - the rel type already distinguishes them
					strSlideXml += `  <${slideItemObj.mtype === 'audio' ? 'a:audioFile' : 'a:videoFile'} r:link="rId${slideItemObj.mediaRid}"/>`
					strSlideXml += '  <p:extLst>'
					strSlideXml += `   <p:ext uri="${OOXML_EXT.media.uri}">`
					strSlideXml += `    <p14:media xmlns:p14="${OOXML_EXT.media.ns}" r:embed="rId${(slideItemObj.mediaRid ?? 0) + 1}"/>`
					strSlideXml += '   </p:ext>'
					strSlideXml += '  </p:extLst>'
					strSlideXml += ' </p:nvPr>'
					strSlideXml += ' </p:nvPicPr>'
					strSlideXml += ` <p:blipFill><a:blip r:embed="rId${(slideItemObj.mediaRid ?? 0) + 2}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` // NOTE: Preview image is required!
					strSlideXml += ' <p:spPr>'
					strSlideXml += `  <a:xfrm${locationAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
					strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
					strSlideXml += ' </p:spPr>'
					strSlideXml += '</p:pic>'
				}
				break

			case SLIDE_OBJECT_TYPES.zoom:
				strSlideXml += genXmlZoom(slideItemObj, slide, sections, { shapeId: idx + 2, x, y, cx, cy, locationAttr })
				break

			case SLIDE_OBJECT_TYPES.chart:
				strSlideXml += '<p:graphicFrame>'
				strSlideXml += ' <p:nvGraphicFramePr>'
				strSlideXml += `   <p:cNvPr id="${idx + 2}" name="${slideItemObj.options.objectName}" descr="${encodeXmlEntities(slideItemObj.options.altText || '')}"/>`
				strSlideXml += '   <p:cNvGraphicFramePr/>'
				strSlideXml += `   <p:nvPr>${genXmlPlaceholder(placeholderObj)}</p:nvPr>`
				strSlideXml += ' </p:nvGraphicFramePr>'
				strSlideXml += ` <p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>`
				strSlideXml += ' <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
				strSlideXml += '  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
				strSlideXml += `   <c:chart r:id="rId${slideItemObj.chartRid}" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>`
				strSlideXml += '  </a:graphicData>'
				strSlideXml += ' </a:graphic>'
				strSlideXml += '</p:graphicFrame>'
				break

			default:
				strSlideXml += ''
				break
		}
	})

	return strSlideXml
}

/**
 * Append the slide-number placeholder after ordinary objects when numbering is configured.
 * Its last position is required by the existing master/layout/slide compatibility behavior.
 */
function genXmlSlideNumber (slide: PresSlide | SlideLayout): string {
	let strSlideXml = ''
	// STEP 4: Add slide numbers (if any) last
	if (slide._slideNumberProps) {
		// Set some defaults (done here b/c SlideNumber canbe added to masters or slides and has numerous entry points)
		if (!slide._slideNumberProps.align) slide._slideNumberProps.align = 'left'

		strSlideXml += '<p:sp>'
		strSlideXml += ' <p:nvSpPr>'
		strSlideXml += '  <p:cNvPr id="25" name="Slide Number Placeholder 0"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
		strSlideXml += '  <p:nvPr><p:ph type="sldNum" sz="quarter" idx="4294967295"/></p:nvPr>'
		strSlideXml += ' </p:nvSpPr>'
		strSlideXml += ' <p:spPr>'
		strSlideXml += '<a:xfrm>' +
			`<a:off x="${getSmartParseNumber(slide._slideNumberProps.x, 'X', slide._presLayout)}" y="${getSmartParseNumber(slide._slideNumberProps.y, 'Y', slide._presLayout)}"/>` +
			`<a:ext cx="${slide._slideNumberProps.w ? getSmartParseNumber(slide._slideNumberProps.w, 'X', slide._presLayout) : '800000'}" cy="${slide._slideNumberProps.h ? getSmartParseNumber(slide._slideNumberProps.h, 'Y', slide._presLayout) : '300000'}"/>` +
			'</a:xfrm>' +
			' <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
			` <a:extLst><a:ext uri="${OOXML_EXT.macWrappingTextBox.uri}"><ma14:wrappingTextBoxFlag val="0" xmlns:ma14="${OOXML_EXT.macWrappingTextBox.ns}"/></a:ext></a:extLst>` +
			'</p:spPr>'
		strSlideXml += '<p:txBody>'
		strSlideXml += '<a:bodyPr'
		if (slide._slideNumberProps.margin && Array.isArray(slide._slideNumberProps.margin)) {
			strSlideXml += ` lIns="${valToPts(slide._slideNumberProps.margin[3] || 0)}"`
			strSlideXml += ` tIns="${valToPts(slide._slideNumberProps.margin[0] || 0)}"`
			strSlideXml += ` rIns="${valToPts(slide._slideNumberProps.margin[1] || 0)}"`
			strSlideXml += ` bIns="${valToPts(slide._slideNumberProps.margin[2] || 0)}"`
		} else if (typeof slide._slideNumberProps.margin === 'number') {
			strSlideXml += ` lIns="${valToPts(slide._slideNumberProps.margin || 0)}"`
			strSlideXml += ` tIns="${valToPts(slide._slideNumberProps.margin || 0)}"`
			strSlideXml += ` rIns="${valToPts(slide._slideNumberProps.margin || 0)}"`
			strSlideXml += ` bIns="${valToPts(slide._slideNumberProps.margin || 0)}"`
		}
		if (slide._slideNumberProps.valign) {
			strSlideXml += ` anchor="${slide._slideNumberProps.valign.replace('top', 't').replace('middle', 'ctr').replace('bottom', 'b')}"`
		}
		strSlideXml += '/>'
		strSlideXml += '  <a:lstStyle><a:lvl1pPr>'
		if (slide._slideNumberProps.fontFace || slide._slideNumberProps.fontSize || slide._slideNumberProps.color) {
			strSlideXml += `<a:defRPr sz="${Math.round((slide._slideNumberProps.fontSize || 12) * 100)}">`
			if (slide._slideNumberProps.color) strSlideXml += genXmlColorSelection(slide._slideNumberProps.color)
			if (slide._slideNumberProps.fontFace) { strSlideXml += `<a:latin typeface="${slide._slideNumberProps.fontFace}"/><a:ea typeface="${slide._slideNumberProps.fontFace}"/><a:cs typeface="${slide._slideNumberProps.fontFace}"/>` }
			strSlideXml += '</a:defRPr>'
		}
		strSlideXml += '</a:lvl1pPr></a:lstStyle>'
		strSlideXml += '<a:p>'
		if (slide._slideNumberProps.align.startsWith('l')) strSlideXml += '<a:pPr algn="l"/>'
		else if (slide._slideNumberProps.align.startsWith('c')) strSlideXml += '<a:pPr algn="ctr"/>'
		else if (slide._slideNumberProps.align.startsWith('r')) strSlideXml += '<a:pPr algn="r"/>'
		else strSlideXml += '<a:pPr algn="l"/>'
		strSlideXml += `<a:fld id="${SLDNUMFLDID}" type="slidenum"><a:rPr b="${slide._slideNumberProps.bold ? 1 : 0}" lang="en-US"/>`
		strSlideXml += `<a:t>${slide._slideNum}</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>`
		strSlideXml += '</p:txBody></p:sp>'
	}

	return strSlideXml
}

/** Close the shape tree and the `p:cSld` wrapper opened by the public orchestrator. */
function genXmlSlideEnd (slide: PresSlide | SlideLayout): string {
	let strSlideXml = ''
	// STEP 5: Close spTree, add any cSld extensions, and finalize slide XML
	strSlideXml += '</p:spTree>'
	strSlideXml += genXmlCreationId(slide)
	strSlideXml += '</p:cSld>'

	return strSlideXml
}

/**
 * Create the `p:cSld` extension carrying this slide's creation id, or `''` when unset
 * - MS-PPTX 2.2.9 / 2.3.1.4: a stable identity for the slide across saves
 * - the generated value is cached on the slide so repeated exports keep the same id
 * @param {PresSlide | SlideLayout} slide - slide object
 * @returns {string} XML string
 */
/**
 * Create the `p:sld` extension list, currently the `p188:commentRel` pointer (MS-PPTX 2.2.10)
 * - Office writes this alongside the slide relationship; without it some builds ignore the part
 * @note belongs after `p:timing`, at the end of the CT_Slide sequence
 * @param {PresSlide | SlideLayout} slide - slide object
 * @returns {string} XML string
 */
export function genXmlSlideExtLst (slide: PresSlide | SlideLayout): string {
	const comments = (slide as PresSlide).comments ?? []
	if (comments.length === 0) return ''

	return (
		`<p:extLst><p:ext uri="${OOXML_EXT.commentRel.uri}">` +
		`<p188:commentRel xmlns:p188="${OOXML_EXT.commentRel.ns}" r:id="rId${slideCommentRelId(slide as PresSlide)}"/>` +
		'</p:ext></p:extLst>'
	)
}

function genXmlCreationId (slide: PresSlide | SlideLayout): string {
	const requested = slide.creationId
	if (!requested) return ''

	if (typeof requested === 'number') {
		// `ST_UnsignedInt`: a non-integer or out-of-range value would trigger the repair dialog
		if (!isFinite(requested) || requested < 0 || requested > 0xffffffff || Math.floor(requested) !== requested) {
			console.warn(`[pptxgenjs] creationId must be an integer between 0 and 4294967295 - "${String(requested)}" ignored`)
			return ''
		}
	} else if (requested !== true) {
		console.warn(`[pptxgenjs] creationId must be \`true\` or an unsigned 32-bit integer - "${String(requested)}" ignored`)
		return ''
	} else {
		// unique per slide and reproducible - see MS_PPTX_ID_BASE
		slide.creationId = MS_PPTX_ID_BASE.creationId + (slide._slideNum ?? 0)
	}

	return `<p:extLst><p:ext uri="${OOXML_EXT.creationId.uri}"><p14:creationId xmlns:p14="${OOXML_EXT.creationId.ns}" val="${String(slide.creationId)}"/></p:ext></p:extLst>`
}

/**
 * Transforms a slide or slideLayout to resulting XML string - Creates `ppt/slide*.xml`
 * @param {PresSlide|SlideLayout} slideObject - slide object created within createSlideObject
 * @return {string} XML string with <p:cSld> as the root
 */
export function slideObjectToXml (slide: PresSlide | SlideLayout, sections: SectionProps[] = []): string {
	let strSlideXml = slide._name ? `<p:cSld name="${encodeXmlEntities(slide._name)}">` : '<p:cSld>'
	strSlideXml += genXmlSlideBackground(slide)
	strSlideXml += genXmlSlideTreeStart()
	strSlideXml += genXmlSlideObjects(slide, sections)
	strSlideXml += genXmlSlideNumber(slide)
	strSlideXml += genXmlSlideEnd(slide)
	return strSlideXml
}

/**
 * Generate XML Paragraph Properties
 * @param {ISlideObject|TextProps} textObj - text object
 * @param {boolean} isDefault - array of default relations
 * @return {string} XML
 */
