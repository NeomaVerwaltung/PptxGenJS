/**
 * PptxGenJS: Zoom objects (MS-PPTX 2.2.15)
 *
 * A zoom is a shape that jumps to a slide (2.10), a section (2.9), or summarises several sections
 * (2.11). The elements live in PowerPoint-2016 namespaces, so each is offered through
 * `mc:AlternateContent`: consumers that understand `p16` get the zoom, everything else renders the
 * `mc:Fallback` - a picture for slide/section zooms, a group shape for a summary zoom.
 */

import { ZOOM_NS } from '../core-enums'
import { ISlideObject, PresSlide, SlideLayout } from '../core-interfaces'
import { encodeXmlEntities, getUuid } from '../gen-utils'

/** Geometry shared by the zoom object and its fallback shape */
interface ZoomGeometry {
	shapeId: number
	x: number
	y: number
	cx: number
	cy: number
	locationAttr: string
}

/**
 * Resolve a zoom's targets to the ids the XML needs
 * - a slide zoom needs the target slide's `p:sldId`, which only the presentation knows
 * - a section zoom needs the GUID from `p14:sectionLst`, so the titles are matched here
 * @param {ISlideObject} obj - the zoom object
 * @param {PresSlide} slide - slide holding the zoom (its `getSlide` resolves slide numbers)
 * @param {Array<{ title: string, _id?: string }>} sections - presentation sections
 * @returns {string[]} resolved target ids, in declaration order
 */
function resolveTargets (obj: ISlideObject, slide: PresSlide, sections: Array<{ title: string, _id?: string }>): string[] {
	const resolved: string[] = []

	;(obj.zoomTargets ?? []).forEach(target => {
		if (typeof target.slideNumber === 'number') {
			const targetSlide = slide.getSlide?.(target.slideNumber)
			if (!targetSlide) {
				console.warn(`[pptxgenjs] zoom target slide ${target.slideNumber} does not exist - zoom target skipped`)
				return
			}
			resolved.push(String(targetSlide._slideId))
			return
		}

		const section = sections.find(item => item.title === target.sectionTitle)
		if (!section?._id) {
			console.warn(`[pptxgenjs] zoom target section "${String(target.sectionTitle)}" does not exist - zoom target skipped`)
			return
		}
		// the GUID must be byte-identical to the one in `p14:sectionLst` or PowerPoint drops the zoom
		resolved.push(section._id)
	})

	return resolved
}

/**
 * Build the `p166:zmPr` zoom properties, shared by all three zoom kinds
 * @param {ISlideObject} obj - the zoom object
 * @param {ZoomGeometry} geom - zoom position and size
 * @returns {string} XML string
 */
function zoomProperties (obj: ISlideObject, geom: ZoomGeometry): string {
	const opts = obj.options ?? {}
	const duration = typeof opts.transitionDur === 'number' && isFinite(opts.transitionDur) && opts.transitionDur >= 0
		? ` p14:transitionDur="${Math.round(opts.transitionDur)}"`
		: ''

	return (
		`<p166:zmPr id="{${getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx')}}" ` +
		`returnToParent="${opts.returnToParent === false ? '0' : '1'}" ` +
		`showBg="${opts.showBg === false ? '0' : '1'}" imageType="preview"${duration}>` +
		coverFill(obj) +
		`<p:spPr>${geometryXml(geom)}</p:spPr>` +
		'</p166:zmPr>'
	)
}

/** The cover image fill, used by both the zoom object and its picture fallback */
function coverFill (obj: ISlideObject): string {
	return `<p:blipFill><a:blip r:embed="rId${obj.zoomRid ?? 0}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
}

/** Position/size XML shared by the zoom and its fallback */
function geometryXml (geom: ZoomGeometry): string {
	return (
		`<a:xfrm${geom.locationAttr}><a:off x="${geom.x}" y="${geom.y}"/><a:ext cx="${geom.cx}" cy="${geom.cy}"/></a:xfrm>` +
		'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
	)
}

/**
 * Create the `mc:AlternateContent` block for a zoom object, or `''` when no target resolved
 * @param {ISlideObject} obj - the zoom object
 * @param {PresSlide | SlideLayout} slide - slide holding the zoom
 * @param {Array<{ title: string, _id?: string }>} sections - presentation sections
 * @param {ZoomGeometry} geom - zoom position and size
 * @returns {string} XML string
 */
export function genXmlZoom (
	obj: ISlideObject,
	slide: PresSlide | SlideLayout,
	sections: Array<{ title: string, _id?: string }>,
	geom: ZoomGeometry
): string {
	const kind = obj.zoomKind ?? 'slide'
	const targets = resolveTargets(obj, slide as PresSlide, sections)
	if (targets.length === 0) {
		console.warn('[pptxgenjs] zoom has no resolvable target - object omitted')
		return ''
	}

	const opts = obj.options ?? {}
	const nsDecls =
		`xmlns:p16="${ZOOM_NS[kind]}" xmlns:p166="${ZOOM_NS.zmPr}" ` +
		'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"'

	let choice = ''
	if (kind === 'slide') {
		choice = `<p16:sldZm><p16:sldZmObj sldId="${targets[0]}">${zoomProperties(obj, geom)}</p16:sldZmObj></p16:sldZm>`
	} else if (kind === 'section') {
		choice = `<p16:sectionZm><p16:sectionZmObj sectionId="${targets[0]}">${zoomProperties(obj, geom)}</p16:sectionZmObj></p16:sectionZm>`
	} else {
		// CT_SummaryZoom requires a layout choice after the objects; `gridLayout` is PowerPoint's default
		// `zmPr@id` identifies the individual zoom object, so each section gets its own
		const objects = targets.map(sectionId => `<p16:summaryZmObj sectionId="${sectionId}">${zoomProperties(obj, geom)}</p16:summaryZmObj>`).join('')
		choice = `<p16:summaryZm>${objects}<p16:gridLayout/></p16:summaryZm>`
	}

	// A summary zoom falls back to a group shape, the other kinds to a picture (MS-PPTX 2.2.15)
	const fallback = kind === 'summary'
		? `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${geom.shapeId}" name="${String(opts.objectName)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
			`<p:grpSpPr><a:xfrm${geom.locationAttr}><a:off x="${geom.x}" y="${geom.y}"/><a:ext cx="${geom.cx}" cy="${geom.cy}"/>` +
			'<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:grpSp>'
		: '<p:pic><p:nvPicPr>' +
			`<p:cNvPr id="${geom.shapeId}" name="${String(opts.objectName)}" descr="${encodeXmlEntities(opts.altText ?? '')}"/>` +
			'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
			coverFill(obj) +
			`<p:spPr>${geometryXml(geom)}</p:spPr></p:pic>`

	return (
		'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
		`<mc:Choice Requires="p16" ${nsDecls}>${choice}</mc:Choice>` +
		`<mc:Fallback>${fallback}</mc:Fallback>` +
		'</mc:AlternateContent>'
	)
}
