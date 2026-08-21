/**
 * PptxGenJS: Content parts (MS-PPTX 2.2.3) and ink (2.2.3.1)
 *
 * A content part embeds markup PresentationML does not define - most commonly InkML for pen
 * annotations - as its own package part, referenced from the slide by `p14:contentPart`. Because that
 * element replaces a shape, MS-PPTX requires it inside `mc:AlternateContent`; the spec further
 * requires a `p:sp` fallback for a plain content part and a `p:pic` fallback for ink, so consumers
 * that cannot render the payload still show something.
 *
 * The payload's content type and relationship type are **not** invented here: they belong to the
 * format being embedded, and writing the wrong ones produces a package PowerPoint reports as
 * damaged. The caller supplies both, as they supply the payload.
 */

import { MS_PPTX_NS } from '../core-enums'
import { ContentPartProps, ISlideObject } from '../core-interfaces'
import { encodeXmlEntities } from '../gen-utils'
import { alternateContent } from './markup-compat'

/** Position and size shared by the content part and its fallback shape */
interface ContentPartGeometry {
	shapeId: number
	x: number
	y: number
	cx: number
	cy: number
}

/**
 * Create the `mc:AlternateContent` block for a content part, or `''` when it has no relationship
 * @param {ISlideObject} obj - the content-part object
 * @param {ContentPartGeometry} geom - position and size
 * @returns {string} XML string
 */
export function genXmlContentPart (obj: ISlideObject, geom: ContentPartGeometry): string {
	const props = (obj.options ?? {}) as ContentPartProps & { objectName?: string }
	if (!obj.contentPartRid) {
		console.warn('[pptxgenjs] a content part needs its payload relationship - object omitted')
		return ''
	}

	const name = String(props.objectName ?? 'Content Part')
	const xfrm = `<p14:xfrm><a:off x="${geom.x}" y="${geom.y}"/><a:ext cx="${geom.cx}" cy="${geom.cy}"/></p14:xfrm>`
	const choice =
		`<p14:contentPart r:id="rId${obj.contentPartRid}" p14:bwMode="auto">` +
		'<p14:nvContentPartPr>' +
		`<p14:cNvPr id="${geom.shapeId}" name="${name}"/>` +
		'<p14:cNvContentPartPr/>' +
		'<p14:nvPr/>' +
		'</p14:nvContentPartPr>' +
		xfrm +
		'</p14:contentPart>'

	return alternateContent({ namespaces: { p14: MS_PPTX_NS.p14 }, choice, fallback: genXmlContentPartFallback(obj, geom) })
}

/**
 * Build the fallback shape MS-PPTX mandates for a content part
 * - ink falls back to `p:pic` showing the supplied raster preview (2.2.3.1)
 * - any other content part falls back to `p:sp`, an empty placeholder of the same size (2.2.3)
 * @param {ISlideObject} obj - the content-part object
 * @param {ContentPartGeometry} geom - position and size
 * @returns {string} XML string
 */
function genXmlContentPartFallback (obj: ISlideObject, geom: ContentPartGeometry): string {
	const props = (obj.options ?? {}) as ContentPartProps & { objectName?: string, altText?: string }
	const name = String(props.objectName ?? 'Content Part')
	const descr = encodeXmlEntities(props.altText ?? '')
	const geometry =
		`<a:xfrm><a:off x="${geom.x}" y="${geom.y}"/><a:ext cx="${geom.cx}" cy="${geom.cy}"/></a:xfrm>` +
		'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'

	// Ink without a raster preview would fall back to nothing, so the preview is required upstream
	if (obj.contentPartKind === 'ink' && obj.coverRid) {
		return (
			'<p:pic><p:nvPicPr>' +
			`<p:cNvPr id="${geom.shapeId}" name="${name}" descr="${descr}"/>` +
			'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
			`<p:blipFill><a:blip r:embed="rId${obj.coverRid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
			`<p:spPr>${geometry}</p:spPr></p:pic>`
		)
	}

	return (
		'<p:sp><p:nvSpPr>' +
		`<p:cNvPr id="${geom.shapeId}" name="${name}" descr="${descr}"/>` +
		'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
		`<p:spPr>${geometry}<a:noFill/></p:spPr>` +
		'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>' +
		'</p:sp>'
	)
}
