/** OOXML relationship-part generation. */

import { CRLF } from '../core-enums'
import { ISlideRel, ISlideRelChart, ISlideRelMedia, PresSlide, SlideLayout } from '../core-interfaces'

type DefaultRelationship = { target: string, type: string }

/** Transform the dynamic and default relationships for one slide-like part. */
function slideObjectRelationsToXml (slide: PresSlide | SlideLayout, defaultRels: DefaultRelationship[]): string {
	let lastRid = 0
	let strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'

	slide._rels.forEach((rel: ISlideRel) => {
		lastRid = Math.max(lastRid, rel.rId)
		if (rel.type.toLowerCase().includes('hyperlink')) {
			if (rel.data === 'slide') strXml += `<Relationship Id="rId${rel.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide${rel.Target}.xml"/>`
			else strXml += `<Relationship Id="rId${rel.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${rel.Target}" TargetMode="External"/>`
		} else if (rel.type.toLowerCase().includes('notesSlide')) {
			strXml += `<Relationship Id="rId${rel.rId}" Target="${rel.Target}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"/>`
		}
	})
	;(slide._relsChart || []).forEach((rel: ISlideRelChart) => {
		lastRid = Math.max(lastRid, rel.rId)
		strXml += `<Relationship Id="rId${rel.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="${rel.Target}"/>`
	})
	;(slide._relsMedia || []).forEach((rel: ISlideRelMedia) => {
		const relRid = rel.rId.toString()
		lastRid = Math.max(lastRid, rel.rId)
		if (rel.type.toLowerCase().includes('image')) {
			strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="' + rel.Target + '"/>'
		} else if (rel.type.toLowerCase().includes('audio')) {
			if (strXml.includes(' Target="' + rel.Target + '"')) strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="' + rel.Target + '"/>'
			else strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" Target="' + rel.Target + '"/>'
		} else if (rel.type.toLowerCase().includes('video')) {
			if (strXml.includes(' Target="' + rel.Target + '"')) strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="' + rel.Target + '"/>'
			else strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="' + rel.Target + '"/>'
		} else if (rel.type.toLowerCase().includes('online')) {
			if (strXml.includes(' Target="' + rel.Target + '"')) strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.microsoft.com/office/2007/relationships/image" Target="' + rel.Target + '"/>'
			else strXml += '<Relationship Id="rId' + relRid + '" Target="' + rel.Target + '" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>'
		}
	})

	defaultRels.forEach((rel, idx) => {
		strXml += `<Relationship Id="rId${lastRid + idx + 1}" Type="${rel.type}" Target="${rel.target}"/>`
	})
	return strXml + '</Relationships>'
}

export function makeXmlSlideLayoutRel (layoutNumber: number, slideLayouts: SlideLayout[]): string {
	return slideObjectRelationsToXml(slideLayouts[layoutNumber - 1], [{ target: '../slideMasters/slideMaster1.xml', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster' }])
}

export function makeXmlSlideRel (slides: PresSlide[], slideLayouts: SlideLayout[], slideNumber: number): string {
	return slideObjectRelationsToXml(slides[slideNumber - 1], [
		{ target: `../slideLayouts/slideLayout${getLayoutIdxForSlide(slides, slideLayouts, slideNumber)}.xml`, type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout' },
		{ target: `../notesSlides/notesSlide${slideNumber}.xml`, type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide' },
	])
}

export function makeXmlNotesSlideRel (slideNumber: number): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>
			<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide${slideNumber}.xml"/>
		</Relationships>`
}

export function makeXmlMasterRel (masterSlide: PresSlide, slideLayouts: SlideLayout[]): string {
	const defaultRels = slideLayouts.map((_layoutDef, idx) => ({ target: `../slideLayouts/slideLayout${idx + 1}.xml`, type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout' }))
	defaultRels.push({ target: '../theme/theme1.xml', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' })
	return slideObjectRelationsToXml(masterSlide, defaultRels)
}

export function makeXmlNotesMasterRel (): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
		</Relationships>`
}

function getLayoutIdxForSlide (slides: PresSlide[], slideLayouts: SlideLayout[], slideNumber: number): number {
	for (let i = 0; i < slideLayouts.length; i++) {
		if (slideLayouts[i]._name === slides[slideNumber - 1]._slideLayout._name) return i + 1
	}
	return 1
}
