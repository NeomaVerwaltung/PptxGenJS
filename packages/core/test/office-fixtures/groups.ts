/**
 * Shape groups and connector shapes - the round-trip fixture for #111. LibreOffice renders both, so
 * unlike the Microsoft-only parts this fixture genuinely exercises the feature.
 */
import type PptxGenJS from '../../src/pptxgen'

export function addGroupsFixture (pptx: PptxGenJS): void {
	const groupSlide = pptx.addSlide()
	groupSlide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.5, w: 2, h: 1, objectName: 'fromBox', fill: { color: '4472C4' } })
	groupSlide.addShape(pptx.ShapeType.rect, { x: 5, y: 2.5, w: 2, h: 1, objectName: 'toBox', fill: { color: 'ED7D31' } })
	groupSlide.addConnector({ x: 2.5, y: 1, w: 2.5, h: 1.5, type: 'bentConnector3', line: { color: '333333', width: 2, endArrowType: 'triangle' }, start: { shape: 'fromBox', site: 3 }, end: { shape: 'toBox', site: 1 } })
	groupSlide.addGroup([
		{ shape: { type: pptx.ShapeType.ellipse, options: { x: 1, y: 1, w: 1, h: 1, fill: { color: '70AD47' } } } },
		{ shape: { type: pptx.ShapeType.rect, options: { x: 2.2, y: 1, w: 1, h: 1, fill: { color: 'FFC000' } } } },
		{ text: { text: 'grouped', options: { x: 1, y: 2.2, w: 2.2, h: 0.5, align: 'center' } } },
	], { x: 0.5, y: 4, w: 4.4, h: 3.4 })
}
