/**
 * The main smoke slide: a zoom, pattern and picture fills, a comment, text runs with an OMML
 * math zone, fields, a picture bullet, locks, a compound line, columns, a hyperlink, a table,
 * and a chart. Must stay ahead of `media`, which it targets by slide number
 * through the zoom and the hyperlink hover.
 */
import type PptxGenJS from '../../src/pptxgen'
import { SAMPLE_PNG } from '../fixtures'

export function addSlideBasicsFixture (pptx: PptxGenJS): void {
	const slide = pptx.addSlide({ transition: { type: 'wipe', direction: 'left', speed: 'slow' } })
	slide.creationId = true
	slide.addZoom({ slideNumber: 2, x: 6, y: 0.5, w: 2, h: 1.2 })
	slide.addShape(pptx.ShapeType.rect, { x: 6, y: 2, w: 1.5, h: 1, fill: { type: 'pattern', pattern: { preset: 'diagCross', color: '0000FF' } } })
	slide.addShape(pptx.ShapeType.rect, { x: 8, y: 2, w: 1.5, h: 1, fill: { type: 'image', image: { data: SAMPLE_PNG, sizing: 'tile' } } })
	slide.addComment({ text: 'consumer smoke comment', author: 'Ada Lovelace', x: 4, y: 2, created: '2026-08-20T09:00:00Z', replies: [{ text: 'ack', author: 'Grace Hopper', created: '2026-08-20T10:00:00Z' }] })
	slide.addText('OOXML consumer smoke test', { x: 0.5, y: 0.5, w: 5, h: 0.5, animation: { type: 'fadeIn' } })
	slide.addText([
		{ text: 'ratio ' },
		{ text: 'a/b', options: { omml: '<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>' } },
	], { x: 0.5, y: 1, w: 5, h: 0.5 })
	slide.addText([{ text: '22/08/2026', options: { field: 'datetime1' } }, { text: ' page ' }, { text: '1', options: { field: 'slidenum' } }], { x: 6, y: 7.2, w: 3, h: 0.4 })
	slide.addText('picture bullet', { x: 0.5, y: 7.2, w: 4, h: 0.4, bullet: { image: SAMPLE_PNG, color: 'FF0000', size: 150 } })
	slide.addShape(pptx.ShapeType.rect, { x: 8, y: 5.5, w: 1.4, h: 0.6, objectName: 'Locked', title: 'Alt title', lock: { noMove: true, noResize: true, noTextEdit: true } })
	slide.addShape(pptx.ShapeType.line, { x: 0.5, y: 6.9, w: 4, h: 0, line: { color: '0000FF', width: 3, compound: 'thickThin', join: 'miter', miterLimit: 400, beginArrowType: 'arrow', beginArrowSize: { width: 'lg', length: 'lg' }, customDash: [{ dash: 400, space: 300 }] } })
	slide.addText('attribute coverage', { x: 6, y: 6.5, w: 3, h: 0.6, capitalization: 'small', upright: true, fontAlign: 'ctr', marginRight: 0.1, underlineLine: { width: 1, color: 'FF0000' }, symbolFontFace: 'Wingdings' })
	slide.addText('column one flows into column two when the box is full', { x: 0.5, y: 5.5, w: 5, h: 1, columns: 2, columnSpacing: 0.25 })
	slide.addShape(pptx.ShapeType.rect, { x: 6.5, y: 5.5, w: 2, h: 0.6, hyperlink: { url: 'https://example.com', tooltip: 'open' }, hyperlinkHover: { slide: 2, tooltip: 'peek' } })
	slide.addTable([['Region', 'Sales'], ['West', '20']], { x: 0.5, y: 1.5, w: 5 })
	slide.addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [10, 20] }], { x: 0.5, y: 3, w: 6, h: 3 })
}
