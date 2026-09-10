/**
 * PptxGenJS: schema-constrained numeric attributes must stay in range
 *
 * Same class as #195: a numeric option reaching an XML attribute whose ECMA-376 simple type has a
 * range, without validation. An out-of-range value is a schema violation PowerPoint reads as damaged
 * content. Ranges below are from the ISO/IEC 29500-4 XSDs (`dml-main.xsd`, `dml-chart.xsd`).
 *
 * Run with: `npm test` (node built-in test runner + tsx)
 */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import pptxgen from '../src/pptxgen'

/** attribute -> inclusive [min, max] per the XSD simple type it is declared as */
const RANGES: Array<[string, RegExp, number, number]> = [
	['a:ln@w (ST_LineWidth)', /<a:ln w="(-?[\d.]+)"/g, 0, 20116800],
	['a:spcPts@val (ST_TextSpacingPoint)', /<a:spcPts val="(-?[\d.]+)"/g, 0, 158400],
	['a:spcPct@val (ST_TextSpacingPercent)', /<a:spcPct val="(-?[\d.]+)"/g, 0, 13200000],
	['buAutoNum@startAt (ST_TextBulletStartAtNum)', /<a:buAutoNum [^>]*startAt="(-?[\d.]+)"/g, 1, 32767],
	['a:alpha@val (ST_PositiveFixedPercentage)', /<a:alpha val="(-?[\d.]+)"/g, 0, 100000],
	['sz (ST_TextFontSize)', /sz="(-?[\d.]+)"/g, 100, 400000],
	['c:gapWidth@val (ST_GapAmount)', /<c:gapWidth val="(-?[\d.]+)"/g, 0, 500],
	['c:gapDepth@val (ST_GapAmount)', /<c:gapDepth val="(-?[\d.]+)"/g, 0, 500],
	['c:overlap@val (ST_Overlap)', /<c:overlap val="(-?[\d.]+)"/g, -100, 100],
	['c:holeSize@val (ST_HoleSize)', /<c:holeSize val="(-?[\d.]+)"/g, 1, 90],
	['c:firstSliceAng@val (ST_FirstSliceAng)', /<c:firstSliceAng val="(-?[\d.]+)"/g, 0, 360],
	['c:size@val (ST_MarkerSize)', /<c:size val="(-?[\d.]+)"/g, 2, 72],
]

/** assert every occurrence of every ranged attribute in `xml` sits inside its schema range */
function assertInRange (xml: string, part: string): void {
	for (const [label, regex, min, max] of RANGES) {
		for (const match of xml.matchAll(regex)) {
			const value = Number(match[1])
			assert.ok(value >= min && value <= max, `${part}: ${label} = ${value}, outside ${min}..${max}`)
		}
	}
}

/** build the deck, then check every slide and chart part */
async function assertDeckInRange (pptx: pptxgen): Promise<void> {
	const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	const parts = zip.file(/ppt\/(slides\/slide|charts\/chart)\d+\.xml$/)
	assert.ok(parts.length > 0, 'no slide or chart parts found')
	for (const part of parts) assertInRange(await part.async('string'), part.name)
}

test('out-of-range text options stay inside their schema types', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('negative spacing', { x: 1, y: 1, w: 4, h: 1, lineSpacing: -12, paraSpaceBefore: -5, paraSpaceAfter: -5 })
	slide.addText('huge spacing', { x: 1, y: 2, w: 4, h: 1, lineSpacing: 99999, lineSpacingMultiple: 9999 })
	slide.addText([{ text: 'negative start', options: { bullet: { type: 'number', numberStartAt: -5 } } }], { x: 1, y: 3, w: 4, h: 1 })
	slide.addText([{ text: 'huge start', options: { bullet: { type: 'number', numberStartAt: 99999 } } }], { x: 1, y: 4, w: 4, h: 1 })
	slide.addText('negative outline', { x: 1, y: 5, w: 4, h: 1, outline: { size: -2, color: '000000' } })

	await assertDeckInRange(pptx)
})

test('out-of-range fill and line options stay inside their schema types', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addShape('rect', { x: 1, y: 1, w: 2, h: 1, fill: { color: '00BFD9', transparency: 150 }, line: { color: 'FF0000', width: -3 } })
	slide.addShape('rect', { x: 1, y: 3, w: 2, h: 1, fill: { color: '00BFD9', transparency: -50 }, line: { color: 'FF0000', width: 99999 } })

	await assertDeckInRange(pptx)
})

test('out-of-range chart options stay inside their schema types', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	const data = [{ name: 'S', labels: ['a', 'b'], values: [1, 2] }]
	slide.addChart('bar', data, { x: 0, y: 0, w: 4, h: 3, barGapWidthPct: 900, barOverlapPct: 200, dataLabelFontSize: -12, showValue: true })
	slide.addChart('bar3D', data, { x: 4, y: 0, w: 4, h: 3, barGapDepthPct: 900, catAxisLabelFontSize: -8, valAxisLabelFontSize: 99999 })
	slide.addChart('doughnut', data, { x: 0, y: 3, w: 4, h: 2, holeSize: 0, firstSliceAng: 400 })
	slide.addChart('line', data, { x: 4, y: 3, w: 4, h: 2, lineDataSymbol: 'circle', lineDataSymbolSize: 1, lineSize: -3, legendFontSize: -10, showLegend: true })

	await assertDeckInRange(pptx)
})
