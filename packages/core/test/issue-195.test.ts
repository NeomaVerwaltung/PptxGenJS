/**
 * PptxGenJS: #195 - out-of-range numeric props must not reach the XML
 * A negative `<a:ext>` is well-formed but violates `ST_PositiveCoordinate`, and PowerPoint
 * answers that with the repair dialog. Same for `sz` (`ST_TextFontSize`, 100..400000) and
 * the `roundRect` adjust value (0..50000).
 *
 * Run with: `npm test` (node built-in test runner + tsx)
 */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import pptxgen from '../src/pptxgen'

async function readSlide (pptx: pptxgen): Promise<string> {
	const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	const file = zip.file('ppt/slides/slide1.xml')
	assert.ok(file, 'missing ppt/slides/slide1.xml')
	return await file.async('string')
}

/** every numeric value of `attr` in the part */
function attrValues (xml: string, attr: string): number[] {
	return [...xml.matchAll(new RegExp(`${attr}="(-?[\\d.]+)"`, 'g'))].map(match => Number(match[1]))
}

test('#195: negative/NaN w+h never emit a negative or non-numeric extent', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('negative height', { x: 1, y: 1, w: 4, h: -0.01 })
	slide.addShape('rect', { x: 1, y: 2, w: -2, h: 1, fill: { color: '00BFD9' } })
	slide.addText('NaN height', { x: 1, y: 3, w: 4, h: Number.NaN })
	slide.addImage({ data: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAADwyuo0AAAADklEQVR4nGP4jwQYkDkANvEX6SAXxcIAAAAASUVORK5CYII=', x: 1, y: 4, w: -3, h: -1 })

	const xml = await readSlide(pptx)
	for (const attr of ['cx', 'cy']) {
		for (const value of attrValues(xml, attr)) {
			assert.ok(Number.isFinite(value) && value >= 0, `${attr}="${value}" violates ST_PositiveCoordinate`)
		}
	}
})

test('#195: negative x/y offsets are still allowed through', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addShape('rect', { x: -1, y: -0.5, w: 2, h: 1, fill: { color: '00BFD9' } })

	assert.match(await readSlide(pptx), /<a:off x="-914400" y="-457200"\/>/)
})

test('#195: out-of-range fontSize is clamped to ST_TextFontSize', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('negative font', { x: 1, y: 1, w: 4, h: 1, fontSize: -12 })
	slide.addText('huge font', { x: 1, y: 2, w: 4, h: 1, fontSize: 99999 })
	slide.addText([{ text: '', options: { breakLine: true, fontSize: -8 } }, { text: 'after break' }], { x: 1, y: 3, w: 4, h: 1 })

	const xml = await readSlide(pptx)
	for (const value of attrValues(xml, 'sz')) {
		assert.ok(value >= 100 && value <= 400000, `sz="${value}" violates ST_TextFontSize`)
	}
})

test('#195: rectRadius adjust value stays within the roundRect 0..50000 range', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addShape('roundRect', { x: 1, y: 1, w: 3, h: 1, rectRadius: -0.4, fill: { color: '00BFD9' } })
	slide.addShape('roundRect', { x: 1, y: 2.2, w: 3, h: 1, rectRadius: 9, fill: { color: '00BFD9' } })
	// a clamped-to-zero extent makes the `adj` divisor zero - the result must still be in range
	slide.addShape('roundRect', { x: 1, y: 3.4, w: 3, h: -1, rectRadius: 0.2, fill: { color: '00BFD9' } })

	const xml = await readSlide(pptx)
	const adjs = [...xml.matchAll(/<a:gd name="adj" fmla="val (-?[\d.]+)"\/>/g)].map(match => Number(match[1]))
	assert.equal(adjs.length, 3)
	for (const value of adjs) {
		assert.ok(value >= 0 && value <= 50000, `adj val ${value} out of roundRect range`)
	}
})
