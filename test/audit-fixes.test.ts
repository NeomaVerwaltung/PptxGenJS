/**
 * Regression tests for the 2026-07 code-practices audit fixes (issues #17-#31).
 * Each test pins the exact bug it guards; see the issue for the original failure scenario.
 *
 * Run with: `npm test` (node built-in test runner + tsx)
 */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import pptxgen from '../src/pptxgen'

async function buildZip (build: (pptx: pptxgen) => void): Promise<JSZip> {
	const pptx = new pptxgen()
	build(pptx)
	const buf = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer
	return await JSZip.loadAsync(buf)
}

// Chart part numbering uses a module-global counter, so chartN/WorksheetN depend on how many
// presentations were built before in this process - always look the parts up by pattern.
function findFile (zip: JSZip, pattern: RegExp): string {
	const name = Object.keys(zip.files).find(f => pattern.test(f))
	assert.ok(name, `no zip entry matching ${pattern}`)
	return name
}

const PNG_1PX = 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

test('#21: embedded chart workbook table ref has no stray apostrophe', async () => {
	const zip = await buildZip(pptx => {
		pptx.addSlide().addChart(pptx.ChartType.bar, [{ name: 'S', labels: ['a'], values: [1] }], { x: 1, y: 1, w: 4, h: 3 })
	})
	const chartZipBuf = await zip.file(findFile(zip, /^ppt\/embeddings\/Microsoft_Excel_Worksheet\d+\.xlsx$/))!.async('nodebuffer')
	const chartZip = await JSZip.loadAsync(chartZipBuf)
	const tableXml = await chartZip.file('xl/tables/table1.xml')!.async('string')
	const ref = /ref="([^"]+)"/.exec(tableXml)![1]
	assert.doesNotMatch(ref, /'/, `table ref "${ref}" must not contain an apostrophe`)
})

test('#26: serAxisLabelPos value is emitted (operator precedence)', async () => {
	const zip = await buildZip(pptx => {
		pptx.addSlide().addChart(
			pptx.ChartType.bar3d,
			[{ name: 'S', labels: ['a', 'b'], values: [1, 2] }],
			{ x: 1, y: 1, w: 4, h: 3, serAxisLabelPos: 'high' }
		)
	})
	const chartXml = await zip.file(findFile(zip, /^ppt\/charts\/chart\d+\.xml$/))!.async('string')
	assert.match(chartXml, /<c:tickLblPos val="high"\/>/)
})

test('#23: bubble chart embedded worksheet preserves zero values', async () => {
	const zip = await buildZip(pptx => {
		pptx.addSlide().addChart(
			pptx.ChartType.bubble,
			[
				{ name: 'X-Axis', values: [1, 2] },
				{ name: 'S1', values: [0, 5], sizes: [0, 3] },
			],
			{ x: 1, y: 1, w: 4, h: 3 }
		)
	})
	const chartZipBuf = await zip.file(findFile(zip, /^ppt\/embeddings\/Microsoft_Excel_Worksheet\d+\.xlsx$/))!.async('nodebuffer')
	const chartZip = await JSZip.loadAsync(chartZipBuf)
	const sheetXml = await chartZip.file('xl/worksheets/sheet1.xml')!.async('string')
	assert.doesNotMatch(sheetXml, /<v><\/v>/, 'zero values must not be blanked to empty cells')
})

test('#18: image hyperlink URL is XML-escaped in slide rels', async () => {
	const zip = await buildZip(pptx => {
		pptx.addSlide().addImage({ data: PNG_1PX, x: 1, y: 1, w: 1, h: 1, hyperlink: { url: 'https://x.com/?a=1&b=2' } })
	})
	const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')
	assert.match(rels, /Target="https:\/\/x\.com\/\?a=1&amp;b=2"/)
})

test('#18: master/layout name is XML-escaped in cSld', async () => {
	const zip = await buildZip(pptx => {
		pptx.defineSlideMaster({ title: 'R&D "Q3" Master', background: { color: 'FFFFFF' } })
		pptx.addSlide({ masterName: 'R&D "Q3" Master' })
	})
	// layout1 is the built-in DEFAULT; the custom master becomes layout2
	const layoutXml = await zip.file('ppt/slideLayouts/slideLayout2.xml')!.async('string')
	assert.match(layoutXml, /name="R&amp;D &quot;Q3&quot; Master"/)
})

test('#19: SVG image + hyperlink produces unique relationship Ids', async () => {
	const svg = 'image/svg+xml;base64,' + Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"/>').toString('base64')
	const zip = await buildZip(pptx => {
		pptx.addSlide().addImage({ data: svg, x: 1, y: 1, w: 1, h: 1, hyperlink: { url: 'https://x.com' } })
	})
	const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')
	const ids = [...rels.matchAll(/Id="(rId\d+)"/g)].map(m => m[1])
	assert.equal(new Set(ids).size, ids.length, `duplicate rIds in: ${ids.join(', ')}`)
})

test('#20: exporting twice produces identical shadow XML (no double unit-conversion)', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addText('shadowed', {
		x: 1, y: 1, w: 3, h: 1,
		shadow: { type: 'outer', blur: 8, offset: 4, angle: 270, opacity: 0.5, color: '000000' },
	})
	const extract = async (buf: Buffer): Promise<string> => {
		const zip = await JSZip.loadAsync(buf)
		return /<a:effectLst>.*?<\/a:effectLst>/s.exec(await zip.file('ppt/slides/slide1.xml')!.async('string'))![0]
	}
	const first = await extract((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	const second = await extract((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	assert.equal(second, first)
	assert.match(first, /dir="16200000"/) // 270deg * 60000 - applied exactly once
})

test('#20: caller-supplied options objects are not mutated', async () => {
	const pptx = new pptxgen()
	const chartOpts = { x: 1, y: 1, w: 4, h: 3, objectName: 'P&L' }
	const before = JSON.stringify(chartOpts)
	const slide = pptx.addSlide()
	slide.addChart(pptx.ChartType.bar, [{ name: 'S', labels: ['a'], values: [1] }], chartOpts)
	assert.equal(JSON.stringify(chartOpts), before, 'addChart must not mutate the options object')

	const shadow = { type: 'outer' as const, blur: 8, offset: 4, angle: 270, opacity: 0.5, color: '000000' }
	slide.addText('t', { x: 1, y: 2, w: 2, h: 0.5, shadow })
	await pptx.write({ outputType: 'nodebuffer' })
	assert.equal(shadow.angle, 270, 'shadow options must not be unit-converted in place')
	assert.equal(shadow.opacity, 0.5)
})

test('#25: multi-type addChart no longer discards the 3rd-arg options', async () => {
	const zip = await buildZip(pptx => {
		pptx.addSlide().addChart(
			[
				{ type: pptx.ChartType.bar, data: [{ name: 'S1', labels: ['a', 'b'], values: [1, 2] }], options: {} },
				{ type: pptx.ChartType.line, data: [{ name: 'S2', labels: ['a', 'b'], values: [3, 4] }], options: { secondaryValAxis: true, secondaryCatAxis: true } },
			],
			// data slot unused in multi-type form; options passed 3rd must be honored
			undefined as never,
			{ x: 1, y: 1, w: 6, h: 3, showLegend: true, title: 'MULTI_TITLE', showTitle: true }
		)
	})
	const chartXml = await zip.file(findFile(zip, /^ppt\/charts\/chart\d+\.xml$/))!.async('string')
	assert.match(chartXml, /MULTI_TITLE/)
})

test('#31: compression option is honored for explicit outputTypes', async () => {
	const pptx = new pptxgen()
	// enough repetitive content that DEFLATE must beat STORE
	const slide = pptx.addSlide()
	for (let i = 0; i < 20; i++) slide.addText('The quick brown fox jumps over the lazy dog. '.repeat(10), { x: 0.1, y: 0.1, w: 9, h: 5 })
	const stored = (await pptx.write({ outputType: 'nodebuffer', compression: false })) as Buffer
	const deflated = (await pptx.write({ outputType: 'nodebuffer', compression: true })) as Buffer
	assert.ok(deflated.length < stored.length, `expected deflated (${deflated.length}) < stored (${stored.length})`)
})
