/**
 * PptxGenJS: Regression tests for fixed issues
 * One check per bug fixed - each fails if the bug comes back.
 *
 * Run with: `npm test` (node built-in test runner + tsx)
 */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import pptxgen from '../src/pptxgen'
import { genTableToSlides } from '../src/gen-tables'

/** 4x2 px PNG - non-square on purpose, so a 1x1 inch default is obvious */
const PNG_4x2 = 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAADwyuo0AAAADklEQVR4nGP4jwQYkDkANvEX6SAXxcIAAAAASUVORK5CYII='

async function writeZip (pptx: pptxgen): Promise<JSZip> {
	return await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
}

/** the chart counter is module-global, so chart part numbering continues across tests */
async function readChart (zip: JSZip): Promise<string> {
	const file = zip.file(/ppt\/charts\/chart\d+\.xml$/)[0]
	assert.ok(file, 'missing chart part')
	return await file.async('string')
}

async function readPart (zip: JSZip, name: string): Promise<string> {
	const file = zip.file(name)
	assert.ok(file, `missing part: ${name}`)
	return await file.async('string')
}

/** Charts embed a whole .xlsx as a single part - open the inner zip to inspect the worksheet */
async function readEmbeddedXlsx (zip: JSZip): Promise<JSZip> {
	const file = zip.file(/ppt\/embeddings\/.*\.xlsx$/)[0]
	assert.ok(file, 'missing embedded chart workbook')
	return await JSZip.loadAsync(await file.async('nodebuffer'))
}

test('#19/#18: SVG image + hyperlink gets unique rIds and an escaped url', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addImage({ data: 'image/svg+xml;base64,PHN2Zy8+', x: 1, y: 1, w: 1, h: 1, hyperlink: { url: 'https://x.com/?a=1&b=2' } })

	const rels = await readPart(await writeZip(pptx), 'ppt/slides/_rels/slide1.xml.rels')
	const ids = [...rels.matchAll(/Id="(rId\d+)"/g)].map(match => match[1])
	assert.equal(new Set(ids).size, ids.length, `duplicate rIds in slide rels: ${ids.join(', ')}`)
	assert.ok(rels.includes('a=1&amp;b=2'), 'hyperlink url was not XML-escaped')
})

test('#20: shadow options are not mutated, so a second export matches the first', async () => {
	const shadow = { type: 'outer' as const, blur: 3, offset: 2, angle: 45, opacity: 0.5, color: '000000' }
	const pptx = new pptxgen()
	pptx.addSlide().addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, fill: { color: 'FF0000' }, shadow })

	const first = await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')
	const second = await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')
	assert.equal(second, first, 'second export produced different shadow XML (options were mutated)')
	assert.equal(shadow.angle, 45, 'caller shadow options were mutated')
	assert.ok(first.includes('dir="2700000"'), 'shadow angle not converted for XML')
})

test('#84: shape effects share one ordered effect list', async () => {
	const shadow = { type: 'outer' as const, color: '000000', opacity: 0.5, blur: 2, offset: 3, angle: 270 }
	const glow = { size: 8, color: '00AAFF', opacity: 0.6 }
	const softEdge = { radius: 4 }
	const reflection = { blur: 2, distance: 3, direction: 90, opacity: 0.4, scaleY: -1 }
	const pptx = new pptxgen()
	pptx.addSlide().addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, fill: { color: 'FF0000' }, shadow, glow, softEdge, reflection })

	const xml = await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')
	const effectList = /<a:effectLst>[\s\S]*?<\/a:effectLst>/.exec(xml)?.[0] ?? ''
	assert.equal((xml.match(/<a:effectLst>/g) ?? []).length, 1, 'effects were emitted in multiple effect lists')
	assert.ok(effectList.indexOf('<a:glow ') < effectList.indexOf('<a:outerShdw '), 'glow must precede outer shadow')
	assert.ok(effectList.indexOf('<a:outerShdw ') < effectList.indexOf('<a:reflection '), 'shadow must precede reflection')
	assert.ok(effectList.indexOf('<a:reflection ') < effectList.indexOf('<a:softEdge '), 'reflection must precede soft edge')
	assert.ok(effectList.includes('stA="40000"'), 'reflection opacity was not converted')
	assert.equal(shadow.angle, 270, 'caller shadow options were mutated')
	assert.equal(glow.size, 8, 'caller glow options were mutated')
	assert.equal(softEdge.radius, 4, 'caller soft-edge options were mutated')
	assert.equal(reflection.direction, 90, 'caller reflection options were mutated')
})

test('#1083: rich text writes one paragraph-properties element per paragraph', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addText([
		{ text: 'Normal ' },
		{ text: 'bold', options: { bold: true } },
		{ text: ' normal' },
	], { x: 1, y: 1, w: 4, h: 1, bullet: { type: 'bullet' } })

	const xml = await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')
	const paragraph = xml.match(/<a:p>[\s\S]*?<\/a:p>/)?.[0] ?? ''
	assert.equal((paragraph.match(/<a:pPr/g) ?? []).length, 1, 'rich text emitted multiple paragraph-properties elements')
	assert.ok(paragraph.includes('<a:t>bold</a:t>'), 'rich-text runs were not preserved')
})

test('#18: slide master name is XML-escaped', async () => {
	const pptx = new pptxgen()
	pptx.defineSlideMaster({ title: 'R&D "Q3" Master', objects: [] })
	pptx.addSlide({ masterName: 'R&D "Q3" Master' })

	const xml = await readPart(await writeZip(pptx), 'ppt/slideMasters/slideMaster1.xml')
	assert.ok(!/name="[^"]*&(?!amp;|quot;|lt;|gt;|apos;)/.test(xml), 'unescaped entity in cSld name')
})

test('#1443: notes master has no placeholder shapes PowerPoint repairs away', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addNotes('Speaker notes')

	const zip = await writeZip(pptx)
	const notesMaster = await readPart(zip, 'ppt/notesMasters/notesMaster1.xml')
	const notesSlide = await readPart(zip, 'ppt/notesSlides/notesSlide1.xml')
	assert.doesNotMatch(notesMaster, /<p:sp>/, 'notes master contains invalid placeholder shapes')
	assert.match(notesMaster, /<p:spTree>[\s\S]*<\/p:spTree>/, 'notes master is missing its shape tree')
	assert.match(notesSlide, /Speaker notes/, 'speaker notes were not preserved')
})

test('#102: negative line deltas use non-negative extents without reversing arrows', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addShape(pptx.ShapeType.line, {
		x: 3,
		y: 3,
		w: -1.533,
		h: -1.218,
		line: { color: '000000', beginArrowType: 'triangle', endArrowType: 'stealth' },
	})

	const xml = await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')
	const transform = /<a:xfrm flipH="1" flipV="1"><a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/><\/a:xfrm>/.exec(xml)
	assert.ok(transform, 'negative line was not normalized with axis flips')
	assert.ok(Number(transform[1]) < 3 * 914400 && Number(transform[2]) < 3 * 914400, 'line offset did not retain its endpoints')
	assert.ok(Number(transform[3]) > 0 && Number(transform[4]) > 0, 'line extents must be non-negative')
	assert.match(xml, /<a:headEnd type="triangle"/, 'line start arrow was not retained')
	assert.match(xml, /<a:tailEnd type="stealth"/, 'line end arrow was not retained')
})

test('#21/#23: bubble chart workbook keeps zeros and has a valid table ref', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addChart(pptx.ChartType.bubble, [
		{ name: 'X-Axis', values: [1, 2, 3] },
		{ name: 'R&D', values: [0, 5, 6], sizes: [0, 2, 3] },
	], { x: 1, y: 1, w: 4, h: 3 })
	slide.addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [1, 2] }], { x: 1, y: 4, w: 4, h: 2 })

	const zip = await writeZip(pptx)
	const bubbleXlsx = await JSZip.loadAsync(await zip.file(/ppt\/embeddings\/.*\.xlsx$/)[0].async('nodebuffer'))
	const sheet = await readPart(bubbleXlsx, 'xl/worksheets/sheet1.xml')
	assert.ok(sheet.includes('<v>0</v>'), 'bubble worksheet dropped a legitimate zero value')
	const table = await readPart(bubbleXlsx, 'xl/tables/table1.xml')
	assert.ok(table.includes('name="R&amp;D"'), 'bubble series name was not XML-escaped')

	for (const name of Object.keys(zip.files).filter(key => key.endsWith('.xlsx'))) {
		const inner = await JSZip.loadAsync(await zip.file(name)!.async('nodebuffer'))
		const tableXml = await readPart(inner, 'xl/tables/table1.xml')
		assert.ok(!/ref="[^"]*'/.test(tableXml), `stray apostrophe in table ref: ${name}`)
	}
})

test('#38: multi-level category chart writes a coherent worksheet', async () => {
	const LABELS = [
		['Gear', 'Berg', 'Motr', 'Swch', 'Plug', 'Cord'],
		['Mech', '', '', 'Elec', '', ''],
	]
	const pptx = new pptxgen()
	pptx.addSlide().addChart(pptx.ChartType.bar, [
		{ name: 'West', labels: LABELS, values: [11, 8, 3, 0, 11, 3] },
		{ name: 'East', labels: LABELS, values: [1, 2, 3, 4, 5, 6] },
	], { x: 1, y: 1, w: 6, h: 4 })

	const xlsx = await readEmbeddedXlsx(await writeZip(pptx))
	const sheet = await readPart(xlsx, 'xl/worksheets/sheet1.xml')
	const strings = await readPart(xlsx, 'xl/sharedStrings.xml')
	const arrStrings = [...strings.matchAll(/<si>(.*?)<\/si>/g)].map(match => /<t[^>]*>([^<]*)<\/t>/.exec(match[1])?.[1] ?? '')

	// series names occupy the two header cells right of the label levels
	assert.ok(sheet.includes('<c r="C1" t="s"><v>1</v></c>'), 'series header cell missing/misplaced')
	assert.equal(arrStrings[1], 'West')
	// outer label of row 2 ("Mech") sits in col A, inner label ("Gear") in col B
	const idxMech = arrStrings.indexOf('Mech')
	const idxGear = arrStrings.indexOf('Gear')
	assert.ok(sheet.includes(`<c r="A2" t="s"><v>${idxMech}</v></c>`), 'outer label cell wrong')
	assert.ok(sheet.includes(`<c r="B2" t="s"><v>${idxGear}</v></c>`), 'inner label cell wrong')
	// zeros survive, and outer labels are merged over their three rows
	assert.ok(sheet.includes('<c r="C5"><v>0</v></c>'), 'multi-cat worksheet dropped a zero value')
	assert.ok(sheet.includes('<mergeCell ref="A2:A4"/>') && sheet.includes('<mergeCell ref="A5:A7"/>'), 'outer label rows were not merged')
})

test('#1466: flat categories use strRef while multi-level categories keep multiLvlStrRef', async () => {
	const flat = new pptxgen()
	flat.addSlide().addChart(flat.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [10, 20] }], { x: 1, y: 1, w: 6, h: 4 })
	const flatChart = await readChart(await writeZip(flat))
	assert.match(flatChart, /<c:cat>\s*<c:strRef>/, 'flat categories were not written as strRef')
	assert.doesNotMatch(flatChart, /<c:multiLvlStrRef>/, 'flat categories used a multi-level reference')

	const multiLevel = new pptxgen()
	multiLevel.addSlide().addChart(multiLevel.ChartType.bar, [{ name: 'Sales', labels: [['Q1', 'Q2'], ['2026', '']], values: [10, 20] }], { x: 1, y: 1, w: 6, h: 4 })
	const multiLevelChart = await readChart(await writeZip(multiLevel))
	assert.match(multiLevelChart, /<c:cat>\s*<c:multiLvlStrRef>/, 'multi-level categories no longer use multiLvlStrRef')
})

test('#1430: embedded workbook preserves per-series data table formats and zeros', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addChart([
		{ type: pptx.ChartType.bar, data: [{ name: 'ABC', labels: ['2012', '2013'], values: [100000, 0] }], options: { dataTableFormatCode: '₹#,##0' } },
		{ type: pptx.ChartType.line, data: [{ name: 'Share', labels: ['2012', '2013'], values: [0.17, 0] }], options: { dataTableFormatCode: '0%' } },
	], [], { x: 1, y: 1, w: 6, h: 4 })

	const xlsx = await readEmbeddedXlsx(await writeZip(pptx))
	const sheet = await readPart(xlsx, 'xl/worksheets/sheet1.xml')
	const styles = await readPart(xlsx, 'xl/styles.xml')
	assert.match(sheet, /<c r="B3" s="1"><v>0<\/v><\/c>/, 'currency zero has no worksheet style')
	assert.match(sheet, /<c r="C3" s="2"><v>0<\/v><\/c>/, 'percentage zero has no worksheet style')
	assert.match(styles, /numFmtId="164" formatCode="₹#,##0"/, 'currency number format is absent')
	assert.match(styles, /numFmtId="165" formatCode="0%"/, 'percentage number format is absent')
})

test('#25: multi-type chart honors the options argument', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addChart(
		[{ type: pptx.ChartType.bar, data: [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [1, 2] }], options: {} }],
		[{ name: 'Sales', labels: ['Q1', 'Q2'], values: [1, 2] }],
		{ x: 1, y: 1, w: 4, h: 3, showLegend: true }
	)

	const chart = await readChart(await writeZip(pptx))
	assert.ok(chart.includes('<c:legend>'), 'options argument was discarded')
})

test('#1188: pie chart titles support italic text', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addChart(pptx.ChartType.pie, [{ name: 'Sales', labels: ['Q1'], values: [1] }], {
		x: 1, y: 1, w: 4, h: 3, showTitle: true, title: 'Sales', titleItalic: true,
	})

	const title = (await readChart(await writeZip(pptx))).match(/<c:title>[\s\S]*?<\/c:title>/)?.[0] ?? ''
	assert.match(title, /<a:rPr[^>]* i="1"/, 'pie title italic was not emitted')
})

test('#1420: chart title and legend set the East Asian font slot', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addChart(pptx.ChartType.pie, [{ name: '状态', labels: ['智能驾驶', '已完成'], values: [5, 25] }], {
		x: 1, y: 1, w: 4, h: 3,
		showTitle: true, title: '销售部门整体业绩', titleFontFace: 'Microsoft YaHei',
		showLegend: true, legendFontFace: 'Microsoft YaHei',
	})

	const chart = await readChart(await writeZip(pptx))
	assert.match(chart, /<a:rPr[\s\S]*?<a:ea typeface="Microsoft YaHei"\/>/, 'title run is missing the East Asian font')
	assert.match(chart, /<c:legend>[\s\S]*?<a:ea\s+typeface="Microsoft YaHei"\/>/, 'legend is missing the East Asian font')
})

test('#1245: scatter axis can cross at zero', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addChart(pptx.ChartType.scatter, [
		{ name: 'X', values: [0, 1] },
		{ name: 'Y', values: [90, 80] },
	], { x: 1, y: 1, w: 4, h: 3, valAxisCrossesAt: 0 })

	assert.ok((await readChart(await writeZip(pptx))).includes('<c:crossesAt val="0"/>'), 'zero was replaced with an invalid axis crossing')
})

test('#1355: a scatter chart keeps a value x-axis in a combo chart', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addChart([
		{ type: pptx.ChartType.bar, data: [{ name: 'Bars', labels: ['Mon', 'Tue'], values: [17, 26] }], options: { barDir: 'bar' } },
		{ type: pptx.ChartType.scatter, data: [{ name: 'X', labels: ['Mon', 'Tue'], values: [1, 2] }, { name: 'Y', labels: ['Mon', 'Tue'], values: [25, 35] }], options: { secondaryValAxis: true, secondaryCatAxis: true } },
	], { x: 1, y: 1, w: 6, h: 3, valAxes: [{}, {}], catAxes: [{}, {}] })

	const chart = await readChart(await writeZip(pptx))
	assert.equal((chart.match(/<c:catAx>/g) ?? []).length, 1, 'scatter x-axis was emitted as a category axis')
	assert.equal((chart.match(/<c:valAx>/g) ?? []).length, 3, 'scatter combo chart is missing a value axis')
})

test('#26: serAxisLabelPos is honored', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addChart(pptx.ChartType.bar3d, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [1, 2] }], {
		x: 1, y: 1, w: 4, h: 3, barDir: 'col', serAxisLabelPos: 'high',
	})

	const chart = await readChart(await writeZip(pptx))
	assert.ok(chart.includes('val="high"'), 'serAxisLabelPos was ignored')
})

test('#976: scatter charts honor catAxisLabelPos', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addChart(pptx.ChartType.scatter, [
		{ name: 'X', values: [1, 2] },
		{ name: 'Y', values: [3, 4] },
	], { x: 1, y: 1, w: 4, h: 3, catAxisLabelPos: 'low' })

	assert.ok((await readChart(await writeZip(pptx))).includes('<c:tickLblPos val="low"/>'), 'scatter category-axis label position was ignored')
})

test('#34: image without w/h is sized from the image itself', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addImage({ data: PNG_4x2, x: 1, y: 1 })

	const xml = await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')
	const pic = /<p:pic>[\s\S]*?<\/p:pic>/.exec(xml)?.[0] ?? ''
	const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(pic)
	assert.ok(ext, 'no image extent found')
	// 4x2 px at 96 DPI = 0.0417 x 0.0208 inch; 1 inch = 914400 EMU
	assert.equal(Number(ext[1]), Math.round((4 / 96) * 914400))
	assert.equal(Number(ext[2]), Math.round((2 / 96) * 914400))
})

test('#1286: contain sizing preserves the ratio of mixed pixel dimensions', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addImage({
		data: PNG_4x2,
		x: '19%', y: '54%', w: 2899, h: 97,
		sizing: { type: 'contain', w: '36%', h: '3%' },
	})

	const xml = await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')
	assert.ok(xml.includes('<a:srcRect l="0" r="0" t="-20047" b="-20047"/>'), 'contain sizing mixed image units and produced invalid crop XML')
})

test('#39: auto-paged tables account for cell margins', async () => {
	const rows = Array.from({ length: 30 }, (_, idx) => [`Row ${idx} cell A`, `Row ${idx} cell B`])
	const pageCount = (margin: number): number => {
		const pptx = new pptxgen()
		pptx.addSlide().addTable(rows, { x: 0.5, y: 0.5, w: 8, autoPage: true, margin })
		return pptx.slides.length
	}

	assert.ok(pageCount(0.5) > pageCount(0), 'large cell margins did not increase the page count')
})

test('#1472: auto-paging one table does not move a sibling table', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	const options = { x: 0.5, y: 3.5, w: 4, autoPage: true }

	slide.addTable(Array.from({ length: 20 }, (_, idx) => [`Long table row ${idx}`]), options)
	options.x = 5
	slide.addTable([['Short table row 1'], ['Short table row 2']], options)

	assert.equal(pptx.slides.length, 2, 'the long table did not create a second slide')
	assert.equal(options.y, 3.5, 'auto-paging mutated the shared table position')

	const xml = await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')
	const secondTable = /<p:graphicFrame>[\s\S]*?<p:cNvPr[^>]*name="Table 1"[\s\S]*?<\/p:graphicFrame>/.exec(xml)?.[0] ?? ''
	assert.ok(secondTable.includes(`<a:off x="${5 * 914400}" y="${3.5 * 914400}"/>`), 'the second table was not kept at its requested position')
})

test('#29: BorderProps accepts `width` (points) alongside the deprecated `pt`', async () => {
	const cellXml = async (border: Record<string, unknown>): Promise<string> => {
		const pptx = new pptxgen()
		pptx.addSlide().addTable([[{ text: 'A', options: { border: [border, border, border, border] } }]], { x: 1, y: 1, w: 4 })
		return await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')
	}

	assert.equal(await cellXml({ color: 'FF0000', width: 3 }), await cellXml({ color: 'FF0000', pt: 3 }), '`width` and `pt` produced different borders')
	assert.ok((await cellXml({ color: 'FF0000', width: 3 })).includes('w="38100"'), '3pt border not emitted')
})

test('#1235: HTML table conversion preserves fractional border widths', async () => {
	class Cell {
		innerText = 'A'
		offsetWidth = 100
		getAttribute (): null { return null }
	}
	class Row {
		cells = [new Cell()]
	}
	const cell = new Cell()
	const row = new Row()
	const globals = globalThis as unknown as Record<string, unknown>
	const original = Object.fromEntries(['document', 'window', 'HTMLTableCellElement', 'HTMLTableRowElement'].map(key => [key, globals[key]]))
	const styles: Record<string, string> = {
		'background-color': 'rgba(0, 0, 0, 0)', 'border-bottom-color': 'rgb(0, 0, 0)', 'border-bottom-width': '0px',
		'border-left-color': 'rgb(0, 0, 0)', 'border-left-width': '0.25px', 'border-right-color': 'rgb(0, 0, 0)', 'border-right-width': '0px',
		'border-top-color': 'rgb(0, 0, 0)', 'border-top-width': '0px', color: 'rgb(0, 0, 0)', 'font-family': 'Arial', 'font-size': '12px',
		'font-weight': 'normal', 'padding-bottom': '0px', 'padding-left': '0px', 'padding-right': '0px', 'padding-top': '0px', 'text-align': 'left', 'vertical-align': 'top',
	}
	Object.assign(globals, {
		document: {
			getElementById: () => ({}),
			querySelector: () => null,
			querySelectorAll: (selector: string) => selector === '#table tr:first-child td' ? [cell] : selector === '#table tbody tr' ? [row] : [],
		},
		window: { getComputedStyle: () => ({ getPropertyValue: (name: string) => styles[name] ?? '' }) },
		HTMLTableCellElement: Cell,
		HTMLTableRowElement: Row,
	})

	try {
		const pptx = new pptxgen()
		genTableToSlides(pptx, 'table', { w: 4 })
		assert.ok((await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')).includes('<a:lnL w="3175"'), 'fractional CSS border was rounded')
	} finally {
		Object.assign(globals, original)
	}
})

test('#29: defineLayout accepts `w`/`h` as aliases of `width`/`height`', () => {
	const pptx = new pptxgen()
	pptx.defineLayout({ name: 'A3_WH', width: 16.5, height: 11.7 })
	pptx.defineLayout({ name: 'A3_SHORT', w: 16.5, h: 11.7 })

	pptx.layout = 'A3_WH'
	const viaWidth = { w: pptx.presLayout.width, h: pptx.presLayout.height }
	pptx.layout = 'A3_SHORT'
	assert.deepEqual({ w: pptx.presLayout.width, h: pptx.presLayout.height }, viaWidth)
})

test('#33: picture/chart/table placeholders emit their `p:ph` type on the layout', async () => {
	const pptx = new pptxgen()
	pptx.defineSlideMaster({
		title: 'PH_MASTER',
		objects: [
			{ placeholder: { options: { name: 'pic1', type: 'pic', x: 0.5, y: 0.5, w: 3, h: 2 }, text: '' } },
			{ placeholder: { options: { name: 'chart1', type: 'chart', x: 4, y: 0.5, w: 3, h: 2 }, text: '' } },
			{ placeholder: { options: { name: 'tbl1', type: 'tbl', x: 0.5, y: 3, w: 3, h: 2 }, text: '' } },
		],
	})
	const slide = pptx.addSlide({ masterName: 'PH_MASTER' })
	slide.addImage({ data: PNG_4x2, placeholder: 'pic1' })

	const zip = await writeZip(pptx)
	const layouts = await Promise.all([1, 2].map(async num => await readPart(zip, `ppt/slideLayouts/slideLayout${num}.xml`)))
	const layout = layouts.find(xml => xml.includes('PH_MASTER')) ?? ''
	for (const type of ['pic', 'chart', 'tbl']) {
		assert.match(layout, new RegExp(`type="${type}"`), `${type} placeholder has no p:ph type`)
	}

	// the image placed into the picture placeholder inherits its position and references the placeholder
	const pic = /<p:pic>[\s\S]*?<\/p:pic>/.exec(await readPart(zip, 'ppt/slides/slide1.xml'))?.[0] ?? ''
	assert.match(pic, /type="pic"/, 'slide image does not reference the picture placeholder')
	assert.ok(pic.includes(`<a:off x="${Math.round(0.5 * 914400)}"`), 'slide image did not inherit the placeholder position')
})

test('#32: masters accept any shape type, tables and media', async () => {
	const pptx = new pptxgen()
	pptx.defineSlideMaster({
		title: 'RICH_MASTER',
		objects: [
			{ shape: { type: pptx.ShapeType.triangle, options: { x: 0.5, y: 0.5, w: 1, h: 1, fill: { color: '00AA00' } } } },
			{ table: { rows: [['A', 'B']], options: { x: 2, y: 0.5, w: 4 } } },
			{ media: { type: 'audio', data: 'audio/mp3;base64,QQ==', x: 7, y: 0.5, w: 1, h: 1 } },
		],
	})
	pptx.addSlide({ masterName: 'RICH_MASTER' })

	const zip = await writeZip(pptx)
	// layout1 is the built-in DEFAULT layout; the defined master gets its own
	const idx = (await Promise.all([1, 2].map(async num => await readPart(zip, `ppt/slideLayouts/slideLayout${num}.xml`)))).findIndex(xml => xml.includes('RICH_MASTER')) + 1
	assert.ok(idx > 0, 'no layout generated for the defined master')
	const layout = await readPart(zip, `ppt/slideLayouts/slideLayout${idx}.xml`)
	assert.ok(layout.includes('prst="triangle"'), 'shape not rendered on the master layout')
	assert.ok(layout.includes('<a:tbl>'), 'table not rendered on the master layout')
	assert.ok(layout.includes('<a:audioFile') || layout.includes('<a:videoFile'), 'media not rendered on the master layout')
	assert.match(await readPart(zip, `ppt/slideLayouts/_rels/slideLayout${idx}.xml.rels`), /media\/media/, 'media rel missing from the layout')
})

test('#28: friendly dataLabelPosition names are translated to OOXML codes', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addChart(pptx.ChartType.bar, [{ name: 'S', labels: ['A'], values: [1] }], { x: 1, y: 1, w: 4, h: 3, showValue: true, dataLabelPosition: 'outsideEnd' })

	assert.ok((await readChart(await writeZip(pptx))).includes('<c:dLblPos val="outEnd"/>'), 'friendly name not translated')
})

test('#28: a dataLabelPosition invalid for the chart type is dropped with a warning', async () => {
	const warnings: string[] = []
	const orig = console.warn
	console.warn = (msg: string) => warnings.push(msg)
	try {
		const pptx = new pptxgen()
		// 'bestFit' is pie-only - on a bar chart it makes PowerPoint offer to repair the file
		pptx.addSlide().addChart(pptx.ChartType.bar, [{ name: 'S', labels: ['A'], values: [1] }], { x: 1, y: 1, w: 4, h: 3, showValue: true, dataLabelPosition: 'bestFit' })

		assert.ok(!(await readChart(await writeZip(pptx))).includes('<c:dLblPos'), 'invalid dLblPos was emitted')
		assert.ok(warnings.some(msg => msg.includes('dataLabelPosition')), `no warning logged: ${warnings.join(' | ')}`)
	} finally {
		console.warn = orig
	}
})

test('#80: pie labels honor their requested position and custom text', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addChart(pptx.ChartType.pie, [{ name: 'Share', labels: ['A', 'B', 'C'], values: [30, 50, 20], dataLabels: ['Alpha & Co', 'Beta', 'Gamma'] }], {
		x: 0.5, y: 0.5, w: 4, h: 3, showPercent: true, dataLabelPosition: 'outsideEnd', dataLabelFontSize: 12,
	})

	const chart = await readChart(await writeZip(pptx))
	assert.ok(chart.includes('<c:dLblPos val="outEnd"/>'), 'pie position was not honored')
	assert.ok(chart.includes('<a:t>Alpha &amp; Co</a:t>'), 'custom pie label was not XML-escaped')
	assert.ok(chart.includes('sz="1200"'), 'custom pie label ignored the configured font size')
	const firstLabel = /<c:dLbl>[\s\S]*?<\/c:dLbl>/.exec(chart)?.[0] ?? ''
	assert.ok(firstLabel.includes('<c:showVal val="0"/>'), 'custom labels must not be combined with numeric values')
})

test('#80: custom labels are additive for series-based charts', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['A', 'B'], values: [10, 20], dataLabels: ['Q1', 'Q2'] }], {
		x: 0.5, y: 0.5, w: 4, h: 3, showValue: true, dataLabelPosition: 'outsideEnd',
	})
	const chart = await readChart(await writeZip(pptx))
	assert.ok(chart.includes('<a:t>Q1</a:t>'), 'custom series label missing')
	assert.ok(chart.includes('<c:showVal val="0"/>'), 'custom series label was combined with the numeric value')
})

test('#31: `compression` is honoured for every outputType, not just STREAM', async () => {
	const build = async (compression: boolean): Promise<number> => {
		const pptx = new pptxgen()
		// repetitive text compresses well, so the two sizes are clearly different
		pptx.addSlide().addText('compress me '.repeat(500), { x: 0.5, y: 0.5, w: 9, h: 5 })
		return ((await pptx.write({ outputType: 'nodebuffer', compression })) as Buffer).byteLength
	}

	assert.ok(await build(true) < await build(false), '`compression: true` was ignored for outputType: nodebuffer')
})

test('#37: a per-series color overrides the chartColors cycle', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addChart(pptx.ChartType.bar, [
		{ name: 'A', labels: ['Q1'], values: [1] },
		{ name: 'B', labels: ['Q1'], values: [2], color: 'FF0000' },
	], { x: 1, y: 1, w: 4, h: 3 })

	const xml = await readChart(await writeZip(pptx))
	const sers = [...xml.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map(match => match[0])
	assert.equal(sers.length, 2)
	assert.ok(sers[1].includes('<a:srgbClr val="FF0000"/>'), 'series color override not applied')
	assert.ok(!sers[0].includes('<a:srgbClr val="FF0000"/>'), 'override leaked into the other series')
})

test('#36: table style flags and style id are emitted in a:tblPr', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addTable([['A', 'B']], { x: 1, y: 1, w: 4, bandRow: true, firstRow: true, tableStyleId: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' })

	const xml = await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')
	assert.ok(xml.includes('<a:tblPr firstRow="1" bandRow="1">'), `tblPr flags missing: ${/<a:tblPr[\s\S]*?tblPr>/.exec(xml)?.[0] ?? xml}`)
	assert.ok(xml.includes('<a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>'), 'table style id missing')
})

test('#36: tables without style options still emit an empty a:tblPr', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addTable([['A', 'B']], { x: 1, y: 1, w: 4 })
	assert.ok((await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')).includes('<a:tblPr/>'))
})

test('#79: autoPageRepeatHeader marks firstRow and respects an explicit override', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addTable(
		[['Header A', 'Header B'], ...Array.from({ length: 40 }, (_, index) => [`Row ${index}`, 'Value'])],
		{ x: 0.5, y: 0.5, w: 8, autoPage: true, autoPageRepeatHeader: true, fontSize: 18 }
	)
	const zip = await writeZip(pptx)
	const slideFiles = zip.file(/ppt\/slides\/slide\d+\.xml/) ?? []
	assert.ok(slideFiles.length > 1, 'table did not paginate')
	for (const slideFile of slideFiles) {
		const xml = await slideFile.async('string')
		assert.ok(xml.includes('<a:tblPr firstRow="1"/>'), `header semantics missing from ${slideFile.name}`)
	}

	const override = new pptxgen()
	override.addSlide().addTable([['Header A', 'Header B'], ['A', 'B']], { x: 0.5, y: 0.5, w: 8, autoPageRepeatHeader: true, firstRow: false })
	assert.ok((await readPart(await writeZip(override), 'ppt/slides/slide1.xml')).includes('<a:tblPr firstRow="0"/>'), 'explicit firstRow:false was ignored')

	const rowspan = new pptxgen()
	rowspan.addSlide().addTable(
		[[{ text: 'Spans all rows', options: { rowspan: 41 } }, 'First'], ...Array.from({ length: 40 }, (_, index) => [`Row ${index}`])],
		{ x: 0.5, y: 0.5, w: 8, h: 2, autoPage: true }
	)
	const rowspanSlides = (await writeZip(rowspan)).file(/ppt\/slides\/slide\d+\.xml/) ?? []
	assert.ok(rowspanSlides.length > 1, 'rowspan table did not paginate')
	for (const slideFile of rowspanSlides) {
		const rows = (await slideFile.async('string')).match(/<a:tr\b[\s\S]*?<\/a:tr>/g) ?? []
		assert.ok(rows.every(row => (row.match(/<a:tc(?:\s|>)/g) ?? []).length === 2), `rowspan changed a column position in ${slideFile.name}`)
	}
})

test('#35: images accept a line/outline and emit it in the picture spPr', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addImage({ data: PNG_4x2, x: 1, y: 1, w: 2, h: 1, line: { color: 'FF0000', width: 2, dashType: 'dash' } })

	const xml = await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')
	const pic = /<p:pic>[\s\S]*?<\/p:pic>/.exec(xml)?.[0] ?? ''
	assert.ok(pic.includes('<a:ln w="25400">'), `picture outline width missing: ${pic}`)
	assert.ok(pic.includes('<a:srgbClr val="FF0000"/>'), 'picture outline color missing')
	assert.ok(pic.includes('<a:prstDash val="dash"/>'), 'picture outline dash type missing')
})

test('#137/#136/#153: presentation, view and document properties reach their parts', async () => {
	const pptx = new pptxgen()
	pptx.documentProps = { category: 'Reports', contentStatus: 'Final', keywords: 'a, b', language: 'de-DE', version: '2.1', manager: 'Ada', template: 'Corp.potx', hyperlinkBase: 'https://x.test', totalEditTime: 42 }
	pptx.slideSizeType = 'screen16x9'
	pptx.photoAlbum = { blackWhite: true, layout: '2pic', frame: 'frameStyle3' }
	pptx.kinsoku = { lang: 'ja-JP', invalidStartChars: ')]}', invalidEndChars: '([{' }
	pptx.printProps = { what: 'handouts4', colorMode: 'gray', frameSlides: true }
	pptx.recentColors = ['FF0000', { scheme: 'accent1' }]
	pptx.viewProps = { lastView: 'sldThumbnailView', showComments: false, zoom: 75, gridSpacing: 0.5, snapToGrid: true, snapToObjects: false }
	const slide = pptx.addSlide()
	slide.addText('one', { x: 1, y: 1 })
	slide.addNotes('speaker note')
	pptx.addSlide().hidden = true

	const zip = await writeZip(pptx)
	const core = await readPart(zip, 'docProps/core.xml')
	for (const el of ['<cp:category>Reports</cp:category>', '<cp:contentStatus>Final</cp:contentStatus>', '<cp:keywords>a, b</cp:keywords>', '<dc:language>de-DE</dc:language>', '<cp:version>2.1</cp:version>']) {
		assert.ok(core.includes(el), `core.xml missing ${el}`)
	}

	// app.xml counts are derived, and each element may appear only once in CT_Properties
	const app = await readPart(zip, 'docProps/app.xml')
	for (const el of ['<Manager>Ada</Manager>', '<Template>Corp.potx</Template>', '<HyperlinkBase>https://x.test</HyperlinkBase>', '<TotalTime>42</TotalTime>', '<Notes>1</Notes>', '<HiddenSlides>1</HiddenSlides>', '<Paragraphs>1</Paragraphs>']) {
		assert.ok(app.includes(el), `app.xml missing ${el}`)
		const tag = /^<(\w+)>/.exec(el)?.[1] ?? ''
		assert.equal((app.match(new RegExp(`<${tag}>`, 'g')) ?? []).length, 1, `app.xml emits <${tag}> more than once`)
	}

	const pres = await readPart(zip, 'ppt/presentation.xml')
	assert.ok(pres.includes(' type="screen16x9"/>'), 'sldSz type missing')
	assert.ok(pres.includes('<p:photoAlbum bw="1" layout="2pic" frame="frameStyle3"/>'), `photoAlbum wrong: ${pres.slice(0, 400)}`)
	assert.ok(pres.includes('<p:kinsoku lang="ja-JP" invalStChars=")]}" invalEndChars="([{"/>'), 'kinsoku wrong')
	assert.ok(pres.includes('<p:prnPr prnWhat="handouts4" clrMode="gray" frameSlides="1"/>'), 'prnPr wrong')
	assert.ok(pres.includes('<p:clrMru><a:srgbClr val="FF0000"/><a:schemeClr val="accent1"/></p:clrMru>'), 'clrMru wrong')
	// CT_Presentation fixes the child order: photoAlbum/kinsoku before defaultTextStyle, prnPr/clrMru after
	const order = ['<p:notesSz', '<p:photoAlbum', '<p:kinsoku', '<p:defaultTextStyle>', '<p:prnPr', '<p:clrMru>'].map(tag => pres.indexOf(tag))
	assert.deepEqual(order, [...order].sort((a, b) => a - b), `CT_Presentation child order violated: ${order.join(',')}`)

	const view = await readPart(zip, 'ppt/viewProps.xml')
	assert.ok(view.includes('lastView="sldThumbnailView"'), 'lastView missing')
	assert.ok(view.includes('showComments="0"'), 'showComments missing')
	assert.ok(view.includes('<p:cSldViewPr snapToGrid="1" snapToObjects="0">'), `snap attrs wrong: ${view}`)
	assert.ok(view.includes('<a:sx n="75" d="100"/>'), 'zoom missing')
	assert.ok(view.includes('<p:gridSpacing cx="457200" cy="457200"/>'), 'gridSpacing missing')

	// a partial kinsoku is dropped: invalStChars/invalEndChars are required on CT_Kinsoku
	const partial = new pptxgen()
	partial.kinsoku = { lang: 'ja-JP' }
	partial.addSlide()
	assert.ok(!(await readPart(await writeZip(partial), 'ppt/presentation.xml')).includes('<p:kinsoku'), 'incomplete kinsoku was emitted')

	// an untouched presentation keeps the previous output: none of these elements appear
	const bare = new pptxgen()
	bare.addSlide()
	const bareZip = await writeZip(bare)
	const barePres = await readPart(bareZip, 'ppt/presentation.xml')
	for (const tag of ['<p:photoAlbum', '<p:kinsoku', '<p:prnPr', '<p:clrMru', 'type="']) assert.ok(!barePres.includes(tag), `default presentation.xml gained ${tag}`)
	const bareView = await readPart(bareZip, 'ppt/viewProps.xml')
	assert.ok(bareView.includes('<a:sx n="136" d="100"/>') && bareView.includes('<p:gridSpacing cx="76200" cy="76200"/>'), 'default viewProps.xml changed')
	assert.ok(!bareView.includes('lastView=') && !bareView.includes('showComments='), 'default viewProps.xml gained attributes')
})

test('#147: table cells emit diagonal borders, 3-D cells, overflow and rtl column order', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addTable(
		[[
			{ text: 'diag', options: { borderDiagonalDown: { color: 'FF0000', width: 2 }, borderDiagonalUp: { type: 'dash' }, cell3D: { material: 'clear' }, fill: { color: 'EEEEEE' } } },
			{ text: '3d', options: { cell3D: { bevel: { preset: 'circle', width: 0.05, height: 0.05 }, material: 'metal', lightRig: { rig: 'threePt', dir: 't' } } } },
			{ text: 'ovf', options: { horzOverflow: 'overflow', anchorCtr: true, textDirection: 'vert270' } },
			{ text: 'mat', options: { cell3D: { material: 'matte' } } },
		]],
		{ x: 0.5, y: 0.5, w: 9, rtl: true }
	)
	const xml = await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')
	const cells = xml.match(/<a:tcPr[\s\S]*?<\/a:tcPr>/g) ?? []
	assert.equal(cells.length, 4, `expected four cells, got ${cells.length}`)

	assert.ok(xml.includes('<a:tblPr rtl="1"'), 'a:tblPr@rtl missing - CT_TableProperties owns rtl, not a:tbl')
	assert.ok(cells[0].includes('<a:lnTlToBr w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>'), `lnTlToBr wrong: ${cells[0]}`)
	assert.ok(cells[0].includes('<a:lnBlToTr') && cells[0].includes('<a:prstDash val="sysDash"/>'), 'lnBlToTr missing or not dashed')
	// CT_TableCellProperties fixes the child sequence, so the first cell carries every one of them
	const seq = ['<a:lnL', '<a:lnR', '<a:lnT', '<a:lnB', '<a:lnTlToBr', '<a:lnBlToTr', '<a:cell3D', '<a:solidFill><a:srgbClr val="EEEEEE"/>'].map(tag => cells[0].indexOf(tag))
	assert.ok(seq.every(idx => idx > -1), `a tcPr child is missing: ${seq.join(',')}`)
	assert.deepEqual(seq, [...seq].sort((a, b) => a - b), `CT_TableCellProperties child order violated: ${seq.join(',')}`)

	assert.ok(cells[1].includes('<a:cell3D prstMaterial="metal"><a:bevel w="45720" h="45720" prst="circle"/><a:lightRig rig="threePt" dir="t"/></a:cell3D>'), `cell3D wrong: ${cells[1]}`)
	assert.ok(cells[1].indexOf('<a:cell3D') > cells[1].indexOf('<a:lnB'), 'cell3D must follow the border elements')
	// `a:bevel` is required by CT_Cell3D, so it is written even when only the material is given
	assert.ok(cells[3].includes('<a:cell3D prstMaterial="matte"><a:bevel/></a:cell3D>'), `material-only cell3D wrong: ${cells[3]}`)

	assert.ok(cells[2].includes('vert="vert270"') && cells[2].includes('anchorCtr="1"') && cells[2].includes('horzOverflow="overflow"'), `cell attrs wrong: ${cells[2]}`)

	// `rig` and `dir` are both required on CT_LightRig, so a partial rig is dropped
	const partial = new pptxgen()
	partial.addSlide().addTable([[{ text: 'y', options: { cell3D: { lightRig: { rig: 'threePt' } as never } } }]], { x: 1, y: 1, w: 4 })
	const partialXml = await readPart(await writeZip(partial), 'ppt/slides/slide1.xml')
	assert.ok(partialXml.includes('<a:cell3D><a:bevel/></a:cell3D>'), 'incomplete lightRig was emitted')

	// a table that asks for none of it is unchanged: no new element, no new attribute
	const bare = new pptxgen()
	bare.addSlide().addTable([['a', 'b']], { x: 1, y: 1, w: 4 })
	const bareXml = await readPart(await writeZip(bare), 'ppt/slides/slide1.xml')
	for (const tag of ['<a:lnTlToBr', '<a:lnBlToTr', '<a:cell3D', 'horzOverflow=', 'anchorCtr=', 'rtl=']) {
		assert.ok(!bareXml.includes(tag), `default table gained ${tag}`)
	}
})

test('#138: blur, fillOverlay, prstShdw, effectDag, blip alpha effects and group fill', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	// one shape carrying every CT_EffectList child, so the fixed sequence is actually exercised
	slide.addShape('rect', {
		x: 0.5, y: 0.5, w: 2, h: 1, fill: { color: 'CCCCCC' },
		blur: { radius: 4, grow: false },
		fillOverlay: { blend: 'mult', fill: { color: 'FF0000', transparency: 50 } },
		glow: { size: 5, color: '00FF00', opacity: 0.4 },
		shadow: { type: 'outer', color: '000000' },
		reflection: { blur: 2 },
		softEdge: { radius: 3 },
	})
	slide.addShape('rect', { x: 3, y: 0.5, w: 2, h: 1, shadow: { type: 'preset', preset: 'shdw7', color: '333333' } })
	slide.addShape('rect', { x: 5.5, y: 0.5, w: 2, h: 1, glow: { size: 4, color: 'FF0000', opacity: 0.5 }, effectDag: { type: 'tree' } })
	slide.addShape('rect', { x: 0.5, y: 2, w: 2, h: 1, fill: { type: 'group' } })
	// a preset shadow with no preset name, and a fill overlay with no fill: both dropped
	slide.addShape('rect', { x: 3, y: 2, w: 2, h: 1, shadow: { type: 'preset', color: '333333' } })
	slide.addShape('rect', { x: 5.5, y: 2, w: 2, h: 1, fillOverlay: { blend: 'over' } as never })
	// `@blend` is required, so a fill overlay with a fill but no blend mode is dropped too
	slide.addShape('rect', { x: 8, y: 2, w: 1, h: 1, fillOverlay: { fill: { color: '00FF00' } } as never })
	slide.addImage({ data: PNG_4x2, x: 0.5, y: 3.5, w: 1, h: 0.5, transparency: 20, alphaEffects: { replace: 60, invert: true, floor: true, ceiling: true } })
	// ST_PositiveFixedPercentage caps at 100%
	slide.addImage({ data: PNG_4x2, x: 2, y: 3.5, w: 1, h: 0.5, alphaEffects: { replace: 150 } })

	const xml = await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')
	const lists = xml.match(/<a:effectLst>[\s\S]*?<\/a:effectLst>/g) ?? []
	assert.ok(lists.length >= 2, `expected effect lists, got ${lists.length}`)

	// CT_EffectList fixes the sequence: blur, fillOverlay, glow, innerShdw, outerShdw, prstShdw, reflection, softEdge
	const all = lists[0]
	assert.ok(all.includes('<a:blur rad="50800" grow="0"/>'), `blur wrong: ${all}`)
	assert.ok(all.includes('<a:fillOverlay blend="mult"><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:fillOverlay>'), `fillOverlay wrong: ${all}`)
	const seq = ['<a:blur', '<a:fillOverlay', '<a:glow', '<a:outerShdw', '<a:reflection', '<a:softEdge'].map(tag => all.indexOf(tag))
	assert.ok(seq.every(idx => idx > -1), `an effect is missing: ${seq.join(',')}`)
	assert.deepEqual(seq, [...seq].sort((a, b) => a - b), `CT_EffectList child order violated: ${seq.join(',')}`)

	// prstShdw sits in the shadow slot, after outerShdw's position in the sequence
	assert.ok(xml.includes('<a:prstShdw prst="shdw7" dist="50800" dir="16200000"><a:srgbClr val="333333"><a:alpha val="75000"/></a:srgbClr></a:prstShdw>'), 'prstShdw missing')
	// a preset shadow with no preset name is dropped: `@prst` is required on CT_PresetShadowEffect
	assert.equal((xml.match(/<a:prstShdw/g) ?? []).length, 1, 'a preset shadow without a preset name was emitted')
	// `@blend` and a fill are both required on CT_FillOverlayEffect
	assert.equal((xml.match(/<a:fillOverlay/g) ?? []).length, 1, 'a fill overlay missing its blend mode or its fill was emitted')
	assert.ok(!xml.includes('blend="undefined"'), 'a fill overlay was emitted without a blend mode')

	// effectLst and effectDag are alternatives in EG_EffectProperties, never both
	const dag = /<a:effectDag[\s\S]*?<\/a:effectDag>/.exec(xml)?.[0] ?? ''
	assert.ok(dag.startsWith('<a:effectDag type="tree">') && dag.includes('<a:glow'), `effectDag wrong: ${dag}`)
	const dagShape = /<p:sp>(?:(?!<\/p:sp>)[\s\S])*<a:effectDag[\s\S]*?<\/p:sp>/.exec(xml)?.[0] ?? ''
	assert.ok(dagShape && !dagShape.includes('<a:effectLst>'), 'a shape emitted both an effectDag and an effectLst')

	assert.ok(xml.includes('<a:grpFill/>'), 'group fill missing')

	const blips = xml.match(/<a:blip [\s\S]*?<\/a:blip>/g) ?? []
	assert.ok(blips[0].includes('<a:alphaModFix amt="80000"/><a:alphaRepl a="60000"/><a:alphaInv/><a:alphaFloor/><a:alphaCeiling/>'), `blip alpha effects wrong: ${blips[0]}`)
	assert.ok(blips[1].includes('<a:alphaRepl a="100000"/>'), `alphaRepl not clamped: ${blips[1]}`)
	// the alpha effects belong on the blip, never in the shape effect list
	for (const list of lists) assert.ok(!/alphaRepl|alphaInv|alphaFloor|alphaCeiling/.test(list), 'an alpha effect leaked into a:effectLst')

	// image effects declared on ImageProps were dropped by the addImage options rebuild
	const imgOnly = new pptxgen()
	imgOnly.addSlide().addImage({ data: PNG_4x2, x: 1, y: 1, w: 1, h: 0.5, glow: { size: 5, color: 'FF0000', opacity: 0.5 }, softEdge: { radius: 3 }, reflection: { blur: 2 } })
	const picXml = await readPart(await writeZip(imgOnly), 'ppt/slides/slide1.xml')
	const pic = /<p:pic>[\s\S]*?<\/p:pic>/.exec(picXml)?.[0] ?? ''
	assert.ok(pic.includes('<a:glow') && pic.includes('<a:softEdge') && pic.includes('<a:reflection'), `image effects never reached the pic: ${pic}`)

	// a shape asking for none of it is unchanged
	const bare = new pptxgen()
	bare.addSlide().addShape('rect', { x: 1, y: 1, w: 2, h: 1, fill: { color: 'CCCCCC' } })
	const bareXml = await readPart(await writeZip(bare), 'ppt/slides/slide1.xml')
	for (const tag of ['<a:blur', '<a:fillOverlay', '<a:prstShdw', '<a:effectDag', '<a:grpFill', '<a:alphaRepl']) {
		assert.ok(!bareXml.includes(tag), `default shape gained ${tag}`)
	}
})

test('#141: shapes and pictures emit p:style theme references', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addShape('rect', { x: 0.5, y: 0.5, w: 2, h: 1, styleRef: { line: 1, fill: 3, effect: 2, font: 'minor' } })
	slide.addShape('rect', { x: 3, y: 0.5, w: 2, h: 1, styleRef: { fill: 1 }, fill: { color: 'FF0000' } })
	slide.addShape('rect', { x: 5.5, y: 0.5, w: 2, h: 1, styleRef: { effect: 1, color: 'phClr' } })
	slide.addImage({ data: PNG_4x2, x: 0.5, y: 2, w: 1, h: 0.5, styleRef: { line: 2, color: 'accent3' } })

	const xml = await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')
	const shapes = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? []
	assert.equal(shapes.length, 3, `expected three shapes, got ${shapes.length}`)

	// CT_ShapeStyle requires all four children, in this order, so one set property emits all four.
	// Indices are 1-based into the theme's `a:fmtScheme` lists; 0 references nothing.
	assert.ok(shapes[0].includes(
		'<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>' +
		'<a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>' +
		'<a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef>' +
		'<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>'
	), `p:style wrong: ${shapes[0]}`)

	// `p:style` follows `p:spPr` and precedes `p:txBody` in the CT_Shape sequence
	const order = ['</p:spPr>', '<p:style>', '</p:style>', '<p:txBody>'].map(tag => shapes[0].indexOf(tag))
	assert.ok(order.every(idx => idx > -1), `a p:sp child is missing: ${order.join(',')}`)
	assert.deepEqual(order, [...order].sort((a, b) => a - b), `CT_Shape child order violated: ${order.join(',')}`)

	// This is what makes a theme swap restyle the shape: a referenced fill and no explicit fill means
	// NO fill element at all - `<a:noFill/>` would override the reference
	const spPr0 = /<p:spPr>[\s\S]*?<\/p:spPr>/.exec(shapes[0])?.[0] ?? ''
	assert.ok(!spPr0.includes('<a:noFill/>') && !spPr0.includes('<a:solidFill>'), `a referenced fill must be left to the theme: ${spPr0}`)

	// an explicit fill still wins, which is OOXML's own precedence
	const spPr1 = /<p:spPr>[\s\S]*?<\/p:spPr>/.exec(shapes[1])?.[0] ?? ''
	assert.ok(spPr1.includes('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>'), 'an explicit fill was dropped')
	// unset references are explicit no-references, never a guessed index
	assert.ok(shapes[1].includes('<a:lnRef idx="0">') && shapes[1].includes('<a:effectRef idx="0">') && shapes[1].includes('<a:fontRef idx="none">'), `unset refs wrong: ${shapes[1]}`)

	// `phClr` is the substitution target a reference resolves, not a colour it can carry
	assert.ok(shapes[2].includes('<a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef>'), 'phClr was not rejected')
	assert.ok(!xml.includes('val="phClr"'), 'phClr reached the output')
	// without a referenced fill the `a:noFill` default still applies
	assert.ok((/<p:spPr>[\s\S]*?<\/p:spPr>/.exec(shapes[2])?.[0] ?? '').includes('<a:noFill/>'), 'the noFill default was suppressed without a fill reference')

	// `p:style` follows `p:spPr` in the CT_Picture sequence too
	const pic = /<p:pic>[\s\S]*?<\/p:pic>/.exec(xml)?.[0] ?? ''
	assert.ok(pic.includes('<a:lnRef idx="2"><a:schemeClr val="accent3"/></a:lnRef>'), `picture p:style wrong: ${pic}`)
	const picOrder = ['</p:spPr>', '<p:style>', '</p:pic>'].map(tag => (pic + '</p:pic>').indexOf(tag))
	assert.deepEqual(picOrder, [...picOrder].sort((a, b) => a - b), `CT_Picture child order violated: ${picOrder.join(',')}`)

	// a shape asking for no references is unchanged, `a:noFill` default included
	const bare = new pptxgen()
	bare.addSlide().addShape('rect', { x: 1, y: 1, w: 2, h: 1 })
	const bareXml = await readPart(await writeZip(bare), 'ppt/slides/slide1.xml')
	assert.ok(!bareXml.includes('<p:style>'), 'default shape gained a p:style')
	assert.ok(bareXml.includes('<a:noFill/>'), 'default shape lost its noFill')
})

test('#149: slide layout and placeholder metadata', async () => {
	const pptx = new pptxgen()
	// one layout carrying every piece of metadata, so the CT_SlideLayout order is exercised
	pptx.defineSlideMaster({
		title: 'SECTION',
		layoutType: 'secHead',
		matchingName: 'Section Header',
		showMasterShapes: false,
		showMasterPlaceholderAnimation: false,
		userDrawn: true,
		colorMapOverride: { bg1: 'dk1', tx1: 'lt1' },
		transition: { type: 'fade', duration: 500 },
		objects: [{ placeholder: { options: { name: 'ttl', type: 'title', x: 1, y: 1, w: 6, h: 1, orient: 'vert', sz: 'half', userDrawn: true }, text: 'Section' } }],
	})
	pptx.defineSlideMaster({ title: 'PLAIN', preserve: false, objects: [{ placeholder: { options: { name: 'b', type: 'body', x: 1, y: 1, w: 6, h: 1 } } }] })
	pptx.addSlide({ masterName: 'SECTION' })

	const zip = await writeZip(pptx)
	const section = (zip.file(/slideLayout\d+\.xml/) ?? []).length
	assert.ok(section >= 2, `expected layout parts, got ${section}`)

	// find the layout carrying our metadata rather than assuming a part number
	let layout = ''
	for (const file of zip.file(/ppt\/slideLayouts\/slideLayout\d+\.xml/) ?? []) {
		const xml = await file.async('string')
		if (xml.includes('type="secHead"')) layout = xml
	}
	assert.ok(layout, 'no layout carried the metadata')

	assert.ok(layout.includes(' preserve="1" type="secHead" matchingName="Section Header" showMasterSp="0" showMasterPhAnim="0" userDrawn="1">'), `p:sldLayout attrs wrong: ${/<p:sldLayout[^>]*>/.exec(layout)?.[0] ?? ''}`)
	// CT_ColorMapping requires all twelve attributes: the two given are used, the rest come from the identity map
	assert.ok(layout.includes('<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'), `clrMapOvr wrong: ${/<p:clrMapOvr>[\s\S]*?<\/p:clrMapOvr>/.exec(layout)?.[0] ?? ''}`)

	// CT_SlideLayout sequence: cSld, clrMapOvr, transition
	const seq = ['<p:cSld', '</p:cSld>', '<p:clrMapOvr>', 'mc:AlternateContent'].map(tag => layout.indexOf(tag))
	assert.ok(seq.every(idx => idx > -1), `a p:sldLayout child is missing: ${seq.join(',')}`)
	assert.deepEqual(seq, [...seq].sort((a, b) => a - b), `CT_SlideLayout child order violated: ${seq.join(',')}`)

	// `p:ph` attributes and `p:nvPr@userDrawn`
	assert.ok(/<p:ph\s+idx="\d+"\s+type="title"\s+orient="vert"\s+sz="half"\s+hasCustomPrompt="1"/.test(layout.replace(/\s+/g, ' ')), `p:ph wrong: ${/<p:ph[\s\S]*?\/>/.exec(layout)?.[0] ?? ''}`)
	assert.ok(layout.includes('<p:nvPr userDrawn="1">'), 'p:nvPr@userDrawn missing')

	// `preserve: false` is the only way to turn off an attribute that has always been written
	let plain = ''
	for (const file of zip.file(/ppt\/slideLayouts\/slideLayout\d+\.xml/) ?? []) {
		const xml = await file.async('string')
		if (xml.includes('name="PLAIN"')) plain = xml
	}
	assert.ok(plain, 'no PLAIN layout')
	assert.ok(!/<p:sldLayout[^>]*preserve=/.test(plain), 'preserve:false was ignored')

	// a layout that asks for none of it is unchanged, including the inherited colour mapping
	const bare = new pptxgen()
	bare.defineSlideMaster({ title: 'BARE', objects: [{ placeholder: { options: { name: 'b', type: 'body', x: 1, y: 1, w: 6, h: 1 } } }] })
	bare.addSlide({ masterName: 'BARE' })
	const bareZip = await writeZip(bare)
	for (const file of bareZip.file(/ppt\/slideLayouts\/slideLayout\d+\.xml/) ?? []) {
		const xml = await file.async('string')
		const tag = /<p:sldLayout[^>]*>/.exec(xml)?.[0] ?? ''
		assert.ok(tag.includes('preserve="1"'), `${file.name} lost preserve="1"`)
		for (const attr of ['type=', 'matchingName=', 'showMasterSp=', 'showMasterPhAnim=', 'userDrawn=']) {
			assert.ok(!tag.includes(attr), `${file.name} gained ${attr}`)
		}
		assert.ok(xml.includes('<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'), `${file.name} lost the inherited colour mapping`)
		assert.ok(!xml.includes('orient=') && !xml.includes(' sz="half"') && !xml.includes('userDrawn='), `${file.name} gained placeholder metadata`)
	}
})

test('#150: media source elements - linked media, audioCd, wavAudioFile', async () => {
	const MP4 = 'video/mp4;base64,AAAAIGZ0eXBpc29tAAACAGlzb21pc28yYXZjMW1wNDE='
	const WAV = 'audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQAAAAA='
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addMedia({ type: 'video', data: MP4, x: 0.5, y: 0.5, w: 2, h: 1.5, isPhoto: true, userDrawn: true })
	slide.addMedia({ type: 'video', link: 'C:/movies/clip.mp4', x: 3, y: 0.5, w: 2, h: 1.5, contentType: 'video/mp4' })
	slide.addMedia({ type: 'audio', link: '/srv/audio/theme.mp3', x: 5.5, y: 0.5, w: 2, h: 1.5 })
	slide.addMedia({ type: 'audioCd', audioCd: { start: { track: 1 }, end: { track: 1, time: 30 } }, x: 0.5, y: 2.5, w: 2, h: 1.5 })
	slide.addMedia({ type: 'wav', data: WAV, x: 3, y: 2.5, w: 2, h: 1.5 })

	const zip = await writeZip(pptx)
	const xml = await readPart(zip, 'ppt/slides/slide1.xml')
	const rels = await readPart(zip, 'ppt/slides/_rels/slide1.xml.rels')

	// EG_Media is a choice: every media frame carries exactly one media element
	const frames = (xml.match(/<p:nvPr[^>]*>[\s\S]*?<\/p:nvPr>/g) ?? []).filter(frame => /a:videoFile|a:audioFile|a:audioCd|a:wavAudioFile/.test(frame))
	assert.equal(frames.length, 5, `expected five media frames, got ${frames.length}`)
	for (const frame of frames) {
		const kinds = (frame.match(/<a:(videoFile|audioFile|audioCd|wavAudioFile)\b/g) ?? []).length
		assert.equal(kinds, 1, `EG_Media is a choice, but a frame carried ${kinds} media elements: ${frame}`)
	}

	// `p:nvPr` attributes default to false, so only the "on" case is written
	assert.ok(xml.includes('<p:nvPr isPhoto="1" userDrawn="1">'), 'isPhoto/userDrawn missing')
	assert.equal((xml.match(/<p:nvPr isPhoto=/g) ?? []).length, 1, 'isPhoto leaked onto other frames')

	// linked media: same three-relationship shape, but external and with no part in the package
	assert.ok(xml.includes('<a:videoFile r:link="rId4" contentType="video/mp4"/>'), `linked video wrong: ${frames[1]}`)
	assert.ok(xml.includes('<a:audioFile r:link="rId7"/>'), `linked audio wrong: ${frames[2]}`)
	const relElements = rels.match(/<Relationship\b[^>]*\/>/g) ?? []
	for (const target of ['C:/movies/clip.mp4', '/srv/audio/theme.mp3']) {
		const matches = relElements.filter(rel => rel.includes(`Target="${target}"`))
		assert.equal(matches.length, 2, `expected a video/audio and a media relationship for ${target}, got ${matches.length}`)
		for (const rel of matches) assert.ok(rel.includes('TargetMode="External"'), `linked media relationship is not external: ${rel}`)
	}
	// nothing was written for the linked files, and no content type was declared for one
	assert.ok(!Object.keys(zip.files).some(name => name.includes('clip.mp4') || name.includes('theme.mp3')), 'a part was written for linked media')
	assert.ok(!(await readPart(zip, '[Content_Types].xml')).includes('Extension="mp3"'), 'a Default Extension was declared for a part that is never written')

	// CT_AudioCD: `a:st`/`a:end` are required, `@track` is required, `@time` defaults to 0
	assert.ok(xml.includes('<a:audioCd><a:st track="1"/><a:end track="1" time="30"/></a:audioCd>'), `audioCd wrong: ${frames[3]}`)
	// CD audio references the drive, so it has no media relationship and no `p14:media`
	assert.ok(!frames[3].includes('p14:media'), 'audioCd emitted a p14:media extension')

	// `a:wavAudioFile` embeds via `r:embed` against an audio relationship, and has no `p14:media`
	assert.ok(/<a:wavAudioFile r:embed="rId\d+" name="[^"]*"\/>/.test(frames[4]), `wavAudioFile wrong: ${frames[4]}`)
	assert.ok(!frames[4].includes('p14:media'), 'wavAudioFile emitted a p14:media extension')
	assert.ok(Object.keys(zip.files).some(name => name.endsWith('.wav')), 'the embedded WAV part is missing')
	assert.ok((await readPart(zip, '[Content_Types].xml')).includes('Extension="wav" ContentType="audio/wav"'), 'the WAV content type is missing')

	// audioCd requires both track numbers - addMedia throws rather than guessing
	assert.throws(() => pptx.addSlide().addMedia({ type: 'audioCd', x: 1, y: 1, w: 1, h: 1 }), /audioCd\.start\.track/, 'audioCd without tracks did not throw')

	// embedded media is unchanged: three relationships, an internal target, and a p14:media
	const bare = new pptxgen()
	bare.addSlide().addMedia({ type: 'video', data: MP4, x: 1, y: 1, w: 2, h: 1.5 })
	const bareZip = await writeZip(bare)
	const bareXml = await readPart(bareZip, 'ppt/slides/slide1.xml')
	assert.ok(bareXml.includes('<p:nvPr>'), 'a default media frame gained an attribute')
	assert.ok(bareXml.includes('<a:videoFile r:link="rId1"/>'), `default video element changed: ${bareXml.match(/<a:videoFile[^>]*>/)?.[0] ?? ''}`)
	assert.ok(bareXml.includes('p14:media') && bareXml.includes('r:embed="rId2"'), 'the embedded media extension changed')
	assert.ok(!(await readPart(bareZip, 'ppt/slides/_rels/slide1.xml.rels')).includes('TargetMode="External"'), 'embedded media became external')
})

test('#157: charts get a style part and a colour-style part', async () => {
	const data = [{ name: 'S1', labels: ['a', 'b'], values: [1, 2] }]
	const pptx = new pptxgen()
	pptx.addSlide().addChart('bar', data, { x: 1, y: 1, w: 6, h: 4 })
	pptx.addSlide().addChart('pie', data, { x: 1, y: 1, w: 6, h: 4, chartStyle: 251, chartColorStyle: { method: 'withinLinear', id: 13, colors: [{ scheme: 'accent3' }, 'FF0000'] } })

	const zip = await writeZip(pptx)
	const chartParts = Object.keys(zip.files).filter(name => /^ppt\/charts\/(chart|colors|style)\d+\.xml$/.test(name)).sort()
	// exactly one style and one colour-style part per chart, numbered to match it. Chart part numbering
	// is process-global (DEPRECATION-PLAN.md F9), so the numbers are read off the package.
	const chartNums = chartParts.filter(name => /\/chart\d+\.xml$/.test(name)).map(name => /(\d+)\.xml$/.exec(name)?.[1] ?? '')
	assert.equal(chartNums.length, 2, `expected two charts, got ${chartNums.join(',')}`)
	for (const num of chartNums) {
		assert.ok(chartParts.includes(`ppt/charts/colors${num}.xml`), `missing colors${num}.xml`)
		assert.ok(chartParts.includes(`ppt/charts/style${num}.xml`), `missing style${num}.xml`)
	}

	// the parts are discovered by relationship type, so the exact URIs matter. The namespace is dated
	// 2012 and the relationships 2011 - that asymmetry is correct, and this pins it.
	const rels = await readPart(zip, `ppt/charts/_rels/chart${chartNums[0]}.xml.rels`)
	assert.ok(rels.includes(`<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors${chartNums[0]}.xml"/>`), `colour-style relationship wrong: ${rels}`)
	assert.ok(rels.includes(`<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style${chartNums[0]}.xml"/>`), `style relationship wrong: ${rels}`)
	assert.ok(rels.includes('rId1'), 'the embedded workbook relationship was displaced')

	// a part with no declared content type is a repair-dialog cause
	const contentTypes = await readPart(zip, '[Content_Types].xml')
	for (const num of chartNums) {
		assert.ok(contentTypes.includes(`<Override PartName="/ppt/charts/colors${num}.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>`), `colors${num}.xml content type missing`)
		assert.ok(contentTypes.includes(`<Override PartName="/ppt/charts/style${num}.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>`), `style${num}.xml content type missing`)
	}

	const colors = await readPart(zip, `ppt/charts/colors${chartNums[0]}.xml`)
	assert.ok(colors.includes('xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"'), 'the cs namespace is wrong')
	assert.ok(colors.includes('meth="cycle" id="10">'), `default colour style wrong: ${colors.slice(0, 300)}`)
	assert.ok(colors.includes('<a:schemeClr val="accent1"/><a:schemeClr val="accent2"/>'), 'the default palette is not the theme accents')
	assert.equal((colors.match(/<cs:variation/g) ?? []).length, 9, 'expected nine luminance variations')

	const style = await readPart(zip, `ppt/charts/style${chartNums[0]}.xml`)
	assert.ok(style.includes('<cs:chartStyle ') && style.includes(' id="201">'), `default style id wrong: ${style.slice(0, 260)}`)
	assert.ok(style.includes('<cs:axisTitle>') && style.includes('<cs:dataPoint>') && style.includes('<cs:plotArea'), 'the style definitions are incomplete')

	// the second chart's overrides
	const colors2 = await readPart(zip, `ppt/charts/colors${chartNums[1]}.xml`)
	assert.ok(colors2.includes('meth="withinLinear" id="13">'), `colour style override ignored: ${colors2.slice(0, 260)}`)
	assert.ok(colors2.includes('<a:schemeClr val="accent3"/><a:srgbClr val="FF0000"/>'), 'a custom palette was ignored')
	assert.ok((await readPart(zip, `ppt/charts/style${chartNums[1]}.xml`)).includes(' id="251">'), 'the style id override was ignored')

	// nothing points at these parts from inside the chart, and the chart itself is unchanged
	const chartXml = await readPart(zip, `ppt/charts/chart${chartNums[0]}.xml`)
	assert.ok(!chartXml.includes('chartStyle') && !chartXml.includes('colorStyle'), 'chartN.xml should not reference the style parts')

	// a deck with no chart gains nothing
	const bare = new pptxgen()
	bare.addSlide().addText('hi', { x: 1, y: 1 })
	const bareZip = await writeZip(bare)
	assert.ok(!Object.keys(bareZip.files).some(name => /colors\d+\.xml|style\d+\.xml/.test(name)), 'a chartless deck gained chart style parts')
	assert.ok(!(await readPart(bareZip, '[Content_Types].xml')).includes('chartstyle'), 'a chartless deck gained a chart style content type')
})

test('#133: custom table style definitions reach ppt/tableStyles.xml', async () => {
	const STYLE_ID = '{A1B2C3D4-1111-2222-3333-444455556666}'
	const pptx = new pptxgen()
	pptx.tableStyles = [
		{
			id: STYLE_ID,
			name: 'NEOMA Blue',
			// one style carrying every part, so the CT_TableStyle child order is actually exercised
			wholeTable: { color: { scheme: 'tx1' }, borders: { top: { color: '4472C4', width: 1 }, insideH: { color: 'D9D9D9', width: 0.5 } } },
			band1H: { fill: { color: 'DEEAF6' } },
			band2H: { fill: { color: 'FFFFFF' } },
			band1V: { fill: { color: 'EEEEEE' } },
			band2V: { fill: { color: 'DDDDDD' } },
			lastCol: { bold: true },
			firstCol: { bold: true },
			lastRow: { bold: true, italic: false },
			seCell: { fill: { color: '111111' } },
			swCell: { fill: { color: '222222' } },
			firstRow: { bold: true, color: 'FFFFFF', fill: { color: '4472C4' } },
			neCell: { fill: { color: '333333' } },
			nwCell: { fill: { color: '444444' } },
		},
		{ id: 'not-a-guid', name: 'Bad' },
		{ id: '{A1B2C3D4-1111-2222-3333-444455556667}', name: '' },
	]
	pptx.addSlide().addTable([['H1', 'H2'], ['a', 'b']], { x: 1, y: 1, w: 6, tableStyleId: STYLE_ID, firstRow: true, bandRow: true })

	const xml = await readPart(await writeZip(pptx), 'ppt/tableStyles.xml')
	assert.ok(xml.includes(`<a:tblStyle styleId="${STYLE_ID}" styleName="NEOMA Blue">`), `tblStyle wrong: ${xml.slice(0, 300)}`)
	// both attributes are required by CT_TableStyle, so a bad id or a missing name is dropped
	assert.equal((xml.match(/<a:tblStyle /g) ?? []).length, 1, 'an invalid table style was emitted')
	assert.ok(!xml.includes('not-a-guid'), 'a malformed style id reached the output')

	// CT_TableStyle fixes this order, and it is neither alphabetical nor intuitive: lastCol before
	// firstCol, and firstRow between swCell and neCell
	const expected = ['wholeTbl', 'band1H', 'band2H', 'band1V', 'band2V', 'lastCol', 'firstCol', 'lastRow', 'seCell', 'swCell', 'firstRow', 'neCell', 'nwCell']
	const emitted = (xml.match(/<a:(wholeTbl|band1H|band2H|band1V|band2V|lastCol|firstCol|lastRow|seCell|swCell|firstRow|neCell|nwCell)>/g) ?? []).map(tag => tag.slice(3, -1))
	assert.deepEqual(emitted, expected, 'CT_TableStyle child order violated')

	// `a:tcTxStyle` precedes `a:tcStyle`, and inside the cell style `a:tcBdr` precedes the fill
	const firstRow = /<a:firstRow>[\s\S]*?<\/a:firstRow>/.exec(xml)?.[0] ?? ''
	assert.ok(firstRow.includes('<a:tcTxStyle b="on"><a:srgbClr val="FFFFFF"/></a:tcTxStyle>'), `firstRow text style wrong: ${firstRow}`)
	assert.ok(firstRow.indexOf('<a:tcTxStyle') < firstRow.indexOf('<a:tcStyle'), 'tcTxStyle must precede tcStyle')
	assert.ok(firstRow.includes('<a:tcStyle><a:fill><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:fill></a:tcStyle>'), `firstRow fill wrong: ${firstRow}`)

	const whole = /<a:wholeTbl>[\s\S]*?<\/a:wholeTbl>/.exec(xml)?.[0] ?? ''
	assert.ok(whole.indexOf('<a:tcBdr>') < whole.indexOf('</a:tcStyle>'), 'tcBdr must sit inside tcStyle')
	assert.ok(whole.includes('<a:top><a:ln w="12700"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:ln></a:top>'), `border wrong: ${whole}`)
	assert.ok(whole.includes('<a:insideH><a:ln w="6350">'), 'insideH border missing')

	// `b`/`i` are ST_OnOffStyleType: an unset property is left to the theme, `false` is written as "off"
	const lastRow = /<a:lastRow>[\s\S]*?<\/a:lastRow>/.exec(xml)?.[0] ?? ''
	assert.ok(lastRow.includes('b="on"') && lastRow.includes('i="off"'), `lastRow on/off wrong: ${lastRow}`)
	assert.ok(!(/<a:band1H>[\s\S]*?<\/a:band1H>/.exec(xml)?.[0] ?? '').includes('b='), 'an unset bold was written')

	// `@def` is what a table with no `tableStyleId` inherits, so a custom style must not repoint it
	assert.ok(xml.includes('def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"'), '@def was repointed at a custom style')

	// with no custom styles the part is the same self-closing stub as before
	const bare = new pptxgen()
	bare.addSlide().addTable([['a']], { x: 1, y: 1, w: 2 })
	const bareXml = await readPart(await writeZip(bare), 'ppt/tableStyles.xml')
	assert.ok(bareXml.trimEnd().endsWith('def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'), `default tableStyles.xml changed: ${bareXml}`)
	assert.ok(!bareXml.includes('<a:tblStyle '), 'a default deck gained a table style')
})
