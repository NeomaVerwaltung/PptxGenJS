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
	assert.ok(layout.includes('<a:videoFile'), 'media not rendered on the master layout')
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

test('#35: images accept a line/outline and emit it in the picture spPr', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addImage({ data: PNG_4x2, x: 1, y: 1, w: 2, h: 1, line: { color: 'FF0000', width: 2, dashType: 'dash' } })

	const xml = await readPart(await writeZip(pptx), 'ppt/slides/slide1.xml')
	const pic = /<p:pic>[\s\S]*?<\/p:pic>/.exec(xml)?.[0] ?? ''
	assert.ok(pic.includes('<a:ln w="25400">'), `picture outline width missing: ${pic}`)
	assert.ok(pic.includes('<a:srgbClr val="FF0000"/>'), 'picture outline color missing')
	assert.ok(pic.includes('<a:prstDash val="dash"/>'), 'picture outline dash type missing')
})
