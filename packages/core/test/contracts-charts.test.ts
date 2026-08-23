/**
 * Chart contracts, for both the ECMA-376 `c:` pipeline and the PowerPoint 2016+ `cx:` one.
 *
 * Unlike golden XML snapshots, these checks document the OOXML that matters and allow harmless
 * serializer changes without regenerating fixture files.
 */
import { test, before } from 'node:test'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import pptxgen from '../src/pptxgen'
import { assertEmbeddedXlsxContracts, assertPptxPackageContracts, readPart } from './pptx-contracts'
import { buildContractFixture, captureWarnings } from './fixtures'

let zip: JSZip

before(async () => {
	zip = await buildContractFixture()
})

test('contract: bar chart keeps its data and chart type', async () => {
	// chart part numbering is process-global, so resolve the part rather than assuming `chart1.xml`
	const chartPart = Object.keys(zip.files).find(name => /^ppt\/charts\/chart\d+\.xml$/.test(name)) ?? ''
	assert.ok(chartPart, 'the fixture presentation has no chart part')
	const xml = await readPart(zip, chartPart)
	assert.match(xml, /<c:barChart>/, 'bar chart missing')
	assert.match(xml, /<c:v>Sales<\/c:v>/, 'series name missing')
	assert.match(xml, /<c:v>Q1<\/c:v>/, 'category label missing')
	assert.match(xml, /<c:v>20<\/c:v>/, 'series value missing')
})

// ChartEx (PowerPoint 2016+) charts - MS-ODRAWXML 2.1
// The package contract is what makes or breaks these: a wrong content type, relationship type, or a
// missing sidecar part makes PowerPoint declare the file damaged rather than degrade gracefully.
/** Chart part ids are global to the process, so the part name is discovered rather than assumed */
async function buildChartEx (type: string, data: object[], opts: object = {}): Promise<{ zip: JSZip, part: string }> {
	const pptx = new pptxgen()
	pptx.addSlide().addChart(type as never, data as never, { x: 0.5, y: 0.5, w: 6, h: 4, ...opts })
	const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	const part = Object.keys(zip.files).find(name => /^ppt\/charts\/chartEx\d+\.xml$/.test(name)) ?? ''
	return { zip, part }
}

/** The chartex part XML for one chart, asserting the part exists under a chartEx name */
async function chartExXml (type: string, data: object[], opts: object = {}): Promise<string> {
	const { zip, part } = await buildChartEx(type, data, opts)
	assert.ok(part, `${type}: chartex parts are named chartExN.xml, not chartN.xml`)
	return await readPart(zip, part)
}

const WATERFALL_DATA = [{ name: 'Cash', labels: ['Start', 'Q1', 'Q2', 'End'], values: [100, 30, -20, 110] }]

test('contract: a chartex chart is a separate part with its own content and relationship types', async () => {
	const { zip: cxZip, part } = await buildChartEx('waterfall', WATERFALL_DATA)
	await assertPptxPackageContracts(cxZip)
	await assertEmbeddedXlsxContracts(cxZip)

	assert.ok(part, 'chartex parts are named chartExN.xml, not chartN.xml')
	assert.equal(Object.keys(cxZip.files).filter(name => /^ppt\/charts\/chart\d+\.xml$/.test(name)).length, 0, 'a chartex chart must not also write an ECMA-376 chart part')

	const contentTypes = await readPart(cxZip, '[Content_Types].xml')
	assert.match(contentTypes, new RegExp(`PartName="/${part}" ContentType="application/vnd\\.ms-office\\.chartex\\+xml"`), 'chartex content type missing')

	const rels = await readPart(cxZip, 'ppt/slides/_rels/slide1.xml.rels')
	assert.match(rels, new RegExp(`Type="http://schemas\\.microsoft\\.com/office/2014/relationships/chartEx" Target="/${part}"`), 'chartex relationship type missing')
	assert.doesNotMatch(rels, /relationships\/chart"/, 'a chartex part must not use the ECMA-376 chart relationship type')
})

test('contract: a chartex chart relates to the style and color parts PowerPoint requires', async () => {
	const { zip: cxZip, part } = await buildChartEx('treemap', [{ name: 'Revenue', labels: ['A', 'B', 'C'], values: [10, 20, 30] }])

	// a chartex layout cannot be laid out without them, and they are the same parts every chart relates to
	const chartRels = await readPart(cxZip, `ppt/charts/_rels/${part.split('/').pop() ?? ''}.rels`)
	const styleRel = /Type="http:\/\/schemas\.microsoft\.com\/office\/2011\/relationships\/chartStyle" Target="(style\d+\.xml)"/.exec(chartRels)
	const colorsRel = /Type="http:\/\/schemas\.microsoft\.com\/office\/2011\/relationships\/chartColorStyle" Target="(colors\d+\.xml)"/.exec(chartRels)
	assert.ok(styleRel, 'chart style relationship missing')
	assert.ok(colorsRel, 'chart color style relationship missing')
	// `cx:externalData` points at rId1, so the workbook must keep that id
	assert.match(chartRels, /Id="rId1" Type="[^"]*relationships\/package"/, 'the embedded workbook must stay rId1')

	const contentTypes = await readPart(cxZip, '[Content_Types].xml')
	assert.match(contentTypes, new RegExp(`/ppt/charts/${styleRel[1]}" ContentType="application/vnd\\.ms-office\\.chartstyle\\+xml"`), 'chart style content type missing')
	assert.match(contentTypes, new RegExp(`/ppt/charts/${colorsRel[1]}" ContentType="application/vnd\\.ms-office\\.chartcolorstyle\\+xml"`), 'chart color style content type missing')
	assert.match(await readPart(cxZip, `ppt/charts/${styleRel[1]}`), /^<\?xml[^>]*\?><cs:chartStyle /, 'chart style part is not a cs:chartStyle document')
	assert.match(await readPart(cxZip, `ppt/charts/${colorsRel[1]}`), /^<\?xml[^>]*\?><cs:colorStyle /, 'chart color style part is not a cs:colorStyle document')
})

test('contract: a chartex frame is offered through mc:AlternateContent with a fallback', async () => {
	const { zip: cxZip } = await buildChartEx('waterfall', WATERFALL_DATA)
	const slideXml = await readPart(cxZip, 'ppt/slides/slide1.xml')

	const block = slideXml.slice(slideXml.indexOf('<mc:AlternateContent'), slideXml.indexOf('</mc:AlternateContent>'))
	assert.ok(block, 'a chartex frame must be wrapped in mc:AlternateContent')
	// the Requires prefix has to be declared on the Choice, or a consumer cannot evaluate the condition
	assert.match(block, /<mc:Choice xmlns:cx1="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2015\/9\/8\/chartex" Requires="cx1">/, 'chartex Choice condition missing')
	assert.match(block, /<a:graphicData uri="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2014\/chartex">\s*<cx:chart [^>]*r:id="rId\d+"\/>/, 'chartex graphicData uri or chart reference missing')
	assert.match(block, /<mc:Fallback><p:sp>/, 'a chartex frame must offer a fallback shape')
	assert.match(block, /<a:spLocks noTextEdit="1"\/>/, 'the fallback must not be editable into something that no longer matches the chart')

	// funnel postdates the 2016 launch layouts and gates on a later schema generation
	const funnelXml = await readPart((await buildChartEx('funnel', [{ name: 'Pipeline', labels: ['Leads', 'Won'], values: [500, 30] }])).zip, 'ppt/slides/slide1.xml')
	assert.match(funnelXml, /xmlns:cx1="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2015\/10\/21\/chartex"/, 'funnel must require its own chartex generation')
})

test('contract: each chartex layout emits the markup PowerPoint keys off', async () => {
	const waterfall = await chartExXml('waterfall', WATERFALL_DATA, { chartExSubtotals: [0, 3] })
	assert.match(waterfall, /<cx:series layoutId="waterfall" uniqueId="\{[0-9A-F-]{36}\}">/, 'waterfall layoutId or uniqueId missing')
	assert.match(waterfall, /<cx:subtotals><cx:idx val="0"\/><cx:idx val="3"\/><\/cx:subtotals>/, 'waterfall subtotals missing')
	// the cell references must resolve against the worksheet workbook.ts writes: labels in A, series in B..
	assert.match(waterfall, /<cx:strDim type="cat"><cx:f>Sheet1!\$A\$2:\$A\$5<\/cx:f>/, 'category dimension does not point at the worksheet label column')
	assert.match(waterfall, /<cx:numDim type="val"><cx:f>Sheet1!\$B\$2:\$B\$5<\/cx:f>/, 'value dimension does not point at the worksheet series column')
	assert.match(waterfall, /<cx:axis id="0"><cx:catScaling gapWidth="0\.5"\/>.*<cx:axis id="1"><cx:valScaling\/>/s, 'waterfall axes missing')

	// a histogram is a clusteredColumn layout that bins raw values, so it carries no category dimension
	const histogram = await chartExXml('histogram', [{ name: 'Ages', labels: ['a', 'b', 'c'], values: [3, 7, 12] }], { chartExBinCount: 4 })
	assert.match(histogram, /<cx:series layoutId="clusteredColumn"/, 'histogram layoutId missing')
	assert.match(histogram, /<cx:binning intervalClosed="r"><cx:binCount val="4"\/><\/cx:binning>/, 'histogram binning missing')
	assert.doesNotMatch(histogram, /<cx:strDim/, 'a histogram bins values itself and must not declare categories')

	// hierarchical layouts size their segments through a `size` dimension and draw no axes
	const treemap = await chartExXml('treemap', [{ name: 'Revenue', labels: ['A', 'B'], values: [10, 20] }], { chartExParentLabels: 'banner' })
	assert.match(treemap, /<cx:numDim type="size">/, 'treemap must size segments through a size dimension')
	assert.match(treemap, /<cx:parentLabelLayout val="banner"\/>/, 'treemap parent label layout missing')
	assert.doesNotMatch(treemap, /<cx:axis /, 'a treemap has no axes')

	const sunburst = await chartExXml('sunburst', [{ name: 'Revenue', labels: ['A', 'B'], values: [10, 20] }])
	assert.match(sunburst, /<cx:series layoutId="sunburst"/, 'sunburst layoutId missing')

	// one cx:data per series, each series naming the one it plots
	const box = await chartExXml('boxWhisker', [
		{ name: 'A', labels: ['x', 'y'], values: [1, 5] },
		{ name: 'B', labels: ['x', 'y'], values: [2, 6] },
	], { chartExMeanLine: true })
	assert.match(box, /<cx:data id="0">.*<cx:data id="1">/s, 'box & whisker needs one data block per series')
	assert.match(box, /<cx:numDim type="val"><cx:f>Sheet1!\$C\$2:\$C\$3<\/cx:f>/, 'the second series must read the second worksheet column')
	assert.match(box, /<cx:dataId val="1"\/>/, 'the second series must name its own data block')
	assert.match(box, /<cx:visibility meanLine="1" meanMarker="1" nonoutliers="0" outliers="1"\/>/, 'box & whisker visibility missing')
	assert.match(box, /<cx:statistics quartileMethod="exclusive"\/>/, 'box & whisker quartile method missing')

	// a funnel is drawn along a single category axis, which PowerPoint numbers 1
	const funnel = await chartExXml('funnel', [{ name: 'Pipeline', labels: ['Leads', 'Won'], values: [500, 30] }])
	assert.match(funnel, /<cx:axis id="1"><cx:catScaling/, 'funnel category axis missing')
	assert.doesNotMatch(funnel, /<cx:axis id="0"/, 'a funnel has no value axis')
})

test('contract: chartex options that PowerPoint would reject are dropped with a warning', async () => {
	let xml = ''
	const warnings = await captureWarnings(async () => {
		// 9 is past the end of a 4-point series; a bin count must be a positive integer
		xml = await chartExXml('waterfall', WATERFALL_DATA, { chartExSubtotals: [0, 9], chartExBinCount: 0, chartExParentLabels: 'sideways' })
	})

	assert.equal(warnings.length, 3, `expected one warning per rejected option, got: ${warnings.join(' | ')}`)
	assert.match(xml, /<cx:subtotals><cx:idx val="0"\/><\/cx:subtotals>/, 'the valid subtotal index must survive')
	assert.doesNotMatch(xml, /val="9"/, 'an out-of-range subtotal index must not be emitted')
})

test('contract: a chartex type cannot be smuggled into a multi-type chart', () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	assert.throws(
		() => slide.addChart([
			{ type: 'bar' as never, data: [{ name: 'A', labels: ['x'], values: [1] }], options: {} },
			{ type: 'waterfall' as never, data: [{ name: 'B', labels: ['x'], values: [2] }], options: {} },
		], []),
		/waterfall.*multi-type/,
		'a chartex layout owns the whole plot area and cannot share one'
	)
})

test('contract: classic charts are untouched by the chartex path', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [10, 20] }], { x: 0.5, y: 0.5, w: 6, h: 4 })
	const barZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	const slideXml = await readPart(barZip, 'ppt/slides/slide1.xml')

	assert.doesNotMatch(slideXml, /mc:AlternateContent/, 'an ECMA-376 chart needs no compatibility wrapper')
	assert.match(slideXml, /<a:graphicData uri="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/chart">/, 'classic chart graphicData uri changed')
	const chartParts = Object.keys(barZip.files).filter(name => /^ppt\/charts\/chart\d+\.xml$/.test(name))
	assert.equal(chartParts.length, 1, 'classic chart part missing')
	assert.equal(Object.keys(barZip.files).filter(name => /chartEx/.test(name)).length, 0, 'no chartex part may appear without a chartex chart')
	// the style and colour parts belong to every chart, so only the chartex content type must be absent
	assert.doesNotMatch(await readPart(barZip, '[Content_Types].xml'), /ms-office\.chartex/, 'no chartex content type may be declared without a chartex chart')
})

test('contract: chartex data that PowerPoint would reject is normalized, not emitted', async () => {
	let ragged = ''
	let multi = ''
	let extra = ''
	let noLabels = ''
	const warnings = await captureWarnings(async () => {
		// a blank point must stay blank: `cx:pt` in a numeric dimension has to be a number, and the range
		// each `cx:f` names has to be exactly as long as its sibling's `ptCount`
		ragged = await chartExXml('waterfall', [{ name: 'Cash', labels: ['a', 'b', 'c', 'd'], values: [10, null, 30] }])
		// chartex cell references assume one label column; the multi-level worksheet layout shifts series right
		multi = await chartExXml('treemap', [{ name: 'R', labels: [['Gear', 'Berg', 'Motr'], ['Mech', '', '']], values: [1, 2, 3] }])
		// only box & whisker plots more than one series
		extra = await chartExXml('waterfall', [{ name: 'A', labels: ['a', 'b'], values: [1, 2] }, { name: 'B', labels: ['a', 'b'], values: [3, 4] }])
		// a histogram bins raw values, so a label-less series is the natural call shape
		noLabels = await chartExXml('histogram', [{ name: 'Ages', values: [1, 2, 3] }])
	})

	assert.doesNotMatch(ragged, />null<|>undefined</, 'a missing point must be blank, not the string "null"')
	assert.match(ragged, /<cx:pt idx="1"><\/cx:pt>/, 'a missing point must still hold its index')
	assert.match(ragged, /<cx:f>Sheet1!\$B\$2:\$B\$4<\/cx:f><cx:lvl ptCount="3"/, 'the value range must be as long as its ptCount')
	assert.match(ragged, /<cx:f>Sheet1!\$A\$2:\$A\$5<\/cx:f><cx:lvl ptCount="4"/, 'the category range must be as long as its ptCount')

	// column A holds the leaf labels, so that is the level the chart must name
	assert.match(multi, /<cx:f>Sheet1!\$A\$2:\$A\$4<\/cx:f><cx:lvl ptCount="3"><cx:pt idx="0">Gear</, 'the leaf label level must address column A')
	assert.match(multi, /type="size"><cx:f>Sheet1!\$B\$2:\$B\$4</, 'values must stay in column B after flattening')
	assert.doesNotMatch(multi, /Mech/, 'the extra label level must be dropped, not emitted')

	assert.equal((extra.match(/<cx:series /g) ?? []).length, 1, 'a waterfall plots one series')
	assert.match(noLabels, /<cx:numDim type="val"><cx:f>Sheet1!\$B\$2:\$B\$4</, 'a label-less histogram must still address its values')

	assert.equal(warnings.length, 2, `expected one warning per normalization, got: ${warnings.join(' | ')}`)

	// box & whisker is the one layout that keeps every series
	const box = await chartExXml('boxWhisker', [{ name: 'A', labels: ['a', 'b'], values: [1, 2] }, { name: 'B', labels: ['a', 'b'], values: [3, 4] }])
	assert.equal((box.match(/<cx:series /g) ?? []).length, 2, 'box & whisker plots one series per distribution')
})

test('contract: chartex options use the cx vocabulary, not the ECMA-376 one', async () => {
	// `cx:legend@pos` takes l/r/t/b only - the library's `tr` has no chartex equivalent
	assert.match(
		await chartExXml('treemap', [{ name: 'R', labels: ['a'], values: [1] }], { showLegend: true, legendPos: 'tr' }),
		/<cx:legend pos="r" align="min" overlay="0"\/>/,
		'top-right must become a top-aligned right legend'
	)

	// `dataLabelPosition` must reach the emitter rather than being validated against `c:dLblPos`
	let honoured = ''
	let rejected = ''
	const warnings = await captureWarnings(async () => {
		honoured = await chartExXml('waterfall', [{ name: 'C', labels: ['a', 'b'], values: [1, 2] }], { showValue: true, dataLabelPosition: 'ctr' })
		rejected = await chartExXml('waterfall', [{ name: 'C', labels: ['a', 'b'], values: [1, 2] }], { showValue: true, dataLabelPosition: 'bestFit' })
	})

	assert.match(honoured, /<cx:dataLabels pos="ctr">/, 'a valid chartex label position must survive')
	assert.equal(warnings.length, 1, `only the unsupported position may warn, got: ${warnings.join(' | ')}`)
	assert.match(rejected, /<cx:dataLabels pos="outEnd">/, 'an unsupported position falls back to the layout default')

	// these layouts colour each point from the colour style; only an explicit series colour overrides it
	const coloured = await chartExXml('funnel', [{ name: 'P', labels: ['a', 'b'], values: [5, 3], color: 'FF0000' }], { chartColors: ['00FF00'] })
	assert.match(coloured, /<cx:spPr><a:solidFill><a:srgbClr val="FF0000"\/><\/a:solidFill><\/cx:spPr>/, 'an explicit series colour must be emitted')
	assert.doesNotMatch(coloured, /00FF00/, 'chartColors does not apply to a chartex layout')
})

test('contract: the chartex fallback shape mirrors the chart frame it replaces', async () => {
	const { zip } = await buildChartEx('waterfall', WATERFALL_DATA, { objectName: 'WF', title: 'Quarterly cash', showTitle: true })
	const slideXml = await readPart(zip, 'ppt/slides/slide1.xml')
	const fallback = slideXml.slice(slideXml.indexOf('<mc:Fallback>'), slideXml.indexOf('</mc:Fallback>'))

	assert.match(fallback, /name="WF"/, 'the fallback must carry the chart name')
	// a chart's `title` is the chart title, not the alt-text title `p:cNvPr@title` carries
	assert.doesNotMatch(fallback, /title="Quarterly cash"/, 'the chart title must not leak into the alt-text title')
	assert.match(fallback, /<a:spLocks noTextEdit="1"\/>/, 'the fallback must never be text-editable')
})
