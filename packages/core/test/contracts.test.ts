/**
 * Semantic contracts for the core slide and chart paths.
 *
 * Unlike golden XML snapshots, these checks document the OOXML that matters and allow harmless
 * serializer changes without regenerating fixture files.
 */
import { test, before } from 'node:test'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import pptxgen from '../src/pptxgen'
import { assertEmbeddedXlsxContracts, assertPptxPackageContracts, readPart } from './pptx-contracts'

let zip: JSZip

before(async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('Contract', { x: 0.5, y: 0.3, w: 6, h: 0.5, fontSize: 18, color: '0000FF', bold: true })
	slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1.2, w: 2, h: 1, fill: { color: 'FF0000' } })
	slide.addTable([['A', 'B'], ['1', '2']], { x: 0.5, y: 2.6, w: 5 })
	slide.addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [10, 20] }], { x: 0.5, y: 4, w: 6, h: 3 })
	zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
})

test('contract: package parts and relationships are coherent', async () => {
	await assertPptxPackageContracts(zip)
	await assertEmbeddedXlsxContracts(zip)
})

test('contract: rejects a part without a declared content type', async () => {
	const invalidZip = await JSZip.loadAsync(await zip.generateAsync({ type: 'nodebuffer' }))
	invalidZip.file('ppt/undeclared.bin', 'invalid')
	await assert.rejects(assertPptxPackageContracts(invalidZip), /package part has no content type/)
})

test('contract: validates relationship references with any legal ID', async () => {
	const invalidZip = await JSZip.loadAsync(await zip.generateAsync({ type: 'nodebuffer' }))
	const slideXml = await readPart(invalidZip, 'ppt/slides/slide1.xml')
	const referencePattern = /r:(id|embed|link)="rId\d+"/
	assert.match(slideXml, referencePattern, 'test presentation has no relationship reference')
	invalidZip.file('ppt/slides/slide1.xml', slideXml.replace(referencePattern, (_match, attribute) => `r:${attribute}="custom-id"`))
	await assert.rejects(assertPptxPackageContracts(invalidZip), /missing custom-id relationship/)
})

test('contract: slide keeps text, shape, and table semantics', async () => {
	const xml = await readPart(zip, 'ppt/slides/slide1.xml')
	assert.match(xml, /<a:t>Contract<\/a:t>/, 'text content missing')
	assert.match(xml, /<a:prstGeom prst="rect">/, 'rectangle shape missing')
	assert.match(xml, /<a:srgbClr val="FF0000"\/>/, 'shape fill missing')
	assert.match(xml, /<a:tbl>/, 'table missing')
	assert.equal([...xml.matchAll(/<a:gridCol /g)].length, 2, 'table grid width changed')
})

test('contract: bar chart keeps its data and chart type', async () => {
	const xml = await readPart(zip, 'ppt/charts/chart1.xml')
	assert.match(xml, /<c:barChart>/, 'bar chart missing')
	assert.match(xml, /<c:v>Sales<\/c:v>/, 'series name missing')
	assert.match(xml, /<c:v>Q1<\/c:v>/, 'category label missing')
	assert.match(xml, /<c:v>20<\/c:v>/, 'series value missing')
})

test('contract: gradient fills reach shapes, lines, table cells, and slide backgrounds', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.background = { type: 'gradient', gradient: { angle: 45, stops: [{ color: 'FFFFFF', position: 0 }, { color: 'E7E6E6', position: 100 }] } }
	slide.addShape(pptx.ShapeType.rect, {
		x: 1,
		y: 1,
		w: 4,
		h: 2,
		fill: { type: 'gradient', gradient: { stops: [{ color: 'FF0000', position: 0 }, { color: '0000FF', position: 100, transparency: 20 }] } },
		line: { type: 'gradient', width: 2, gradient: { type: 'radial', stops: [{ color: '00FF00', position: 0 }, { color: '000000', position: 100 }] } },
	})
	slide.addTable([[{ text: 'grad', options: { fill: { type: 'gradient', gradient: { stops: [{ color: '111111', position: 0 }, { color: '222222', position: 100 }] } } } }]], { x: 1, y: 4, w: 4 })
	const gradZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)

	await assertPptxPackageContracts(gradZip)
	const xml = await readPart(gradZip, 'ppt/slides/slide1.xml')
	assert.doesNotMatch(xml, /NaN|undefined/, 'gradient options must not leak invalid attribute values')
	assert.match(xml, /<p:bg><p:bgPr><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">.*?<a:lin ang="2700000" scaled="0"\/><\/a:gradFill><a:effectLst\/><\/p:bgPr><\/p:bg>/, 'slide background gradient missing')
	assert.match(xml, /<a:gs pos="100000"><a:srgbClr val="0000FF"><a:alpha val="80000"\/><\/a:srgbClr><\/a:gs>/, 'shape gradient stop transparency missing')
	assert.match(xml, /<a:ln w="25400"><a:gradFill [^>]*>.*?<a:path path="circle">/, 'line gradient missing')
	// per CT_TableCellProperties the fill follows the cell line elements
	assert.match(xml, /<\/a:lnB><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="111111"\/>/, 'table cell gradient missing')
	assert.equal([...xml.matchAll(/<a:gradFill /g)].length, 4, 'one gradient per requested fill site')
})

test('contract: solid, none, and string fills are unchanged by gradient support', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.background = { color: 'FFFF00' }
	slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, fill: { color: 'FF0000', transparency: 50 }, line: { color: '00FF00', width: 1 } })
	slide.addTable([[{ text: 'solid', options: { fill: { color: '112233' } } }]], { x: 1, y: 3, w: 4 })
	const xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /gradFill/, 'no gradient emitted for solid fills')
	assert.match(xml, /<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFF00"\/><\/a:solidFill><a:effectLst\/><\/p:bgPr><\/p:bg>/, 'solid background changed')
	assert.match(xml, /<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"\/><\/a:srgbClr><\/a:solidFill>/, 'solid shape fill changed')
	assert.match(xml, /<a:ln w="12700"><a:solidFill><a:srgbClr val="00FF00"\/><\/a:solidFill><a:prstDash val="solid"\/><\/a:ln>/, 'solid line changed')
	assert.match(xml, /<a:solidFill><a:srgbClr val="112233"\/><\/a:solidFill>/, 'solid table cell fill changed')
})

test('contract: OMML math runs are emitted as an a14 math zone with a plain-text fallback', async () => {
	const FRAC = '<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>'
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText([{ text: 'before ' }, { text: 'a/b', options: { omml: FRAC } }, { text: ' after' }], { x: 1, y: 1, w: 6, h: 1 })
	slide.addText([{ text: '', options: { omml: FRAC } }], { x: 1, y: 3, w: 6, h: 1 })
	const mathZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(mathZip)
	const xml = await readPart(mathZip, 'ppt/slides/slide1.xml')

	// ECMA-376 Part 3: an extension to the text-body content model is offered via mc:AlternateContent
	assert.match(
		xml,
		/<mc:AlternateContent xmlns:mc="http:\/\/schemas\.openxmlformats\.org\/markup-compatibility\/2006"><mc:Choice xmlns:a14="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2010\/main" Requires="a14"><a14:m>/,
		'math zone is not offered through mc:AlternateContent'
	)
	// the OMML payload is raw XML, never entity-encoded
	assert.match(xml, /<m:oMath xmlns:m="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/math" xmlns:w="[^"]+"><m:f>/, 'OMML root/namespaces missing')
	assert.doesNotMatch(xml, /&lt;m:/, 'OMML must not be entity-encoded')
	// `text` becomes the fallback for consumers that do not understand a14
	assert.match(xml, /<mc:Fallback><a:r>.*?<a:t>a\/b<\/a:t><\/a:r><\/mc:Fallback>/, 'plain-text fallback missing')
	// math-only run: empty text yields an empty fallback, and the run still survives
	assert.match(xml, /<mc:Fallback\/><\/mc:AlternateContent>/, 'math-only run should emit an empty fallback')
	assert.equal([...xml.matchAll(/<a14:m>/g)].length, 2, 'expected one math zone per math run')

	// a paragraph permits exactly one pPr, and it must precede the runs
	const para = /<a:p>(?:(?!<\/a:p>)[\s\S])*a14:m[\s\S]*?<\/a:p>/.exec(xml)?.[0] ?? ''
	assert.equal([...para.matchAll(/<a:pPr[ >]/g)].length, 1, 'a paragraph with math must still emit exactly one pPr')
	assert.match(para, /^<a:p><a:pPr/, 'pPr must be the first child of the paragraph')
	assert.match(para, /<a:t>before <\/a:t>[\s\S]*<a14:m>[\s\S]*<a:t> after<\/a:t>/, 'math must stay between its sibling plain runs')
})

test('contract: OMML input is normalized without touching plain text output', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	// caller-supplied root + namespace must not be duplicated
	slide.addText([{ text: '', options: { omml: '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath>' } }], { x: 1, y: 1, w: 4, h: 1 })
	// an already-wrapped math zone passes through
	slide.addText([{ text: '', options: { omml: '<a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>y</m:t></m:r></m:oMath></a14:m>' } }], { x: 1, y: 2, w: 4, h: 1 })
	// blank/whitespace omml falls back to a normal text run
	slide.addText([{ text: 'plain', options: { omml: '   ' } }], { x: 1, y: 3, w: 4, h: 1 })
	const xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')

	assert.equal([...xml.matchAll(/xmlns:m="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/math"/g)].length, 2, 'namespace declared once per math zone')
	assert.equal([...xml.matchAll(/<a14:m[ >]/g)].length, 2, 'expected exactly two math zones')
	assert.match(xml, /<a:r><a:rPr[^>]*>(?:(?!<\/a:r>)[\s\S])*<a:t>plain<\/a:t><\/a:r>/, 'blank omml must fall back to a plain run')

	// shape-level `omml` applies to a single-run `addText(string)`, but must never be cloned into
	// every run of a multi-run call - math is a per-run payload
	const shapeLevel = new pptxgen()
	const shapeSlide = shapeLevel.addSlide()
	shapeSlide.addText('a/b', { x: 1, y: 1, w: 4, h: 1, omml: '<m:r><m:t>x</m:t></m:r>' })
	shapeSlide.addText([{ text: 'one' }, { text: 'two' }], { x: 1, y: 3, w: 4, h: 1, omml: '<m:r><m:t>x</m:t></m:r>' })
	const shapeXml = await readPart(await JSZip.loadAsync((await shapeLevel.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	assert.equal([...shapeXml.matchAll(/<a14:m[ >]/g)].length, 1, 'shape-level omml must not be cloned across runs')
	assert.match(shapeXml, /<a:t>one<\/a:t>/, 'sibling runs keep their plain text')
	assert.match(shapeXml, /<a:t>two<\/a:t>/, 'sibling runs keep their plain text')

	// text-only presentations must be untouched by math support
	const plain = new pptxgen()
	plain.addSlide().addText('no math here', { x: 1, y: 1, w: 4, h: 1 })
	const plainXml = await readPart(await JSZip.loadAsync((await plain.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	assert.doesNotMatch(plainXml, /mc:AlternateContent|a14:m|m:oMath/, 'text-only output must not gain math markup')
})

test('contract: chartTrackingRefBased is on by default and can be turned off', async () => {
	const withChart = (pptx: pptxgen): void => {
		pptx.addSlide().addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1'], values: [10] }], { x: 1, y: 1, w: 4, h: 3 })
	}
	// @note chart part numbering is process-global, so resolve the part rather than assuming `chart1.xml`
	const chartPart = (source: JSZip): string => {
		const name = Object.keys(source.files).find(file => /^ppt\/charts\/chart\d+\.xml$/.test(file))
		assert.ok(name, 'chart part missing')
		return name
	}

	// DEFAULT: PowerPoint writes this on every presentation it creates, so PptxGenJS does too
	const on = new pptxgen()
	assert.equal(on.chartTrackingRefBased, true, 'chartTrackingRefBased should default to true')
	withChart(on)
	const onZip = await JSZip.loadAsync((await on.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(onZip)
	const onXml = await readPart(onZip, 'ppt/presProps.xml')
	// MS-PPTX 2.2 requires extensions to sit inside an extLst wrapper with the URI from 2.2.12
	assert.match(
		onXml,
		/<p:presentationPr [^>]*><p:extLst><p:ext uri="\{FD5EFAAD-0ECE-453E-9831-46B23BE46B34\}"><p15:chartTrackingRefBased xmlns:p15="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2012\/main" val="1"\/><\/p:ext><\/p:extLst><\/p:presentationPr>$/,
		'chartTrackingRefBased extension missing or malformed'
	)

	// OFF: the part goes back to the empty element written before this property existed
	const off = new pptxgen()
	off.chartTrackingRefBased = false
	withChart(off)
	const offZip = await JSZip.loadAsync((await off.write({ outputType: 'nodebuffer' })) as Buffer)
	const offXml = await readPart(offZip, 'ppt/presProps.xml')
	assert.doesNotMatch(offXml, /extLst|chartTrackingRefBased/, 'disabling must remove the extension')
	assert.match(offXml, /<p:presentationPr [^>]*\/>$/, 'disabled presProps.xml must be the empty element')

	// This is a presentation-level flag: it must not touch chart output either way, which is what
	// makes the default safe to change - slides stay visually identical
	assert.equal(await readPart(onZip, chartPart(onZip)), await readPart(offZip, chartPart(offZip)), 'chart XML changed')

	// a non-boolean from plain JS must not silently flip the flag in either direction
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	try {
		const bogus = new pptxgen()
		;(bogus as unknown as { chartTrackingRefBased: unknown }).chartTrackingRefBased = 'no'
		assert.equal(bogus.chartTrackingRefBased, true, 'a non-boolean must leave the default in place')
	} finally {
		console.warn = origWarn
	}
	assert.ok(warnings.some(warning => warning.includes('chartTrackingRefBased must be a boolean')), 'a non-boolean must warn')
})

test('contract: media playback options drive the slide timing tree', async () => {
	const MP3 = 'audio/mp3;base64,QQ=='
	const MP4 = 'video/mp4;base64,QQ=='

	// A: no playback options -> no `p:timing` at all, so existing decks are unchanged
	const plain = new pptxgen()
	const plainSlide = plain.addSlide()
	plainSlide.addMedia({ type: 'video', data: MP4, x: 1, y: 1, w: 3, h: 2 })
	const plainXml = await readPart(await JSZip.loadAsync((await plain.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	assert.doesNotMatch(plainXml, /<p:timing>/, 'media without playback options must not add a timing tree')

	// B: playback options -> one media node per media shape
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addMedia({ type: 'video', data: MP4, x: 1, y: 1, w: 3, h: 2, autoplay: true, loop: true, fullScreen: true, mute: true })
	slide.addMedia({ type: 'audio', data: MP3, x: 5, y: 1, w: 2, h: 2 })
	const zipWithTiming = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(zipWithTiming)
	const xml = await readPart(zipWithTiming, 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /NaN|undefined/, 'playback options must not leak invalid attribute values')
	// CT_Slide sequence requires `p:timing` after `p:clrMapOvr`
	assert.match(xml, /<\/p:clrMapOvr><p:timing>/, 'timing tree is in the wrong position')
	assert.match(xml, /<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>/, 'tmRoot node missing')
	assert.match(
		xml,
		/<p:video fullScrn="1"><p:cMediaNode mute="1"><p:cTn id="2" fill="hold" display="0" repeatCount="indefinite"><p:stCondLst><p:cond delay="0"\/><\/p:stCondLst><\/p:cTn><p:tgtEl><p:spTgt spid="\d+"\/><\/p:tgtEl><\/p:cMediaNode><\/p:video>/,
		'video playback node missing or malformed'
	)
	// audio has no playback options of its own, so it gets no node
	assert.equal([...xml.matchAll(/<p:(video|audio)[ >]/g)].length, 1, 'only media with playback options gets a node')
	// `spTgt@spid` must match the media shape's `p:cNvPr@id`
	const spid = /<p:spTgt spid="(\d+)"\/>/.exec(xml)?.[1]
	assert.ok(spid, 'no media target')
	assert.match(xml, new RegExp(`<p:cNvPr id="${spid}" name="Media 0"`), 'timing target does not match the media shape id')
	// audio references a:audioFile, video a:videoFile
	assert.match(xml, /<a:audioFile r:link="rId\d+"\/>/, 'audio must reference a:audioFile')
	assert.match(xml, /<a:videoFile r:link="rId\d+"\/>/, 'video must reference a:videoFile')
})

test('contract: invalid media playback combinations are dropped before XML', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let xml = ''
	try {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		// fullScreen is a `p:video` attribute - meaningless for audio
		slide.addMedia({ type: 'audio', data: 'audio/mp3;base64,QQ==', x: 1, y: 1, w: 2, h: 2, autoplay: true, fullScreen: true })
		// online videos are played by the embed, not the timing tree
		slide.addMedia({ type: 'online', link: 'https://www.youtube.com/embed/Dph6ynRVyUc', x: 4, y: 1, w: 3, h: 2, autoplay: true, loop: true })
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	} finally {
		console.warn = origWarn
	}

	assert.equal(warnings.filter(w => w.includes('`fullScreen` applies to `type:"video"` only')).length, 1, 'audio fullScreen must warn')
	assert.equal(warnings.filter(w => w.includes('not supported for `type:"online"`')).length, 1, 'online playback options must warn')
	assert.doesNotMatch(xml, /fullScrn/, 'fullScreen must not be emitted for audio')
	assert.equal([...xml.matchAll(/<p:audio[ >]/g)].length, 1, 'the audio node is still emitted for autoplay')
	assert.equal([...xml.matchAll(/<p:video[ >]/g)].length, 0, 'online video must not get a timing node')
})
test('contract: slide-show and image/view-mode extensions land in presProps.xml', async () => {
	const pptx = new pptxgen()
	pptx.slideShow = { mode: 'browse', loop: true, showNarration: false, useTimings: false, browseMode: true, laserColor: 'FF0000' }
	pptx.defaultImageDpi = 220
	pptx.discardImageEditData = true
	pptx.readonlyRecommended = true
	pptx.addSlide().addText('show options', { x: 1, y: 1, w: 4, h: 1 })
	const showZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(showZip)
	const xml = await readPart(showZip, 'ppt/presProps.xml')

	assert.doesNotMatch(xml, /NaN|undefined/, 'presentation properties must not leak invalid values')
	// CT_PresentationProperties sequence: showPr precedes extLst
	assert.match(xml, /<p:showPr loop="1" showNarration="0" useTimings="0"><p:browse showScrollbar="1"\/>/, 'showPr attributes/choice missing')
	assert.match(xml, /<\/p:showPr><p:extLst>/, 'showPr must precede the presentationPr extLst')

	// showPr extensions (MS-PPTX 2.2.6)
	assert.match(xml, /<p:ext uri="\{F99C55AA-B7CB-42B0-86F8-08522FDF87E8\}"><p14:browseMode xmlns:p14="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2010\/main" val="1"\/><\/p:ext>/, 'browseMode ext missing')
	assert.match(xml, /<p:ext uri="\{EC167BDD-8182-4AB7-AECC-EB403E3ABB37\}"><p14:laserClr xmlns:p14="[^"]+"><a:srgbClr val="FF0000"\/><\/p14:laserClr><\/p:ext>/, 'laserClr ext missing')

	// presentationPr extensions (MS-PPTX 2.2.7 and 2.2.16)
	assert.match(xml, /<p:ext uri="\{E76CE94A-603C-4142-B9EB-6D1370010A27\}"><p14:discardImageEditData xmlns:p14="[^"]+" val="1"\/><\/p:ext>/, 'discardImageEditData ext missing')
	assert.match(xml, /<p:ext uri="\{D31A062A-798A-4329-ABDD-BBA856620510\}"><p14:defaultImageDpi xmlns:p14="[^"]+" val="220"\/><\/p:ext>/, 'defaultImageDpi ext missing')
	assert.match(xml, /<p:ext uri="\{1BD7E111-0CB8-44D6-8891-C1BB2F81B7CC\}"><p1710:readonlyRecommended xmlns:p1710="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2017\/10\/main" val="1"\/><\/p:ext>/, 'readonlyRecommended ext missing')
})

test('contract: each presentation property emits only its own extension', async () => {
	// `chartTrackingRefBased` is on by default, so it is the whole baseline of this part
	const plain = new pptxgen()
	plain.addSlide()
	const plainXml = await readPart(await JSZip.loadAsync((await plain.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/presProps.xml')
	assert.doesNotMatch(plainXml, /showPr|defaultImageDpi|discardImageEditData|readonlyRecommended/, 'no unrequested extension may appear')
	assert.equal([...plainXml.matchAll(/<p:ext /g)].length, 1, 'only the default chart-tracking extension is expected')

	// turning everything off leaves the empty element written before these properties existed
	const off = new pptxgen()
	off.chartTrackingRefBased = false
	off.addSlide()
	const offXml = await readPart(await JSZip.loadAsync((await off.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/presProps.xml')
	assert.match(offXml, /<p:presentationPr [^>]*\/>$/, 'presProps.xml must be the empty element when nothing is set')

	// a single property adds a single extension
	const dpiOnly = new pptxgen()
	dpiOnly.chartTrackingRefBased = false
	dpiOnly.defaultImageDpi = 0 // 0 is meaningful: "do not compress"
	dpiOnly.addSlide()
	const dpiXml = await readPart(await JSZip.loadAsync((await dpiOnly.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/presProps.xml')
	assert.match(dpiXml, /<p14:defaultImageDpi xmlns:p14="[^"]+" val="0"\/>/, 'a DPI of 0 must still be written')
	assert.doesNotMatch(dpiXml, /showPr|discardImageEditData|readonlyRecommended|chartTracking/, 'unrelated extensions must not appear')

	// default mode with no extensions still writes a valid showPr choice, and no empty extLst
	const presentOnly = new pptxgen()
	presentOnly.chartTrackingRefBased = false
	presentOnly.slideShow = { loop: true }
	presentOnly.addSlide()
	const presentXml = await readPart(await JSZip.loadAsync((await presentOnly.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/presProps.xml')
	assert.match(presentXml, /<p:showPr loop="1"><p:present\/><\/p:showPr>/, 'default show mode must be present')
	assert.doesNotMatch(presentXml, /<p:extLst>/, 'no extLst when nothing needs one')
})

test('contract: invalid presentation-property input is dropped with a warning', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let xml = ''
	try {
		const pptx = new pptxgen()
		pptx.slideShow = { mode: 'cinema' as unknown as 'kiosk' }
		pptx.defaultImageDpi = -1
		;(pptx as unknown as { readonlyRecommended: unknown }).readonlyRecommended = 'yes'
		pptx.addSlide()
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/presProps.xml')
	} finally {
		console.warn = origWarn
	}

	assert.ok(warnings.some(w => w.includes('slideShow.mode must be')), 'invalid mode must warn')
	assert.ok(warnings.some(w => w.includes('defaultImageDpi must be a number >= 0')), 'negative dpi must warn')
	assert.match(xml, /<p:showPr><p:present\/><\/p:showPr>/, 'invalid mode falls back to present')
	assert.doesNotMatch(xml, /defaultImageDpi/, 'negative dpi must not be emitted')
	assert.doesNotMatch(xml, /readonlyRecommended/, 'a non-boolean must not enable the extension')
	assert.doesNotMatch(xml, /NaN/, 'no NaN attributes')
})
test('contract: base slide transitions land after clrMapOvr with their directional attribute', async () => {
	const pptx = new pptxgen()
	pptx.addSlide({ transition: { type: 'wipe', direction: 'left', speed: 'slow' } })
	pptx.addSlide({ transition: { type: 'split', orient: 'vertical', direction: 'in' } })
	pptx.addSlide({ transition: { type: 'wheel', spokes: 8 } })
	pptx.addSlide({ transition: { type: 'fade', thruBlk: true, advClick: false, advTm: 5000 } })
	pptx.addSlide({ transition: { type: 'strips', direction: 'bottomRight' } })
	const transZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(transZip)
	const [wipe, split, wheel, fade, strips] = await Promise.all(
		[1, 2, 3, 4, 5].map(async num => await readPart(transZip, `ppt/slides/slide${num}.xml`))
	)

	// CT_Slide sequence: cSld, clrMapOvr, transition, timing
	assert.match(wipe, /<\/p:clrMapOvr><p:transition spd="slow"><p:wipe dir="l"\/><\/p:transition><\/p:sld>$/, 'wipe transition missing or misplaced')
	assert.match(split, /<p:transition><p:split orient="vert" dir="in"\/><\/p:transition>/, 'split needs both orient and dir')
	assert.match(wheel, /<p:transition><p:wheel spokes="8"\/><\/p:transition>/, 'wheel spokes missing')
	assert.match(fade, /<p:transition advClick="0" advTm="5000"><p:fade thruBlk="1"\/><\/p:transition>/, 'fade attributes missing')
	assert.match(strips, /<p:transition><p:strips dir="rd"\/><\/p:transition>/, 'friendly corner direction not translated')
	// no transition set -> nothing emitted
	const plain = new pptxgen()
	plain.addSlide()
	const plainXml = await readPart(await JSZip.loadAsync((await plain.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	assert.doesNotMatch(plainXml, /p:transition|mc:AlternateContent/, 'a slide without a transition must be unchanged')
})

test('contract: modern transitions and millisecond durations carry a base fallback', async () => {
	const pptx = new pptxgen()
	const morph = pptx.addSlide()
	morph.transition = { type: 'morph', duration: 1200, advClick: false }
	pptx.addSlide({ transition: { type: 'conveyor', direction: 'rightup', duration: 900 } })
	// a base transition with a millisecond duration also needs the p14 extension, so it is wrapped too
	pptx.addSlide({ transition: { type: 'push', direction: 'up', duration: 700, speed: 'fast' } })
	const modernZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(modernZip)
	const [morphXml, conveyorXml, pushXml] = await Promise.all(
		[1, 2, 3].map(async num => await readPart(modernZip, `ppt/slides/slide${num}.xml`))
	)

	// morph lives in p16 and falls back to fade
	assert.match(morphXml, /<mc:Choice [^>]*Requires="p16"><p:transition p14:dur="1200" advClick="0"><p16:morph\/><\/p:transition><\/mc:Choice>/, 'morph choice missing')
	assert.match(morphXml, /<mc:Fallback><p:transition advClick="0"><p:fade\/><\/p:transition><\/mc:Fallback>/, 'morph fallback missing')
	assert.match(morphXml, /xmlns:p16="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2016\/main"/, 'p16 namespace missing')
	assert.match(morphXml, /xmlns:p14="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2010\/main"/, 'p14 namespace missing for p14:dur')

	// conveyor lives in p14 and falls back to push
	assert.match(conveyorXml, /<mc:Choice [^>]*Requires="p14"><p:transition p14:dur="900"><p14:conveyor dir="ru"\/><\/p:transition><\/mc:Choice><mc:Fallback><p:transition><p:push\/><\/p:transition><\/mc:Fallback>/, 'conveyor choice/fallback missing')

	// `p14:dur` never appears in a fallback (its namespace is exactly what the fallback consumer lacks),
	// and `spd` is what the fallback keeps instead
	assert.match(pushXml, /<mc:Choice [^>]*Requires="p14"><p:transition p14:dur="700"><p:push dir="u"\/><\/p:transition><\/mc:Choice>/, 'duration must be offered via a Choice')
	assert.match(pushXml, /<mc:Fallback><p:transition spd="fast"><p:push dir="u"\/><\/p:transition><\/mc:Fallback>/, 'fallback must use spd, not p14:dur')
	const fallbacks = [morphXml, conveyorXml, pushXml].map(xml => /<mc:Fallback>[\s\S]*?<\/mc:Fallback>/.exec(xml)?.[0] ?? '')
	fallbacks.forEach(fb => { assert.doesNotMatch(fb, /p14:|p16:/, 'a fallback must not use extension markup') })
})

test('contract: invalid transition input is dropped with a warning', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let xmls: string[] = []
	try {
		const pptx = new pptxgen()
		pptx.addSlide({ transition: { type: 'nonsense' as unknown as 'fade' } })
		pptx.addSlide({ transition: { type: 'wipe', direction: 'in' } }) // wipe takes l/r/u/d only
		pptx.addSlide({ transition: { type: 'wheel', spokes: 5 as unknown as 4 } })
		pptx.addSlide({ transition: { type: 'push', speed: 'turbo' as unknown as 'fast', duration: NaN, advTm: NaN } })
		const zipInvalid = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
		xmls = await Promise.all([1, 2, 3, 4].map(async num => await readPart(zipInvalid, `ppt/slides/slide${num}.xml`)))
	} finally {
		console.warn = origWarn
	}

	assert.ok(warnings.some(w => w.includes('unknown slide transition "nonsense"')), 'unknown type must warn')
	assert.ok(warnings.some(w => w.includes('does not accept direction "in"')), 'invalid direction must warn')
	assert.ok(warnings.some(w => w.includes('spokes must be one of')), 'invalid spokes must warn')
	assert.ok(warnings.some(w => w.includes('speed must be \'slow\' | \'med\' | \'fast\'')), 'invalid speed must warn')

	assert.doesNotMatch(xmls[0], /p:transition/, 'unknown type must emit no transition')
	assert.match(xmls[1], /<p:transition><p:wipe\/><\/p:transition>/, 'invalid direction is dropped, transition kept')
	assert.match(xmls[2], /<p:transition><p:wheel\/><\/p:transition>/, 'invalid spokes are dropped, transition kept')
	assert.match(xmls[3], /<p:transition><p:push\/><\/p:transition>/, 'invalid speed/duration/advTm are dropped')
	xmls.forEach(xml => { assert.doesNotMatch(xml, /NaN/, 'no NaN attributes') })
})
test('contract: animation presets build a mainSeq with click groups and a build list', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('one', { x: 1, y: 1, w: 3, h: 1, animation: { type: 'fadeIn' } })
	slide.addText('two', { x: 1, y: 2, w: 3, h: 1, animation: { type: 'wipeIn', direction: 'left', trigger: 'withPrevious', duration: 800 } })
	slide.addShape(pptx.ShapeType.rect, { x: 1, y: 3, w: 3, h: 1, animation: { type: 'zoomOut', trigger: 'afterPrevious', delay: 250 } })
	const animZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(animZip)
	const xml = await readPart(animZip, 'ppt/slides/slide1.xml')
	const timing = /<p:timing>[\s\S]*<\/p:timing>/.exec(xml)?.[0] ?? ''

	assert.doesNotMatch(timing, /NaN|undefined/, 'animation options must not leak invalid attribute values')
	assert.match(timing, /<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">/, 'mainSeq missing')

	// preset id/class/subtype per effect, and the trigger drives nodeType
	assert.match(timing, /presetID="10" presetClass="entr" presetSubtype="0"[^>]*nodeType="clickEffect"/, 'fadeIn effect missing')
	assert.match(timing, /presetID="22" presetClass="entr" presetSubtype="8"[^>]*nodeType="withEffect"/, 'wipeIn left effect missing')
	assert.match(timing, /presetID="53" presetClass="exit" presetSubtype="32"[^>]*nodeType="afterEffect"/, 'zoomOut effect missing')
	assert.match(timing, /<p:animEffect transition="in" filter="fade">/, 'fade filter missing')
	assert.match(timing, /<p:animEffect transition="in" filter="wipe\(left\)">/, 'wipe filter missing')
	assert.match(timing, /<p:animEffect transition="out" filter="zoom\(out\)">/, 'exit effect must animate out')
	assert.match(timing, /<p:cTn id="\d+" dur="800"\/>/, 'custom duration missing')
	assert.match(timing, /<p:stCondLst><p:cond delay="250"\/><\/p:stCondLst>/, 'delay missing')

	// an exit effect hides its target when it ends; an entrance shows it as it starts
	assert.match(timing, /<p:to><p:strVal val="hidden"\/><\/p:to>/, 'exit effect must set visibility hidden')
	assert.match(timing, /<p:to><p:strVal val="visible"\/><\/p:to>/, 'entrance effect must set visibility visible')

	// `withPrevious`/`afterPrevious` join the click group before them, so there is one click group here
	assert.equal([...timing.matchAll(/<p:cond delay="indefinite"\/><p:cond evt="onBegin"/g)].length, 1, 'expected a single click group')
	// every animated shape is listed in the build list
	assert.match(timing, /<p:bldLst><p:bldP spid="2" grpId="0"\/><p:bldP spid="3" grpId="0"\/><p:bldP spid="4" grpId="0"\/><\/p:bldLst>/, 'build list missing')
	// node ids must be unique across the whole tree
	const ids = [...timing.matchAll(/<p:cTn id="(\d+)"/g)].map(match => match[1])
	assert.equal(new Set(ids).size, ids.length, `duplicate timing node ids: ${ids.join(',')}`)

	// no animation -> no timing tree
	const plain = new pptxgen()
	plain.addSlide().addText('static', { x: 1, y: 1, w: 3, h: 1 })
	const plainXml = await readPart(await JSZip.loadAsync((await plain.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	assert.doesNotMatch(plainXml, /p:timing/, 'a slide without animations or media playback must be unchanged')
})

test('contract: animations and media playback share one timing tree', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('fades', { x: 1, y: 1, w: 3, h: 1, animation: { type: 'fadeIn' } })
	slide.addMedia({ type: 'video', data: 'video/mp4;base64,QQ==', x: 5, y: 1, w: 2, h: 2, autoplay: true, loop: true })
	const xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	const timing = /<p:timing>[\s\S]*<\/p:timing>/.exec(xml)?.[0] ?? ''

	assert.equal([...xml.matchAll(/<p:timing>/g)].length, 1, 'a slide must have exactly one timing tree')
	// the media node is a sibling of mainSeq under tmRoot
	assert.match(timing, /<\/p:seq><p:video><p:cMediaNode>/, 'media node must follow mainSeq inside tmRoot')
	assert.match(timing, /repeatCount="indefinite"/, 'media loop lost')
	const ids = [...timing.matchAll(/<p:cTn id="(\d+)"/g)].map(match => match[1])
	assert.equal(new Set(ids).size, ids.length, 'animation and media nodes must not reuse node ids')
})

test('contract: invalid animation input is dropped with a warning', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let xml = ''
	try {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		slide.addText('bad preset', { x: 1, y: 1, w: 3, h: 1, animation: { type: 'explode' as unknown as 'fadeIn' } })
		slide.addText('bad trigger', { x: 1, y: 2, w: 3, h: 1, animation: { type: 'fadeIn', trigger: 'whenever' as unknown as 'onClick' } })
		slide.addText('bad direction', { x: 1, y: 3, w: 3, h: 1, animation: { type: 'wipeIn', direction: 'sideways' } })
		slide.addText('bad timings', { x: 1, y: 4, w: 3, h: 1, animation: { type: 'fadeIn', delay: -5, duration: NaN } })
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	} finally {
		console.warn = origWarn
	}

	assert.ok(warnings.some(w => w.includes('unknown animation "explode"')), 'unknown preset must warn')
	assert.ok(warnings.some(w => w.includes('animation trigger must be')), 'unknown trigger must warn')
	assert.ok(warnings.some(w => w.includes('does not accept direction "sideways"')), 'invalid direction must warn')

	// the two bad-preset/bad-trigger shapes are skipped; the other two still animate
	assert.equal([...xml.matchAll(/presetID=/g)].length, 2, 'only valid animations are emitted')
	assert.match(xml, /presetSubtype="1"/, 'invalid direction falls back to the preset default')
	assert.match(xml, /<p:stCondLst><p:cond delay="0"\/><\/p:stCondLst>/, 'negative delay falls back to 0')
	assert.match(xml, /dur="500"/, 'non-finite duration falls back to the default')
	assert.doesNotMatch(xml, /NaN/, 'no NaN attributes')
})
/**
 * Minimal stand-in for an EOT font: the only bytes PptxGenJS inspects are the magic number
 * `0x504C` at offset 34 (MS-EOT 2.1), so a 64-byte buffer with that field set is enough to
 * exercise the part/relationship/content-type plumbing without shipping a font file.
 */
function fakeEot (): Uint8Array {
	const bytes = new Uint8Array(64)
	bytes[34] = 0x4c
	bytes[35] = 0x50
	bytes[36] = 0x77 // marker so the test can prove the bytes reached the part verbatim
	return bytes
}

test('contract: embedded fonts add font parts, relationships, and a content type', async () => {
	const eot = fakeEot()
	const base64 = Buffer.from(eot).toString('base64')
	const pptx = new pptxgen()
	pptx.addFont({ fontFace: 'Custom Sans', data: base64 })
	pptx.addFont({ fontFace: 'Custom Sans', data: eot, style: 'bold' })
	pptx.addFont({ fontFace: 'Other & Co', data: eot.buffer, style: 'italic' })
	pptx.addSlide().addText('embedded', { x: 1, y: 1, w: 4, h: 1, fontFace: 'Custom Sans' })
	const fontZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)

	// package-level coherence: every part declared, every relationship resolvable
	await assertPptxPackageContracts(fontZip)

	assert.match(await readPart(fontZip, '[Content_Types].xml'), /<Default Extension="fntdata" ContentType="application\/x-fontdata"\/>/, 'fntdata content type missing')

	const rels = await readPart(fontZip, 'ppt/_rels/presentation.xml.rels')
	const fontRels = [...rels.matchAll(/<Relationship Id="(rId\d+)" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/font" Target="(fonts\/font\d+\.fntdata)"\/>/g)]
	assert.equal(fontRels.length, 3, 'expected one font relationship per registered style')
	assert.deepEqual(fontRels.map(rel => rel[2]), ['fonts/font1.fntdata', 'fonts/font2.fntdata', 'fonts/font3.fntdata'], 'font targets must be sequential')

	// the presentation lists one embeddedFont per typeface, with a child element per style
	const presentation = await readPart(fontZip, 'ppt/presentation.xml')
	const relIds = fontRels.map(rel => rel[1])
	assert.match(
		presentation,
		new RegExp(`<p:embeddedFontLst><p:embeddedFont><p:font typeface="Custom Sans"/><p:regular r:id="${relIds[0]}"/><p:bold r:id="${relIds[1]}"/></p:embeddedFont>` +
			`<p:embeddedFont><p:font typeface="Other &amp; Co"/><p:italic r:id="${relIds[2]}"/></p:embeddedFont></p:embeddedFontLst>`),
		'embeddedFontLst missing or malformed'
	)
	// CT_Presentation sequence: embeddedFontLst precedes defaultTextStyle
	assert.match(presentation, /<\/p:embeddedFontLst><p:defaultTextStyle>/, 'embeddedFontLst must precede defaultTextStyle')

	// the font bytes reach the part untouched
	const part = fontZip.file('ppt/fonts/font1.fntdata')
	assert.ok(part, 'font part missing')
	const stored = await part.async('uint8array')
	assert.equal(stored.length, eot.length, 'font part length changed')
	assert.equal(stored[36], 0x77, 'font bytes were altered')
})

test('contract: presentations without addFont are unchanged', async () => {
	const plain = new pptxgen()
	plain.addSlide().addText('no fonts', { x: 1, y: 1, w: 4, h: 1 })
	const plainZip = await JSZip.loadAsync((await plain.write({ outputType: 'nodebuffer' })) as Buffer)

	assert.equal(Object.keys(plainZip.files).filter(file => file.includes('fonts')).length, 0, 'no font parts or folder may be created')
	assert.doesNotMatch(await readPart(plainZip, '[Content_Types].xml'), /fntdata/, 'no fntdata content type when unused')
	assert.doesNotMatch(await readPart(plainZip, 'ppt/_rels/presentation.xml.rels'), /relationships\/font/, 'no font relationship when unused')
	assert.doesNotMatch(await readPart(plainZip, 'ppt/presentation.xml'), /embeddedFont/, 'no embeddedFontLst when unused')
})

test('contract: invalid addFont input is rejected before it reaches the package', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let zipNoFonts: JSZip
	try {
		const pptx = new pptxgen()
		pptx.addFont(undefined as unknown as { fontFace: string, data: string })
		pptx.addFont({ fontFace: '   ', data: Buffer.from(fakeEot()).toString('base64') })
		pptx.addFont({ fontFace: 'No Data', data: '' })
		pptx.addFont({ fontFace: 'Bad Style', data: fakeEot(), style: 'heavy' as unknown as 'bold' })
		// a raw TTF starts with the 0x00010000 version tag and has no EOT magic number
		const ttf = new Uint8Array(64)
		ttf[1] = 0x01
		pptx.addFont({ fontFace: 'Raw TTF', data: ttf })
		pptx.addSlide()
		zipNoFonts = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	} finally {
		console.warn = origWarn
	}

	assert.ok(warnings.some(w => w.includes('an object is required')), 'non-object must warn')
	assert.ok(warnings.some(w => w.includes('`fontFace` is required')), 'blank fontFace must warn')
	assert.ok(warnings.some(w => w.includes('`data` must be base64 or binary')), 'missing data must warn')
	assert.ok(warnings.some(w => w.includes('`style` must be')), 'invalid style must warn')
	assert.ok(warnings.some(w => w.includes('does not look like EOT data')), 'non-EOT data must warn')

	assert.equal(Object.keys(zipNoFonts.files).filter(file => file.includes('fonts')).length, 0, 'rejected fonts must not create parts')
	assert.doesNotMatch(await readPart(zipNoFonts, 'ppt/presentation.xml'), /embeddedFont/, 'rejected fonts must not be listed')
})

test('contract: re-registering a font style replaces it rather than duplicating the part', async () => {
	const first = fakeEot()
	const second = fakeEot()
	second[37] = 0x99
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let replaceZip: JSZip
	try {
		const pptx = new pptxgen()
		pptx.addFont({ fontFace: 'Custom Sans', data: first })
		pptx.addFont({ fontFace: 'Custom Sans', data: second })
		pptx.addSlide()
		replaceZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	} finally {
		console.warn = origWarn
	}

	assert.ok(warnings.some(w => w.includes('is already embedded')), 're-registration must warn')
	assert.equal(Object.keys(replaceZip.files).filter(file => file.endsWith('.fntdata')).length, 1, 'expected a single font part')
	const stored = await (replaceZip.file('ppt/fonts/font1.fntdata') as JSZip.JSZipObject).async('uint8array')
	assert.equal(stored[37], 0x99, 'the latest registration must win')
	await assertPptxPackageContracts(replaceZip)
})
test('contract: slide and notes guides are written with EMU positions and a required color', async () => {
	const pptx = new pptxgen()
	pptx.guides = [{ orientation: 'vert', position: 5 }, { orientation: 'horz', position: 3.75, color: 'FF0000' }]
	pptx.notesGuides = [{ orientation: 'horz', position: 1 }]
	pptx.addSection({ title: 'Intro' })
	pptx.addSlide({ sectionTitle: 'Intro' })
	const guideZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(guideZip)
	const xml = await readPart(guideZip, 'ppt/presentation.xml')

	assert.doesNotMatch(xml, /NaN|undefined/, 'guides must not leak invalid attribute values')
	// 5 inches = 4572000 EMU; `orient` is omitted for horizontal guides (the schema default)
	const slideGuides = new RegExp(
		'<p:ext uri="\\{EFAFB233-063F-42B5-8137-9DF3F51BA10A\\}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main">' +
		'<p15:guide orient="vert" pos="4572000"><p15:clr><a:srgbClr val="A4A3A4"/></p15:clr></p15:guide>' +
		'<p15:guide pos="3429000"><p15:clr><a:srgbClr val="FF0000"/></p15:clr></p15:guide>' +
		'</p15:sldGuideLst></p:ext>'
	)
	assert.match(xml, slideGuides, 'slide guide list missing or malformed')
	assert.match(xml, /<p:ext uri="\{2D200454-40CA-4A62-9FC3-DE9A4176ACB9\}"><p15:notesGuideLst xmlns:p15="[^"]+"><p15:guide pos="914400">/, 'notes guide list missing')

	// guides do not require sections, and sections do not force an empty guide list
	const noSections = new pptxgen()
	noSections.guides = [{ orientation: 'horz', position: 2 }]
	noSections.addSlide()
	const noSectionsXml = await readPart(await JSZip.loadAsync((await noSections.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/presentation.xml')
	assert.match(noSectionsXml, /<p15:sldGuideLst/, 'guides must be written without sections')
	assert.doesNotMatch(noSectionsXml, /sectionLst/, 'no sections were added')

	const sectionsOnly = new pptxgen()
	sectionsOnly.addSection({ title: 'Only' })
	sectionsOnly.addSlide({ sectionTitle: 'Only' })
	const sectionsOnlyXml = await readPart(await JSZip.loadAsync((await sectionsOnly.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/presentation.xml')
	assert.match(sectionsOnlyXml, /sectionLst/, 'sections missing')
	assert.doesNotMatch(sectionsOnlyXml, /GuideLst/, 'no guide list should be written when no guides are set')

	// nothing set at all -> no extLst
	const plain = new pptxgen()
	plain.addSlide()
	assert.doesNotMatch(await readPart(await JSZip.loadAsync((await plain.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/presentation.xml'), /p:extLst/, 'no extLst without sections or guides')
})

test('contract: section ids are stable across exports', async () => {
	const pptx = new pptxgen()
	pptx.addSection({ title: 'Intro' })
	pptx.addSlide({ sectionTitle: 'Intro' })

	const sectionId = async (): Promise<string | undefined> => {
		const zipOut = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
		return /<p14:section name="Intro" id="(\{[^"]+\})"/.exec(await readPart(zipOut, 'ppt/presentation.xml'))?.[1]
	}

	const first = await sectionId()
	assert.match(String(first), /^\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}$/, 'section id must be a braced GUID')
	// anything that references a section (ex: a section zoom) stores this GUID, so it must not change
	assert.equal(await sectionId(), first, 'section id changed between exports')
})

test('contract: invalid guides are dropped with a warning', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let xml = ''
	try {
		const pptx = new pptxgen()
		pptx.guides = [
			{ orientation: 'diagonal' as unknown as 'horz', position: 2 },
			{ orientation: 'horz', position: NaN },
			{ orientation: 'vert', position: -1 },
			{ orientation: 'vert', position: 4 },
		]
		pptx.addSlide()
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/presentation.xml')
	} finally {
		console.warn = origWarn
	}

	assert.equal(warnings.filter(w => w.includes('guide orientation must be')).length, 1, 'invalid orientation must warn')
	assert.equal(warnings.filter(w => w.includes('guide position must be a number')).length, 2, 'invalid positions must warn')
	assert.equal([...xml.matchAll(/<p15:guide[ >]/g)].length, 1, 'only the valid guide is written')
	assert.doesNotMatch(xml, /NaN/, 'no NaN attributes')

	// every guide invalid -> no guide list at all, rather than an empty one
	const allBad = new pptxgen()
	allBad.guides = [{ orientation: 'horz', position: NaN }]
	allBad.addSlide()
	const origWarn2 = console.warn
	console.warn = () => {}
	try {
		const allBadXml = await readPart(await JSZip.loadAsync((await allBad.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/presentation.xml')
		assert.doesNotMatch(allBadXml, /GuideLst/, 'no guide list when every guide was rejected')
	} finally {
		console.warn = origWarn2
	}
})
test('contract: table modIds are unique per slide and stable across exports', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addTable([['a']], { x: 1, y: 1, w: 3 })
	slide.addTable([['b']], { x: 1, y: 3, w: 3 })
	pptx.addSlide().addTable([['c']], { x: 1, y: 1, w: 3 })

	const modIds = async (): Promise<string[][]> => {
		const zipOut = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
		return await Promise.all([1, 2].map(async num => {
			const xml = await readPart(zipOut, `ppt/slides/slide${num}.xml`)
			return [...xml.matchAll(/<p14:modId xmlns:p14="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2010\/main" val="(\d+)"\/>/g)].map(match => match[1])
		}))
	}

	const [slide1, slide2] = await modIds()
	// MS-PPTX 2.3.1.19: each modId must be unique on the slide - a constant would break tracking
	assert.equal(slide1.length, 2, 'expected one modId per table')
	assert.equal(new Set(slide1).size, 2, `modIds must differ within a slide: ${slide1.join(',')}`)
	// derived from a base plus a per-slide index, so values are unique but reproducible
	// base + the shape's own index on the slide
	assert.deepEqual(slide1, ['1579011935', '1579011936'], 'unexpected modId sequence')
	assert.deepEqual(slide2, ['1579011935'], 'each slide numbers from its own shape indexes')
	slide1.concat(slide2).forEach(id => {
		assert.ok(Number(id) >= 1 && Number(id) <= 0xffffffff, `modId ${id} outside ST_UnsignedInt`)
	})

	// repeated exports must not renumber existing shapes
	assert.deepEqual(await modIds(), [slide1, slide2], 'modIds changed between exports')
})

test('contract: slide creationId is opt-in, accepts a caller value, and is stable', async () => {
	const pptx = new pptxgen()
	const auto = pptx.addSlide()
	auto.creationId = true
	const explicit = pptx.addSlide()
	explicit.creationId = 4242
	pptx.addSlide() // no creationId

	const zipOut = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(zipOut)
	const [xml1, xml2, xml3] = await Promise.all([1, 2, 3].map(async num => await readPart(zipOut, `ppt/slides/slide${num}.xml`)))

	// the extension belongs to `p:cSld`, immediately after the shape tree
	assert.match(
		xml1,
		/<\/p:spTree><p:extLst><p:ext uri="\{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E\}"><p14:creationId xmlns:p14="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2010\/main" val="\d+"\/><\/p:ext><\/p:extLst><\/p:cSld>/,
		'creationId extension missing or misplaced'
	)
	const generated = /<p14:creationId[^>]*val="(\d+)"/.exec(xml1)?.[1]
	assert.ok(Number(generated) >= 1 && Number(generated) <= 0xffffffff, 'generated creationId outside ST_UnsignedInt')
	// derived from the slide number, so it is unique per slide and reproducible
	assert.equal(generated, '2147483649', 'unexpected generated creationId')
	assert.match(xml2, /<p14:creationId xmlns:p14="[^"]+" val="4242"\/>/, 'caller-supplied creationId not used')
	assert.doesNotMatch(xml3, /creationId/, 'creationId must be opt-in per slide')

	// stable across exports
	const again = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	assert.equal(/<p14:creationId[^>]*val="(\d+)"/.exec(await readPart(again, 'ppt/slides/slide1.xml'))?.[1], generated, 'creationId changed between exports')
})

test('contract: invalid creationId values are dropped with a warning', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let xmls: string[] = []
	try {
		const pptx = new pptxgen()
		pptx.addSlide().creationId = 1.5
		pptx.addSlide().creationId = 0x100000000
		pptx.addSlide().creationId = -1
		pptx.addSlide().creationId = 'yes' as unknown as number
		const zipInvalid = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
		xmls = await Promise.all([1, 2, 3, 4].map(async num => await readPart(zipInvalid, `ppt/slides/slide${num}.xml`)))
	} finally {
		console.warn = origWarn
	}

	assert.equal(warnings.filter(w => w.includes('creationId must be an integer between')).length, 3, 'out-of-range values must warn')
	assert.equal(warnings.filter(w => w.includes('creationId must be `true` or an unsigned')).length, 1, 'a non-number must warn')
	xmls.forEach((xml, idx) => {
		assert.doesNotMatch(xml, /creationId/, `slide ${idx + 1} must not carry an invalid creationId`)
		assert.doesNotMatch(xml, /NaN/, 'no NaN attributes')
	})
})

test('contract: zoom objects target slides and sections through mc:AlternateContent', async () => {
	const pptx = new pptxgen()
	pptx.addSection({ title: 'Intro' })
	pptx.addSection({ title: 'Results' })
	const hub = pptx.addSlide({ sectionTitle: 'Intro' })
	hub.addZoom({ slideNumber: 3, x: 1, y: 1, w: 3, h: 2, transitionDur: 900 })
	hub.addSectionZoom({ sectionTitle: 'Results', x: 5, y: 1, w: 3, h: 2, returnToParent: false, showBg: false })
	hub.addSummaryZoom({ sectionTitles: ['Intro', 'Results'], x: 1, y: 4, w: 8, h: 3 })
	pptx.addSlide({ sectionTitle: 'Results' })
	pptx.addSlide({ sectionTitle: 'Results' })

	const zoomZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(zoomZip)
	const xml = await readPart(zoomZip, 'ppt/slides/slide1.xml')
	const presentation = await readPart(zoomZip, 'ppt/presentation.xml')

	assert.doesNotMatch(xml, /NaN|undefined/, 'zoom options must not leak invalid attribute values')
	assert.equal([...xml.matchAll(/<mc:AlternateContent/g)].length, 3, 'expected one AlternateContent per zoom')

	// each kind lives in its own namespace, bound to `p16` inside the Choice
	assert.match(xml, /<mc:Choice [^>]*xmlns:p16="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2016\/slidezoom"[^>]*Requires="p16"><p16:sldZm>/, 'slide zoom namespace/element wrong')
	assert.match(xml, /<mc:Choice [^>]*xmlns:p16="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2016\/sectionzoom"[^>]*Requires="p16"><p16:sectionZm>/, 'section zoom namespace/element wrong')
	assert.match(xml, /<mc:Choice [^>]*xmlns:p16="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2016\/summaryzoom"[^>]*Requires="p16"><p16:summaryZm>/, 'summary zoom namespace/element wrong')

	// a slide zoom targets the target slide's own `p:sldId`, not its ordinal
	const targetSldId = /<p:sldId id="(\d+)" r:id="rId4"\/>/.exec(presentation)?.[1]
	assert.ok(targetSldId, 'third slide id not found')
	assert.match(xml, new RegExp(`<p16:sldZmObj sldId="${targetSldId}">`), 'slide zoom must reference the target slide id')

	// section zooms must reuse the exact GUIDs written into p14:sectionLst
	const sectionIds = Object.fromEntries([...presentation.matchAll(/<p14:section name="([^"]+)" id="(\{[^"]+\})"/g)].map(match => [match[1], match[2]]))
	assert.ok(sectionIds.Intro && sectionIds.Results, 'section ids missing from presentation.xml')
	assert.match(xml, new RegExp(`<p16:sectionZmObj sectionId="${sectionIds.Results.replace(/[{}]/g, char => `\\${char}`)}">`), 'section zoom GUID does not match sectionLst')
	// a summary zoom carries one object per section, in order, plus the required layout choice
	const summary = /<p16:summaryZm>[\s\S]*?<\/p16:summaryZm>/.exec(xml)?.[0] ?? ''
	assert.deepEqual(
		[...summary.matchAll(/<p16:summaryZmObj sectionId="(\{[^"]+\})"/g)].map(match => match[1]),
		[sectionIds.Intro, sectionIds.Results],
		'summary zoom sections wrong or out of order'
	)
	assert.match(summary, /<p16:gridLayout\/><\/p16:summaryZm>/, 'CT_SummaryZoom requires a layout choice after its objects')

	// zoom behaviour flags and duration
	assert.match(xml, /returnToParent="1" showBg="1" imageType="preview" p14:transitionDur="900"/, 'slide zoom properties wrong')
	assert.match(xml, /returnToParent="0" showBg="0" imageType="preview"/, 'section zoom flags not honoured')
	// every zoom object needs its own id
	const zoomIds = [...xml.matchAll(/<p166:zmPr id="(\{[^"]+\})"/g)].map(match => match[1])
	assert.equal(zoomIds.length, 4, 'expected four zoom objects (1 slide, 1 section, 2 summary sections)')
	assert.equal(new Set(zoomIds).size, 4, 'zoom object ids must be unique')

	// fallbacks: a picture for slide/section, a group shape for summary (MS-PPTX 2.2.15)
	assert.equal([...xml.matchAll(/<mc:Fallback><p:pic>/g)].length, 2, 'slide and section zooms need a picture fallback')
	assert.equal([...xml.matchAll(/<mc:Fallback><p:grpSp>/g)].length, 1, 'a summary zoom needs a group-shape fallback')

	// each zoom's cover image is a real relationship used by both the Choice and the Fallback
	const rels = await readPart(zoomZip, 'ppt/slides/_rels/slide1.xml.rels')
	const coverRids = [...xml.matchAll(/<a:blip r:embed="(rId\d+)"\/>/g)].map(match => match[1])
	assert.ok(coverRids.length >= 4, 'cover fills missing')
	new Set(coverRids).forEach(rid => {
		assert.match(rels, new RegExp(`<Relationship Id="${rid}" Type="[^"]*\\/image"`), `${rid} has no image relationship`)
	})
})

test('contract: zoom targets that cannot be resolved are dropped with a warning', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let xml = ''
	try {
		const pptx = new pptxgen()
		pptx.addSection({ title: 'Real' })
		const slide = pptx.addSlide({ sectionTitle: 'Real' })
		slide.addZoom({ slideNumber: 99, x: 1, y: 1, w: 2, h: 1 })
		slide.addSectionZoom({ sectionTitle: 'Nope', x: 1, y: 2, w: 2, h: 1 })
		slide.addZoom({ slideNumber: 0, x: 1, y: 3, w: 2, h: 1 })
		slide.addSummaryZoom({ sectionTitles: [], x: 1, y: 4, w: 2, h: 1 })
		// this one resolves, so the slide is not empty
		slide.addSectionZoom({ sectionTitle: 'Real', x: 4, y: 1, w: 2, h: 1 })
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	} finally {
		console.warn = origWarn
	}

	assert.ok(warnings.some(w => w.includes('zoom target slide 99 does not exist')), 'missing slide must warn')
	assert.ok(warnings.some(w => w.includes('zoom target section "Nope" does not exist')), 'missing section must warn')
	assert.ok(warnings.some(w => w.includes('`slideNumber` must be a whole number >= 1')), 'invalid slide number must warn')
	assert.ok(warnings.some(w => w.includes('`sectionTitles` must list at least one section')), 'empty summary must warn')

	// only the resolvable zoom survives, and nothing half-written is emitted
	assert.equal([...xml.matchAll(/<mc:AlternateContent/g)].length, 1, 'only the resolvable zoom may be emitted')
	assert.match(xml, /<p16:sectionZmObj sectionId="\{[0-9a-f-]{36}\}">/, 'the resolvable section zoom is missing')
	assert.doesNotMatch(xml, /sldZm|summaryZm/, 'unresolved zooms must not be emitted')
})

test('contract: presentations without zooms are unchanged', async () => {
	const plain = new pptxgen()
	plain.addSlide().addText('no zooms', { x: 1, y: 1, w: 3, h: 1 })
	const plainZip = await JSZip.loadAsync((await plain.write({ outputType: 'nodebuffer' })) as Buffer)
	const xml = await readPart(plainZip, 'ppt/slides/slide1.xml')
	assert.doesNotMatch(xml, /mc:AlternateContent|p16:|p166:/, 'a slide without zooms must not gain zoom markup')
	assert.equal(Object.keys(plainZip.files).filter(file => /^ppt\/media\/.+/.test(file)).length, 0, 'no cover images may be created')
})

test('contract: comments add author and comment parts wired to the slide', async () => {
	const pptx = new pptxgen()
	pptx.commentAuthors = [{ name: 'Ada Lovelace', initials: 'AL', userId: 'ada@example.com', providerId: 'AD' }]
	const reviewed = pptx.addSlide()
	reviewed.addText('review me', { x: 1, y: 1, w: 4, h: 1 })
	reviewed.addComment({
		text: 'Check this figure',
		author: 'Ada Lovelace',
		x: 4,
		y: 2,
		created: '2026-08-20T09:00:00Z',
		replies: [{ text: 'Fixed', author: 'Grace Hopper', created: '2026-08-20T10:00:00Z' }],
	})
	reviewed.addComment({ text: 'Resolved thread', author: 'Grace Hopper', resolved: true, created: '2026-08-20T11:00:00Z' })
	pptx.addSlide().addText('no comments here', { x: 1, y: 1, w: 3, h: 1 })

	const cmZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	// validates that every part is declared and every relationship resolves
	await assertPptxPackageContracts(cmZip)

	// only the slide with comments gets a part
	assert.ok(cmZip.file('ppt/comments/commentSlide1.xml'), 'comment part missing')
	assert.equal(cmZip.file('ppt/comments/commentSlide2.xml'), null, 'a slide without comments must not get a part')

	// content types (MS-PPTX 2.1.5 / 2.1.6)
	const types = await readPart(cmZip, '[Content_Types].xml')
	assert.match(types, /<Override PartName="\/ppt\/authors\.xml" ContentType="application\/vnd\.ms-powerpoint\.authors\+xml"\/>/, 'authors content type missing')
	assert.match(types, /<Override PartName="\/ppt\/comments\/commentSlide1\.xml" ContentType="application\/vnd\.ms-powerpoint\.comments\+xml"\/>/, 'comments content type missing')

	// the author part is reached from the presentation; the comment part only from its slide
	assert.match(await readPart(cmZip, 'ppt/_rels/presentation.xml.rels'), /Type="http:\/\/schemas\.microsoft\.com\/office\/2018\/10\/relationships\/authors" Target="authors\.xml"/, 'authors relationship missing')
	const slideRels = await readPart(cmZip, 'ppt/slides/_rels/slide1.xml.rels')
	const commentRelId = /<Relationship Id="(rId\d+)" Type="http:\/\/schemas\.microsoft\.com\/office\/2018\/10\/relationships\/comments" Target="\.\.\/comments\/commentSlide1\.xml"\/>/.exec(slideRels)?.[1]
	assert.ok(commentRelId, 'slide comment relationship missing')

	// the `p188:commentRel` pointer must name that same relationship, and sit in the p:sld extLst
	const slideXml = await readPart(cmZip, 'ppt/slides/slide1.xml')
	assert.match(
		slideXml,
		new RegExp(`<p:extLst><p:ext uri="\\{6950BFC3-D8DA-4A85-94F7-54DA5524770B\\}"><p188:commentRel xmlns:p188="[^"]+" r:id="${commentRelId}"\\/><\\/p:ext><\\/p:extLst><\\/p:sld>$`),
		'commentRel pointer missing, misplaced, or pointing at the wrong relationship'
	)

	// authors: the declared one keeps its metadata, the reply-only one is added with derived initials
	const authors = await readPart(cmZip, 'ppt/authors.xml')
	assert.match(authors, /<p188:author id="(\{[0-9a-f-]{36}\})" name="Ada Lovelace" initials="AL" userId="ada@example\.com" providerId="AD"\/>/, 'declared author lost its metadata')
	assert.match(authors, /<p188:author id="\{[0-9a-f-]{36}\}" name="Grace Hopper" initials="GH" userId="" providerId="None"\/>/, 'author named only by a comment was not added')
	assert.equal([...authors.matchAll(/<p188:author /g)].length, 2, 'authors must be deduped')

	// comments: anchor moniker, position, thread, and text
	const comments = await readPart(cmZip, 'ppt/comments/commentSlide1.xml')
	const slideId = /<p:sldId id="(\d+)"/.exec(await readPart(cmZip, 'ppt/presentation.xml'))?.[1]
	assert.match(comments, new RegExp(`<pc:sldMkLst><pc:docMkLst><pc:docMk/></pc:docMkLst><pc:sldMk sldId="${slideId}"/></pc:sldMkLst>`), 'comment anchor moniker wrong')
	assert.match(comments, /<p188:pos x="3657600" y="1828800"\/>/, 'comment anchor position wrong (4in, 2in)')
	assert.match(comments, /<p188:replyLst><p188:reply [^>]*created="2026-08-20T10:00:00Z">.*?<a:t>Fixed<\/a:t>/, 'reply missing')
	assert.match(comments, /created="2026-08-20T11:00:00Z" status="resolved">/, 'resolved status missing')
	assert.match(comments, /<a:t>Check this figure<\/a:t>/, 'comment text missing')
	assert.equal([...comments.matchAll(/<p188:cm /g)].length, 2, 'expected two comments on the slide')

	// derived ids must be valid GUIDs - a mnemonic letter like `r` is not a hex digit
	const ids = [...authors.matchAll(/ id="(\{[^"]+\})"/g), ...comments.matchAll(/ id="(\{[^"]+\})"/g)].map(match => match[1])
	assert.ok(ids.length >= 5, 'ids not found')
	ids.forEach(id => {
		assert.match(id, /^\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}$/, `${id} is not a valid GUID`)
	})
})

test('contract: comments are reproducible when created timestamps are supplied', async () => {
	const build = (): pptxgen => {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		slide.addComment({ text: 'stable', author: 'Ada', created: '2026-08-20T09:00:00Z' })
		return pptx
	}
	const partOf = async (pptx: pptxgen, name: string): Promise<string> =>
		await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), name)

	assert.equal(await partOf(build(), 'ppt/comments/commentSlide1.xml'), await partOf(build(), 'ppt/comments/commentSlide1.xml'), 'comment part is not reproducible')
	assert.equal(await partOf(build(), 'ppt/authors.xml'), await partOf(build(), 'ppt/authors.xml'), 'author part is not reproducible')
})

test('contract: invalid comment input is rejected and unused comments cost nothing', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let comments = ''
	try {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		slide.addComment({ text: '   ', author: 'Ada' })
		slide.addComment({ text: 'no author', author: '' })
		// a half-specified anchor would place the comment at the slide origin
		slide.addComment({ text: 'partial anchor', author: 'Ada', x: 3, created: '2026-08-20T09:00:00Z' })
		comments = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/comments/commentSlide1.xml')
	} finally {
		console.warn = origWarn
	}

	assert.ok(warnings.some(w => w.includes('addComment: `text` is required')), 'blank text must warn')
	assert.ok(warnings.some(w => w.includes('addComment: `author` is required')), 'blank author must warn')
	assert.ok(warnings.some(w => w.includes('anchored comment needs both `x` and `y`')), 'partial anchor must warn')
	assert.equal([...comments.matchAll(/<p188:cm /g)].length, 1, 'only the valid comment may be emitted')
	assert.doesNotMatch(comments, /<p188:pos/, 'a partial anchor must not be written')
	assert.doesNotMatch(comments, /NaN|undefined/, 'no invalid attribute values')

	// nothing is added to a presentation that never calls addComment()
	const plain = new pptxgen()
	plain.addSlide().addText('quiet', { x: 1, y: 1, w: 2, h: 1 })
	const plainZip = await JSZip.loadAsync((await plain.write({ outputType: 'nodebuffer' })) as Buffer)
	assert.equal(Object.keys(plainZip.files).filter(file => /authors\.xml|comments\//.test(file)).length, 0, 'no comment parts or folder may be created')
	assert.doesNotMatch(await readPart(plainZip, '[Content_Types].xml'), /authors|comments/, 'no comment content types when unused')
	assert.doesNotMatch(await readPart(plainZip, 'ppt/slides/slide1.xml'), /commentRel/, 'no commentRel pointer when unused')
})

test('contract: unit-suffixed lengths reach the slide XML as EMU', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('Metrisch', { x: '2.54cm', y: '25.4mm', w: '72pt', h: '1in' })
	slide.addTable([['A', 'B']], { x: '2.54cm', y: 3, colW: ['2.54cm', '2.54cm'], rowH: '2.54cm' })
	const unitZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	const xml = await readPart(unitZip, 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /NaN/, 'no coordinate resolved to NaN')
	assert.match(xml, /<a:off x="914400" y="914400"\/><a:ext cx="914400" cy="914400"\/>/, 'cm/mm/pt/in all resolve to one inch')
	// colW/rowH stay typed `number`; the shared parse means an untyped caller gets the length, not a NaN in the XML
	assert.equal((xml.match(/<a:gridCol w="914400"\/>/g) ?? []).length, 2, 'colW tolerates suffixed lengths')
	assert.match(xml, /<a:tr h="914400">/, 'rowH tolerates suffixed lengths')
})
test('contract: multi-column text boxes emit numCol and spcCol', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('two columns of flowing text', { x: 1, y: 1, w: 6, h: 2, columns: 2, columnSpacing: 0.25 })
	slide.addText('three columns, default gap', { x: 1, y: 4, w: 6, h: 2, columns: 3 })
	slide.addText('single column stays plain', { x: 1, y: 6, w: 6, h: 1, columns: 1 })
	const colZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(colZip)
	const xml = await readPart(colZip, 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /NaN|undefined/, 'column options must not leak invalid values')
	// 0.25in = 228600 EMU; ECMA-376 21.1.2.1.1 puts numCol/spcCol before rtlCol
	assert.match(xml, /<a:bodyPr [^>]*numCol="2" spcCol="228600" rtlCol="0"/, 'two-column body properties missing')
	assert.match(xml, /<a:bodyPr [^>]*numCol="3" rtlCol="0"/, 'three-column body properties missing')
	// `numCol="1"` is the schema default, so writing it would be noise
	assert.equal([...xml.matchAll(/numCol="1"/g)].length, 0, 'a single column must not be written')
	assert.equal([...xml.matchAll(/numCol="/g)].length, 2, 'expected exactly two multi-column text boxes')
})

test('contract: invalid column options are dropped with a warning', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let xml = ''
	try {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		// ECMA-376 allows 1-16; anything else makes a:bodyPr unparseable
		slide.addText('too many', { x: 1, y: 1, w: 4, h: 1, columns: 17 })
		slide.addText('too few', { x: 1, y: 2, w: 4, h: 1, columns: 0 })
		slide.addText('not a number', { x: 1, y: 3, w: 4, h: 1, columns: 'two' as unknown as number })
		slide.addText('negative gap', { x: 1, y: 4, w: 4, h: 1, columns: 2, columnSpacing: -1 })
		slide.addText('gap without columns', { x: 1, y: 5, w: 4, h: 1, columnSpacing: 0.5 })
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	} finally {
		console.warn = origWarn
	}

	assert.equal(warnings.filter(w => w.includes('`columns` must be a whole number between 1 and 16')).length, 3, 'out-of-range columns must warn')
	assert.ok(warnings.some(w => w.includes('`columnSpacing` must be a number >= 0')), 'negative spacing must warn')
	assert.ok(warnings.some(w => w.includes('`columnSpacing` has no effect without `columns`')), 'orphan spacing must warn')

	assert.equal([...xml.matchAll(/numCol="/g)].length, 1, 'only the valid column count survives')
	assert.match(xml, /numCol="2" rtlCol="0"/, 'a rejected spacing must leave the column count intact')
	assert.doesNotMatch(xml, /spcCol=/, 'no invalid spacing may be written')
})

test('contract: text boxes without column options are unchanged', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addText('plain', { x: 1, y: 1, w: 4, h: 1 })
	const xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	assert.doesNotMatch(xml, /numCol|spcCol/, 'a text box without columns must not gain column attributes')
})
const HOVER_PNG = 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkAAIAAAoAAv/lxKUAAAAASUVORK5CYII='
const HOVER_WAV = 'audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQAAAAA='

test('contract: mouse-over actions use the right element for each host', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addShape(pptx.ShapeType.rect, {
		x: 1, y: 1, w: 2, h: 1,
		hyperlink: { url: 'https://example.com', tooltip: 'go' },
		hyperlinkHover: { slide: 2, tooltip: 'peek' },
	})
	slide.addImage({ data: HOVER_PNG, x: 4, y: 1, w: 1, h: 1, hyperlinkHover: { url: 'https://hover.test' } })
	slide.addText([{ text: 'link', options: { hyperlink: { url: 'https://a.test' }, hyperlinkHover: { slide: 2 } } }], { x: 1, y: 3, w: 3, h: 1 })
	pptx.addSlide()

	const hlZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(hlZip)
	const xml = await readPart(hlZip, 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /NaN|undefined/, 'hyperlink options must not leak invalid values')
	// DrawingML names the same concept differently by host: `hlinkHover` in p:cNvPr,
	// `hlinkMouseOver` in a:rPr. Swapping them yields markup PowerPoint cannot parse.
	assert.match(xml, /<a:hlinkClick [^>]*tooltip="go"[^>]*\/><a:hlinkHover r:id="rId\d+" action="ppaction:\/\/hlinksldjump" tooltip="peek"\/>/, 'shape hover must use a:hlinkHover, after the click link')
	assert.match(xml, /<a:hlinkHover r:id="rId\d+" invalidUrl="" action="" tgtFrame="" tooltip="" history="1"\/>/, 'image hover link missing')
	assert.match(xml, /<a:hlinkMouseOver r:id="rId\d+" action="ppaction:\/\/hlinksldjump" tooltip=""\/>/, 'text-run hover must use a:hlinkMouseOver')
	assert.equal([...xml.matchAll(/<a:hlinkMouseOver/g)].length, 1, 'exactly one run-level hover expected')
	assert.equal([...xml.matchAll(/<a:hlinkHover/g)].length, 2, 'exactly two shape-level hovers expected')
	// the run-level hover element must not appear on a shape, nor vice versa
	assert.doesNotMatch(xml, /<p:cNvPr[^>]*>(?:(?!<\/p:cNvPr>)[\s\S])*<a:hlinkMouseOver/, 'a:hlinkMouseOver must not appear in p:cNvPr')

	// every link resolves to a relationship
	const rels = await readPart(hlZip, 'ppt/slides/_rels/slide1.xml.rels')
	;[...xml.matchAll(/<a:hlink\w+ r:id="(rId\d+)"/g)].map(match => match[1]).forEach(rid => {
		assert.match(rels, new RegExp(`<Relationship Id="${rid}"`), `${rid} has no relationship`)
	})
})

test('contract: action sounds and click attributes are emitted', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText([{
		text: 'noisy',
		options: { hyperlink: { url: 'https://a.test', highlightClick: true, stopSoundsOnClick: true, sound: { data: HOVER_WAV, name: 'ding.wav' } } },
	}], { x: 1, y: 1, w: 3, h: 1 })
	const sndZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	// proves the wav part is declared with a content type and its relationship resolves
	await assertPptxPackageContracts(sndZip)
	const xml = await readPart(sndZip, 'ppt/slides/slide1.xml')

	assert.match(xml, /<a:hlinkClick [^>]*highlightClick="1" endSnd="1">/, 'click attributes must be settable')
	assert.match(xml, /<a:snd r:embed="rId\d+" name="ding\.wav"\/>/, 'action sound missing')
	assert.equal(Object.keys(sndZip.files).filter(file => /^ppt\/media\/.+\.wav$/.test(file)).length, 1, 'wav part missing')
	assert.match(await readPart(sndZip, '[Content_Types].xml'), /Extension="wav"/, 'wav content type missing')

	// both default to false in the schema, so an ordinary link writes neither
	const plain = new pptxgen()
	plain.addSlide().addText([{ text: 'quiet', options: { hyperlink: { url: 'https://a.test' } } }], { x: 1, y: 1, w: 3, h: 1 })
	const plainXml = await readPart(await JSZip.loadAsync((await plain.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	assert.doesNotMatch(plainXml, /highlightClick|endSnd|a:snd|hlinkHover|hlinkMouseOver/, 'a plain link must not gain hover, sound, or click attributes')
})

test('contract: invalid hover and sound input is dropped with a warning', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let xml = ''
	try {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, hyperlinkHover: {} })
		slide.addShape(pptx.ShapeType.rect, { x: 4, y: 1, w: 2, h: 1, hyperlinkHover: { url: 'https://b.test', sound: {} } })
		slide.addShape(pptx.ShapeType.rect, { x: 1, y: 3, w: 2, h: 1, hyperlinkHover: { url: 'https://c.test', sound: { data: 'not-base64' } } })
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	} finally {
		console.warn = origWarn
	}

	assert.ok(warnings.some(w => w.includes('hyperlink requires either `url` or `slide`')), 'a target-less link must warn')
	assert.ok(warnings.some(w => w.includes('`sound` requires `data` or `path`')), 'a sound without data must warn')
	assert.ok(warnings.some(w => w.includes('`sound.data` lacks a base64 header')), 'bad sound data must warn')

	assert.equal([...xml.matchAll(/<a:hlinkHover/g)].length, 2, 'links with a target survive; the target-less one does not')
	assert.doesNotMatch(xml, /<a:snd /, 'no invalid sound may be written')
	assert.doesNotMatch(xml, /r:id="rId0"/, 'no link may reference a non-existent relationship')
})
const FILL_PNG = 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkAAIAAAoAAv/lxKUAAAAASUVORK5CYII='

test('contract: pattern fills emit a:pattFill with both colors', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, fill: { type: 'pattern', pattern: { preset: 'diagCross', color: '0000FF', backColor: 'FFFF00' } } })
	// defaults: black on white, per ECMA-376
	slide.addShape(pptx.ShapeType.rect, { x: 4, y: 1, w: 2, h: 1, fill: { type: 'pattern', pattern: { preset: 'smGrid' } } })
	slide.addTable([[{ text: 'p', options: { fill: { type: 'pattern', pattern: { preset: 'wave', color: 'FF0000' } } } }]], { x: 1, y: 3, w: 4 })
	const patZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(patZip)
	const xml = await readPart(patZip, 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /NaN|undefined/, 'pattern options must not leak invalid values')
	assert.match(xml, /<a:pattFill prst="diagCross"><a:fgClr><a:srgbClr val="0000FF"\/><\/a:fgClr><a:bgClr><a:srgbClr val="FFFF00"\/><\/a:bgClr><\/a:pattFill>/, 'pattern fill missing')
	assert.match(xml, /<a:pattFill prst="smGrid"><a:fgClr><a:srgbClr val="000000"\/><\/a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"\/><\/a:bgClr><\/a:pattFill>/, 'pattern defaults wrong')
	// a table cell fill needs the whole fill object, not just a color
	assert.match(xml, /<a:tcPr[\s\S]*?<a:pattFill prst="wave">/, 'table cell pattern fill missing')
	assert.equal([...xml.matchAll(/<a:pattFill /g)].length, 3, 'expected three pattern fills')
})

test('contract: picture fills emit a:blipFill wired to an image relationship', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, fill: { type: 'image', image: { data: FILL_PNG } } })
	slide.addShape(pptx.ShapeType.rect, { x: 4, y: 1, w: 2, h: 1, fill: { type: 'image', image: { data: FILL_PNG, sizing: 'tile', scale: 50, alignment: 'ctr', rotateWithShape: false } } })
	slide.addTable([[{ text: 'i', options: { fill: { type: 'image', image: { data: FILL_PNG } } } }]], { x: 1, y: 3, w: 4 })
	const picZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	// proves each r:embed resolves and every media part is declared
	await assertPptxPackageContracts(picZip)
	const xml = await readPart(picZip, 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /NaN|undefined/, 'picture-fill options must not leak invalid values')
	assert.match(xml, /<a:blipFill rotWithShape="1"><a:blip r:embed="rId\d+"\/><a:stretch><a:fillRect\/><\/a:stretch><\/a:blipFill>/, 'stretch picture fill missing')
	assert.match(xml, /<a:blipFill rotWithShape="0"><a:blip r:embed="rId\d+"\/><a:tile tx="0" ty="0" sx="50000" sy="50000" flip="none" algn="ctr"\/><\/a:blipFill>/, 'tiled picture fill missing')
	assert.match(xml, /<a:tcPr[\s\S]*?<a:blipFill /, 'table cell picture fill missing')

	// each fill's r:embed must name a real image relationship, and the bytes must be in the package
	const rels = await readPart(picZip, 'ppt/slides/_rels/slide1.xml.rels')
	const embeds = [...xml.matchAll(/<a:blip r:embed="(rId\d+)"\/>/g)].map(match => match[1])
	assert.equal(embeds.length, 3, 'expected one blip per picture fill')
	embeds.forEach(rid => {
		assert.match(rels, new RegExp(`<Relationship Id="${rid}" Type="[^"]*\\/image" Target="\\.\\./media/[^"]+"\\/>`), `${rid} has no image relationship`)
	})
	assert.equal(Object.keys(picZip.files).filter(file => /^ppt\/media\/.+/.test(file)).length, 3, 'image parts missing')
})

test('contract: invalid pattern and picture fills degrade instead of writing broken XML', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let xml = ''
	try {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		// `prst` is an enum, so an unknown value would make the element unparseable
		slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, fill: { type: 'pattern', pattern: { preset: 'tartan' as unknown as 'wave' } } })
		slide.addShape(pptx.ShapeType.rect, { x: 4, y: 1, w: 2, h: 1, fill: { type: 'pattern' } })
		// a picture fill with no image would leave `a:blip` dangling, which reads as damage
		slide.addShape(pptx.ShapeType.rect, { x: 1, y: 3, w: 2, h: 1, fill: { type: 'image', image: {} } })
		slide.addShape(pptx.ShapeType.rect, { x: 4, y: 3, w: 2, h: 1, fill: { type: 'image', image: { data: 'not-base64' } } })
		slide.addShape(pptx.ShapeType.rect, { x: 1, y: 5, w: 2, h: 1, fill: { type: 'image', image: { data: FILL_PNG, sizing: 'tile', alignment: 'middle' as unknown as 'ctr' } } })
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	} finally {
		console.warn = origWarn
	}

	assert.ok(warnings.some(w => w.includes('unknown fill pattern "tartan"')), 'unknown preset must warn')
	assert.ok(warnings.some(w => w.includes('requires `fill.pattern.preset`')), 'missing preset must warn')
	assert.ok(warnings.some(w => w.includes('requires `fill.image.data` or `fill.image.path`')), 'missing image must warn')
	assert.ok(warnings.some(w => w.includes('lacks a base64 header')), 'bad image data must warn')
	assert.ok(warnings.some(w => w.includes('unknown tile alignment "middle"')), 'bad alignment must warn')

	assert.match(xml, /<a:pattFill prst="pct50">/, 'an unknown preset falls back to pct50')
	assert.equal([...xml.matchAll(/<a:blip r:embed=/g)].length, 1, 'only the valid picture fill may emit a blip')
	assert.match(xml, /algn="tl"/, 'an unknown tile alignment falls back to tl')
	assert.doesNotMatch(xml, /r:embed="rId0"/, 'no fill may reference a non-existent relationship')
})

test('contract: existing fill types are unaffected by pattern and picture support', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, fill: { color: 'FF0000', transparency: 50 } })
	slide.addShape(pptx.ShapeType.rect, { x: 4, y: 1, w: 2, h: 1, fill: { type: 'gradient', gradient: { stops: [{ color: 'FF0000', position: 0 }, { color: '0000FF', position: 100 }] } } })
	slide.addTable([[{ text: 'c', options: { fill: { color: '112233' } } }]], { x: 1, y: 3, w: 4 })
	const xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /pattFill|blipFill/, 'solid and gradient fills must not gain pattern or picture markup')
	assert.match(xml, /<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"\/><\/a:srgbClr><\/a:solidFill>/, 'solid fill changed')
	assert.match(xml, /<a:gradFill rotWithShape="1">/, 'gradient fill changed')
	assert.match(xml, /<a:solidFill><a:srgbClr val="112233"\/><\/a:solidFill>/, 'table cell solid fill changed')
})
const CP_PNG = 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkAAIAAAoAAv/lxKUAAAAASUVORK5CYII='
const CP_INK = '<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:trace>0 0, 10 10</inkml:trace></inkml:ink>'
const CP_INK_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'

test('contract: content parts embed a payload with the fallback the spec requires', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addContentPart({ data: CP_INK, contentType: 'application/inkml+xml', relationshipType: CP_INK_REL, ink: true, cover: CP_PNG, x: 1, y: 1, w: 3, h: 2, fileName: 'ink1.xml' })
	slide.addContentPart({ data: '<custom/>', contentType: 'application/xml', relationshipType: 'http://example.com/rel/custom', x: 5, y: 1, w: 2, h: 1 })

	const cpZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	// every part declared, every relationship resolvable, no duplicate ids
	await assertPptxPackageContracts(cpZip)
	const xml = await readPart(cpZip, 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /NaN|undefined/, 'content-part options must not leak invalid values')
	assert.equal([...xml.matchAll(/<p14:contentPart /g)].length, 2, 'expected two content parts')
	assert.match(xml, /<mc:Choice xmlns:p14="[^"]+" Requires="p14"><p14:contentPart r:id="rId\d+" p14:bwMode="auto"><p14:nvContentPartPr><p14:cNvPr id="\d+" name="Ink 1"\/><p14:cNvContentPartPr\/><p14:nvPr\/><\/p14:nvContentPartPr><p14:xfrm>/, 'content-part choice malformed')

	// MS-PPTX 2.2.3.1 requires a picture fallback for ink, 2.2.3 a shape fallback otherwise
	assert.match(xml, /<mc:Fallback><p:pic>[\s\S]*?<a:blip r:embed="rId\d+"\/>/, 'ink must fall back to a picture showing its preview')
	assert.match(xml, /<mc:Fallback><p:sp>[\s\S]*?<a:noFill\/>/, 'a non-ink content part must fall back to a shape')

	// the payload lands next to its slide with the caller's content type declared
	assert.ok(cpZip.file('ppt/slides/contentParts/ink1.xml'), 'ink payload part missing')
	assert.equal(await readPart(cpZip, 'ppt/slides/contentParts/ink1.xml'), CP_INK, 'payload must be written verbatim')
	const types = await readPart(cpZip, '[Content_Types].xml')
	assert.match(types, /<Override PartName="\/ppt\/slides\/contentParts\/ink1\.xml" ContentType="application\/inkml\+xml"\/>/, 'caller content type not declared')
	assert.match(types, /ContentType="application\/xml"\/>/, 'second content type not declared')

	// the caller's relationship type is used verbatim, and no rId is reused
	const rels = await readPart(cpZip, 'ppt/slides/_rels/slide1.xml.rels')
	assert.ok(rels.includes(`<Relationship Id="rId1" Type="${CP_INK_REL}" Target="contentParts/ink1.xml"/>`), 'ink relationship missing or wrong type')
	assert.match(rels, /Type="http:\/\/example\.com\/rel\/custom" Target="contentParts\/contentPart2\.xml"/, 'custom relationship type not honoured')
	const ids = [...rels.matchAll(/<Relationship Id="(rId\d+)"/g)].map(match => match[1])
	assert.equal(new Set(ids).size, ids.length, `duplicate relationship ids: ${ids.join(',')}`)
})

test('contract: content parts without a verifiable package contract are refused', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let zipOut: JSZip
	try {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		slide.addContentPart({ data: '', contentType: 'application/xml', relationshipType: 'http://x.test/r' })
		slide.addContentPart({ data: '<a/>', contentType: '', relationshipType: 'http://x.test/r' })
		slide.addContentPart({ data: '<a/>', contentType: 'application/xml', relationshipType: '' })
		// ink without a raster preview would fall back to nothing
		slide.addContentPart({ data: CP_INK, contentType: 'application/inkml+xml', relationshipType: CP_INK_REL, ink: true })
		zipOut = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	} finally {
		console.warn = origWarn
	}

	assert.ok(warnings.some(w => w.includes('`data` is required')), 'missing payload must warn')
	assert.ok(warnings.some(w => w.includes('`contentType` is required')), 'missing content type must warn')
	assert.ok(warnings.some(w => w.includes('`relationshipType` is required')), 'missing relationship type must warn')
	assert.ok(warnings.some(w => w.includes('ink requires a base64 `cover` image')), 'ink without a preview must warn')

	const xml = await readPart(zipOut, 'ppt/slides/slide1.xml')
	assert.doesNotMatch(xml, /p14:contentPart/, 'no content part may be emitted from rejected input')
	assert.equal(Object.keys(zipOut.files).filter(file => /contentParts\/.+/.test(file)).length, 0, 'no payload parts may be written')
})

test('contract: presentations without content parts are unchanged', async () => {
	const plain = new pptxgen()
	plain.addSlide().addText('none', { x: 1, y: 1, w: 3, h: 1 })
	const plainZip = await JSZip.loadAsync((await plain.write({ outputType: 'nodebuffer' })) as Buffer)
	assert.equal(Object.keys(plainZip.files).filter(file => file.includes('contentParts')).length, 0, 'no content-part folder may be created')
	assert.doesNotMatch(await readPart(plainZip, 'ppt/slides/slide1.xml'), /contentPart/, 'a slide without content parts must be unchanged')
})

test('contract: the remaining a:bodyPr, a:pPr, and a:rPr attributes are reachable', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('all the attributes', {
		x: 1, y: 1, w: 4, h: 2,
		// a:bodyPr (ECMA-376 21.1.2.1.1)
		upright: true, textRotate: 15, anchorCenter: true, spaceFirstLastPara: true,
		compatLineSpacing: true, forceAntiAlias: true, horizontalOverflow: 'clip', verticalOverflow: 'ellipsis',
		// a:pPr (21.1.2.2.7)
		marginRight: 0.25, defaultTabSize: 0.5, fontAlign: 'ctr',
		eastAsianLineBreak: false, latinLineBreak: false, hangingPunctuation: false,
		// a:rPr (21.1.2.3.9)
		capitalization: 'small', normalizeHeight: true, noProof: true, dirty: true,
		symbolFontFace: 'Wingdings', latinFontFace: 'Georgia', eastAsianFontFace: 'MS Gothic', complexScriptFontFace: 'Arial',
		underlineLine: { width: 1.5, color: 'FF0000', dashType: 'dash' },
	})
	const attrZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(attrZip)
	const xml = await readPart(attrZip, 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /NaN|undefined/, 'attributes must not leak invalid values')
	// 15deg -> 900000 (60000ths); 0.25in -> 228600 EMU; 0.5in -> 457200 EMU
	assert.match(xml, /<a:bodyPr wrap="square" upright="1" rot="900000" anchorCtr="1" spcFirstLastPara="1" compatLnSpc="1" forceAA="1" horzOverflow="clip" vertOverflow="ellipsis"/, 'bodyPr attributes missing')
	assert.match(xml, /<a:pPr marR="228600" defTabSz="457200" fontAlgn="ctr" eaLnBrk="0" latinLnBrk="0" hangingPunct="0"/, 'pPr attributes missing')
	assert.match(xml, /<a:rPr lang="en-US" cap="small" normalizeH="1" noProof="1" dirty="1">/, 'rPr attributes missing')

	// per-script typefaces can differ; `a:sym` follows them
	assert.match(xml, /<a:latin typeface="Georgia"[^>]*\/><a:ea typeface="MS Gothic"[^>]*\/><a:cs typeface="Arial"[^>]*\/><a:sym typeface="Wingdings"\/>/, 'per-script fonts or sym missing')
	// a:uLn carries its own line properties, distinct from the underline colour
	assert.match(xml, /<a:uLn w="19050"><a:solidFill><a:srgbClr val="FF0000"\/><\/a:solidFill><a:prstDash val="dash"\/><\/a:uLn>/, 'uLn missing')
})

test('contract: rPr children follow the schema sequence', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addText('ordered', {
		x: 1, y: 1, w: 4, h: 1,
		outline: { size: 1, color: '000000' },
		color: 'FF0000',
		glow: { size: 4, color: 'FFFF00', opacity: 0.5 },
		highlight: '00FF00',
		underline: { style: 'sng', color: '0000FF' },
		underlineLine: 'text',
		fontFace: 'Arial',
		symbolFontFace: 'Wingdings',
	})
	const xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')

	// CT_TextCharacterProperties order: ln, fill, effect, highlight, uLn, uFill, latin, ea, cs, sym.
	// The glow effect used to be written last, after a:uFill, which is invalid against the schema.
	const order = ['a:ln', 'a:solidFill', 'a:effectLst', 'a:highlight', 'a:uLnTx', 'a:uFill', 'a:latin', 'a:ea', 'a:cs', 'a:sym']
	const rPr = /<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/.exec(xml)?.[0] ?? ''
	assert.ok(rPr, 'rPr not found')
	const positions = order.map(tag => ({ tag, at: rPr.indexOf(`<${tag}`) }))
	positions.forEach(entry => { assert.ok(entry.at > -1, `${entry.tag} missing from rPr`) })
	positions.reduce((prev, entry) => {
		assert.ok(entry.at > prev.at, `${entry.tag} must follow ${prev.tag} in CT_TextCharacterProperties`)
		return entry
	})
})

test('contract: text without the new attributes is byte-identical', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addText('plain', { x: 1, y: 1, w: 3, h: 1 })
	const xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')

	assert.match(xml, /<a:bodyPr wrap="square" rtlCol="0" anchor="ctr">/, 'default bodyPr changed')
	assert.match(xml, /<a:rPr lang="en-US" dirty="0">/, 'default rPr changed - `dirty` must stay 0')
	assert.doesNotMatch(xml, /upright|rot=|anchorCtr|spcFirstLastPara|compatLnSpc|forceAA|horzOverflow|vertOverflow/, 'no bodyPr attribute may appear unasked')
	assert.doesNotMatch(xml, /marR=|defTabSz=|fontAlgn=|eaLnBrk=|latinLnBrk=|hangingPunct=/, 'no pPr attribute may appear unasked')
	assert.doesNotMatch(xml, /cap=|normalizeH|noProof|a:uLn|a:sym/, 'no rPr attribute may appear unasked')
})
