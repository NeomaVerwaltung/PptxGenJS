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
