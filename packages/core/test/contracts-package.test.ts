/**
 * Package-level contracts: parts, relationships, content types, and the presentation-wide
 * features that reach them - fonts, guides, sections, comments, content parts, presentation props.
 *
 * Unlike golden XML snapshots, these checks document the OOXML that matters and allow harmless
 * serializer changes without regenerating fixture files.
 */
import { test, before } from 'node:test'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import pptxgen from '../src/pptxgen'
import { assertEmbeddedXlsxContracts, assertPptxPackageContracts, readPart } from './pptx-contracts'
import { CUSTOM_XML_REL, SAMPLE_INK, SAMPLE_PNG, buildContractFixture, captureWarnings, fakeEot } from './fixtures'

let zip: JSZip

before(async () => {
	zip = await buildContractFixture()
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
	const warnings = await captureWarnings(async () => {
		const bogus = new pptxgen()
			;(bogus as unknown as { chartTrackingRefBased: unknown }).chartTrackingRefBased = 'no'
		assert.equal(bogus.chartTrackingRefBased, true, 'a non-boolean must leave the default in place')
	})
	assert.ok(warnings.some(warning => warning.includes('chartTrackingRefBased must be a boolean')), 'a non-boolean must warn')
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
	let xml = ''
	const warnings = await captureWarnings(async () => {
		const pptx = new pptxgen()
		pptx.slideShow = { mode: 'cinema' as unknown as 'kiosk' }
		pptx.defaultImageDpi = -1
		;(pptx as unknown as { readonlyRecommended: unknown }).readonlyRecommended = 'yes'
		pptx.addSlide()
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/presProps.xml')
	})

	assert.ok(warnings.some(w => w.includes('slideShow.mode must be')), 'invalid mode must warn')
	assert.ok(warnings.some(w => w.includes('defaultImageDpi must be a number >= 0')), 'negative dpi must warn')
	assert.match(xml, /<p:showPr><p:present\/><\/p:showPr>/, 'invalid mode falls back to present')
	assert.doesNotMatch(xml, /defaultImageDpi/, 'negative dpi must not be emitted')
	assert.doesNotMatch(xml, /readonlyRecommended/, 'a non-boolean must not enable the extension')
	assert.doesNotMatch(xml, /NaN/, 'no NaN attributes')
})

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
	let zipNoFonts: JSZip
	const warnings = await captureWarnings(async () => {
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
	})

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
	let replaceZip: JSZip
	const warnings = await captureWarnings(async () => {
		const pptx = new pptxgen()
		pptx.addFont({ fontFace: 'Custom Sans', data: first })
		pptx.addFont({ fontFace: 'Custom Sans', data: second })
		pptx.addSlide()
		replaceZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	})

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
	let xml = ''
	const warnings = await captureWarnings(async () => {
		const pptx = new pptxgen()
		pptx.guides = [
			{ orientation: 'diagonal' as unknown as 'horz', position: 2 },
			{ orientation: 'horz', position: NaN },
			{ orientation: 'vert', position: -1 },
			{ orientation: 'vert', position: 4 },
		]
		pptx.addSlide()
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/presentation.xml')
	})

	assert.equal(warnings.filter(w => w.includes('guide orientation must be')).length, 1, 'invalid orientation must warn')
	assert.equal(warnings.filter(w => w.includes('guide position must be a number')).length, 2, 'invalid positions must warn')
	assert.equal([...xml.matchAll(/<p15:guide[ >]/g)].length, 1, 'only the valid guide is written')
	assert.doesNotMatch(xml, /NaN/, 'no NaN attributes')

	// every guide invalid -> no guide list at all, rather than an empty one
	const allBad = new pptxgen()
	allBad.guides = [{ orientation: 'horz', position: NaN }]
	allBad.addSlide()
	await captureWarnings(async () => {
		const allBadXml = await readPart(await JSZip.loadAsync((await allBad.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/presentation.xml')
		assert.doesNotMatch(allBadXml, /GuideLst/, 'no guide list when every guide was rejected')
	})
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
	let xmls: string[] = []
	const warnings = await captureWarnings(async () => {
		const pptx = new pptxgen()
		pptx.addSlide().creationId = 1.5
		pptx.addSlide().creationId = 0x100000000
		pptx.addSlide().creationId = -1
		pptx.addSlide().creationId = 'yes' as unknown as number
		const zipInvalid = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
		xmls = await Promise.all([1, 2, 3, 4].map(async num => await readPart(zipInvalid, `ppt/slides/slide${num}.xml`)))
	})

	assert.equal(warnings.filter(w => w.includes('creationId must be an integer between')).length, 3, 'out-of-range values must warn')
	assert.equal(warnings.filter(w => w.includes('creationId must be `true` or an unsigned')).length, 1, 'a non-number must warn')
	xmls.forEach((xml, idx) => {
		assert.doesNotMatch(xml, /creationId/, `slide ${idx + 1} must not carry an invalid creationId`)
		assert.doesNotMatch(xml, /NaN/, 'no NaN attributes')
	})
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
	let comments = ''
	const warnings = await captureWarnings(async () => {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		slide.addComment({ text: '   ', author: 'Ada' })
		slide.addComment({ text: 'no author', author: '' })
		// a half-specified anchor would place the comment at the slide origin
		slide.addComment({ text: 'partial anchor', author: 'Ada', x: 3, created: '2026-08-20T09:00:00Z' })
		comments = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/comments/commentSlide1.xml')
	})

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

test('contract: content parts embed a payload with the fallback the spec requires', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addContentPart({ data: SAMPLE_INK, contentType: 'application/inkml+xml', relationshipType: CUSTOM_XML_REL, ink: true, cover: SAMPLE_PNG, x: 1, y: 1, w: 3, h: 2, fileName: 'ink1.xml' })
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
	assert.equal(await readPart(cpZip, 'ppt/slides/contentParts/ink1.xml'), SAMPLE_INK, 'payload must be written verbatim')
	const types = await readPart(cpZip, '[Content_Types].xml')
	assert.match(types, /<Override PartName="\/ppt\/slides\/contentParts\/ink1\.xml" ContentType="application\/inkml\+xml"\/>/, 'caller content type not declared')
	assert.match(types, /ContentType="application\/xml"\/>/, 'second content type not declared')

	// the caller's relationship type is used verbatim, and no rId is reused
	const rels = await readPart(cpZip, 'ppt/slides/_rels/slide1.xml.rels')
	assert.ok(rels.includes(`<Relationship Id="rId1" Type="${CUSTOM_XML_REL}" Target="contentParts/ink1.xml"/>`), 'ink relationship missing or wrong type')
	assert.match(rels, /Type="http:\/\/example\.com\/rel\/custom" Target="contentParts\/contentPart2\.xml"/, 'custom relationship type not honoured')
	const ids = [...rels.matchAll(/<Relationship Id="(rId\d+)"/g)].map(match => match[1])
	assert.equal(new Set(ids).size, ids.length, `duplicate relationship ids: ${ids.join(',')}`)
})

test('contract: content parts without a verifiable package contract are refused', async () => {
	let zipOut: JSZip
	const warnings = await captureWarnings(async () => {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		slide.addContentPart({ data: '', contentType: 'application/xml', relationshipType: 'http://x.test/r' })
		slide.addContentPart({ data: '<a/>', contentType: '', relationshipType: 'http://x.test/r' })
		slide.addContentPart({ data: '<a/>', contentType: 'application/xml', relationshipType: '' })
		// ink without a raster preview would fall back to nothing
		slide.addContentPart({ data: SAMPLE_INK, contentType: 'application/inkml+xml', relationshipType: CUSTOM_XML_REL, ink: true })
		zipOut = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	})

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
