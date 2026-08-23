/**
 * Slide-level contracts: transitions, animations, media playback, and zoom objects.
 *
 * Unlike golden XML snapshots, these checks document the OOXML that matters and allow harmless
 * serializer changes without regenerating fixture files.
 */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import pptxgen from '../src/pptxgen'
import { captureWarnings } from './fixtures'
import { assertPptxPackageContracts, readPart } from './pptx-contracts'

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
	let xml = ''
	const warnings = await captureWarnings(async () => {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		// fullScreen is a `p:video` attribute - meaningless for audio
		slide.addMedia({ type: 'audio', data: 'audio/mp3;base64,QQ==', x: 1, y: 1, w: 2, h: 2, autoplay: true, fullScreen: true })
		// online videos are played by the embed, not the timing tree
		slide.addMedia({ type: 'online', link: 'https://www.youtube.com/embed/Dph6ynRVyUc', x: 4, y: 1, w: 3, h: 2, autoplay: true, loop: true })
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	})

	assert.equal(warnings.filter(w => w.includes('`fullScreen` applies to `type:"video"` only')).length, 1, 'audio fullScreen must warn')
	assert.equal(warnings.filter(w => w.includes('not supported for `type:"online"`')).length, 1, 'online playback options must warn')
	assert.doesNotMatch(xml, /fullScrn/, 'fullScreen must not be emitted for audio')
	assert.equal([...xml.matchAll(/<p:audio[ >]/g)].length, 1, 'the audio node is still emitted for autoplay')
	assert.equal([...xml.matchAll(/<p:video[ >]/g)].length, 0, 'online video must not get a timing node')
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
	let xmls: string[] = []
	const warnings = await captureWarnings(async () => {
		const pptx = new pptxgen()
		pptx.addSlide({ transition: { type: 'nonsense' as unknown as 'fade' } })
		pptx.addSlide({ transition: { type: 'wipe', direction: 'in' } }) // wipe takes l/r/u/d only
		pptx.addSlide({ transition: { type: 'wheel', spokes: 5 as unknown as 4 } })
		pptx.addSlide({ transition: { type: 'push', speed: 'turbo' as unknown as 'fast', duration: NaN, advTm: NaN } })
		const zipInvalid = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
		xmls = await Promise.all([1, 2, 3, 4].map(async num => await readPart(zipInvalid, `ppt/slides/slide${num}.xml`)))
	})

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
	let xml = ''
	const warnings = await captureWarnings(async () => {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		slide.addText('bad preset', { x: 1, y: 1, w: 3, h: 1, animation: { type: 'explode' as unknown as 'fadeIn' } })
		slide.addText('bad trigger', { x: 1, y: 2, w: 3, h: 1, animation: { type: 'fadeIn', trigger: 'whenever' as unknown as 'onClick' } })
		slide.addText('bad direction', { x: 1, y: 3, w: 3, h: 1, animation: { type: 'wipeIn', direction: 'sideways' } })
		slide.addText('bad timings', { x: 1, y: 4, w: 3, h: 1, animation: { type: 'fadeIn', delay: -5, duration: NaN } })
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	})

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
	let xml = ''
	const warnings = await captureWarnings(async () => {
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
	})

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
