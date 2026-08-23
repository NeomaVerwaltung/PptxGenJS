/**
 * Shape contracts: fills, editing locks, non-visual properties, and mouse-over actions.
 *
 * Unlike golden XML snapshots, these checks document the OOXML that matters and allow harmless
 * serializer changes without regenerating fixture files.
 */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import pptxgen from '../src/pptxgen'
import { assertPptxPackageContracts, readPart } from './pptx-contracts'
import { SAMPLE_PNG, SAMPLE_WAV, captureWarnings } from './fixtures'

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

test('contract: mouse-over actions use the right element for each host', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addShape(pptx.ShapeType.rect, {
		x: 1, y: 1, w: 2, h: 1,
		hyperlink: { url: 'https://example.com', tooltip: 'go' },
		hyperlinkHover: { slide: 2, tooltip: 'peek' },
	})
	slide.addImage({ data: SAMPLE_PNG, x: 4, y: 1, w: 1, h: 1, hyperlinkHover: { url: 'https://hover.test' } })
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
	assert.doesNotMatch(xml, /<p:cNvPr[^>]*[^/]>(?:(?!<\/p:cNvPr>)[\s\S])*<a:hlinkMouseOver/, 'a:hlinkMouseOver must not appear in p:cNvPr')
	// and the run-level element does belong to a:rPr
	assert.match(xml, /<a:rPr[^>]*>(?:(?!<\/a:rPr>)[\s\S])*<a:hlinkMouseOver/, 'a:hlinkMouseOver must sit inside a:rPr')

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
		options: { hyperlink: { url: 'https://a.test', highlightClick: true, stopSoundsOnClick: true, sound: { data: SAMPLE_WAV, name: 'ding.wav' } } },
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
	let xml = ''
	const warnings = await captureWarnings(async () => {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, hyperlinkHover: {} })
		slide.addShape(pptx.ShapeType.rect, { x: 4, y: 1, w: 2, h: 1, hyperlinkHover: { url: 'https://b.test', sound: {} } })
		slide.addShape(pptx.ShapeType.rect, { x: 1, y: 3, w: 2, h: 1, hyperlinkHover: { url: 'https://c.test', sound: { data: 'not-base64' } } })
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	})

	assert.ok(warnings.some(w => w.includes('hyperlink requires either `url` or `slide`')), 'a target-less link must warn')
	assert.ok(warnings.some(w => w.includes('`sound` requires `data` or `path`')), 'a sound without data must warn')
	assert.ok(warnings.some(w => w.includes('`sound.data` lacks a base64 header')), 'bad sound data must warn')

	assert.equal([...xml.matchAll(/<a:hlinkHover/g)].length, 2, 'links with a target survive; the target-less one does not')
	assert.doesNotMatch(xml, /<a:snd /, 'no invalid sound may be written')
	assert.doesNotMatch(xml, /r:id="rId0"/, 'no link may reference a non-existent relationship')
})

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
	slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, fill: { type: 'image', image: { data: SAMPLE_PNG } } })
	slide.addShape(pptx.ShapeType.rect, { x: 4, y: 1, w: 2, h: 1, fill: { type: 'image', image: { data: SAMPLE_PNG, sizing: 'tile', scale: 50, alignment: 'ctr', rotateWithShape: false } } })
	slide.addTable([[{ text: 'i', options: { fill: { type: 'image', image: { data: SAMPLE_PNG } } } }]], { x: 1, y: 3, w: 4 })
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
	let xml = ''
	const warnings = await captureWarnings(async () => {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		// `prst` is an enum, so an unknown value would make the element unparseable
		slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, fill: { type: 'pattern', pattern: { preset: 'tartan' as unknown as 'wave' } } })
		slide.addShape(pptx.ShapeType.rect, { x: 4, y: 1, w: 2, h: 1, fill: { type: 'pattern' } })
		// a picture fill with no image would leave `a:blip` dangling, which reads as damage
		slide.addShape(pptx.ShapeType.rect, { x: 1, y: 3, w: 2, h: 1, fill: { type: 'image', image: {} } })
		slide.addShape(pptx.ShapeType.rect, { x: 4, y: 3, w: 2, h: 1, fill: { type: 'image', image: { data: 'not-base64' } } })
		slide.addShape(pptx.ShapeType.rect, { x: 1, y: 5, w: 2, h: 1, fill: { type: 'image', image: { data: SAMPLE_PNG, sizing: 'tile', alignment: 'middle' as unknown as 'ctr' } } })
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	})

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

test('contract: editing locks and non-visual properties are reachable', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addShape(pptx.ShapeType.rect, {
		x: 1, y: 1, w: 2, h: 1,
		objectName: 'Locked Box', title: 'Alt title', hidden: true,
		lock: { noMove: true, noResize: true, noSelect: true, noTextEdit: true },
	})
	slide.addImage({ data: SAMPLE_PNG, x: 4, y: 1, w: 1, h: 1, lock: { noCrop: true, preferRelativeResize: true }, title: 'Pic title' })
	slide.addTable([['a']], { x: 1, y: 3, w: 4, lock: { noSelect: true } })
	const lockZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(lockZip)
	const xml = await readPart(lockZip, 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /NaN|undefined/, 'lock options must not leak invalid values')
	// `title` is the alt-text title, distinct from `descr`
	// self-closes because this shape has no hyperlink children
	assert.match(xml, /<p:cNvPr id="2" name="Locked Box" title="Alt title" hidden="1"\/>/, 'shape non-visual props missing')
	assert.match(xml, /<a:spLocks noSelect="1" noMove="1" noResize="1" noTextEdit="1"\/>/, 'shape locks missing')
	// the picture keeps the `noChangeAspect` the library has always emitted, and adds the caller's
	assert.match(xml, /<p:cNvPicPr preferRelativeResize="1"><a:picLocks noChangeAspect="1" noCrop="1"\/><\/p:cNvPicPr>/, 'picture locks missing')
	assert.match(xml, /title="Pic title"/, 'picture alt title missing')
	// the table frame keeps its default noGrp
	assert.match(xml, /<a:graphicFrameLocks noGrp="1" noSelect="1"\/>/, 'graphic frame locks missing')
})

test('contract: locks that do not apply to an object are dropped with a warning', async () => {
	let xml = ''
	const warnings = await captureWarnings(async () => {
		const pptx = new pptxgen()
		// `noTextEdit` is not a graphicFrameLocks attribute; `noCrop` is not a spLocks one
		pptx.addSlide().addTable([['a']], { x: 1, y: 1, w: 3, lock: { noTextEdit: true, noSelect: true } })
		pptx.addSlide().addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, lock: { noCrop: true, noMove: true } })
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	})

	assert.equal(warnings.filter(w => w.includes('does not apply to this object type')).length, 2, 'inapplicable locks must warn')
	// the applicable lock still lands; the inapplicable one is not written
	assert.match(xml, /<a:graphicFrameLocks noGrp="1" noSelect="1"\/>/, 'the applicable lock must survive')
	assert.doesNotMatch(xml, /noTextEdit/, 'an inapplicable lock must not be emitted')
})

test('contract: objects without locks keep the output they always had', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1 })
	slide.addImage({ data: SAMPLE_PNG, x: 4, y: 1, w: 1, h: 1 })
	slide.addTable([['a']], { x: 1, y: 3, w: 3 })
	const xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /a:spLocks|title=|hidden=|preferRelativeResize/, 'no lock or non-visual attribute may appear unasked')
	// the two locks the library has always written are unchanged
	assert.match(xml, /<p:cNvPicPr><a:picLocks noChangeAspect="1"\/><\/p:cNvPicPr>/, 'default picture lock changed')
	assert.match(xml, /<a:graphicFrameLocks noGrp="1"\/>/, 'default frame lock changed')
})
