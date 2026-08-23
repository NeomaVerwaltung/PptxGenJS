/**
 * Text contracts: runs, paragraphs, fields, bullets, columns, and the OMML math zone.
 *
 * Unlike golden XML snapshots, these checks document the OOXML that matters and allow harmless
 * serializer changes without regenerating fixture files.
 */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import pptxgen from '../src/pptxgen'
import { assertPptxPackageContracts, readPart } from './pptx-contracts'
import { SAMPLE_PNG, captureWarnings } from './fixtures'

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
	let xml = ''
	const warnings = await captureWarnings(async () => {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		// ECMA-376 allows 1-16; anything else makes a:bodyPr unparseable
		slide.addText('too many', { x: 1, y: 1, w: 4, h: 1, columns: 17 })
		slide.addText('too few', { x: 1, y: 2, w: 4, h: 1, columns: 0 })
		slide.addText('not a number', { x: 1, y: 3, w: 4, h: 1, columns: 'two' as unknown as number })
		slide.addText('negative gap', { x: 1, y: 4, w: 4, h: 1, columns: 2, columnSpacing: -1 })
		slide.addText('gap without columns', { x: 1, y: 5, w: 4, h: 1, columnSpacing: 0.5 })
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	})

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

test('contract: text fields emit a:fld with a cached value', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addText([
		{ text: '22/08/2026', options: { field: 'datetime1' } },
		{ text: ' page ' },
		{ text: '1', options: { field: 'slidenum' } },
	], { x: 1, y: 1, w: 4, h: 1 })
	const fldZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(fldZip)
	const xml = await readPart(fldZip, 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /NaN|undefined/, 'field options must not leak invalid values')
	// the cached `a:t` is what a consumer that does not refresh renders
	assert.match(xml, /<a:fld id="\{[0-9a-f]{8}-0000-0000-0000-[0-9a-f]{12}\}" type="datetime1"><a:rPr[^>]*>[\s\S]*?<a:t>22\/08\/2026<\/a:t><\/a:fld>/, 'datetime field missing')
	assert.match(xml, /<a:fld id="\{[0-9a-f-]+\}" type="slidenum">[\s\S]*?<a:t>1<\/a:t><\/a:fld>/, 'slide-number field missing')
	// fields sit alongside ordinary runs in the same paragraph
	assert.match(xml, /<\/a:fld><a:r>[\s\S]*?<a:t> page <\/a:t><\/a:r><a:fld/, 'a field must coexist with plain runs')
	// ids are GUIDs and unique per field
	const ids = [...xml.matchAll(/<a:fld id="(\{[^"]+\})"/g)].map(match => match[1])
	assert.equal(new Set(ids).size, 2, 'each field needs its own id')

	// an unknown type degrades to plain text rather than emitting an invalid enum
	let badXml = ''
	const warnings = await captureWarnings(async () => {
		const bad = new pptxgen()
		bad.addSlide().addText([{ text: 'x', options: { field: 'lunchtime' as unknown as 'slidenum' } }], { x: 1, y: 1, w: 2, h: 1 })
		badXml = await readPart(await JSZip.loadAsync((await bad.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	})
	assert.ok(warnings.some(w => w.includes('unknown text field "lunchtime"')), 'unknown field type must warn')
	assert.doesNotMatch(badXml, /a:fld/, 'an unknown field type must not be emitted')
	assert.match(badXml, /<a:t>x<\/a:t>/, 'the text still renders as a plain run')
})

test('contract: bullets support colour, size, font, and pictures', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('picture bullet', { x: 1, y: 1, w: 4, h: 1, bullet: { image: SAMPLE_PNG, color: 'FF0000', size: 150 } })
	slide.addText('char bullet', { x: 1, y: 3, w: 4, h: 1, bullet: { characterCode: '25BA', fontFace: 'Wingdings', sizePts: 14 } })
	slide.addText('numbered', { x: 1, y: 5, w: 4, h: 1, bullet: { type: 'number', color: '0000FF' } })
	const buZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(buZip)
	const xml = await readPart(buZip, 'ppt/slides/slide1.xml')

	// a picture bullet replaces the character/number bullet and needs an image relationship
	assert.match(xml, /<a:buClr><a:srgbClr val="FF0000"\/><\/a:buClr><a:buSzPct val="150000"\/><a:buBlip><a:blip r:embed="rId\d+"\/><\/a:buBlip>/, 'picture bullet missing')
	const rid = /<a:blip r:embed="(rId\d+)"\/><\/a:buBlip>/.exec(xml)?.[1]
	assert.ok(rid, 'picture bullet has no relationship')
	assert.match(await readPart(buZip, 'ppt/slides/_rels/slide1.xml.rels'), new RegExp(`<Relationship Id="${rid}" Type="[^"]*\\/image"`), 'bullet image relationship missing')
	// `a:buSzPts` is in 100ths of a point and replaces `a:buSzPct`
	assert.match(xml, /<a:buSzPts val="1400"\/><a:buFont typeface="Wingdings"\/><a:buChar char="&#x25BA;"\/>/, 'char bullet font/size missing')
	// numbered bullets keep the major-latin fallback when no face is given
	assert.match(xml, /<a:buClr><a:srgbClr val="0000FF"\/><\/a:buClr><a:buSzPct val="100000"\/><a:buFont typeface="\+mj-lt"\/><a:buAutoNum/, 'numbered bullet changed')
})

test('contract: rtlCol follows rtlMode, and kumimoji is settable', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('rtl', { x: 1, y: 1, w: 4, h: 1, columns: 2, rtlMode: true, kumimoji: true })
	slide.addText('ltr', { x: 1, y: 3, w: 4, h: 1, columns: 2 })
	slide.addText('override', { x: 1, y: 5, w: 4, h: 1, rtlMode: true, rtlColumns: false })
	const xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')

	// an RTL text box flows its columns right-to-left; an LTR one is unchanged
	assert.match(xml, /<a:bodyPr wrap="square" numCol="2" rtlCol="1"/, 'rtlCol must follow rtlMode')
	assert.match(xml, /<a:bodyPr wrap="square" numCol="2" rtlCol="0"/, 'a non-RTL box must stay rtlCol="0"')
	assert.match(xml, /kumimoji="1"/, 'kumimoji missing')
	// an explicit rtlColumns wins over rtlMode
	assert.equal([...xml.matchAll(/rtlCol="0"/g)].length, 2, 'the explicit override must produce rtlCol="0"')
})

test('contract: text without fields, bullet extras, or RTL is unchanged', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addText('plain', { x: 1, y: 1, w: 3, h: 1, bullet: { characterCode: '2022' } })
	const xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /a:fld|a:buClr|a:buBlip|a:buSzPts|kumimoji/, 'nothing new may appear unasked')
	// the historical bullet output: 100% size, no font, then the char
	assert.match(xml, /<a:buSzPct val="100000"\/><a:buChar char="&#x2022;"\/>/, 'default bullet output changed')
	assert.match(xml, /rtlCol="0"/, 'default rtlCol changed')
})
