/**
 * PptxGenJS: Unit tests for utility methods
 * Run with: `npm test` (node built-in test runner + tsx)
 */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import {
	debugLog,
	isDebugEnabled,
	getSmartParseNumber,
	getUuid,
	encodeXmlEntities,
	inch2Emu,
	valToPts,
	convertRotationDegrees,
	componentToHex,
	rgbToHex,
	createColorElement,
	createGlowElement,
	genXmlColorSelection,
	getNewRelId,
	correctShadowOptions,
	base64ToBytes,
	binaryStringToBase64,
	bytesToBase64,
	utf8ToBase64,
} from '../src/gen-utils'
import { PresLayout, PresSlide, ShadowProps } from '../src/core-interfaces'
import { DEF_FONT_COLOR } from '../src/core-enums'

// 10in x 7.5in layout expressed in EMU
const LAYOUT = { name: 'TEST', width: 9144000, height: 6858000 } as PresLayout

test('inch2Emu', () => {
	assert.equal(inch2Emu(1), 914400)
	assert.equal(inch2Emu(0.5), 457200)
	assert.equal(inch2Emu('2'), 1828800)
	assert.equal(inch2Emu('1in'), 914400)
	assert.equal(inch2Emu(200), 200, 'values > 100 are assumed to be EMU already')
})

test('valToPts', () => {
	assert.equal(valToPts(1), 12700)
	assert.equal(valToPts('2'), 25400)
	assert.equal(valToPts('not-a-number'), 0)
	assert.equal(valToPts(undefined as unknown as number), 0)
})

test('convertRotationDegrees', () => {
	assert.equal(convertRotationDegrees(0), 0)
	assert.equal(convertRotationDegrees(90), 5400000)
	assert.equal(convertRotationDegrees(360), 21600000)
	assert.equal(convertRotationDegrees(361), 60000, 'wraps values over 360')
	assert.equal(convertRotationDegrees(undefined as unknown as number), 0)
})

test('componentToHex', () => {
	assert.equal(componentToHex(0), '00')
	assert.equal(componentToHex(15), '0f')
	assert.equal(componentToHex(255), 'ff')
})

test('rgbToHex', () => {
	assert.equal(rgbToHex(255, 0, 0), 'FF0000')
	assert.equal(rgbToHex(0, 255, 0), '00FF00')
	assert.equal(rgbToHex(0, 0, 0), '000000')
})

test('encodeXmlEntities', () => {
	assert.equal(encodeXmlEntities('a & b < c > d "e" \'f\''), 'a &amp; b &lt; c &gt; d &quot;e&quot; &apos;f&apos;')
	assert.equal(encodeXmlEntities(null as unknown as string), '')
	assert.equal(encodeXmlEntities(undefined as unknown as string), '')
})

test('getSmartParseNumber', () => {
	assert.equal(getSmartParseNumber(1, 'X', LAYOUT), 914400, 'small numbers are inches')
	assert.equal(getSmartParseNumber(914400, 'X', LAYOUT), 914400, 'large numbers are already EMU')
	assert.equal(getSmartParseNumber('50%', 'X', LAYOUT), 4572000, 'percent of width')
	assert.equal(getSmartParseNumber('50%', 'Y', LAYOUT), 3429000, 'percent of height')
	assert.equal(getSmartParseNumber('garbage', 'X', LAYOUT), 0)
})

test('getUuid', () => {
	assert.match(getUuid('xxxxxxxx'), /^[0-9a-f]{8}$/)
	assert.match(getUuid('y'), /^[89ab]$/, 'the "y" nibble is constrained per RFC4122')
	assert.notEqual(getUuid('xxxxxxxx-xxxx'), getUuid('xxxxxxxx-xxxx'), 'values are random')
})

test('base64 helpers preserve binary and UTF-8 data without Buffer', () => {
	assert.equal(bytesToBase64(Uint8Array.from([0, 1, 255])), 'AAH/')
	assert.deepEqual(base64ToBytes('data:application/octet-stream;base64,AAH/'), Uint8Array.from([0, 1, 255]))
	assert.equal(utf8ToBase64('Grüße'), 'R3LDvMOfZQ==')
	assert.equal(binaryStringToBase64('\x00\xff'), 'AP8=')
})

test('createColorElement: hex', () => {
	assert.equal(createColorElement('FF0000'), '<a:srgbClr val="FF0000"/>')
	assert.equal(createColorElement('#ff0000'), '<a:srgbClr val="FF0000"/>', 'strips # and uppercases')
	assert.equal(createColorElement('FF0000', '<a:alpha val="50000"/>'), '<a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr>')
})

test('createColorElement: scheme color', () => {
	assert.equal(createColorElement('accent1'), '<a:schemeClr val="accent1"/>')
})

test('createColorElement: invalid falls back to default font color', () => {
	const orig = console.warn
	console.warn = () => {} // silence the expected warning
	try {
		assert.equal(createColorElement('NOPE'), '<a:srgbClr val="000000"/>')
	} finally {
		console.warn = orig
	}
})

test('createGlowElement', () => {
	assert.equal(
		createGlowElement({ size: 8, color: 'FFFFFF', opacity: 0.75 }, { size: 8, color: 'FFFFFF', opacity: 0.75 }),
		'<a:glow rad="101600"><a:srgbClr val="FFFFFF"><a:alpha val="75000"/></a:srgbClr></a:glow>'
	)
})

test('genXmlColorSelection', () => {
	assert.equal(genXmlColorSelection('FF0000'), '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')
	assert.equal(
		genXmlColorSelection({ type: 'solid', color: 'FF0000', transparency: 50 }),
		'<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>'
	)
	assert.equal(genXmlColorSelection({ type: 'none', color: 'FF0000' }), '', 'unsupported fill types are not emitted')
})

test('genXmlColorSelection: gradient fill', () => {
	assert.equal(
		genXmlColorSelection({ type: 'gradient', gradient: { stops: [{ color: 'FF0000', position: 0 }, { color: '0000FF', position: 100 }] } }),
		'<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>',
		'default is a 90deg linear gradient that rotates with the shape'
	)
	assert.equal(
		genXmlColorSelection({
			type: 'gradient',
			gradient: {
				type: 'radial',
				rotateWithShape: false,
				stops: [{ color: '0000FF', position: 100 }, { color: 'accent1', position: 0, transparency: 25 }],
			},
		}),
		'<a:gradFill rotWithShape="0"><a:gsLst><a:gs pos="0"><a:schemeClr val="accent1"><a:alpha val="75000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst>' +
			'<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>',
		'stops are sorted by position; scheme colors, transparency, and radial geometry are supported'
	)
	assert.equal(
		genXmlColorSelection({ type: 'gradient', gradient: { angle: -90, scaled: true, stops: [{ color: 'FF0000', position: 0 }, { color: '0000FF', position: 100 }] } }),
		'<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>',
		'negative angles normalize into ST_PositiveFixedAngle'
	)
})

test('genXmlColorSelection: gradient input validation', () => {
	const orig = console.warn
	const warnings: string[] = []
	console.warn = (msg: string) => warnings.push(msg)
	try {
		assert.equal(
			genXmlColorSelection({ type: 'gradient', gradient: { stops: [{ color: '00FF00', position: 0 }] } }),
			'<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>',
			'fewer than 2 stops degrades to a solid fill'
		)
		assert.equal(
			genXmlColorSelection({ type: 'gradient' } as unknown as Parameters<typeof genXmlColorSelection>[0]),
			`<a:solidFill><a:srgbClr val="${DEF_FONT_COLOR}"/></a:solidFill>`,
			'a missing gradient degrades to a solid fill'
		)
		assert.equal(warnings.length, 3, 'each degraded gradient warns, plus the invalid-color warning for the missing gradient')
	} finally {
		console.warn = orig
	}

	// out-of-range/non-finite stop positions and angles must never reach the XML as NaN or > 100000
	const xml = genXmlColorSelection({
		type: 'gradient',
		gradient: { angle: NaN, stops: [{ color: 'FF0000', position: 999 }, { color: '0000FF', position: -50 }, { color: '00FF00', position: 'x' as unknown as number }] },
	})
	assert.doesNotMatch(xml, /NaN/, 'no NaN attributes')
	assert.equal(xml, '<a:gradFill rotWithShape="1"><a:gsLst>' +
		'<a:gs pos="0"><a:srgbClr val="0000FF"/></a:gs><a:gs pos="0"><a:srgbClr val="00FF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="FF0000"/></a:gs>' +
		'</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>', 'positions clamp to 0-100 and a non-finite angle falls back to 90deg')
})

test('getNewRelId', () => {
	const slide = { _rels: [1, 2], _relsChart: [1], _relsMedia: [] } as unknown as PresSlide
	assert.equal(getNewRelId(slide), 4, 'sum of all rel arrays + 1')
})

test('correctShadowOptions', () => {
	assert.equal(correctShadowOptions(undefined as unknown as ShadowProps), undefined)
	assert.equal(correctShadowOptions('nope' as unknown as ShadowProps), undefined)
	assert.equal(correctShadowOptions({ type: 'bogus' } as unknown as ShadowProps)?.type, 'outer', 'invalid type corrected')
	assert.equal(correctShadowOptions({ type: 'outer', angle: 400 } as ShadowProps)?.angle, 270, 'out-of-range angle corrected')
	assert.equal(correctShadowOptions({ type: 'outer', opacity: 2 } as ShadowProps)?.opacity, 0.75, 'out-of-range opacity corrected')
	assert.equal(correctShadowOptions({ type: 'outer', color: '#FF0000' } as ShadowProps)?.color, 'FF0000', 'strips leading #')
})

test('debugLog: silent unless PPTXGENJS_DEBUG or NODE_DEBUG is set', () => {
	const { PPTXGENJS_DEBUG, NODE_DEBUG } = process.env
	const calls: unknown[][] = []
	const orig = console.debug
	console.debug = (...args: unknown[]) => calls.push(args)
	try {
		delete process.env.PPTXGENJS_DEBUG
		delete process.env.NODE_DEBUG
		assert.equal(isDebugEnabled(), false)
		debugLog('quiet')
		assert.equal(calls.length, 0)

		process.env.NODE_DEBUG = 'http,pptxgenjs'
		assert.equal(isDebugEnabled(), true)
		debugLog('loud')
		assert.deepEqual(calls, [['[pptxgenjs]', 'loud']])

		process.env.NODE_DEBUG = 'http'
		assert.equal(isDebugEnabled(), false, 'NODE_DEBUG matched on a partial section name')

		process.env.PPTXGENJS_DEBUG = '1'
		assert.equal(isDebugEnabled(), true)
	} finally {
		console.debug = orig
		if (PPTXGENJS_DEBUG === undefined) delete process.env.PPTXGENJS_DEBUG; else process.env.PPTXGENJS_DEBUG = PPTXGENJS_DEBUG
		if (NODE_DEBUG === undefined) delete process.env.NODE_DEBUG; else process.env.NODE_DEBUG = NODE_DEBUG
	}
})
