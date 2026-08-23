/**
 * Sample data shared by the contract test files.
 *
 * These are inputs, not assertions: a 1x1 PNG, a silent WAV, a minimal EOT and InkML payload, and
 * the small presentation several package-level contracts inspect. They live here so splitting the
 * contracts by feature area does not mean duplicating a base64 blob into five files.
 */
import JSZip from 'jszip'
import pptxgen from '../src/pptxgen'

/** 1x1 transparent PNG, as a data URI - the smallest thing `addImage()` accepts */
export const SAMPLE_PNG = 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkAAIAAAoAAv/lxKUAAAAASUVORK5CYII='

/** 4x2 px PNG, for effects that need more than one pixel to be visible */
export const SAMPLE_PNG_4X2 = 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAADwyuo0AAAADklEQVR4nGP4jwQYkDkANvEX6SAXxcIAAAAASUVORK5CYII='

/** Empty 8-bit mono WAV, as a data URI */
export const SAMPLE_WAV = 'audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQAAAAA='

/** One-byte MP4 stand-in: nothing inspects the bytes, only the relationship and content type */
export const SAMPLE_MP4 = 'video/mp4;base64,QQ=='

/** One-byte MP3 stand-in */
export const SAMPLE_MP3 = 'audio/mp3;base64,QQ=='

/** Smallest well-formed InkML document, for content-part payloads */
export const SAMPLE_INK = '<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:trace>0 0, 10 10</inkml:trace></inkml:ink>'

/** Relationship type a custom-XML content part is reached by */
export const CUSTOM_XML_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'

/**
 * Minimal stand-in for an EOT font: the only bytes PptxGenJS inspects are the magic number
 * `0x504C` at offset 34 (MS-EOT 2.1), so a 64-byte buffer with that field set is enough to
 * exercise the part/relationship/content-type plumbing without shipping a font file.
 */
export function fakeEot (): Uint8Array {
	const bytes = new Uint8Array(64)
	bytes[34] = 0x4c
	bytes[35] = 0x50
	bytes[36] = 0x77 // marker so the test can prove the bytes reached the part verbatim
	return bytes
}

/**
 * One slide carrying a text box, a shape, a table, and a chart - the fixture the package-level
 * contracts inspect. Built per test file rather than shared across them, so the files stay
 * independent and can run in any order.
 */
export async function buildContractFixture (): Promise<JSZip> {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('Contract', { x: 0.5, y: 0.3, w: 6, h: 0.5, fontSize: 18, color: '0000FF', bold: true })
	slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1.2, w: 2, h: 1, fill: { color: 'FF0000' } })
	slide.addTable([['A', 'B'], ['1', '2']], { x: 0.5, y: 2.6, w: 5 })
	slide.addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [10, 20] }], { x: 0.5, y: 4, w: 6, h: 3 })
	return await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
}

/**
 * Run `body` with `console.warn` captured, and return what it warned.
 *
 * Several contracts assert that bad input is dropped *with* a warning rather than silently, so the
 * capture/restore dance appeared in every one of them. Restores in a `finally`, so a throwing body
 * does not leave the console patched for the rest of the file.
 */
export async function captureWarnings (body: () => Promise<void> | void): Promise<string[]> {
	const warnings: string[] = []
	const original = console.warn
	console.warn = (message: unknown) => warnings.push(String(message))
	try {
		await body()
	} finally {
		console.warn = original
	}
	return warnings
}
