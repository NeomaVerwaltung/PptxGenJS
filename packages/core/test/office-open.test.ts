/**
 * Opens a generated presentation with LibreOffice and converts it to PDF.
 * Run explicitly with PPTXGENJS_OFFICE_BIN set to libreoffice or soffice.
 */
import assert from 'node:assert/strict'
import { execFile as execFileCallback } from 'node:child_process'
import { mkdtemp, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { test } from 'node:test'
import pptxgen from '../src/pptxgen'

const officeBinary = process.env.PPTXGENJS_OFFICE_BIN
if (!officeBinary) throw new Error('Set PPTXGENJS_OFFICE_BIN to the LibreOffice executable before running npm run test:office')

const execFile = promisify(execFileCallback)

test('office: LibreOffice opens and converts a generated presentation', async () => {
	const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-office-'))
	const presentationPath = join(directory, 'smoke.pptx')

	try {
		const pptx = new pptxgen()
		// a minimal EOT stand-in: only the magic number at offset 34 is inspected
		const eot = new Uint8Array(64)
		eot[34] = 0x4c
		eot[35] = 0x50
		pptx.addFont({ fontFace: 'Smoke Sans', data: eot })
		pptx.slideShow = { mode: 'browse', loop: true, browseMode: true, laserColor: 'FF0000' }
		pptx.defaultImageDpi = 220
		pptx.readonlyRecommended = true
		pptx.guides = [{ orientation: 'vert', position: 5 }, { orientation: 'horz', position: 3.75 }]
		pptx.notesGuides = [{ orientation: 'horz', position: 1 }]
		const slide = pptx.addSlide({ transition: { type: 'wipe', direction: 'left', speed: 'slow' } })
		slide.creationId = true
		slide.addZoom({ slideNumber: 2, x: 6, y: 0.5, w: 2, h: 1.2 })
		slide.addShape(pptx.ShapeType.rect, { x: 6, y: 2, w: 1.5, h: 1, fill: { type: 'pattern', pattern: { preset: 'diagCross', color: '0000FF' } } })
		slide.addShape(pptx.ShapeType.rect, { x: 8, y: 2, w: 1.5, h: 1, fill: { type: 'image', image: { data: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkAAIAAAoAAv/lxKUAAAAASUVORK5CYII=', sizing: 'tile' } } })
		slide.addComment({ text: 'consumer smoke comment', author: 'Ada Lovelace', x: 4, y: 2, created: '2026-08-20T09:00:00Z', replies: [{ text: 'ack', author: 'Grace Hopper', created: '2026-08-20T10:00:00Z' }] })
		slide.addText('OOXML consumer smoke test', { x: 0.5, y: 0.5, w: 5, h: 0.5, animation: { type: 'fadeIn' } })
		slide.addText([
			{ text: 'ratio ' },
			{ text: 'a/b', options: { omml: '<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>' } },
		], { x: 0.5, y: 1, w: 5, h: 0.5 })
		slide.addShape(pptx.ShapeType.rect, { x: 8, y: 5.5, w: 1.4, h: 0.6, objectName: 'Locked', title: 'Alt title', lock: { noMove: true, noResize: true, noTextEdit: true } })
		slide.addShape(pptx.ShapeType.line, { x: 0.5, y: 6.9, w: 4, h: 0, line: { color: '0000FF', width: 3, compound: 'thickThin', join: 'miter', miterLimit: 400, beginArrowType: 'arrow', beginArrowSize: { width: 'lg', length: 'lg' }, customDash: [{ dash: 400, space: 300 }] } })
		slide.addText('attribute coverage', { x: 6, y: 6.5, w: 3, h: 0.6, capitalization: 'small', upright: true, fontAlign: 'ctr', marginRight: 0.1, underlineLine: { width: 1, color: 'FF0000' }, symbolFontFace: 'Wingdings' })
		slide.addText('column one flows into column two when the box is full', { x: 0.5, y: 5.5, w: 5, h: 1, columns: 2, columnSpacing: 0.25 })
		slide.addShape(pptx.ShapeType.rect, { x: 6.5, y: 5.5, w: 2, h: 0.6, hyperlink: { url: 'https://example.com', tooltip: 'open' }, hyperlinkHover: { slide: 2, tooltip: 'peek' } })
		slide.addTable([['Region', 'Sales'], ['West', '20']], { x: 0.5, y: 1.5, w: 5 })
		slide.addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [10, 20] }], { x: 0.5, y: 3, w: 6, h: 3 })
		const mediaSlide = pptx.addSlide()
		mediaSlide.addMedia({ type: 'video', data: 'video/mp4;base64,QQ==', x: 0.5, y: 0.5, w: 4, h: 2.5, autoplay: true, loop: true, mute: true })
		mediaSlide.addMedia({ type: 'audio', data: 'audio/mp3;base64,QQ==', x: 5, y: 0.5, w: 2, h: 2, autoplay: true })
		mediaSlide.addText('animated with media', { x: 0.5, y: 3.5, w: 5, h: 0.5, animation: { type: 'wipeIn', direction: 'left', trigger: 'afterPrevious' } })
		const gradientSlide = pptx.addSlide()
		gradientSlide.background = { type: 'gradient', gradient: { angle: 45, stops: [{ color: 'FFFFFF', position: 0 }, { color: 'E7E6E6', position: 100 }] } }
		gradientSlide.addShape(pptx.ShapeType.rect, {
			x: 0.5,
			y: 0.5,
			w: 4,
			h: 2,
			fill: { type: 'gradient', gradient: { stops: [{ color: 'FF0000', position: 0 }, { color: '0000FF', position: 100, transparency: 20 }] } },
			line: { type: 'gradient', width: 2, gradient: { type: 'radial', stops: [{ color: '00FF00', position: 0 }, { color: '000000', position: 100 }] } },
		})
		gradientSlide.addTable([[{ text: 'gradient cell', options: { fill: { type: 'gradient', gradient: { stops: [{ color: '111111', position: 0 }, { color: '888888', position: 100 }] } } } }]], { x: 0.5, y: 3, w: 4 })
		pptx.addSlide({ transition: { type: 'morph', duration: 1200 } }).addText('modern transition', { x: 0.5, y: 0.5, w: 5, h: 0.5 })
		await writeFile(presentationPath, (await pptx.write({ outputType: 'nodebuffer' })) as Buffer)

		await execFile(officeBinary, ['--headless', '--convert-to', 'pdf', '--outdir', directory, presentationPath], { timeout: 60_000 })
		assert.ok((await stat(join(directory, 'smoke.pdf'))).size > 0, 'LibreOffice did not produce a PDF')
	} finally {
		await rm(directory, { recursive: true, force: true })
	}
})

