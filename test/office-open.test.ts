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
		const slide = pptx.addSlide({ transition: { type: 'wipe', direction: 'left', speed: 'slow' } })
		slide.addText('OOXML consumer smoke test', { x: 0.5, y: 0.5, w: 5, h: 0.5, animation: { type: 'fadeIn' } })
		slide.addText([
			{ text: 'ratio ' },
			{ text: 'a/b', options: { omml: '<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>' } },
		], { x: 0.5, y: 1, w: 5, h: 0.5 })
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

