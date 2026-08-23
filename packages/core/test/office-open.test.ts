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

const execFile = promisify(execFileCallback)

/** 4x2 px PNG */
const PNG_4x2 = 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAADwyuo0AAAADklEQVR4nGP4jwQYkDkANvEX6SAXxcIAAAAASUVORK5CYII='

// The skip above must not let `npm run test:office` pass without doing anything. npm sets
// `npm_lifecycle_event` to the script name, so this fires for that entry point only - `npm test`
// and the `check` job load this file to compile it, and are expected to skip.
test('office: the test:office script provides a LibreOffice binary', { skip: process.env.npm_lifecycle_event === 'test:office' ? false : 'only meaningful for `npm run test:office`' }, () => {
	assert.ok(officeBinary, 'PPTXGENJS_OFFICE_BIN is unset - `npm run test:office` would silently skip')
})

// Skipped rather than failed when LibreOffice is absent, so this file can live in the normal `test`
// script and be compiled on every `npm run check` - it used to be transformed only by the CI-only
// `test:office` script, which is how a duplicate declaration reached CI unnoticed
test('office: LibreOffice opens and converts a generated presentation', { skip: officeBinary ? false : 'set PPTXGENJS_OFFICE_BIN to run' }, async () => {
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
		slide.addText([{ text: '22/08/2026', options: { field: 'datetime1' } }, { text: ' page ' }, { text: '1', options: { field: 'slidenum' } }], { x: 6, y: 7.2, w: 3, h: 0.4 })
		slide.addText('picture bullet', { x: 0.5, y: 7.2, w: 4, h: 0.4, bullet: { image: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkAAIAAAoAAv/lxKUAAAAASUVORK5CYII=', color: 'FF0000', size: 150 } })
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
		// diagonal cell borders, 3-D cells and vertical cell text - the round-trip fixture for #147
		const cellSlide = pptx.addSlide()
		cellSlide.addTable(
			[[
				{ text: 'diagonals', options: { borderDiagonalDown: { color: 'FF0000', width: 2 }, borderDiagonalUp: { type: 'dash', color: '0000FF' } } },
				{ text: 'bevelled', options: { cell3D: { bevel: { preset: 'circle', width: 0.05, height: 0.05 }, material: 'metal', lightRig: { rig: 'threePt', dir: 't' } } } },
				{ text: 'rotated', options: { textDirection: 'vert270', anchorCtr: true, horzOverflow: 'overflow' } },
			]],
			{ x: 0.5, y: 0.5, w: 8, h: 1.5, rtl: true, border: [{ type: 'solid', color: '999999', width: 1 }, { type: 'solid', color: '999999', width: 1 }, { type: 'solid', color: '999999', width: 1 }, { type: 'solid', color: '999999', width: 1 }] }
		)
		// the DrawingML effect vocabulary - the round-trip fixture for #138
		const effectSlide = pptx.addSlide()
		effectSlide.addShape(pptx.ShapeType.rect, {
			x: 0.5, y: 0.5, w: 3, h: 1.5, fill: { color: 'CCCCCC' },
			blur: { radius: 3 },
			fillOverlay: { blend: 'mult', fill: { color: 'FF0000', transparency: 40 } },
			shadow: { type: 'preset', preset: 'shdw7', color: '333333' },
		})
		effectSlide.addShape(pptx.ShapeType.ellipse, { x: 4.5, y: 0.5, w: 2, h: 1.5, fill: { color: '4472C4' }, glow: { size: 6, color: 'FFFF00', opacity: 0.6 }, effectDag: { type: 'sib' } })
		effectSlide.addImage({ data: PNG_4x2, x: 0.5, y: 2.5, w: 2, h: 1, alphaEffects: { invert: true } })
		// theme style references - the round-trip fixture for #141
		const styleSlide = pptx.addSlide()
		styleSlide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.5, w: 3, h: 1.5, styleRef: { line: 1, fill: 3, effect: 2, font: 'minor' } })
		styleSlide.addShape(pptx.ShapeType.ellipse, { x: 4.5, y: 0.5, w: 3, h: 1.5, styleRef: { fill: 1, color: 'accent2' } })
		styleSlide.addImage({ data: PNG_4x2, x: 0.5, y: 2.5, w: 2, h: 1, styleRef: { line: 2 } })
		// layout metadata - the round-trip fixture for #149
		pptx.defineSlideMaster({
			title: 'SECTION HEADER',
			layoutType: 'secHead',
			matchingName: 'Section Header',
			colorMapOverride: { bg1: 'dk1', tx1: 'lt1' },
			objects: [{ placeholder: { options: { name: 'secTitle', type: 'title', x: 0.5, y: 2, w: 9, h: 1.5 }, text: 'Section' } }],
		})
		pptx.addSlide({ masterName: 'SECTION HEADER' }).addText('section', { placeholder: 'secTitle' })
		// media source elements - the round-trip fixture for #150. The linked target is external, so it
		// need not exist: the deck only has to open.
		const mediaSourceSlide = pptx.addSlide()
		mediaSourceSlide.addMedia({ type: 'video', link: 'clip.mp4', x: 0.5, y: 0.5, w: 3, h: 2, contentType: 'video/mp4' })
		mediaSourceSlide.addMedia({ type: 'audioCd', audioCd: { start: { track: 1 }, end: { track: 1, time: 30 } }, x: 4.5, y: 0.5, w: 2, h: 2 })
		mediaSourceSlide.addMedia({ type: 'wav', data: 'audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQAAAAA=', x: 0.5, y: 3, w: 2, h: 1.5 })
		// chart style parts - the round-trip fixture for #157
		pptx.addSlide().addChart('bar', [{ name: 'Series 1', labels: ['a', 'b', 'c'], values: [4, 7, 2] }], { x: 0.5, y: 0.5, w: 6, h: 3, chartStyle: 201, chartColorStyle: { method: 'cycle', id: 10 } })
		// custom table style - the round-trip fixture for #133, proving banding renders
		pptx.tableStyles = [{
			id: '{A1B2C3D4-1111-2222-3333-444455556666}',
			name: 'Office Smoke',
			wholeTable: { borders: { top: { color: '4472C4', width: 1 }, bottom: { color: '4472C4', width: 1 }, insideH: { color: 'D9D9D9', width: 0.5 } } },
			band1H: { fill: { color: 'DEEAF6' } },
			firstRow: { bold: true, color: 'FFFFFF', fill: { color: '4472C4' } },
		}]
		pptx.addSlide().addTable(
			[['Region', 'Sales'], ['West', '20'], ['East', '31'], ['North', '18']],
			{ x: 0.5, y: 0.5, w: 8, tableStyleId: '{A1B2C3D4-1111-2222-3333-444455556666}', firstRow: true, bandRow: true }
		)
		// ChartEx charts are a separate part type wrapped in `mc:AlternateContent`. A consumer that does not
		// understand the `cx1` namespace must take the fallback shape and still open the file cleanly - which
		// is what this asserts; it does not prove the chartex layouts themselves render.
		const chartExSlide = pptx.addSlide()
		chartExSlide.addChart(pptx.ChartType.waterfall, [{ name: 'Cash flow', labels: ['Start', 'Q1', 'Q2', 'End'], values: [100, 30, -20, 110] }], { x: 0.5, y: 0.5, w: 4.5, h: 3, showTitle: true, title: 'Waterfall', showValue: true, chartExSubtotals: [0, 3] })
		chartExSlide.addChart(pptx.ChartType.treemap, [{ name: 'Revenue', labels: ['A', 'B', 'C'], values: [10, 20, 30] }], { x: 5.2, y: 0.5, w: 4.5, h: 3 })
		chartExSlide.addChart(pptx.ChartType.funnel, [{ name: 'Pipeline', labels: ['Leads', 'Qualified', 'Won'], values: [500, 120, 30] }], { x: 0.5, y: 3.8, w: 4.5, h: 3, showLegend: true })
		chartExSlide.addChart(pptx.ChartType.boxWhisker, [
			{ name: 'Alpha', labels: ['x', 'y', 'z'], values: [1, 5, 9] },
			{ name: 'Beta', labels: ['x', 'y', 'z'], values: [2, 6, 4] },
		], { x: 5.2, y: 3.8, w: 4.5, h: 3, chartExMeanLine: true })
		// picture recolour - the round-trip fixture for #131
		const recolorSlide = pptx.addSlide()
		recolorSlide.addImage({ data: PNG_4x2, x: 0.5, y: 0.5, w: 2, h: 1, recolor: { duotone: ['1F3864', 'DEEAF6'] } })
		recolorSlide.addImage({ data: PNG_4x2, x: 3, y: 0.5, w: 2, h: 1, recolor: { grayscale: true, brightness: 20, contrast: -15 } })
		recolorSlide.addImage({ data: PNG_4x2, x: 5.5, y: 0.5, w: 2, h: 1, recolor: { colorChange: { from: 'FFFFFF', to: 'FF0000' } } })
		// groups and connectors - the round-trip fixture for #111. LibreOffice renders both, so unlike
		// the Microsoft-only parts this fixture genuinely exercises the feature.
		const groupSlide = pptx.addSlide()
		groupSlide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.5, w: 2, h: 1, objectName: 'fromBox', fill: { color: '4472C4' } })
		groupSlide.addShape(pptx.ShapeType.rect, { x: 5, y: 2.5, w: 2, h: 1, objectName: 'toBox', fill: { color: 'ED7D31' } })
		groupSlide.addConnector({ x: 2.5, y: 1, w: 2.5, h: 1.5, type: 'bentConnector3', line: { color: '333333', width: 2, endArrowType: 'triangle' }, start: { shape: 'fromBox', site: 3 }, end: { shape: 'toBox', site: 1 } })
		groupSlide.addGroup([
			{ shape: { type: pptx.ShapeType.ellipse, options: { x: 1, y: 1, w: 1, h: 1, fill: { color: '70AD47' } } } },
			{ shape: { type: pptx.ShapeType.rect, options: { x: 2.2, y: 1, w: 1, h: 1, fill: { color: 'FFC000' } } } },
			{ text: { text: 'grouped', options: { x: 1, y: 2.2, w: 2.2, h: 0.5, align: 'center' } } },
		], { x: 0.5, y: 4, w: 4.4, h: 3.4 })
		pptx.addSlide({ transition: { type: 'morph', duration: 1200 } }).addText('modern transition', { x: 0.5, y: 0.5, w: 5, h: 0.5 })
		await writeFile(presentationPath, (await pptx.write({ outputType: 'nodebuffer' })) as Buffer)

		await execFile(officeBinary, ['--headless', '--convert-to', 'pdf', '--outdir', directory, presentationPath], { timeout: 60_000 })
		assert.ok((await stat(join(directory, 'smoke.pdf'))).size > 0, 'LibreOffice did not produce a PDF')
	} finally {
		await rm(directory, { recursive: true, force: true })
	}
})

