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
import { OFFICE_FIXTURES } from './office-fixtures'

const officeBinary = process.env.PPTXGENJS_OFFICE_BIN

const execFile = promisify(execFileCallback)

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
		// One fixture per feature area, applied in order - `office-fixtures/index.ts` is the list
		for (const addFixture of OFFICE_FIXTURES) addFixture(pptx)
		await writeFile(presentationPath, (await pptx.write({ outputType: 'nodebuffer' })) as Buffer)

		await execFile(officeBinary, ['--headless', '--convert-to', 'pdf', '--outdir', directory, presentationPath], { timeout: 60_000 })
		assert.ok((await stat(join(directory, 'smoke.pdf'))).size > 0, 'LibreOffice did not produce a PDF')
	} finally {
		await rm(directory, { recursive: true, force: true })
	}
})

