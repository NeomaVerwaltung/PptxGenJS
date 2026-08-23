/**
 * The round-trip fixture registry, not the round trip itself.
 *
 * `office-open.test.ts` builds its deck from `office-fixtures/` and skips unless LibreOffice is
 * installed, so a fixture module that is never listed - or that throws - would report success
 * forever. These assertions run on every `npm test`: they fail when a module is missing from the
 * list, and they execute every fixture against a real presentation so a broken one surfaces here
 * rather than in the CI-only consumer job.
 */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readdirSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import pptxgen from '../src/pptxgen'
import { OFFICE_FIXTURES } from './office-fixtures'

const DIR = join(import.meta.dirname, 'office-fixtures')

/** `picture-recolor.ts` exports `addPictureRecolorFixture` */
function expectedExport (fileName: string): string {
	const camel = fileName.replace(/\.ts$/, '').split('-').map(part => part[0].toUpperCase() + part.slice(1)).join('')
	return `add${camel}Fixture`
}

test('office-fixtures: every module in the directory is wired into the list', async () => {
	const modules = readdirSync(DIR).filter(name => name.endsWith('.ts') && name !== 'index.ts').sort()
	assert.ok(modules.length > 0, 'no fixture modules found')

	const listed = new Set(OFFICE_FIXTURES.map(fixture => fixture.name))
	const missing: string[] = []
	for (const fileName of modules) {
		const exported = expectedExport(fileName)
		const loaded = await import(join(DIR, fileName)) as Record<string, unknown>
		assert.equal(typeof loaded[exported], 'function', `${fileName} must export ${exported}()`)
		if (!listed.has(exported)) missing.push(`${fileName} -> ${exported}`)
	}

	assert.deepEqual(missing, [], 'these fixture modules exist but are not in office-fixtures/index.ts:\n' + missing.join('\n'))
	assert.equal(OFFICE_FIXTURES.length, modules.length, 'the list and the directory disagree on how many fixtures there are')
})

test('office-fixtures: the whole registry builds a readable presentation', async () => {
	const pptx = new pptxgen()
	for (const addFixture of OFFICE_FIXTURES) addFixture(pptx)
	const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)

	// the deck LibreOffice is handed: every fixture ran, and the package is at least well formed
	assert.ok(zip.file('ppt/presentation.xml'), 'the fixture deck is not a presentation')
	const slides = Object.keys(zip.files).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))
	assert.ok(slides.length >= OFFICE_FIXTURES.length - 2, `expected about one slide per fixture, got ${slides.length}`)
})
