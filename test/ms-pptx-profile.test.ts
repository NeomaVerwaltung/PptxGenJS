/**
 * MS-PPTX conformance profile checks.
 *
 * `docs/ms-pptx-profile.md` records which OOXML extensions this library emits. These tests keep that
 * profile honest: every extension must come from the `OOXML_EXT` registry, every registry entry must
 * be well formed, and every extension that reaches a generated package must sit inside an `extLst`
 * (or `mc:AlternateContent`) as MS-PPTX 2.2 requires.
 */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readdir, readFile } from 'node:fs/promises'
import { join } from 'node:path'
import JSZip from 'jszip'
import pptxgen from '../src/pptxgen'
import { OOXML_EXT } from '../src/core-enums'

const SRC_DIR = new URL('../src', import.meta.url).pathname
/** `src/bld` holds committed build output, not source */
const SKIP_DIRS = new Set(['bld', 'vendor'])

async function sourceFiles (dir: string): Promise<string[]> {
	const entries = await readdir(dir, { withFileTypes: true })
	const files = await Promise.all(entries.map(async entry => {
		const full = join(dir, entry.name)
		if (entry.isDirectory()) return SKIP_DIRS.has(entry.name) ? [] : await sourceFiles(full)
		return entry.name.endsWith('.ts') ? [full] : []
	}))
	return files.flat()
}

test('profile: every extension URI comes from the OOXML_EXT registry', async () => {
	const registered = new Set(Object.values(OOXML_EXT).map(ext => ext.uri))
	const offenders: string[] = []

	for (const file of await sourceFiles(SRC_DIR)) {
		// the registry itself holds the literals; comments may quote sample XML
		if (file.endsWith('core-enums.ts')) continue
		const source = (await readFile(file, 'utf8'))
			.replace(/\/\*[\s\S]*?\*\//g, comment => comment.replace(/[^\n]/g, ' '))
			.replace(/(^|[^:])\/\/[^\n]*/g, match => match.replace(/[^\n]/g, ' '))

		source.split('\n').forEach((line, idx) => {
			for (const match of line.matchAll(/uri="(\{[0-9A-Fa-f-]{36}\})"/g)) {
				offenders.push(`${file.replace(SRC_DIR, 'src')}:${idx + 1} ${match[1]}`)
			}
		})
	}

	assert.deepEqual(
		offenders,
		[],
		'extension URIs must be referenced through `OOXML_EXT` so they stay recorded in docs/ms-pptx-profile.md:\n' + offenders.join('\n')
	)
	assert.ok(registered.size > 0, 'registry is empty')
})

test('profile: registry entries are well formed and unique', () => {
	const uris = new Set<string>()

	Object.entries(OOXML_EXT).forEach(([name, ext]) => {
		assert.match(ext.uri, /^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$/, `${name}: uri must be a braced GUID`)
		assert.match(ext.ns, /^http:\/\/schemas\.(microsoft\.com|openxmlformats\.org)\//, `${name}: namespace must be an OOXML or Microsoft schema URI`)
		assert.equal(uris.has(ext.uri), false, `${name}: duplicate URI ${ext.uri}`)
		uris.add(ext.uri)
	})
})

test('profile: every extension in a generated package is wrapped as MS-PPTX 2.2 requires', async () => {
	// exercise the paths that emit extensions: theme, sections, notes, tables, charts, media, svg, hyperlink color
	const pptx = new pptxgen()
	pptx.addSection({ title: 'Profile' })
	const slide = pptx.addSlide({ sectionTitle: 'Profile' })
	slide.addText([{ text: 'link', options: { hyperlink: { url: 'https://example.com' }, color: '0000FF' } }], { x: 1, y: 1, w: 3, h: 1 })
	slide.addTable([['a', 'b']], { x: 1, y: 2, w: 4 })
	slide.addChart(pptx.ChartType.pie, [{ name: 'Share', labels: ['A', 'B'], values: [60, 40] }], { x: 1, y: 3, w: 4, h: 3, showLeaderLines: true })
	slide.addMedia({ type: 'video', data: 'video/mp4;base64,QQ==', x: 6, y: 1, w: 2, h: 2 })
	slide.addNotes('notes')
	const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)

	const registered = new Set(Object.values(OOXML_EXT).map(ext => ext.uri))
	const parts = Object.keys(zip.files).filter(name => name.endsWith('.xml'))
	assert.ok(parts.length > 0, 'no XML parts generated')

	let seen = 0
	for (const name of parts) {
		const xml = await (zip.file(name) as JSZip.JSZipObject).async('string')
		for (const match of xml.matchAll(/<(\w+):ext uri="(\{[0-9A-Fa-f-]{36}\})"/g)) {
			seen++
			const [, prefix, uri] = match
			assert.ok(registered.has(uri), `${name} emits unregistered extension ${uri}`)
			// MS-PPTX 2.2: an extension element must be a child of an extLst of the same namespace prefix
			assert.match(xml, new RegExp(`<${prefix}:extLst>(?:(?!</${prefix}:extLst>)[\\s\\S])*${uri.replace(/[{}]/g, m => `\\${m}`)}`), `${name}: extension ${uri} is not inside a <${prefix}:extLst>`)
		}
	}
	assert.ok(seen >= 5, `expected the sample deck to exercise several extensions, saw ${seen}`)
})
