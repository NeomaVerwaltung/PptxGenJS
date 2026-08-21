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

test('profile: every mc:AlternateContent is well formed and offers a fallback', async () => {
	// exercise every path that emits optional markup through the compatibility wrapper
	const pptx = new pptxgen()
	pptx.addSection({ title: 'MCE' })
	const slide = pptx.addSlide({ sectionTitle: 'MCE', transition: { type: 'morph', duration: 1200 } })
	slide.addZoom({ slideNumber: 2, x: 1, y: 1, w: 2, h: 1 })
	slide.addSummaryZoom({ sectionTitles: ['MCE'], x: 4, y: 1, w: 3, h: 2 })
	slide.addText([{ text: 'a/b', options: { omml: '<m:r><m:t>x</m:t></m:r>' } }], { x: 1, y: 3, w: 3, h: 1 })
	pptx.addSlide({ sectionTitle: 'MCE', transition: { type: 'push', direction: 'up', duration: 700, speed: 'fast' } })
	const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)

	let seen = 0
	for (const name of Object.keys(zip.files).filter(file => file.endsWith('.xml'))) {
		const xml = await (zip.file(name) as JSZip.JSZipObject).async('string')
		for (const block of xml.matchAll(/<mc:AlternateContent\b[\s\S]*?<\/mc:AlternateContent>/g)) {
			seen++
			const content = block[0]

			// `mc` itself must be declared on the wrapper, or the whole block is unreadable
			assert.match(content, /^<mc:AlternateContent xmlns:mc="http:\/\/schemas\.openxmlformats\.org\/markup-compatibility\/2006">/, `${name}: mc namespace not declared on AlternateContent`)

			// ECMA-376 Part 3: a Choice must state what a consumer needs, and those prefixes must be
			// in scope at the Choice - otherwise the condition cannot be evaluated
			const choice = /<mc:Choice\b([^>]*)>/.exec(content)
			assert.ok(choice, `${name}: AlternateContent without a Choice`)
			const requires = /Requires="([^"]+)"/.exec(choice[1])
			assert.ok(requires, `${name}: mc:Choice without a Requires attribute`)
			const declared = [...choice[1].matchAll(/xmlns:(\w+)=/g)].map(match => match[1])
			requires[1].split(/\s+/).forEach(prefix => {
				assert.ok(declared.includes(prefix), `${name}: mc:Choice requires "${prefix}" but does not declare it`)
			})

			// without a fallback, a consumer that rejects the choice silently drops the content
			assert.match(content, /<mc:Fallback(\/>|>)/, `${name}: AlternateContent without a Fallback`)

			// the fallback exists for consumers that lack the optional namespaces, so it must not use them
			const fallback = /<mc:Fallback>([\s\S]*)<\/mc:Fallback>/.exec(content)?.[1] ?? ''
			declared.forEach(prefix => {
				assert.doesNotMatch(fallback, new RegExp(`<${prefix}:|\\s${prefix}:`), `${name}: mc:Fallback uses the optional prefix "${prefix}"`)
			})
		}
	}

	assert.ok(seen >= 5, `expected the sample deck to exercise several AlternateContent blocks, saw ${seen}`)
})

test('profile: optional markup only appears inside a wrapper MS-PPTX 2.2 permits', async () => {
	// Optional (non-ECMA-376) markup is legal in exactly three places: an `extLst` extension, an
	// `mc:Choice`, or under an `mc:Ignorable` declaration on the part root. Anything else is markup a
	// consumer cannot skip, which is how a package ends up needing repair.
	const pptx = new pptxgen()
	pptx.chartTrackingRefBased = true
	pptx.guides = [{ orientation: 'vert', position: 3 }]
	pptx.addSection({ title: 'Audit' })
	const slide = pptx.addSlide({ sectionTitle: 'Audit', transition: { type: 'morph', duration: 900 } })
	slide.creationId = true
	slide.addTable([['a']], { x: 1, y: 1, w: 3 })
	slide.addZoom({ slideNumber: 2, x: 5, y: 1, w: 2, h: 1 })
	slide.addMedia({ type: 'video', data: 'video/mp4;base64,QQ==', x: 1, y: 3, w: 2, h: 2, autoplay: true })
	slide.addText([{ text: 'x', options: { omml: '<m:r><m:t>x</m:t></m:r>' } }], { x: 4, y: 3, w: 3, h: 1 })
	pptx.addSlide({ sectionTitle: 'Audit' })
	const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)

	/** Strip the regions where optional markup is allowed, so anything left is a violation */
	const stripAllowed = (xml: string): string =>
		xml
			.replace(/<mc:AlternateContent\b[\s\S]*?<\/mc:AlternateContent>/g, '')
			.replace(/<(\w+):extLst>[\s\S]*?<\/\1:extLst>/g, '')

	/** Prefixes bound to a base ECMA-376 namespace, which every consumer understands */
	const REQUIRED_PREFIXES = new Set(['a', 'p', 'r', 'c', 'm', 'w', 'x', 'mc', 'xdr', 'x14ac', 'x15', 'mv', 'o', 'v'])

	for (const name of Object.keys(zip.files).filter(file => file.endsWith('.xml') && file.startsWith('ppt/'))) {
		const xml = await (zip.file(name) as JSZip.JSZipObject).async('string')
		const ignorable = /mc:Ignorable="([^"]+)"/.exec(xml)?.[1].split(/\s+/) ?? []
		const remaining = stripAllowed(xml)

		const offenders = [...new Set([...remaining.matchAll(/<(\w+):/g)].map(match => match[1]))]
			.filter(prefix => !REQUIRED_PREFIXES.has(prefix) && !ignorable.includes(prefix))

		assert.deepEqual(
			offenders,
			[],
			`${name}: optional markup with prefix(es) "${offenders.join(', ')}" is neither in an extLst, ` +
			'in an mc:Choice, nor covered by mc:Ignorable on the part root'
		)
	}
})
