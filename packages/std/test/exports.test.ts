/**
 * The package's shape, not its behaviour.
 *
 * Helpers live in category directories and are re-exported by the root barrel, with one subpath
 * per category. A new category is easy to add and just as easy to forget to wire up - these
 * assertions fail when the barrel and the categories drift apart.
 */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import * as barrel from '../src/index'
import * as charts from '../src/charts'
import * as layout from '../src/layout'
import * as text from '../src/text'
import * as tables from '../src/tables'

const SRC = join(import.meta.dirname, '..', 'src')
const PACKAGE = JSON.parse(readFileSync(join(import.meta.dirname, '..', 'package.json'), 'utf8'))

/** Every `src/<category>/` directory - the source of truth these assertions compare against */
const categories = readdirSync(SRC, { withFileTypes: true })
	.filter(entry => entry.isDirectory())
	.map(entry => entry.name)
	.sort()

test('every category directory is reachable as its own subpath', () => {
	assert.deepEqual(categories, ['charts', 'layout', 'tables', 'text'], 'a new category needs the wiring below')
	for (const category of categories) {
		const subpath = PACKAGE.exports[`./${category}`]
		assert.ok(subpath, `exports is missing ./${category} - consumers cannot import it`)
		assert.equal(subpath.types, `./dist/${category}/index.d.ts`)
		assert.equal(subpath.import, `./dist/${category}/index.mjs`)
		assert.equal(subpath.require, `./dist/${category}/index.cjs`)
	}
})

test('every category has a rollup entry, so its subpath is actually built', () => {
	const config = readFileSync(join(SRC, '..', 'rollup.config.mjs'), 'utf8')
	for (const category of categories) {
		assert.match(config, new RegExp(`'${category}/index': 'src/${category}/index.ts'`), `rollup has no entry for ${category}`)
	}
})

test('the root barrel re-exports every category and nothing else', () => {
	const fromCategories = [...Object.keys(layout), ...Object.keys(charts), ...Object.keys(text), ...Object.keys(tables)].sort()
	assert.deepEqual(Object.keys(barrel).sort(), fromCategories, 'the barrel and the categories have drifted')
})

test('the barrel defines nothing of its own', () => {
	const source = readFileSync(join(SRC, 'index.ts'), 'utf8')
	assert.doesNotMatch(source, /^(export )?(function|const|class|let|var) /m, 'helpers belong in a category, not in the barrel')
})

test('every export is a callable helper', () => {
	assert.deepEqual(Object.keys(barrel).sort(), ['checkOverflow', 'cm', 'column', 'cssColorToHex', 'fitText', 'grid', 'gridFor', 'measureText', 'paginateTable', 'pt', 'registerFontMetrics', 'row', 'tableFromHtml', 'waterfall'])
	for (const [name, value] of Object.entries(barrel)) assert.equal(typeof value, 'function', `${name} should be a function`)
})

test('the barrel and the subpaths hand back the same functions', () => {
	assert.equal(barrel.grid, layout.grid)
	assert.equal(barrel.gridFor, layout.gridFor)
	assert.equal(barrel.waterfall, charts.waterfall)
	assert.equal(barrel.measureText, text.measureText)
	assert.equal(barrel.paginateTable, tables.paginateTable)
})

test('every exports target is published', () => {
	const targets = Object.values(PACKAGE.exports).flatMap(entry => Object.values(entry as Record<string, string>))
	assert.ok(targets.length > 0)
	for (const target of targets) {
		assert.match(target, /^\.\/dist\//, `${target} is outside dist/, so \`files\` would not publish it`)
	}
	assert.deepEqual(PACKAGE.files, ['dist'], 'the publish list must still cover every exports target')
	assert.equal(PACKAGE.sideEffects, false, 'pure helpers: bundlers must be free to drop unused categories')
})

test('the core is a peer, not a runtime dependency', () => {
	// A `dependencies` entry would let a consumer end up with a second copy of the core at another
	// version, while these helpers act on the slide objects the consumer's own instance created.
	// `waterfall` needs per-series `color: 'transparent'` (4.1.0) for the riser and per-point
	// `dataLabels` (4.2.0) for the signed labels. A wider range would report a satisfied peer and
	// then render the wrong chart, so the floor tracks the newest feature in use.
	assert.equal(PACKAGE.peerDependencies['@neo-ma/pptxgenjs'], '^4.2.0')
	assert.equal(PACKAGE.dependencies, undefined, 'std must stay dependency-free at runtime')
	// Declared as a dev dependency too, so this package resolves the core's types on its own rather
	// than by accident of sharing a workspace with it - at the same floor it promises consumers.
	assert.equal(PACKAGE.devDependencies['@neo-ma/pptxgenjs'], PACKAGE.peerDependencies['@neo-ma/pptxgenjs'])

	const sources = categories.flatMap(category =>
		readdirSync(join(SRC, category)).map(file => readFileSync(join(SRC, category, file), 'utf8'))
	)
	for (const source of sources) {
		// Relative imports stay allowed - helpers may share code inside the package. What must never
		// appear is a runtime import of a *package*, which is how the core would sneak in as a dep.
		const packageImport = /^import (?!type )[^']*from '[^.][^']*'/m.exec(source)
		assert.equal(packageImport, null, `runtime package import found - the core must only be imported as a type: ${packageImport?.[0]}`)
	}
})
