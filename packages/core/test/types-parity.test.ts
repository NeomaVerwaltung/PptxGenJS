/**
 * Parity between the internal model and the published type surface.
 *
 * `src/core-interfaces.ts` is what the library compiles against; `types/index.d.ts` is what
 * consumers get (`package.json` → `exports.types`) and what `@neo-ma/pptxgenjs-std` typechecks
 * against. They are maintained by hand and nothing linked them, so options could reach the runtime
 * while remaining unreachable from TypeScript - `reflection` and `softEdge` did exactly that.
 *
 * What counts as public:
 * - `_`-prefixed props are internal by naming convention
 * - anything whose declaration is tagged `@internal` is internal by declaration
 * - everything else on an exported interface must appear on the same-named interface in the d.ts
 */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'

const SRC = new URL('../src/core-interfaces.ts', import.meta.url)
const DTS = new URL('../types/index.d.ts', import.meta.url)

interface Declaration {
	name: string
	internal: boolean
	props: Map<string, boolean>
}

/**
 * Parse exported interfaces and their immediate props
 * - nested object literals are skipped: their members belong to the parent prop, not the interface
 */
function parseInterfaces (source: string): Map<string, Declaration> {
	const found = new Map<string, Declaration>()
	const lines = source.split('\n')
	let current: Declaration | undefined
	let depth = 0
	let docInternal = false
	let inDoc = false

	lines.forEach(raw => {
		const line = raw.trim()

		if (line.startsWith('/**')) { inDoc = true; docInternal = line.includes('@internal') }
		else if (inDoc) {
			if (line.includes('@internal')) docInternal = true
			if (line.endsWith('*/')) inDoc = false
		}

		const open = /^export interface (\w+)/.exec(line)
		if (open && depth === 0) {
			current = { name: open[1], internal: docInternal, props: new Map() }
			found.set(open[1], current)
			depth = 1
			docInternal = false
			return
		}
		if (!current) { if (!inDoc && line && !line.startsWith('*')) docInternal = false; return }

		// track nesting so members of an inline object literal are not read as interface props
		const opens = (line.match(/\{/g) ?? []).length
		const closes = (line.match(/\}/g) ?? []).length
		const prop = /^(\w+)\??\s*[:(]/.exec(line)
		if (prop && depth === 1) current.props.set(prop[1], docInternal)

		depth += opens - closes
		if (depth <= 0) { current = undefined; depth = 0 }
		if (!inDoc) docInternal = false
	})

	return found
}

test('parity: every public option in the model is reachable from the published types', async () => {
	const [src, dts] = await Promise.all([readFile(SRC, 'utf8'), readFile(DTS, 'utf8')])
	const model = parseInterfaces(src)
	const published = parseInterfaces(dts)
	assert.ok(model.size > 50, `expected to parse the model, got ${model.size} interfaces`)
	assert.ok(published.size > 50, `expected to parse the published types, got ${published.size} interfaces`)

	// Self-check: a parser that found no props would make every assertion below vacuous, so prove
	// it reads real members - and that it does not mistake a nested object literal for one
	assert.ok(model.get('ShapeProps')?.props.has('fill'), 'parser failed to read ShapeProps.fill')
	assert.ok(published.get('ShapeProps')?.props.has('fill'), 'parser failed to read the published ShapeProps.fill')
	assert.equal(model.get('TextPropsOptions')?.props.has('anchor'), false, '`anchor` belongs to the nested _bodyProp, not to TextPropsOptions')

	const gaps: string[] = []
	model.forEach(decl => {
		if (decl.internal) return
		const target = published.get(decl.name)
		// an interface absent from the published types is only acceptable if it is internal
		if (!target) {
			gaps.push(`${decl.name} (whole interface missing from types/index.d.ts)`)
			return
		}
		decl.props.forEach((propInternal, prop) => {
			if (propInternal || prop.startsWith('_')) return
			if (!target.props.has(prop)) gaps.push(`${decl.name}.${prop}`)
		})
	})

	assert.deepEqual(
		gaps,
		[],
		'these are settable at runtime but unreachable from TypeScript - add them to types/index.d.ts, ' +
		'or mark the declaration `@internal` if that is deliberate:\n  ' + gaps.join('\n  ')
	)
})
