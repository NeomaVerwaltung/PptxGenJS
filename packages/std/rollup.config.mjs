import typescript from 'rollup-plugin-typescript2'

// One entry per category so each is reachable as its own subpath (`.../layout`, `.../charts`).
// Adding a category means adding its `src/<name>/index.ts` here and to `exports` in package.json.
const entries = {
	index: 'src/index.ts',
	'layout/index': 'src/layout/index.ts',
	'charts/index': 'src/charts/index.ts',
	'text/index': 'src/text/index.ts',
	'tables/index': 'src/tables/index.ts',
}

// The helpers only compose the public API - `@neo-ma/pptxgenjs` is a peer used for types,
// never imported at runtime, so there is nothing to bundle and nothing to externalise.
const shared = {
	input: entries,
	external: ['@neo-ma/pptxgenjs'],
	plugins: [typescript({ typescript: require('typescript'), tsconfigOverride: { compilerOptions: { declaration: false, noEmit: false } } })],
}

export default [
	{
		...shared,
		output: { dir: './dist', format: 'es', entryFileNames: '[name].mjs', chunkFileNames: '[name]-[hash].mjs' },
	},
	{
		...shared,
		// `.cjs` rather than `.js`: the package has no `"type"`, so an extension-typed file keeps
		// both builds resolvable by plain Node without a bundler.
		output: { dir: './dist', format: 'cjs', entryFileNames: '[name].cjs', chunkFileNames: '[name]-[hash].cjs' },
	},
]
