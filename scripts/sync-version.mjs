// Keeps the version strings outside package.json in sync. Run by the `version`
// npm lifecycle hook, so `npm version <level>` bumps all three at once.
import { readFileSync, writeFileSync } from 'node:fs'

const version = JSON.parse(readFileSync('package.json', 'utf8')).version
const edits = [
	['src/pptxgen.ts', /^const VERSION = '.*'$/m, `const VERSION = '${version}'`],
	['types/index.d.ts', /^\/\/ Type definitions for pptxgenjs .*$/m, `// Type definitions for pptxgenjs ${version}`],
]

for (const [file, pattern, replacement] of edits) {
	const before = readFileSync(file, 'utf8')
	const after = before.replace(pattern, replacement)
	if (after === before) throw new Error(`sync-version: no version line matched in ${file}`)
	writeFileSync(file, after)
}
