// Keeps the version strings outside package.json in sync. Run by the `version` npm lifecycle
// hook of the package being released, so `npm version <level>` bumps all three at once.
// Paths are resolved against the cwd, which npm sets to that package's directory.
import { readFileSync, writeFileSync } from 'node:fs'

const version = JSON.parse(readFileSync('package.json', 'utf8')).version
const edits = [
	['src/pptxgen.ts', /^const VERSION = '.*'$/m, `const VERSION = '${version}'`],
	['types/index.d.ts', /^\/\/ Type definitions for pptxgenjs .*$/m, `// Type definitions for pptxgenjs ${version}`],
]

// Idempotent: the release workflow runs this explicitly after bumping, which may be a second
// run if npm already fired the `version` lifecycle hook. Guard on the pattern, not on a diff.
for (const [file, pattern, replacement] of edits) {
	const before = readFileSync(file, 'utf8')
	if (!pattern.test(before)) throw new Error(`sync-version: no version line matched in ${file}`)
	writeFileSync(file, before.replace(pattern, replacement))
}
