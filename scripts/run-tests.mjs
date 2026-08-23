#!/usr/bin/env node
/**
 * Run a package's test files through `node --test`.
 *
 * The file list used to be hardcoded in `package.json`, which meant a new test file silently did not
 * run until someone remembered to register it. A shell glob is not a safe replacement: npm runs
 * scripts through `cmd.exe` on Windows, which does not expand globs, and Node only learned to expand
 * them itself in v21 while this repo supports v20. Discovery therefore happens here - recursively, so
 * a test file in a subdirectory is found too.
 *
 * Extra arguments are passed through to `node` (the coverage script adds
 * `--experimental-test-coverage`).
 */
import { readdirSync } from 'node:fs'
import { spawnSync } from 'node:child_process'
import { join } from 'node:path'

const TEST_DIR = 'test'

function findTests (dir) {
	return readdirSync(dir, { withFileTypes: true }).flatMap(entry => {
		const path = join(dir, entry.name)
		if (entry.isDirectory()) return findTests(path)
		return entry.name.endsWith('.test.ts') ? [path] : []
	})
}

const files = findTests(TEST_DIR).sort()
if (files.length === 0) {
	console.error(`no *.test.ts files found under ${TEST_DIR}/`)
	process.exit(1)
}

const { status } = spawnSync(process.execPath, ['--import', 'tsx', '--test', ...process.argv.slice(2), ...files], { stdio: 'inherit' })
process.exit(status ?? 1)
