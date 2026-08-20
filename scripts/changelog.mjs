// Prepends a Keep a Changelog section for a release, built from the commits that touched the
// package being released. Run by the Release workflow's dispatch path before the release commit.
//
// The tag carries the version, so nothing else needs to be passed in.
//
// Usage: node scripts/changelog.mjs --dir packages/core --tag v4.2.1
//        node scripts/changelog.mjs --self-check
import { readFileSync, writeFileSync } from 'node:fs'
import { execFileSync } from 'node:child_process'

const REPO = 'https://github.com/NeomaVerwaltung/PptxGenJS'
const CHANGELOG = 'CHANGELOG.md'

// Conventional-commit type -> Keep a Changelog heading. Types absent here (docs, chore, ci,
// test, build, style) are release plumbing, not user-facing, and are left out.
const HEADINGS = { feat: 'Added', fix: 'Fixed', perf: 'Changed', refactor: 'Changed', revert: 'Changed' }
const ORDER = ['Added', 'Fixed', 'Changed']

/** One commit -> `{ heading, text }`, or null when the commit does not belong in a changelog */
export function classify (subject, hash) {
	const parsed = /^(?<type>[a-z]+)(?:\((?<scope>[^)]*)\))?(?<breaking>!)?: (?<summary>.+)$/.exec(subject)
	if (!parsed?.groups) return null
	const { type, breaking, summary } = parsed.groups
	if (type === 'chore' && summary.startsWith('release')) return null
	const heading = breaking ? 'Changed' : HEADINGS[type]
	if (!heading) return null
	const capitalized = summary.charAt(0).toUpperCase() + summary.slice(1)
	const link = hash ? ` ([${hash}](${REPO}/commit/${hash}))` : ''
	return { heading, text: `- ${breaking ? '**Breaking:** ' : ''}${capitalized}${link}` }
}

/**
 * Render the section body for one release from already-classified entries.
 *
 * The heading is labelled with the tag, not the bare version: two packages share this file and
 * only the tag tells `v4.3.0` from `std-v4.3.0`. It is also what makes the re-run check exact.
 */
export function renderSection (tag, date, entries) {
	const lines = [`## [${tag}](${REPO}/releases/tag/${tag}) - ${date}`, '']
	const used = ORDER.filter(heading => entries.some(entry => entry.heading === heading))
	if (used.length === 0) {
		lines.push('### Changed', '', '- Maintenance release; no user-facing changes.', '')
		return lines.join('\n')
	}
	for (const heading of used) {
		lines.push(`### ${heading}`, '')
		for (const entry of entries.filter(item => item.heading === heading)) lines.push(entry.text)
		lines.push('')
	}
	return lines.join('\n')
}

/** Insert a section above the newest existing one, keeping the file's intro intact */
export function insertSection (changelog, section) {
	const index = changelog.indexOf('\n## [')
	if (index === -1) return `${changelog.trimEnd()}\n\n${section}`
	return `${changelog.slice(0, index + 1)}${section}\n${changelog.slice(index + 1)}`
}

function git (...args) {
	return execFileSync('git', args, { encoding: 'utf8' }).trim()
}

function selfCheck () {
	const assert = (actual, expected, label) => {
		const [a, b] = [JSON.stringify(actual), JSON.stringify(expected)]
		if (a !== b) throw new Error(`changelog self-check: ${label}\n  expected ${b}\n  actual   ${a}`)
	}
	assert(classify('feat: add zoom objects', 'abc1234')?.heading, 'Added', 'feat maps to Added')
	assert(classify('fix(charts): stop repair prompt', 'abc1234')?.heading, 'Fixed', 'scoped fix maps to Fixed')
	assert(classify('refactor!: drop legacy option', 'abc1234')?.heading, 'Changed', 'breaking maps to Changed')
	assert(classify('refactor!: drop legacy option', '').text, '- **Breaking:** Drop legacy option', 'breaking is marked and capitalized')
	assert(classify('chore: release v4.2.0', 'abc1234'), null, 'release commits are skipped')
	assert(classify('docs: fix a typo', 'abc1234'), null, 'plumbing types are skipped')
	assert(classify('not a conventional commit', 'abc1234'), null, 'unparseable subjects are skipped')
	assert(renderSection('v1.0.0', '2026-01-01', []).includes('no user-facing changes'), true, 'empty release still gets a section')
	assert(
		insertSection('# Changelog\n\nintro\n\n## [v1](x) - old\n', '## [v2](y) - new\n'),
		'# Changelog\n\nintro\n\n## [v2](y) - new\n\n## [v1](x) - old\n',
		'new section goes above the newest existing one'
	)
	assert(insertSection('# Changelog\n\nintro\n', '## [v1](y) - new\n'), '# Changelog\n\nintro\n\n## [v1](y) - new\n', 'first-ever section appends')
	console.log('changelog self-check: ok')
}

function main () {
	const args = process.argv.slice(2)
	if (args.includes('--self-check')) return selfCheck()

	const read = name => {
		const at = args.indexOf(`--${name}`)
		if (at === -1 || !args[at + 1]) throw new Error(`changelog: --${name} is required`)
		return args[at + 1]
	}
	const dir = read('dir')
	const tag = read('tag')
	const date = new Date().toISOString().slice(0, 10)

	const changelog = readFileSync(CHANGELOG, 'utf8')
	if (changelog.includes(`## [${tag}]`)) {
		console.log(`changelog: ${tag} is already documented - nothing to do`)
		return
	}

	// The previous release of *this* package: tags share the repo, so match the same prefix.
	const prefix = tag.replace(/[0-9].*$/, '')
	const previous = git('tag', '--list', `${prefix}[0-9]*`, '--sort=-v:refname').split('\n').filter(Boolean)[0]
	const range = previous ? `${previous}..HEAD` : 'HEAD'
	const log = git('log', range, '--no-merges', '--format=%h%x00%s', '--', dir)

	const entries = log
		.split('\n')
		.filter(Boolean)
		.map(line => {
			const [hash, subject] = line.split('\0')
			return classify(subject, hash)
		})
		.filter(entry => entry !== null)

	writeFileSync(CHANGELOG, insertSection(changelog, renderSection(tag, date, entries)))
	console.log(`changelog: added ${tag} (${entries.length} entries from ${range} in ${dir})`)
}

main()
