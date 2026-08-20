/**
 * The peer floor, checked against the core that is actually installed.
 *
 * Every other std test imports the core by relative path, so it always exercises the newest source
 * in this repo. This file imports it by package name instead, which is what a consumer resolves -
 * and what the `std-peer-floor` CI job points at the oldest core the manifest claims to support.
 * If a helper starts relying on a newer core feature, this fails there before it ships.
 */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import pptxgen from '@neo-ma/pptxgenjs'
import { waterfall } from '../src/charts'

const PACKAGE = JSON.parse(readFileSync(join(import.meta.dirname, '..', 'package.json'), 'utf8'))

const chart = async (): Promise<string> => {
	const pres = new pptxgen()
	waterfall(pres.addSlide(), { labels: ['Up', 'Down'], values: [40, -15] }, { x: 1, y: 1, w: 6, h: 4, showValue: true })
	const zip = await JSZip.loadAsync((await pres.write({ outputType: 'nodebuffer' })) as Buffer)
	const name = Object.keys(zip.files).find(file => /^ppt\/charts\/chart\d+\.xml$/.test(file))
	assert.ok(name, 'chart part missing')
	const part = await zip.file(name)?.async('string')
	assert.ok(part, 'chart part unreadable')
	return part
}

test('the installed core satisfies the declared peer range', () => {
	const floor = PACKAGE.peerDependencies['@neo-ma/pptxgenjs']
	const major = Number(floor.replace(/^\^/, '').split('.')[0])
	const installed = new pptxgen().version
	assert.equal(
		Number(installed.split('.')[0]), major,
		`installed core ${installed} is outside the declared peer range ${floor}`
	)
})

test('the installed core supports transparent series - the waterfall riser depends on it', async () => {
	const xml = await chart()
	const riser = xml.split('<c:ser>')[1]
	assert.match(
		riser.slice(0, riser.indexOf('<c:cat>')), /<a:noFill\/>/,
		'the riser series is not transparent: this core ignores per-series `color: "transparent"` (added in 4.1.0), so every bar sits on a visible block'
	)
})

test('the installed core supports per-point data labels - the waterfall labels depend on them', async () => {
	const xml = await chart()
	assert.match(
		xml, /<a:t>-15<\/a:t>/,
		'the decrease is not labelled with its sign: this core ignores per-point `dataLabels` (added in 4.2.0), so labels show the stacked magnitude instead of the delta'
	)
})
