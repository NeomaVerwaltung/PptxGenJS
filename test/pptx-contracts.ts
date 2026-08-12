/**
 * Executable contracts for generated OOXML packages.
 *
 * These deliberately validate package semantics instead of comparing whole XML documents. A
 * harmless serializer change should not require replacing a checked-in XML baseline, while a
 * broken relationship, missing part, duplicate rId, or malformed document still fails here.
 */
import assert from 'node:assert/strict'
import { posix } from 'node:path'
import JSZip from 'jszip'
import { XMLValidator } from 'fast-xml-parser'

const REQUIRED_PARTS = [
	'[Content_Types].xml',
	'_rels/.rels',
	'ppt/presentation.xml',
	'ppt/_rels/presentation.xml.rels',
]

/** Read a generated OOXML part and fail with its package path when it is absent. */
export async function readPart (zip: JSZip, name: string): Promise<string> {
	const part = zip.file(name)
	assert.ok(part, `missing package part: ${name}`)
	return await part.async('string')
}

function sourcePartForRelationships (relationshipPart: string): string | undefined {
	if (relationshipPart === '_rels/.rels') return undefined
	const marker = '/_rels/'
	const markerIndex = relationshipPart.lastIndexOf(marker)
	if (markerIndex === -1 || !relationshipPart.endsWith('.rels')) return undefined
	return relationshipPart.slice(0, markerIndex) + '/' + relationshipPart.slice(markerIndex + marker.length, -'.rels'.length)
}

function resolveTarget (sourcePart: string | undefined, target: string): string {
	if (target.startsWith('/')) return target.slice(1)
	const baseDir = sourcePart ? posix.dirname(sourcePart) : '.'
	return posix.normalize(posix.join(baseDir, target)).replace(/^\.\//, '')
}

/**
 * Validate the stable OOXML package invariants that whole-document snapshots used to cover.
 * Keep feature-specific XML assertions next to the test that creates that feature.
 */
export async function assertPptxPackageContracts (zip: JSZip): Promise<void> {
	for (const name of REQUIRED_PARTS) assert.ok(zip.file(name), `missing required package part: ${name}`)

	const packageFiles = new Set(Object.keys(zip.files))
	for (const name of packageFiles) {
		if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue
		const xml = await readPart(zip, name)
		assert.equal(XMLValidator.validate(xml), true, `malformed XML in ${name}`)
	}

	const contentTypes = await readPart(zip, '[Content_Types].xml')
	for (const match of contentTypes.matchAll(/<Override\s+PartName="([^"]+)"/g)) {
		const partName = match[1].replace(/^\//, '')
		assert.ok(packageFiles.has(partName), `content type points to missing package part: ${partName}`)
	}

	const relationshipIdsBySource = new Map<string, Set<string>>()
	for (const name of packageFiles) {
		if (!name.endsWith('.rels')) continue
		const sourcePart = sourcePartForRelationships(name)
		if (sourcePart) assert.ok(packageFiles.has(sourcePart), `relationship source is missing: ${sourcePart}`)

		const relationshipXml = await readPart(zip, name)
		const ids = new Set<string>()
		for (const match of relationshipXml.matchAll(/<Relationship\s+([^>]+?)\/?\s*>/g)) {
			const attributes = match[1]
			const id = /\bId="([^"]+)"/.exec(attributes)?.[1]
			const target = /\bTarget="([^"]+)"/.exec(attributes)?.[1]
			const targetMode = /\bTargetMode="([^"]+)"/.exec(attributes)?.[1]
			assert.ok(id, `relationship without Id in ${name}`)
			assert.ok(target, `relationship without Target in ${name}`)
			assert.ok(!ids.has(id), `duplicate relationship Id ${id} in ${name}`)
			ids.add(id)
			if (targetMode === 'External') continue

			const targetPart = resolveTarget(sourcePart, target)
			assert.ok(packageFiles.has(targetPart), `relationship target is missing: ${name} -> ${targetPart}`)
		}
		if (sourcePart) relationshipIdsBySource.set(sourcePart, ids)
	}

	for (const [sourcePart, relationshipIds] of relationshipIdsBySource) {
		const sourceXml = await readPart(zip, sourcePart)
		for (const match of sourceXml.matchAll(/\br:(?:id|embed|link)="(rId\d+)"/g)) {
			assert.ok(relationshipIds.has(match[1]), `missing ${match[1]} relationship for ${sourcePart}`)
		}
	}
}
