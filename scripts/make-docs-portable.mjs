import { readdir, readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'

const output = 'docs/.vitepress/dist'

async function relativize(directory) {
	for (const entry of await readdir(directory, { withFileTypes: true })) {
		const file = join(directory, entry.name)
		if (entry.isDirectory()) await relativize(file)
		else if (entry.name.endsWith('.html') || entry.name.endsWith('.js')) {
			const content = await readFile(file, 'utf8')
			await writeFile(file, content.replaceAll('"/assets/', '"./assets/').replaceAll('"/neoma-', '"./neoma-'))
		}
	}
}

await relativize(output)
