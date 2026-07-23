/**
 * PptxGenJS: Media Methods
 */

import { IMG_BROKEN } from './core-enums'
import { PresSlide, SlideLayout, ISlideRelMedia } from './core-interfaces'

/**
 * Encode Image/Audio/Video into base64
 * @param {PresSlide | SlideLayout} layout - slide layout
 * @return {Promise} promise
 */
export function encodeSlideMediaRels(layout: PresSlide | SlideLayout): Array<Promise<string>> {
	// STEP 1: Detect real Node runtime once
	const isNode = typeof process !== 'undefined' && !!process.versions?.node && process.release?.name === 'node'
	// These will be filled only when we’re in Node
	let fs: typeof import('node:fs') | undefined
	let http: typeof import('node:http') | undefined
	let https: typeof import('node:https') | undefined

	// STEP 2: Lazy-load Node built-ins if needed
	const loadNodeDeps = isNode
		? async () => {
			; ({ default: fs } = await import('node:fs')); ({ default: http } = await import('node:http')); ({ default: https } = await import('node:https'))
		}
		: async () => { }
	// Immediately start it when we know we’re in Node
	if (isNode) loadNodeDeps()

	// STEP 3: Prepare promises list
	const imageProms: Array<Promise<string>> = []

	// A: Capture all audio/image/video candidates for encoding (filtering online/pre-encoded)
	const candidateRels = layout._relsMedia.filter(
		rel => rel.type !== 'online' && !rel.data && (!rel.path || (rel.path && !rel.path.includes('preencoded')))
	)

	// B: PERF: Mark dupes (same `path`) to avoid loading the same media over-and-over!
	const unqPaths: string[] = []
	candidateRels.forEach(rel => {
		const relPath = rel.path ?? ''
		if (!unqPaths.includes(relPath)) {
			rel.isDuplicate = false
			unqPaths.push(relPath)
		} else {
			rel.isDuplicate = true
		}
	})

	// STEP 4: Read/Encode each unique media item
	candidateRels
		.filter(rel => !rel.isDuplicate)
		.forEach(rel => {
			imageProms.push(
				(async () => {
					if (!https) await loadNodeDeps()

					const relPath = rel.path ?? ''

					// ────────────  NODE LOCAL FILE  ────────────
					if (isNode && fs && relPath.indexOf('http') !== 0) {
						try {
							const bitmap = fs.readFileSync(relPath)
							rel.data = Buffer.from(bitmap).toString('base64')
							candidateRels
								.filter(dupe => dupe.isDuplicate && dupe.path === rel.path)
								.forEach(dupe => (dupe.data = rel.data))
							return 'done'
						} catch (ex) {
							rel.data = IMG_BROKEN
							candidateRels
								.filter(dupe => dupe.isDuplicate && dupe.path === rel.path)
								.forEach(dupe => (dupe.data = rel.data))
							throw new Error(`ERROR: Unable to read media: "${rel.path}"\n${String(ex)}`)
						}
					}

					// ────────────  NODE HTTP(S)  ────────────
					if (isNode && https && http && relPath.startsWith('http')) {
						const reqMod = relPath.startsWith('https:') ? https : http
						return await new Promise<string>((resolve, reject) => {
							const markBroken = (err: Error): void => {
								rel.data = IMG_BROKEN
								candidateRels
									.filter(dupe => dupe.isDuplicate && dupe.path === rel.path)
									.forEach(dupe => (dupe.data = rel.data))
								reject(err)
							}
							const req = reqMod.get(relPath, res => {
								if (!res.statusCode || res.statusCode < 200 || res.statusCode > 299) {
									res.resume() // drain so the socket is freed
									markBroken(new Error(`ERROR! HTTP status ${res.statusCode} loading image: ${rel.path}`))
									return
								}
								let raw = ''
								res.setEncoding('binary') // IMPORTANT: Only binary encoding works
								res.on('data', chunk => (raw += chunk))
								res.on('end', () => {
									rel.data = Buffer.from(raw, 'binary').toString('base64')
									candidateRels
										.filter(dupe => dupe.isDuplicate && dupe.path === rel.path)
										.forEach(dupe => (dupe.data = rel.data))
									resolve('done')
								})
								res.on('error', () => markBroken(new Error(`ERROR! Unable to load image (response error): ${rel.path}`)))
							})
							// Without this listener a DNS/connection/TLS failure is an uncaught 'error' event that kills the process
							req.on('error', err => markBroken(new Error(`ERROR! Unable to load image (request error): ${rel.path}\n${String(err)}`)))
						})
					}

					// ────────────  BROWSER  ────────────
					return await new Promise<string>((resolve, reject) => {
						// A: build request
						const xhr = new XMLHttpRequest()
						xhr.onload = () => {
							// status 0 = non-HTTP schemes (file://); anything outside 2xx is an error page, not image bytes
							if (xhr.status !== 0 && (xhr.status < 200 || xhr.status > 299)) {
								rel.data = IMG_BROKEN
								candidateRels
									.filter(dupe => dupe.isDuplicate && dupe.path === rel.path)
									.forEach(dupe => (dupe.data = rel.data))
								reject(new Error(`ERROR! HTTP status ${xhr.status} loading image: ${rel.path}`))
								return
							}
							const reader = new FileReader()
							reader.onloadend = () => {
								if (typeof reader.result === 'string') rel.data = reader.result
								candidateRels
									.filter(dupe => dupe.isDuplicate && dupe.path === rel.path)
									.forEach(dupe => (dupe.data = rel.data))
								if (!rel.isSvgPng) {
									resolve('done')
								} else {
									createSvgPngPreview(rel)
										.then(() => resolve('done'))
										.catch(reject)
								}
							}
							reader.readAsDataURL(xhr.response)
						}
						xhr.onerror = () => {
							rel.data = IMG_BROKEN
							candidateRels
								.filter(dupe => dupe.isDuplicate && dupe.path === rel.path)
								.forEach(dupe => (dupe.data = rel.data))
							reject(new Error(`ERROR! Unable to load image (xhr.onerror): ${rel.path}`))
						}
						// B: execute request
						xhr.open('GET', relPath)
						xhr.responseType = 'blob'
						xhr.send()
					})
				})(),
			)
		})

	// STEP 5: SVG-PNG previews
	// ......: "SVG:" base64 data still requires a png to be generated
	// ......: (`isSvgPng` flag this as the preview image, not the SVG itself)
	layout._relsMedia
		.filter(rel => rel.isSvgPng && rel.data)
		.forEach(rel => {
			// Must push synchronously: the caller copies this array as soon as we return, so an async push would be lost (race)
			if (isNode) {
				// SVG preview rendering needs a DOM canvas; not supported in Node (https://github.com/gitbrent/PptxGenJS/issues/401)
				rel.data = IMG_BROKEN
				imageProms.push(Promise.resolve('done'))
			} else {
				imageProms.push(createSvgPngPreview(rel))
			}
		})

	return imageProms
}

/**
 * Create SVG preview image
 * @param {ISlideRelMedia} rel - slide rel
 * @return {Promise} promise
 */
async function createSvgPngPreview(rel: ISlideRelMedia): Promise<string> {
	return await new Promise((resolve, reject) => {
		// A: Create
		const image = new Image()

		// Shared error handler (also wired to `image.onerror`); the reason string is informational only
		const handleError = (): void => {
			rel.data = IMG_BROKEN
			reject(new Error(`ERROR! Unable to load image (image.onerror): ${rel.path}`))
		}

		// B: Set onload event
		image.onload = () => {
			// First: Check for any errors: This is the best method (try/catch wont work, etc.)
			if (image.width + image.height === 0) {
				handleError()
			}
			const canvas = document.createElement('canvas')
			const ctx = canvas.getContext('2d')
			if (!ctx) {
				handleError()
				return
			}
			canvas.width = image.width
			canvas.height = image.height
			ctx.drawImage(image, 0, 0)
			// Users running on local machine will get the following error:
			// "SecurityError: Failed to execute 'toDataURL' on 'HTMLCanvasElement': Tainted canvases may not be exported."
			// when the canvas.toDataURL call executes below.
			try {
				rel.data = canvas.toDataURL(rel.type)
				resolve('done')
			} catch (_ex) {
				handleError()
			}
		}
		image.onerror = handleError

		// C: Load image
		image.src = typeof rel.data === 'string' ? rel.data : IMG_BROKEN
	})
}
