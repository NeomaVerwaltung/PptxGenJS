/**
 * Presentation-level properties: an embedded font, slide-show mode, image DPI, read-only
 * recommendation, and slide/notes guides.
 */
import type PptxGenJS from '../../src/pptxgen'
import { fakeEot } from '../fixtures'

export function addPresentationFixture (pptx: PptxGenJS): void {
	pptx.addFont({ fontFace: 'Smoke Sans', data: fakeEot() })
	pptx.slideShow = { mode: 'browse', loop: true, browseMode: true, laserColor: 'FF0000' }
	pptx.defaultImageDpi = 220
	pptx.readonlyRecommended = true
	pptx.guides = [{ orientation: 'vert', position: 5 }, { orientation: 'horz', position: 3.75 }]
	pptx.notesGuides = [{ orientation: 'horz', position: 1 }]
}
