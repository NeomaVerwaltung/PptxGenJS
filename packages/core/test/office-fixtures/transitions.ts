/**
 * A modern (morph) slide transition, which carries a base-transition fallback.
 */
import type PptxGenJS from '../../src/pptxgen'

export function addTransitionsFixture (pptx: PptxGenJS): void {
	pptx.addSlide({ transition: { type: 'morph', duration: 1200 } }).addText('modern transition', { x: 0.5, y: 0.5, w: 5, h: 0.5 })
}
