/**
 * Linked video, audio CD, and WAV media sources - the round-trip fixture for #150. The linked
 * target is external, so it need not exist: the deck only has to open.
 */
import type PptxGenJS from '../../src/pptxgen'
import { SAMPLE_WAV } from '../fixtures'

export function addMediaSourcesFixture (pptx: PptxGenJS): void {
	const mediaSourceSlide = pptx.addSlide()
	mediaSourceSlide.addMedia({ type: 'video', link: 'clip.mp4', x: 0.5, y: 0.5, w: 3, h: 2, contentType: 'video/mp4' })
	mediaSourceSlide.addMedia({ type: 'audioCd', audioCd: { start: { track: 1 }, end: { track: 1, time: 30 } }, x: 4.5, y: 0.5, w: 2, h: 2 })
	mediaSourceSlide.addMedia({ type: 'wav', data: SAMPLE_WAV, x: 0.5, y: 3, w: 2, h: 1.5 })
}
