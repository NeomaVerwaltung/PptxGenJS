/**
 * Embedded audio and video with playback options, plus an animation that shares their timing tree.
 * This is slide 2, which `slide-basics` targets with a zoom and a hyperlink hover.
 */
import type PptxGenJS from '../../src/pptxgen'
import { SAMPLE_MP3, SAMPLE_MP4 } from '../fixtures'

export function addMediaFixture (pptx: PptxGenJS): void {
	const mediaSlide = pptx.addSlide()
	mediaSlide.addMedia({ type: 'video', data: SAMPLE_MP4, x: 0.5, y: 0.5, w: 4, h: 2.5, autoplay: true, loop: true, mute: true })
	mediaSlide.addMedia({ type: 'audio', data: SAMPLE_MP3, x: 5, y: 0.5, w: 2, h: 2, autoplay: true })
	mediaSlide.addText('animated with media', { x: 0.5, y: 3.5, w: 5, h: 0.5, animation: { type: 'wipeIn', direction: 'left', trigger: 'afterPrevious' } })
}
