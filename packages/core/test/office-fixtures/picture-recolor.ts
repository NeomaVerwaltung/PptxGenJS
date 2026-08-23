/**
 * Picture recolour effects - the round-trip fixture for #131.
 */
import type PptxGenJS from '../../src/pptxgen'
import { SAMPLE_PNG_4X2 } from '../fixtures'

export function addPictureRecolorFixture (pptx: PptxGenJS): void {
	const recolorSlide = pptx.addSlide()
	recolorSlide.addImage({ data: SAMPLE_PNG_4X2, x: 0.5, y: 0.5, w: 2, h: 1, recolor: { duotone: ['1F3864', 'DEEAF6'] } })
	recolorSlide.addImage({ data: SAMPLE_PNG_4X2, x: 3, y: 0.5, w: 2, h: 1, recolor: { grayscale: true, brightness: 20, contrast: -15 } })
	recolorSlide.addImage({ data: SAMPLE_PNG_4X2, x: 5.5, y: 0.5, w: 2, h: 1, recolor: { colorChange: { from: 'FFFFFF', to: 'FF0000' } } })
}
