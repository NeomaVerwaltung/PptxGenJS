/**
 * The DrawingML effect vocabulary (#138) and theme style references (#141).
 */
import type PptxGenJS from '../../src/pptxgen'
import { SAMPLE_PNG_4X2 } from '../fixtures'

export function addShapeEffectsFixture (pptx: PptxGenJS): void {
	const effectSlide = pptx.addSlide()
	effectSlide.addShape(pptx.ShapeType.rect, {
		x: 0.5, y: 0.5, w: 3, h: 1.5, fill: { color: 'CCCCCC' },
		blur: { radius: 3 },
		fillOverlay: { blend: 'mult', fill: { color: 'FF0000', transparency: 40 } },
		shadow: { type: 'preset', preset: 'shdw7', color: '333333' },
	})
	effectSlide.addShape(pptx.ShapeType.ellipse, { x: 4.5, y: 0.5, w: 2, h: 1.5, fill: { color: '4472C4' }, glow: { size: 6, color: 'FFFF00', opacity: 0.6 }, effectDag: { type: 'sib' } })
	effectSlide.addImage({ data: SAMPLE_PNG_4X2, x: 0.5, y: 2.5, w: 2, h: 1, alphaEffects: { invert: true } })
	const styleSlide = pptx.addSlide()
	styleSlide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.5, w: 3, h: 1.5, styleRef: { line: 1, fill: 3, effect: 2, font: 'minor' } })
	styleSlide.addShape(pptx.ShapeType.ellipse, { x: 4.5, y: 0.5, w: 3, h: 1.5, styleRef: { fill: 1, color: 'accent2' } })
	styleSlide.addImage({ data: SAMPLE_PNG_4X2, x: 0.5, y: 2.5, w: 2, h: 1, styleRef: { line: 2 } })
}
