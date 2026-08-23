/**
 * Gradient fills on a slide background, a shape, a line, and a table cell.
 */
import type PptxGenJS from '../../src/pptxgen'

export function addGradientsFixture (pptx: PptxGenJS): void {
	const gradientSlide = pptx.addSlide()
	gradientSlide.background = { type: 'gradient', gradient: { angle: 45, stops: [{ color: 'FFFFFF', position: 0 }, { color: 'E7E6E6', position: 100 }] } }
	gradientSlide.addShape(pptx.ShapeType.rect, {
		x: 0.5,
		y: 0.5,
		w: 4,
		h: 2,
		fill: { type: 'gradient', gradient: { stops: [{ color: 'FF0000', position: 0 }, { color: '0000FF', position: 100, transparency: 20 }] } },
		line: { type: 'gradient', width: 2, gradient: { type: 'radial', stops: [{ color: '00FF00', position: 0 }, { color: '000000', position: 100 }] } },
	})
	gradientSlide.addTable([[{ text: 'gradient cell', options: { fill: { type: 'gradient', gradient: { stops: [{ color: '111111', position: 0 }, { color: '888888', position: 100 }] } } } }]], { x: 0.5, y: 3, w: 4 })
}
