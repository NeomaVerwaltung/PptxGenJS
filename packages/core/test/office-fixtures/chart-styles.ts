/**
 * Chart style and chart colour-style parts - the round-trip fixture for #157.
 */
import type PptxGenJS from '../../src/pptxgen'

export function addChartStylesFixture (pptx: PptxGenJS): void {
	pptx.addSlide().addChart('bar', [{ name: 'Series 1', labels: ['a', 'b', 'c'], values: [4, 7, 2] }], { x: 0.5, y: 0.5, w: 6, h: 3, chartStyle: 201, chartColorStyle: { method: 'cycle', id: 10 } })
}
