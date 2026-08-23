/**
 * Diagonal cell borders, 3-D cells, and vertical cell text - the round-trip fixture for #147.
 */
import type PptxGenJS from '../../src/pptxgen'

export function addTableCellsFixture (pptx: PptxGenJS): void {
	const cellSlide = pptx.addSlide()
	cellSlide.addTable(
		[[
			{ text: 'diagonals', options: { borderDiagonalDown: { color: 'FF0000', width: 2 }, borderDiagonalUp: { type: 'dash', color: '0000FF' } } },
			{ text: 'bevelled', options: { cell3D: { bevel: { preset: 'circle', width: 0.05, height: 0.05 }, material: 'metal', lightRig: { rig: 'threePt', dir: 't' } } } },
			{ text: 'rotated', options: { textDirection: 'vert270', anchorCtr: true, horzOverflow: 'overflow' } },
		]],
		{ x: 0.5, y: 0.5, w: 8, h: 1.5, rtl: true, border: [{ type: 'solid', color: '999999', width: 1 }, { type: 'solid', color: '999999', width: 1 }, { type: 'solid', color: '999999', width: 1 }, { type: 'solid', color: '999999', width: 1 }] }
	)
}
