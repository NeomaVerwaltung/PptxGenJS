/**
 * A custom table style definition - the round-trip fixture for #133, proving banding renders.
 */
import type PptxGenJS from '../../src/pptxgen'

export function addTableStylesFixture (pptx: PptxGenJS): void {
	pptx.tableStyles = [{
		id: '{A1B2C3D4-1111-2222-3333-444455556666}',
		name: 'Office Smoke',
		wholeTable: { borders: { top: { color: '4472C4', width: 1 }, bottom: { color: '4472C4', width: 1 }, insideH: { color: 'D9D9D9', width: 0.5 } } },
		band1H: { fill: { color: 'DEEAF6' } },
		firstRow: { bold: true, color: 'FFFFFF', fill: { color: '4472C4' } },
	}]
	pptx.addSlide().addTable(
		[['Region', 'Sales'], ['West', '20'], ['East', '31'], ['North', '18']],
		{ x: 0.5, y: 0.5, w: 8, tableStyleId: '{A1B2C3D4-1111-2222-3333-444455556666}', firstRow: true, bandRow: true }
	)
}
