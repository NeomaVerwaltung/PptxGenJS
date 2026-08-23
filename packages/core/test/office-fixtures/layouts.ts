/**
 * Slide layout and placeholder metadata - the round-trip fixture for #149.
 */
import type PptxGenJS from '../../src/pptxgen'

export function addLayoutsFixture (pptx: PptxGenJS): void {
	pptx.defineSlideMaster({
		title: 'SECTION HEADER',
		layoutType: 'secHead',
		matchingName: 'Section Header',
		colorMapOverride: { bg1: 'dk1', tx1: 'lt1' },
		objects: [{ placeholder: { options: { name: 'secTitle', type: 'title', x: 0.5, y: 2, w: 9, h: 1.5 }, text: 'Section' } }],
	})
	pptx.addSlide({ masterName: 'SECTION HEADER' }).addText('section', { placeholder: 'secTitle' })
}
