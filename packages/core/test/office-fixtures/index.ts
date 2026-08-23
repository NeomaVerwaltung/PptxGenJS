/**
 * Every round-trip fixture the LibreOffice consumer test builds, in the order it applies them.
 *
 * One file per feature area: adding a feature adds a file and one line here, rather than another
 * block inside a single shared function. The order matters - `slide-basics` targets `media` by
 * slide number - so new entries go at the end.
 */
import { addPresentationFixture } from './presentation'
import { addSlideBasicsFixture } from './slide-basics'
import { addMediaFixture } from './media'
import { addGradientsFixture } from './gradients'
import { addTableCellsFixture } from './table-cells'
import { addShapeEffectsFixture } from './shape-effects'
import { addLayoutsFixture } from './layouts'
import { addMediaSourcesFixture } from './media-sources'
import { addChartStylesFixture } from './chart-styles'
import { addTableStylesFixture } from './table-styles'
import { addChartexFixture } from './chartex'
import { addPictureRecolorFixture } from './picture-recolor'
import { addGroupsFixture } from './groups'
import { addTransitionsFixture } from './transitions'

export const OFFICE_FIXTURES = [
	addPresentationFixture,
	addSlideBasicsFixture,
	addMediaFixture,
	addGradientsFixture,
	addTableCellsFixture,
	addShapeEffectsFixture,
	addLayoutsFixture,
	addMediaSourcesFixture,
	addChartStylesFixture,
	addTableStylesFixture,
	addChartexFixture,
	addPictureRecolorFixture,
	addGroupsFixture,
	addTransitionsFixture,
]
