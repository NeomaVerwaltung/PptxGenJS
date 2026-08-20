/**
 * PptxGenJS: Slide transition XML (`<p:transition>`)
 *
 * ECMA-376 19.3.1.50 `CT_SlideTransition` defines the base transition set. PowerPoint 2010+ adds
 * further transitions in its own namespaces; MS-PPTX 2.2.1 requires those to be offered through
 * `mc:AlternateContent` so that consumers which do not understand the extension read a base
 * transition from the `mc:Fallback` instead of failing.
 */

import { MS_PPTX_NS } from '../core-enums'
import { alternateContent } from './markup-compat'
import { PresSlide, SlideTransitionProps } from '../core-interfaces'

/** Which directional attribute an inner transition element accepts (ECMA-376 19.3.1.50) */
type DirKind = 'none' | 'side' | 'orient' | 'eight' | 'corner' | 'inout' | 'split' | 'wheel' | 'thruBlk'

/** Base ECMA-376 transitions: element name -> the directional attribute it takes */
const BASE_TRANSITIONS: Record<string, DirKind> = {
	blinds: 'orient',
	checker: 'orient',
	circle: 'none',
	comb: 'orient',
	cover: 'eight',
	cut: 'thruBlk',
	diamond: 'none',
	dissolve: 'none',
	fade: 'thruBlk',
	newsflash: 'none',
	none: 'none',
	plus: 'none',
	pull: 'eight',
	push: 'side',
	random: 'none',
	randomBar: 'orient',
	split: 'split',
	strips: 'corner',
	wedge: 'none',
	wheel: 'wheel',
	wipe: 'side',
	zoom: 'inout',
}

/** PowerPoint 2010+ transitions: [namespace prefix, directional attribute, base transition to fall back to] */
const MODERN_TRANSITIONS: Record<string, ['p14' | 'p16', DirKind, string]> = {
	conveyor: ['p14', 'eight', 'push'],
	doors: ['p14', 'orient', 'split'],
	ferris: ['p14', 'eight', 'push'],
	flash: ['p14', 'none', 'fade'],
	flip: ['p14', 'eight', 'push'],
	flythrough: ['p14', 'inout', 'zoom'],
	gallery: ['p14', 'eight', 'push'],
	glitter: ['p14', 'eight', 'dissolve'],
	honeycomb: ['p14', 'none', 'dissolve'],
	morph: ['p16', 'none', 'fade'],
	pan: ['p14', 'side', 'push'],
	prism: ['p14', 'eight', 'push'],
	reveal: ['p14', 'side', 'fade'],
	ripple: ['p14', 'none', 'dissolve'],
	shred: ['p14', 'none', 'dissolve'],
	switch: ['p14', 'eight', 'push'],
	vortex: ['p14', 'side', 'push'],
	warp: ['p14', 'inout', 'zoom'],
	wheelReverse: ['p14', 'none', 'wheel'],
	window: ['p14', 'orient', 'split'],
}

/** Directions each `DirKind` accepts, as the OOXML tokens they must be written as */
const VALID_DIRS: Record<DirKind, string[]> = {
	none: [],
	side: ['l', 'r', 'u', 'd'],
	orient: ['horz', 'vert'],
	eight: ['l', 'r', 'u', 'd', 'lu', 'ru', 'ld', 'rd'],
	corner: ['lu', 'ru', 'ld', 'rd'],
	inout: ['in', 'out'],
	split: ['in', 'out'],
	wheel: [],
	thruBlk: [],
}

/** Friendly names accepted for `direction`, mapped to their OOXML token */
const DIR_ALIASES: Record<string, string> = {
	left: 'l', right: 'r', up: 'u', down: 'd',
	horizontal: 'horz', vertical: 'vert',
	leftup: 'lu', rightup: 'ru', leftdown: 'ld', rightdown: 'rd',
	topleft: 'lu', topright: 'ru', bottomleft: 'ld', bottomright: 'rd',
}

/** `wheel` spoke counts PowerPoint offers */
const WHEEL_SPOKES = [1, 2, 3, 4, 8]

/** Resolve a caller-supplied direction to its OOXML token (aliases and casing are accepted) */
function normalizeDir (value?: string): string | undefined {
	if (!value || typeof value !== 'string') return undefined
	const key = value.replace(/[-_\s]/g, '').toLowerCase()
	return DIR_ALIASES[key] ?? key
}

/**
 * Build the `dir` attribute for a transition element, dropping a direction the element cannot take
 * @param {DirKind} kind - directional attribute the element accepts
 * @param {SlideTransitionProps} props - transition props
 * @returns {string} XML attribute (may be empty)
 */
function dirAttribute (kind: DirKind, props: SlideTransitionProps): string {
	const dir = normalizeDir(props.direction)
	if (!dir) return ''
	if (!VALID_DIRS[kind].includes(dir)) {
		console.warn(`[pptxgenjs] transition "${String(props.type)}" does not accept direction "${String(props.direction)}" - valid values are ${VALID_DIRS[kind].join(', ') || '(none)'}; value ignored`)
		return ''
	}
	return ` dir="${dir}"`
}

/**
 * Build the inner element of a base transition, ex: `<p:wipe dir="l"/>`
 * @param {string} name - transition element name
 * @param {DirKind} kind - directional attribute the element accepts
 * @param {SlideTransitionProps} props - transition props
 * @returns {string} XML string
 */
function baseInner (name: string, kind: DirKind, props: SlideTransitionProps): string {
	if (kind === 'split') {
		// CT_SplitTransition takes both: `orient` (horz|vert) and `dir` (in|out)
		const dir = normalizeDir(props.direction)
		const orient = normalizeDir(props.orient) ?? (dir === 'vert' ? 'vert' : 'horz')
		return `<p:${name} orient="${orient === 'vert' ? 'vert' : 'horz'}" dir="${dir === 'in' ? 'in' : 'out'}"/>`
	}
	if (kind === 'wheel') {
		if (props.spokes !== undefined && !WHEEL_SPOKES.includes(props.spokes)) {
			console.warn(`[pptxgenjs] transition "wheel" spokes must be one of ${WHEEL_SPOKES.join(', ')} - value ignored`)
			return `<p:${name}/>`
		}
		return `<p:${name}${props.spokes ? ` spokes="${props.spokes}"` : ''}/>`
	}
	if (kind === 'thruBlk') return `<p:${name}${props.thruBlk ? ' thruBlk="1"' : ''}/>`
	if (kind === 'none') return `<p:${name}/>`
	return `<p:${name}${dirAttribute(kind, props)}/>`
}

/**
 * Build the attributes of `<p:transition>`
 * - `p14:dur` is the millisecond duration PowerPoint 2010+ uses; it supersedes the coarse `spd`, so
 *   it only ever appears in an `mc:Choice`. The `mc:Fallback` keeps `spd`, which every consumer reads.
 * @param {SlideTransitionProps} props - transition props
 * @param {boolean} modern - whether these attributes are for the `mc:Choice` (PowerPoint 2010+)
 * @returns {string} XML attributes
 */
function transitionAttributes (props: SlideTransitionProps, modern: boolean): string {
	let attrs = ''

	const hasDuration = typeof props.duration === 'number' && isFinite(props.duration) && props.duration > 0
	if (props.speed && !['slow', 'med', 'fast'].includes(props.speed)) {
		if (modern) console.warn(`[pptxgenjs] transition speed must be 'slow' | 'med' | 'fast' - "${String(props.speed)}" ignored`)
	} else if (props.speed) {
		if (!(modern && hasDuration)) attrs += ` spd="${props.speed}"`
	}
	if (modern && hasDuration) attrs += ` p14:dur="${Math.round(props.duration as number)}"`
	if (props.advClick === false) attrs += ' advClick="0"'
	if (typeof props.advTm === 'number' && isFinite(props.advTm) && props.advTm >= 0) attrs += ` advTm="${Math.round(props.advTm)}"`

	return attrs
}

/**
 * Create `<p:transition>` for a slide, or `''` when the slide has no transition
 * @note belongs after `<p:clrMapOvr>` and before `<p:timing>` in the CT_Slide sequence
 * @param {PresSlide} slide - slide object
 * @returns {string} XML string
 */
export function genXmlTransition (slide: PresSlide): string {
	const props = slide.transition
	if (!props?.type) return ''

	const type = String(props.type)
	const modern = MODERN_TRANSITIONS[type]
	if (!modern && !BASE_TRANSITIONS[type]) {
		console.warn(`[pptxgenjs] unknown slide transition "${type}" - transition ignored`)
		return ''
	}

	const choiceAttrs = transitionAttributes(props, true)
	const fallbackAttrs = transitionAttributes(props, false)
	const inner = modern
		? `<${modern[0]}:${type}${modern[1] === 'none' ? '' : dirAttribute(modern[1], props)}/>`
		: baseInner(type, BASE_TRANSITIONS[type], props)
	// A base transition falls back to itself; a modern one falls back to its closest base effect
	const fallbackType = modern ? modern[2] : type
	const fallback = baseInner(fallbackType, BASE_TRANSITIONS[fallbackType], modern ? { type: props.type } : props)

	// Plain transition: nothing here needs an extension namespace
	if (!modern && !choiceAttrs.includes('p14:dur')) return `<p:transition${fallbackAttrs}>${inner}</p:transition>`

	// Anything that needs `p14`/`p16` is offered with a fallback, per ECMA-376 Part 3 / MS-PPTX 2.2.1
	const prefix = modern ? modern[0] : 'p14'
	const namespaces: Record<string, string> = { [prefix]: MS_PPTX_NS[prefix] }
	// `p14:dur` needs its namespace in scope even when the effect itself is p16
	if (choiceAttrs.includes('p14:dur')) namespaces.p14 = MS_PPTX_NS.p14

	return alternateContent({
		namespaces,
		// only the effect's own prefix decides whether a consumer can render the choice
		requires: [prefix],
		choice: `<p:transition${choiceAttrs}>${inner}</p:transition>`,
		fallback: `<p:transition${fallbackAttrs}>${fallback}</p:transition>`,
	})
}
