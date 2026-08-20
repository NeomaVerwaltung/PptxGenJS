/**
 * PptxGenJS: Slide animation and media playback timing tree (`<p:timing>`)
 *
 * A slide has at most one timing tree. Object animations live in a `mainSeq` sequence and media
 * playback nodes are its siblings under the `tmRoot` common time node (ECMA-376 19.5).
 *
 * Scope: the animation presets below are the ones expressible as a visibility `p:set` plus a
 * `p:animEffect` filter (ECMA-376 19.5.35 `CT_TLAnimateEffectBehavior`). Motion-path and emphasis
 * effects need `p:animMotion`/`p:anim` behaviours and are deliberately not supported yet.
 */

import { SLIDE_OBJECT_TYPES } from '../core-enums'
import { AnimationProps, PresSlide } from '../core-interfaces'

/** Default effect length in milliseconds, matching PowerPoint's own default */
const DEF_DURATION = 500

/**
 * Supported presets: `presetID` and `presetClass` are the values PowerPoint writes for the effect,
 * and `filter` is the `p:animEffect@filter` that produces it. `appear`/`disappear` are instant, so
 * they have no filter at all.
 */
interface Preset {
	/** `p:cTn@presetID` PowerPoint writes for this effect */
	id: number
	/** `p:cTn@presetClass` */
	class: 'entr' | 'exit'
	/** `p:animEffect@filter`; absent for the instant `appear`/`disappear` effects */
	filter?: string
	/** direction name -> [`presetSubtype`, `filter`]; both encode the direction, so they stay together */
	directions?: Record<string, [number, string]>
	/** which `directions` key applies when the caller gives none */
	defaultDirection?: string
}

const WIPE_DIRECTIONS: Record<string, [number, string]> = {
	up: [1, 'wipe(up)'],
	right: [2, 'wipe(right)'],
	down: [4, 'wipe(down)'],
	left: [8, 'wipe(left)'],
}
const ZOOM_DIRECTIONS: Record<string, [number, string]> = {
	in: [16, 'zoom(in)'],
	out: [32, 'zoom(out)'],
}

const PRESETS: Record<string, Preset> = {
	appear: { id: 1, class: 'entr' },
	disappear: { id: 1, class: 'exit' },
	fadeIn: { id: 10, class: 'entr', filter: 'fade' },
	fadeOut: { id: 10, class: 'exit', filter: 'fade' },
	wipeIn: { id: 22, class: 'entr', directions: WIPE_DIRECTIONS, defaultDirection: 'up' },
	wipeOut: { id: 22, class: 'exit', directions: WIPE_DIRECTIONS, defaultDirection: 'up' },
	zoomIn: { id: 53, class: 'entr', directions: ZOOM_DIRECTIONS, defaultDirection: 'in' },
	zoomOut: { id: 53, class: 'exit', directions: ZOOM_DIRECTIONS, defaultDirection: 'out' },
}

/** `p:cTn@nodeType` for each trigger */
const NODE_TYPES: Record<string, string> = {
	onClick: 'clickEffect',
	withPrevious: 'withEffect',
	afterPrevious: 'afterEffect',
}

/** An animation resolved against the object it belongs to */
interface ResolvedAnimation {
	/** `p:spTgt@spid` - the target's `p:cNvPr@id` */
	spid: number
	preset: Preset
	presetSubtype: number
	filter?: string
	nodeType: string
	trigger: string
	delay: number
	duration: number
}

/** Coerce a caller-supplied millisecond value, dropping negatives and non-finite input */
function msValue (value: unknown, fallback: number): number {
	if (typeof value !== 'number' || !isFinite(value) || value < 0) return fallback
	return Math.round(value)
}

/**
 * Resolve every animation on a slide against its target shape
 * - invalid presets/triggers warn and are skipped, so they can never reach the XML
 * @param {PresSlide} slide - slide object
 * @returns {ResolvedAnimation[]} resolved animations in declaration order
 */
function resolveAnimations (slide: PresSlide): ResolvedAnimation[] {
	const resolved: ResolvedAnimation[] = []

	;(slide._slideObjects ?? []).forEach((obj, idx) => {
		const anim = obj.options?.animation as AnimationProps | undefined
		if (!anim) return

		const preset = PRESETS[anim.type]
		if (!preset) {
			console.warn(`[pptxgenjs] unknown animation "${String(anim.type)}" - supported presets are ${Object.keys(PRESETS).join(', ')}; animation ignored`)
			return
		}
		const trigger = anim.trigger ?? 'onClick'
		if (!NODE_TYPES[trigger]) {
			console.warn(`[pptxgenjs] animation trigger must be 'onClick' | 'withPrevious' | 'afterPrevious' - "${String(anim.trigger)}" ignored`)
			return
		}

		let presetSubtype = 0
		let filter = preset.filter
		if (preset.directions) {
			const dir = anim.direction && preset.directions[anim.direction] ? anim.direction : undefined
			if (anim.direction && !dir) {
				console.warn(`[pptxgenjs] animation "${anim.type}" does not accept direction "${String(anim.direction)}" - valid values are ${Object.keys(preset.directions).join(', ')}; value ignored`)
			}
			const [subtype, directionFilter] = preset.directions[dir ?? String(preset.defaultDirection)]
			presetSubtype = subtype
			filter = directionFilter
		}

		resolved.push({
			// media shapes derive their `cNvPr@id` from `mediaRid`; every other object uses its index
			spid: obj._type === SLIDE_OBJECT_TYPES.media ? (obj.mediaRid ?? 0) + 2 : idx + 2,
			preset,
			presetSubtype,
			filter,
			nodeType: NODE_TYPES[trigger],
			trigger,
			delay: msValue(anim.delay, 0),
			duration: msValue(anim.duration, DEF_DURATION),
		})
	})

	return resolved
}

/**
 * Build the behaviours of one effect: a visibility `p:set` plus the optional `p:animEffect` filter
 * @param {ResolvedAnimation} anim - resolved animation
 * @param {number} nodeId - first node id available to this effect
 * @returns {string} XML string
 */
function effectBehaviours (anim: ResolvedAnimation, nodeId: number): string {
	const isExit = anim.preset.class === 'exit'
	const target = `<p:tgtEl><p:spTgt spid="${anim.spid}"/></p:tgtEl>`
	// An exit effect hides the shape when it ends; an entrance shows it when it starts
	const setDelay = isExit ? anim.duration : 0
	let xml =
		'<p:set><p:cBhvr>' +
		`<p:cTn id="${nodeId}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${setDelay}"/></p:stCondLst></p:cTn>` +
		target +
		'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>' +
		'</p:cBhvr>' +
		`<p:to><p:strVal val="${isExit ? 'hidden' : 'visible'}"/></p:to>` +
		'</p:set>'

	if (anim.filter) {
		xml +=
			`<p:animEffect transition="${isExit ? 'out' : 'in'}" filter="${anim.filter}">` +
			`<p:cBhvr><p:cTn id="${nodeId + 1}" dur="${anim.duration}"/>${target}</p:cBhvr>` +
			'</p:animEffect>'
	}

	return xml
}

/**
 * Build the `mainSeq` sequence holding every object animation on the slide
 * @param {ResolvedAnimation[]} anims - resolved animations
 * @returns {string} XML string
 */
function mainSequence (anims: ResolvedAnimation[]): { xml: string, nextNodeId: number } {
	// Each `onClick` effect starts a click group; `withPrevious`/`afterPrevious` effects join the
	// group before them, which is what makes them run with (or after) that click's animation.
	const groups: ResolvedAnimation[][] = []
	anims.forEach(anim => {
		if (anim.trigger === 'onClick' || groups.length === 0) groups.push([anim])
		else groups[groups.length - 1].push(anim)
	})

	let nodeId = 3
	let xml = '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'

	groups.forEach(group => {
		xml += `<p:par><p:cTn id="${nodeId++}" fill="hold">`
		xml += '<p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst>'
		xml += '<p:childTnLst>'
		xml += `<p:par><p:cTn id="${nodeId++}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>`

		group.forEach(anim => {
			const effectId = nodeId
			// reserve the effect node plus its two behaviour nodes
			nodeId += 3
			xml +=
				'<p:par>' +
				`<p:cTn id="${effectId}" presetID="${anim.preset.id}" presetClass="${anim.preset.class}" presetSubtype="${anim.presetSubtype}" fill="hold" grpId="0" nodeType="${anim.nodeType}">` +
				`<p:stCondLst><p:cond delay="${anim.delay}"/></p:stCondLst>` +
				`<p:childTnLst>${effectBehaviours(anim, effectId + 1)}</p:childTnLst>` +
				'</p:cTn></p:par>'
		})

		xml += '</p:childTnLst></p:cTn></p:par>'
		xml += '</p:childTnLst></p:cTn></p:par>'
	})

	xml += '</p:childTnLst></p:cTn>'
	xml += '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
	xml += '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
	xml += '</p:seq>'

	return { xml, nextNodeId: nodeId }
}

/**
 * Build the `<p:video>`/`<p:audio>` media playback nodes for a slide
 * - `p:spTgt@spid` targets the media shape's `p:cNvPr@id`, which the media emitter derives from `mediaRid`
 * @param {PresSlide} slide - slide object
 * @param {number} startNodeId - first node id available to the media nodes
 * @returns {string} XML string
 */
function mediaNodes (slide: PresSlide, startNodeId: number): string {
	let nodeId = startNodeId

	return (slide._slideObjects ?? [])
		.filter(obj => obj._type === SLIDE_OBJECT_TYPES.media && obj.mtype !== 'online')
		.filter(obj => obj.options?.autoplay || obj.options?.loop || obj.options?.fullScreen || obj.options?.mute)
		.map(obj => {
			const opts = obj.options ?? {}
			const tag = obj.mtype === 'audio' ? 'p:audio' : 'p:video'
			const fullScrn = tag === 'p:video' && opts.fullScreen ? ' fullScrn="1"' : ''
			return (
				`<${tag}${fullScrn}>` +
				`<p:cMediaNode${opts.mute ? ' mute="1"' : ''}>` +
				`<p:cTn id="${nodeId++}" fill="hold" display="0"${opts.loop ? ' repeatCount="indefinite"' : ''}>` +
				`<p:stCondLst><p:cond delay="${opts.autoplay ? '0' : 'indefinite'}"/></p:stCondLst>` +
				'</p:cTn>' +
				`<p:tgtEl><p:spTgt spid="${(obj.mediaRid ?? 0) + 2}"/></p:tgtEl>` +
				'</p:cMediaNode>' +
				`</${tag}>`
			)
		})
		.join('')
}

/**
 * Create the slide timing tree (`<p:timing>`) for object animations and media playback
 * - returns `''` when the slide has neither, so slides without them are unchanged
 * @note belongs after `<p:transition>` in the CT_Slide sequence
 * @param {PresSlide} slide - slide object
 * @returns {string} XML
 */
export function genXmlTiming (slide: PresSlide): string {
	const anims = resolveAnimations(slide)
	// node id 1 is tmRoot; the mainSeq (id 2) and its effects allocate from there
	const sequence = anims.length > 0 ? mainSequence(anims) : { xml: '', nextNodeId: 2 }
	const media = mediaNodes(slide, sequence.nextNodeId)
	if (!sequence.xml && !media) return ''

	let xml = '<p:timing><p:tnLst><p:par>'
	xml += '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
	xml += sequence.xml
	xml += media
	xml += '</p:childTnLst></p:cTn>'
	xml += '</p:par></p:tnLst>'
	// The build list tells PowerPoint which shapes participate in the click sequence
	if (anims.length > 0) {
		const spids = [...new Set(anims.map(anim => anim.spid))]
		xml += `<p:bldLst>${spids.map(spid => `<p:bldP spid="${spid}" grpId="0"/>`).join('')}</p:bldLst>`
	}
	xml += '</p:timing>'

	return xml
}
