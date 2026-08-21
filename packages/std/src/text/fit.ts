import { measureText, type MeasureProps, type Measurement } from './measure'

export interface FitTextArea {
	/** Available width, inches */
	w: number
	/** Available height, inches */
	h: number
}

export interface FitTextProps extends Omit<MeasureProps, 'w' | 'fontSize'> {
	/** Smallest size to consider, points @default 8 */
	min?: number
	/** Largest size to consider, points @default 40 */
	max?: number
	/**
	 * Text-box margin subtracted from the area before fitting, inches.
	 * The core's default cell margin is 0.05 vertical / 0.1 horizontal. @default 0
	 */
	margin?: number | [number, number, number, number]
}

export interface FitTextResult extends Measurement {
	/** The largest size in range whose wrapped text fits the area, points */
	fontSize: number
	/** True when even `min` overflows - the text is too long for the area at any size in range */
	overflows: boolean
}

/** TRBL inches, from a number or an explicit tuple */
function resolveMargin (margin: FitTextProps['margin']): [number, number, number, number] {
	if (margin === undefined) return [0, 0, 0, 0]
	if (typeof margin === 'number') return [margin, margin, margin, margin]
	return margin
}

/**
 * Largest font size at which text fits a box.
 *
 * `fit: 'shrink'` makes PowerPoint do this at render time, but the result is not readable back -
 * so it cannot inform anything else on the slide. This computes it up front.
 *
 * @example
 * const { fontSize, h } = fitText({ w: 4, h: 2 }, headline)
 * slide.addText(headline, { x: 1, y: 1, w: 4, h: 2, fontSize })
 */
export function fitText (area: FitTextArea, text: string, props: FitTextProps = {}): FitTextResult {
	const { min = 8, max = 40, margin, ...measureProps } = props

	if (!(area.w > 0) || !(area.h > 0)) throw new Error(`fitText: area must have positive w and h (got ${area.w}x${area.h})`)
	if (!(min > 0)) throw new Error(`fitText: min must be > 0 (got ${min})`)
	if (max < min) throw new Error(`fitText: max ${max} is below min ${min}`)

	const [top, right, bottom, left] = resolveMargin(margin)
	const w = area.w - left - right
	const h = area.h - top - bottom
	if (!(w > 0) || !(h > 0)) throw new Error(`fitText: margin leaves no room in ${area.w}x${area.h}`)

	const measure = (fontSize: number): Measurement => measureText(text, { ...measureProps, fontSize, w })
	const fits = (fontSize: number): boolean => {
		const result = measure(fontSize)
		return result.h <= h && result.w <= w
	}

	// Whole points only, which is what the PowerPoint UI offers. Binary search over the range: the
	// fits/does-not-fit boundary is monotonic in font size.
	let low = Math.ceil(min)
	let high = Math.floor(max)
	if (!fits(low)) {
		const result = measure(low)
		return { ...result, fontSize: low, overflows: true }
	}

	while (low < high) {
		const mid = Math.ceil((low + high) / 2)
		if (fits(mid)) low = mid
		else high = mid - 1
	}

	return { ...measure(low), fontSize: low, overflows: false }
}
