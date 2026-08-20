import type PptxGenJS from '@neo-ma/pptxgenjs'

/**
 * The slice of `Slide` this helper needs.
 *
 * Structural rather than `PptxGenJS.Slide` so the helper composes the public API without
 * binding to a class - any object with a conforming `addChart` works, including the class
 * the core exports and any test double.
 */
export interface ChartSlide {
	addChart: (type: 'bar', data: object[], options?: PptxGenJS.IChartOpts) => unknown
}

export interface WaterfallProps {
	/** Category label per delta */
	labels: string[]
	/** Signed change per category - the bar spans from the running total to running + value */
	values: number[]
	/** Append a bar spanning 0 to the final running total, labelled with this string */
	total?: string
	/** Bar color for positive deltas @default '2E7D32' */
	increaseColor?: string
	/** Bar color for negative deltas @default 'C62828' */
	decreaseColor?: string
	/** Series name shown in the legend for positive deltas @default 'Increase' */
	increaseName?: string
	/** Series name shown in the legend for negative deltas @default 'Decrease' */
	decreaseName?: string
}

/**
 * Add a waterfall (bridge) chart.
 *
 * PowerPoint has no waterfall chart type reachable through ECMA-376 `c:barChart`,
 * so this is the standard construction: a stacked bar chart whose first series is
 * transparent and carries each bar up to its starting value.
 *
 * `options` is passed straight through to `addChart`, so position, axes, titles
 * and data labels all work as usual - only `barDir`, `barGrouping` and
 * `chartColors` are fixed by the construction.
 *
 * @example
 * waterfall(slide, {
 *   labels: ['Start', 'Sales', 'Costs', 'Tax'],
 *   values: [100, 45, -30, -12],
 *   total: 'Net',
 * }, { x: 0.5, y: 0.5, w: 9, h: 4.5, showValue: true })
 */
export function waterfall<T extends ChartSlide> (slide: T, props: WaterfallProps, options: Partial<PptxGenJS.IChartOpts> = {}): T {
	const { labels, values, total, increaseColor = '2E7D32', decreaseColor = 'C62828', increaseName = 'Increase', decreaseName = 'Decrease' } = props

	if (labels.length !== values.length) throw new Error(`waterfall: labels (${labels.length}) and values (${values.length}) must be the same length`)
	if (values.length === 0) throw new Error('waterfall: at least one value is required')
	const bad = values.findIndex(val => !Number.isFinite(val))
	if (bad !== -1) throw new Error(`waterfall: values[${bad}] is not a finite number`)

	const base: number[] = []
	const up: number[] = []
	const down: number[] = []
	const upLabels: string[] = []
	const downLabels: string[] = []

	let running = 0
	values.forEach(delta => {
		const start = running
		running += delta
		// A bar hangs below its start when the delta is negative, so the transparent
		// riser stops at the lower of the two ends and the visible bar covers |delta|.
		base.push(Math.min(start, running))
		up.push(delta >= 0 ? delta : 0)
		down.push(delta < 0 ? -delta : 0)
		upLabels.push(delta >= 0 ? String(delta) : '')
		downLabels.push(delta < 0 ? String(delta) : '')
	})

	const chartLabels = [...labels]
	if (total !== undefined) {
		// The total bar is absolute: it rises from zero, coloured by the sign of the result.
		chartLabels.push(total)
		base.push(Math.min(0, running))
		up.push(running >= 0 ? running : 0)
		down.push(running < 0 ? -running : 0)
		upLabels.push(running >= 0 ? String(running) : '')
		downLabels.push(running < 0 ? String(running) : '')
	}

	slide.addChart('bar', [
		{ name: '', labels: chartLabels, values: base, color: 'transparent' },
		{ name: increaseName, labels: chartLabels, values: up, color: increaseColor, dataLabels: upLabels },
		{ name: decreaseName, labels: chartLabels, values: down, color: decreaseColor, dataLabels: downLabels },
	], { ...options, barDir: 'col', barGrouping: 'stacked' })

	return slide
}
