/**
 * ChartEx charts are a separate part type wrapped in `mc:AlternateContent`. A consumer that does
 * not understand the `cx1` namespace must take the fallback shape and still open the file cleanly -
 * this does not prove the chartex layouts themselves render.
 */
import type PptxGenJS from '../../src/pptxgen'

export function addChartexFixture (pptx: PptxGenJS): void {
	const chartExSlide = pptx.addSlide()
	chartExSlide.addChart(pptx.ChartType.waterfall, [{ name: 'Cash flow', labels: ['Start', 'Q1', 'Q2', 'End'], values: [100, 30, -20, 110] }], { x: 0.5, y: 0.5, w: 4.5, h: 3, showTitle: true, title: 'Waterfall', showValue: true, chartExSubtotals: [0, 3] })
	chartExSlide.addChart(pptx.ChartType.treemap, [{ name: 'Revenue', labels: ['A', 'B', 'C'], values: [10, 20, 30] }], { x: 5.2, y: 0.5, w: 4.5, h: 3 })
	chartExSlide.addChart(pptx.ChartType.funnel, [{ name: 'Pipeline', labels: ['Leads', 'Qualified', 'Won'], values: [500, 120, 30] }], { x: 0.5, y: 3.8, w: 4.5, h: 3, showLegend: true })
	chartExSlide.addChart(pptx.ChartType.boxWhisker, [
		{ name: 'Alpha', labels: ['x', 'y', 'z'], values: [1, 5, 9] },
		{ name: 'Beta', labels: ['x', 'y', 'z'], values: [2, 6, 4] },
	], { x: 5.2, y: 3.8, w: 4.5, h: 3, chartExMeanLine: true })
}
