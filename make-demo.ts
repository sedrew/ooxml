// Generates chart-layout-demo.pptx with this branch's ooxml-core. Usage: bun make-demo.ts out.pptx
import { writeFileSync } from 'node:fs';
import {
	PptxHandler, PresentationBuilder, setChartAreaFormat, setChartAxis, setChartGroupOptions, setChartDataPointGradient,
} from 'ooxml-core/pptx';
import type { ChartPptxElement } from 'ooxml-core/pptx';

const out = process.argv[2];
const BG = '#DCE6F2';
const cats = ['Unit A', 'Unit B', 'Unit C', 'Unit D', 'Benchmark'];
const series = [
	{ name: 'Series 1', values: [0.88, 0.46, 0.55, 0.73, 0.83], color: '#B4C7E7' },
	{ name: 'Series 2', values: [0.76, 0.49, 0.64, 0.89, 0.75], gradientFill: { angle: 0, stops: [{ color: '#003173', position: 0 }, { color: '#8FB4D9', position: 100 }] } },
];
const reportLook = {
	barDirection: 'bar' as const, gapWidth: 35, overlap: 10,
	chartArea: { fill: 'none', border: 'none', roundedCorners: false },
	plotArea: { fill: 'none', border: 'none' },
	axes: { catAx: { visible: false, orientation: 'maxMin' as const }, valAx: { visible: false, min: 0, max: 1 } },
};
const { handler, data, createSlide } = await PresentationBuilder.create();
function slide(title: string, left: string, right: string) {
	return createSlide('Blank')
		.addShape('rect', { x: 0, y: 0, width: 1280, height: 720, fill: { type: 'solid', color: BG } })
		.addText(title, { x: 40, y: 16, width: 1200, height: 50, fontSize: 26, bold: true, color: '#1F2937' })
		.addText(left, { x: 40, y: 70, width: 580, height: 60, fontSize: 14, color: '#374151' })
		.addText(right, { x: 660, y: 70, width: 580, height: 60, fontSize: 14, color: '#374151' });
}
const L = { x: 40, y: 140, width: 580, height: 540 }, Rb = { x: 660, y: 140, width: 580, height: 540 };
data.slides.push(
	slide('Bar: gapWidth, overlap, hidden axes, transparent chart area',
		'Before: fields ignored (gapWidth 150, no overlap, axes shown, no area spPr, no roundedCorners)',
		'After: gapWidth 35, overlap 10, axes hidden (catAx maxMin, valAx 0..1), chart/plot area noFill + ln noFill, roundedCorners 0')
		.addChart('bar', { categories: cats, series, barDirection: 'bar' }, L)
		.addChart('bar', { categories: cats, series, ...reportLook }, Rb).build(),
	slide('Doughnut: firstSliceAng and holeSize',
		'Before: holeSize always 50, firstSliceAng never written (0)',
		'After: firstSliceAngle 90, holeSize 20, no border, square corners')
		.addChart('doughnut', { categories: ['A', 'B', 'C', 'D'], series: [{ name: 'N', values: [45, 25, 20, 10] }] }, L)
		.addChart('doughnut', { categories: ['A', 'B', 'C', 'D'], series: [{ name: 'N', values: [45, 25, 20, 10] }], firstSliceAngle: 90, holeSize: 20, chartArea: { fill: 'none', border: 'none', roundedCorners: false } }, Rb).build(),
	slide('Pie: chart-area gradient fill and border',
		'Before: chart-area spPr never written',
		'After: chart-area gradient + #1F4E79 border, firstSliceAngle 270, plot area noFill')
		.addChart('pie', { categories: ['X', 'Y', 'Z'], series: [{ name: 'S', values: [5, 3, 2] }] }, L)
		.addChart('pie', { categories: ['X', 'Y', 'Z'], series: [{ name: 'S', values: [5, 3, 2] }], firstSliceAngle: 270,
			chartArea: { fill: { angle: 90, stops: [{ color: '#FFFFFF', position: 0 }, { color: '#BDD7EE', position: 100 }] }, border: '#1F4E79', roundedCorners: false }, plotArea: { fill: 'none' } }, Rb).build(),
	slide('Loaded chart: edited with setChart* after reload',
		'Left: saved and reloaded without edits (chart part unchanged)',
		'Right: same chart after setChartGroupOptions, setChartAreaFormat, setChartAxis({ visible: false })')
		.addChart('bar', { categories: cats, series }, L)
		.addChart('bar', { categories: cats, series }, Rb).build(),
);
const first = await handler.save(data.slides);
const h2 = new PptxHandler();
const loaded = await h2.load(first.buffer as ArrayBuffer);
const charts = loaded.slides[3].elements.filter((e) => e.type === 'chart') as ChartPptxElement[];
const right = charts[1]!;
setChartGroupOptions(right, { gapWidth: 60, overlap: -20 });
setChartAreaFormat(right, 'chart', { fill: 'none', border: 'none', roundedCorners: false });
setChartAreaFormat(right, 'plot', { fill: '#FFFFFF', border: '#9CA3AF' });
setChartAxis(right, 'valAx', { visible: false });
setChartDataPointGradient(right, 0, 4, { type: 'radial', stops: [{ color: '#FFFFFF', position: 0 }, { color: '#C00000', position: 100 }] });
loaded.slides[3].isDirty = true;
writeFileSync(out, await h2.save(loaded.slides));
console.log('written', out);
