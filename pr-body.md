## What does this change?

New charts (`addChart` / `createChartElement` -> `buildChartSpaceXml`) now write `barGapWidth`, `barOverlap`, `firstSliceAngle`, `doughnutHoleSize`, `axes[].deleted`, the chart-area and plot-area fill / gradient / border and `roundedCorners` from the model instead of constants, keeping today's output for unset fields. The same fields get an SDK surface and are reconciled on save for loaded charts without touching an unedited chart part.

Commits:

- `fix(pptx): write c:scaling children in schema order`: an axis with both `min` and `max` was written `orientation, min, max`, which fails `dml-chart.xsd` (`CT_Scaling` is `logBase, orientation, max, min`). Affects generated and loaded charts; a node already in order is left as it is.
- `feat(pptx): write chart spacing, axis visibility and area format`: generator and loaded-chart save path (`chart-group-options.ts`, `chart-area-format.ts`, `chart-space-format.ts`, `chart-axis-deleted.ts`). Adds `PptxChartStyle.chartAreaBorder` / `plotAreaBorder` (`a:ln` colour or `'none'`), parsed by the same helper the save path uses; without it a transparent, borderless chart area (what PowerPoint templates commonly author) cannot be expressed.
- `feat(pptx): add chart spacing and area options to the chart SDK`: `ChartInput` fields (`gapWidth`, `overlap`, `firstSliceAngle`, `holeSize`, `chartArea`, `plotArea`, `axes`), `ChartBuilder` methods (`gapWidth()`, `overlap()`, `firstSliceAngle()`, `holeSize()`, `chartArea()`, `plotArea()`, `axis()`), `setChartGroupOptions()`, `setChartAreaFormat()`, and `setChartAxis(..., { visible })`.
- `test(pptx): ...`: unit and round-trip coverage.

Behaviour:

- Unset fields produce exactly the previous XML (`gapWidth` 150, `holeSize` 50, `delete` 0, no `overlap`, `firstSliceAng`, area `spPr` or `roundedCorners`).
- Elements follow the schema sequences (`CHART_CONTAINER_CHILD_ORDER`, `CT_ChartSpace`, `CT_PlotArea`, `CT_*Ax`). Only containers whose schema has the element get it: no `overlap` on `bar3D`, no `firstSliceAng` on `pie3D` / `ofPie`, `holeSize` on doughnut only.
- SDK input is validated against `ST_GapAmount` (0..500), `ST_Overlap` (-100..100), `ST_FirstSliceAng` (0..360) and `ST_HoleSize` (1..90) with `RangeError`, and an option the chart type has no element for throws. The generator itself clamps out-of-range model values (for example a parsed `holeSize="0"`) rather than emit invalid XML.
- Loaded charts: an element is rewritten only when the model holds a different value, removed only when the model cleared a value the parser had read, and an unparseable authored value is left alone. Area fills keep an unchanged themed colour (`a:schemeClr` + transforms) and the rest of `c:spPr` (`a:effectLst`, `a:ln` attributes). Combo charts keep their per-container spacing as authored (they are consolidated and re-split on save).
- `roundedCorners`: PowerPoint draws rounded corners when the element is absent, so square corners need an explicit `false`; JSDoc updated accordingly.

## Type of change

- [ ] Bug fix (non-breaking)
- [x] New feature (non-breaking)
- [ ] Breaking change
- [ ] Docs / tooling / CI only

## Area

pptx

- [x] The change is made once in the area that owns it, not copied per format
- [x] DOM-free logic stays in `src/<area>/`; custom elements stay in `packages/ui`
- [x] Unsupported behaviour is reported honestly (no claim of Office parity without evidence)
- [ ] Extraction provenance is recorded in `PROVENANCE.md` for any module moved here (nothing moved)

## Testing

- [x] Generator: each field produces its element in schema order, unset fields keep the previous values (`chart-xml-generator-layout.test.ts`)
- [x] Loaded-chart reconcilers: untouched nodes kept by identity, edits in place, themed fill kept when only the border changes, insert positions, cleared values removed (`chart-group-options.test.ts`, `chart-space-format.test.ts`)
- [x] SDK validation and builders (`chart-layout-operations.test.ts`)
- [x] Save/reload round trips through `PptxHandler`, including a byte-for-byte check of an untouched authored layout after a values-only edit (`__tests__/integration/chart-layout-roundtrip.test.ts`)
- [x] Regression test for the scaling order that fails without the fix (`chart-axis-scaling-order.test.ts`)
- [x] `bun run test:package` passes

Schema check: generated bar (all fields), doughnut, pie (gradient chart area), bar3D, ofPie and default charts validate with `xmllint --schema schemas/ecma-376-transitional/dml-chart.xsd`. Before the scaling fix the bar chart with `min`/`max` did not.

Checked in Microsoft PowerPoint 16.93 for Mac: the demo deck below opens without a repair prompt and renders as intended. Note that on this build a generated chart without `c:spPr` already shows no fill and no border, so the visible difference there is the spacing, axes and the explicit area format; LibreOffice and PowerPoint for Windows were not available to compare.

## Checks run locally

- [x] `bun run lint` (oxlint on the changed files)
- [x] `bun run fmt:check` (on the changed files; `main` itself is not clean under the installed oxfmt, so I did not reformat other files)
- [x] `bun run typecheck`
- [x] `bun run test` (1297 files, 21731 tests)

## Conventional Commits

- [x] My commit messages follow [Conventional Commits](https://www.conventionalcommits.org) (`scripts/check-conventional-commits.mjs`: OK)

## Screenshots / recordings

![Before / after, rendered by PowerPoint](https://raw.githubusercontent.com/sedrew/ooxml/pr-assets/chart-generator-model-fields/powerpoint-overview.png)

Full-size slides: [1](https://raw.githubusercontent.com/sedrew/ooxml/pr-assets/chart-generator-model-fields/powerpoint-slide-1.png), [2](https://raw.githubusercontent.com/sedrew/ooxml/pr-assets/chart-generator-model-fields/powerpoint-slide-2.png), [3](https://raw.githubusercontent.com/sedrew/ooxml/pr-assets/chart-generator-model-fields/powerpoint-slide-3.png), [4](https://raw.githubusercontent.com/sedrew/ooxml/pr-assets/chart-generator-model-fields/powerpoint-slide-4.png).

Left of each slide: the chart without the new fields (what the generator writes today); right: with them. Slide 4 edits a reloaded chart through `setChartGroupOptions` / `setChartAreaFormat` / `setChartAxis` / `setChartDataPointGradient`. Demo deck: [chart-layout-demo.pptx](https://github.com/sedrew/ooxml/raw/pr-assets/chart-generator-model-fields/chart-layout-demo.pptx), generated by [make-demo.ts](https://github.com/sedrew/ooxml/blob/pr-assets/chart-generator-model-fields/make-demo.ts) (assets live on a separate branch of my fork, not in this PR).

## Out of scope / follow-ups

- Embedded workbook for new charts (`ppt/embeddings/*.xlsx`, `c:externalData`, `c:strRef` / `c:numRef` with `c:f`) so "Edit Data" works: separate PR.
- Per-container spacing edits on combo charts.
- `holeSize="0"` (written by PowerPoint for some decks) is outside `ST_HoleSize` and is not accepted by the SDK; a loaded chart that has it keeps it untouched.
