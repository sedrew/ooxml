# PR assets: chart spacing, axis visibility and area format

Supporting files for the pull request "feat(pptx): honour chart spacing, axis visibility and area format in new charts"
(branch `feat/pptx-chart-generator-model-fields`). Not part of the change itself.

- `chart-layout-demo.pptx`: deck generated with the branch's `ooxml-core` by `make-demo.ts`.
- `powerpoint-slide-1.png` .. `powerpoint-slide-4.png`, `powerpoint-overview.png`: that deck opened in Microsoft
  PowerPoint 16.93 for Mac (AppleScript), exported to PDF by PowerPoint and rasterised with `pdftoppm`. PowerPoint
  opened the file without a repair prompt. The decimal commas on the axes come from the rendering machine's locale.

Each slide shows the same chart twice: left without the new fields (what the generator writes on `main`), right with
them. Slide 4 edits a reloaded chart with `setChartGroupOptions`, `setChartAreaFormat`, `setChartAxis` and
`setChartDataPointGradient`; the left chart on that slide was saved and reloaded without edits.
