# Rebuild Guidelines

## Reference Analysis

Capture these properties before drawing:

- Canvas aspect ratio and page margins.
- Major panels and their accent colors.
- Text hierarchy: title, module title, block labels, equations, captions.
- Connector semantics: solid arrows, dashed feedback, braces, ellipses, summation nodes.
- Repeated motifs: stacked frames, neural blocks, graph nodes, cubes, heatmaps, small charts.
- Typography: use Times New Roman for academic figures unless the user specifies another venue style.

## Panel Inventory Template

For a complex model figure like the AATN reference image, write a short inventory before coding:

```text
Canvas: wide landscape, two bands separated by dashed blue horizontal rule.
Top band: input stack -> repeated encoder blocks -> latent space -> reconstruction -> output stack; loss box feeds reconstruction by dashed arrow.
Bottom band: 2x2 module grid.
Panel 1 MSC-Y: purple accent, three convolution branches, GN/ReLU, TAP, Y-gate, Film, small feature stack.
Panel 2 VGD: green accent, y-slices, graph construction, edge weights, graph diffusion, TAP feedback, Msg/Output.
Panel 3 AHPool: cyan/blue accent, input cube, feature transform, attention map, pooling strip, TAP weight.
Panel 4 TAP: orange accent, vertical thickness scale, small charts, W/Yfeat tables, Y-gate, heatmap.
Captions: (a) overall framework, (b) submodule structure.
```

This inventory is the contract for the drawing script. If an item is not in the inventory, it is likely to be omitted or drawn inconsistently.

## Coordinate Strategy

Use a reference coordinate system in pixels or normalized units, then convert to Visio page inches:

```powershell
$PageW = 16.0
$PageH = 12.0
$RefW = 1448.0
$RefH = 1086.0
function VX([double]$x) { $PageW * $x / $RefW }
function VY([double]$y) { $PageH - ($PageH * $y / $RefH) }
```

Draw from top-left bounds as `RectTL(x, y, w, h)` so the script remains readable against screenshots.

## Visual Matching Heuristics

- Match structure before decoration. Panel placement and flow arrows matter more than texture detail.
- Use simplified native motifs for dense image-like content: stacked rounded rectangles, small dots, mini heatmaps, line charts, and graph nodes.
- Preserve scientific labels exactly when legible. If text is unreadable, use a close placeholder and mention it.
- Use muted panel backgrounds and colored outlines instead of heavy fills.
- Keep equations editable as text when practical. Use plain approximations if Visio equation objects are unavailable.
- Use color as semantic grouping, not decoration. Reuse one accent per module and keep internal operator boxes low-saturation.
- Prefer native approximations over screenshots for cubes, graphs, and heatmaps. The goal is editability plus visual equivalence, not pixel-perfect raster tracing.
- If the reference contains many tiny repeated elements, create 2-3 reusable helper functions rather than hand placing each occurrence.

## Scientific Figure Style Tokens

Use these defaults unless the reference clearly differs:

- Font: Times New Roman.
- Main border: 0.9-1.2 pt.
- Internal border: 0.5-0.8 pt.
- Rounded rectangle radius: small, usually 4-8 px equivalent.
- Arrowheads: consistent filled end arrows for forward flow; dashed lines for feedback/loss constraints.
- Backgrounds: white or very pale panel tints.
- Captions: black, bold, centered; keep below the figure bands.

## Reconstruction Order

1. Page size and canvas background.
2. Major bands, panel boxes, captions, and separators.
3. Main dataflow arrows and module containers.
4. Text-bearing process boxes.
5. Repeated motifs: stacks, cubes, graphs, heatmaps, mini charts.
6. Equations, legends, axis labels, small annotations.
7. Grouping, font normalization, preview export, package inspection.

Do not optimize small graphics before the panel grid and arrows are correct.

## Avoid These Failure Modes

- Inserting the whole reference image as the final page.
- Using a single ungrouped mass of hundreds of shapes with no structure.
- Recoloring by global color replacement when modules share colors with different meanings.
- Leaving Visio locked in the background after a timeout.
- Claiming success after a script times out without checking file timestamp or package contents.
- Letting text overflow boxes after changing fonts.
- Losing semantic editability by converting equations, graph nodes, or tables into a pasted crop.
- Making all modules the same palette when the reference uses color to distinguish submodules.

## Verification Rubric

Score the result before delivery:

- Structure: panel locations and flow match the reference.
- Semantics: every named module and caption exists as editable text.
- Style: colors, line weights, typography, and spacing are consistent.
- Editability: no full-page image; major objects are native shapes.
- Robustness: target file has a backup; preview/package checks are recorded.

If any category is weak, either fix it or state the limitation explicitly.

## Delivery Expectations

Final response should include:

- Target `.vsdx` path.
- Backup path.
- Preview path if exported.
- Whether the file is native editable Visio shapes.
- Any caveats: unreadable labels, skipped verification, or Visio automation issues.
