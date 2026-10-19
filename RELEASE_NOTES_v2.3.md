# SlideScope v2.3 — Theme styles and connector arrowheads

## Changes

- Resolve shape fillRef and lnRef through the presentation theme's style matrix, including each reference's placeholder color. Direct fills, no-fill settings, line widths, colors, dashes, and end decorations override inherited style values independently.
- Render straight connectors as inspectable SVG using slide coordinates. Horizontal and vertical lines retain a nonzero SVG viewport while preserving their actual zero-height or zero-width endpoints.
- Render open arrow, triangle, stealth, diamond, and oval end decorations with width/length settings; preserve separate start and end decorations. The supplied decks visually exercise open arrows, stealth and triangles. Synthetic parser tests cover inheritance and overrides; all end geometries are not independently compared against Office exports.
- Preserve preset dash patterns on straight connectors. Center their strokes on the source geometry instead of placing them below/right of the source position as CSS borders did.
- No new runtime dependencies, installations, build steps, external services, or slide-specific rules.

## Full-frame comparison against v2.2

All 71 PowerPoint-exported slides were graded at native resolution with no cropping, including logos and footers. Baseline screenshots are the retained, validated v2.2 final captures. Candidate screenshots use the same Chromium 153.0.8010.0 binary, font substitutions, device scale 1, and capture script. The script checks every slide's viewport dimensions and origin before saving.

Average RGB absolute error on a 0–255 scale (lower is better):

| Test set | v2.2 | v2.3 |
| --- | ---: | ---: |
| Original 12 slides, 1536 × 864 | 10.059 | 9.918 |
| Additional 59 slides, 960 × 720 | 12.494 | 12.039 |
| All 71, equal weight per slide | 12.083 | 11.681 |

Relative error reduction: 1.40% on the first deck, 3.64% on the second, and 3.33% across all 71 slides. These are image-error changes, not percentages of PowerPoint compatibility.

36 slides improve, 33 are unchanged within 0.001 MAE, and two regress slightly. Full precision results are in results.json; per-slide scores and 3×3 regional scores are in each deck's CSV files.

| Additional-deck slide | v2.2 | v2.3 | Change, after − before |
| --- | ---: | ---: | ---: |
| 6 | 12.945 | 11.000 | -1.945 |
| 7 | 13.352 | 11.469 | -1.883 |
| 11 | 8.903 | 8.658 | -0.245 |
| 33 | 30.996 | 24.449 | -6.548 |
| 35 | 27.003 | 24.296 | -2.707 |
| 2 | 4.952 | 4.954 | +0.002 |
| 48 | 17.767 | 17.768 | +0.001 |

Slide 11's yellow double arrow now appears. Slide 33 regains its pink and blue fills and line decorations. Slide 35 regains theme-styled labels, arrows, and frame colors. The original deck's master rule now renders centered on its source coordinates. The report includes all full-slide comparisons, not just highlights.

## Verification

- 79 Python tests pass on Linux/Python 3.12.14, including the real-world presentation and six new independent style-cascade tests.
- All 71 slides render without uncaught JavaScript errors; all 71 thumbnails load.
- Native-object selection and ArrowRight navigation pass on both decks.
- Slide 11's zero-height double arrow has a visible SVG viewport, two end decorations, the resolved theme color, and working inspector selection.
- JavaScript syntax validation passes.

## Limitations

This remains an approximate PowerPoint renderer. Slide 33 still exposes grouped text sizing/rotation issues, a cube approximated as a rectangle, and image-effect differences. Some diagrams, charts, theme effects, and EMF details remain approximate. Built-in table styles absent from the package still need fallback definitions. Curved/bent connector routing and custom dash arrays are not addressed here.

The QA host lacks Microsoft fonts. Carlito substitutes for Calibri and Arimo for Arial identically in both versions; other fonts use system fallbacks. These are QA-only substitutions, not packaged dependencies, and are not pixel-identical to the Microsoft fonts. Windows, native GDI, and the VS Code integrated browser still require workstation validation. This run used a separate headless Chromium for development.

## Run

Extract the source ZIP, open the SlideScope folder, run `python server.py`, then open http://127.0.0.1:8000 in the VS Code integrated browser. The existing requirements and installation rules remain unchanged.

## Implementation references

- https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fillreference
- https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linereference
- https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tailend
