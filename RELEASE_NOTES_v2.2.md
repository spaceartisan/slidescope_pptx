# SlideScope v2.2 — Table styles, bullets, and color fidelity

## Changes

- Read table styles embedded in the PPTX: whole-table defaults, row/column bands, headers, edge rows/columns, and corner overrides. Explicit cell fills, margins, borders, and run formatting take precedence.
- Preserve styled table text colors and font emphasis. Render borders as centered SVG rules, so they do not consume cell text space. Preserve merged cell spans.
- Inherit bullet fonts and follow-text color settings. Common Wingdings square bullets render as geometry without requiring Wingdings to be installed; raw paragraph character codes remain inspectable.
- Correct DrawingML tint direction and blend with white in linear RGB. This restores pale band shading instead of dark, saturated table colors.
- Visual grading now fails when a reference slide lacks a before or after screenshot; it no longer silently grades a partial deck.

## Full-frame verification

Both versions were freshly captured with the same Chromium 153.0.8010.0 setup, device scale 1, and identical fonts. All 71 PowerPoint-exported PNGs were graded at their native dimensions. Nothing was cropped, including footers and logos.

| Test set | v2.1 RGB MAE | v2.2 RGB MAE |
| --- | ---: | ---: |
| Original 12 slides, 1536 × 864 | 10.059 | 10.059 |
| Additional 59 slides, 960 × 720 | 13.018 | 12.494 |
| All 71 slides, equal weight per slide | 12.518 | 12.083 |

The second deck improves by 4.02%. Across all 71 slides, 33 improve, 34 are unchanged within 0.001 MAE, and four regress slightly.

| Second-deck slide | Before | After | Difference (after − before) |
| --- | ---: | ---: | ---: |
| 12 | 8.684 | 7.503 | -1.180 |
| 21 | 25.016 | 17.280 | -7.737 |
| 28 | 27.544 | 14.815 | -12.729 |
| 32 | 20.042 | 14.404 | -5.638 |
| 51 | 13.899 | 13.242 | -0.656 |
| 6 | 12.941 | 12.945 | +0.004 |
| 7 | 13.349 | 13.352 | +0.003 |
| 29 | 18.900 | 18.925 | +0.025 |
| 31 | 13.914 | 13.941 | +0.027 |

Four tiny regressions remain on slides 6, 7, 29, and 31. The bullet corrections are retained because the source specifies square Wingdings bullets; these figures are not presented as a no-regression result. Detailed regional scores are in each deck's regions.csv.

73 Python tests passed on Linux/Python 3.12.14, including the original real-world presentation and six new independent tests. All 71 slides rendered without uncaught JavaScript errors. All thumbnails loaded; clicking native objects and navigating with the keyboard passed on both decks. JavaScript syntax validation passed.

## Font and platform limits

The QA host does not contain the Microsoft fonts. For both baseline and candidate captures, Carlito substitutes for Calibri and Arimo substitutes for Arial. These QA-only substitutions are identical across versions. They are not included in the application and are not pixel-identical to the Microsoft fonts. Compare this release only against its freshly captured baseline, not absolute scores from earlier environments. Windows rendering, including native GDI, still requires a local run.

The existing public-PyPI runtime requirements are unchanged. There are no new runtime packages, Node.js requirements, browser downloads, external applications, or build steps for the user. The separate development screenshot tooling is not part of the runnable release.

## Remaining rendering gaps

This release does not provide PowerPoint parity. Built-in table styles omitted from the PPTX's tableStyles.xml still need fallback definitions (visible in slide 51). Some theme effects, custom table borders, text layout, symbol glyphs beyond common square bullets, chart styling, and EMF text/diagram details remain approximate. The original 12-slide deck is preserved, not newly improved by this release.

## Run

Extract the source ZIP, open the SlideScope folder, then run `python server.py`. Open http://127.0.0.1:8000 in the VS Code integrated browser and load a presentation as before.

## Implementation references

- DrawingML table styles: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tablestyle
- DrawingML tint semantics: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tint
- Wingdings symbol encoding: https://www.unicode.org/L2/L2011/11196-n4022-wingdings.pdf
- Carlito: https://github.com/googlefonts/carlito (documentation only; no GitHub dependency installation)
