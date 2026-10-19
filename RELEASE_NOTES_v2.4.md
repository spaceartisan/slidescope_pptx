# SlideScope v2.4 — Text direction, scripts, and wrapping

## Changes

- Fix a CSS conflict that made run spans wrap even when their PowerPoint text frame explicitly set wrap="none". Runs now inherit paragraph whitespace behavior; ordinary wrapping remains enabled where requested.
- Preserve text-body direction and text-specific rotation separately from shape rotation. Use native CSS vertical writing for vertical labels, including the supplied decks' vert, vert270, and eaVert cases.
- Preserve run baseline offsets and direct paragraph baseline defaults, including an explicit zero reset. Display subscript/superscript glyphs at an approximate 65% size while retaining original text, declared font size, and baseline values in the inspectable model.
- Keep script glyphs from increasing paragraph leading. This fixes an intermediate regression where a subscript pushed later bullets downward.
- Add browser layout checks that measure rendered character positions, wrapping, script size and offsets, paragraph leading, and vertical text flow.

No new runtime packages, build steps, or external services. These are general text-formatting rules, not slide-specific corrections. Group scaling was investigated but is not changed in this release: the targeted examples were explained by wrapping, script, and direction properties.

## Full-frame comparison against v2.3

All 71 PowerPoint-exported slides were graded at native resolution, including footers and logos, with no cropping. The baseline is the retained, validated v2.3 final capture set. Candidate captures use the same Chromium 153.0.8010.0 binary, fonts, device scale 1, and screenshot script. Every canvas's dimensions and origin are checked before capture.

Mean RGB absolute error, 0–255; lower is better:

| Test set | v2.3 | v2.4 |
| --- | ---: | ---: |
| Original 12 slides, 1536 × 864 | 9.918 | 9.916 |
| Additional 59 slides, 960 × 720 | 12.039 | 11.964 |
| All 71, equal weight per slide | 11.681 | 11.618 |

12 slides improve, 59 are unchanged within 0.001 MAE, and 0 regress beyond that tolerance. The original 12-slide deck is essentially unchanged (9.918 → 9.916); the small change is on slide 2. Small score differences may include capture variability. Relative mean-error reduction on the additional deck: 0.62%. This is an image-error metric, not a percentage of PowerPoint compatibility.

| Additional-deck slide | v2.3 | v2.4 | Change, after − before |
| --- | ---: | ---: | ---: |
| 10 | 11.709 | 11.665 | -0.044 |
| 12 | 7.504 | 7.359 | -0.145 |
| 16 | 13.094 | 12.019 | -1.075 |
| 19 | 11.083 | 10.986 | -0.097 |
| 31 | 13.941 | 13.796 | -0.145 |
| 33 | 24.449 | 24.159 | -0.290 |
| 36 | 5.972 | 5.832 | -0.140 |
| 37 | 5.253 | 5.127 | -0.126 |
| 54 | 8.920 | 8.812 | -0.108 |
| 56 | 9.523 | 7.435 | -2.088 |

Every slide's comparison is included, alongside per-slide scores.csv and 3×3 regions.csv. Full precision scores are in results.json. Native PNG triples for slides 10, 16, 19, 31, 33, and 56 are in details/59/.

## Verification

83 Python tests passed on Linux/Python 3.12.14, including the real-world presentation and four new independent text-property tests. All 71 slides render without uncaught JavaScript errors; all thumbnails load. Native-object selection and keyboard navigation pass on both decks. The previously fixed double arrow remains inspectable and correctly styled.

Eight browser layout assertions pass on both loaded test sessions. These check actual browser character rectangles, not just source strings. The standalone development helper tools/browser_text_checks.js is included; load it after viewer.js in the running app and call `await runSlideScopeTextChecks()`. It adds no runtime dependencies. JavaScript syntax validation also passes.

## Remaining limitations

This is still an approximate PowerPoint renderer. The 65% script glyph scale is a rendering approximation; font-specific Office metrics may differ. Body direction/rotation is read from the local text body; comprehensive layout/master direction inheritance, table-cell direction overrides, and all complex-script/WordArt modes are not fully validated. Group transforms and some text spacing still need work.

Slide 33 still has an approximate cube, a white-backed scatter image that differs from the PowerPoint image effect, and font/spacing differences. Slide 31's embedded equations remain missing. Charts, theme effects, EMF details, and built-in table styles absent from the package retain earlier limitations. These unchanged gaps remain visible in the full-slide evidence.

The QA host lacks Microsoft fonts. Carlito substitutes for Calibri and Arimo for Arial identically in both versions, with system fallbacks for other fonts. These QA-only fonts are not packaged and are not pixel-identical to Microsoft's fonts. Windows, native GDI, and the VS Code integrated browser still require workstation validation. This run used separate headless Chromium development tooling.

## Run

Extract the source ZIP, open SlideScope, run `python server.py`, and open http://127.0.0.1:8000 in the VS Code integrated browser. Upload the presentation as before. Existing Python 3.12 / official public-PyPI wheel requirements are unchanged.

## Implementation references

- https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.vertical
- https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.rotation
- https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textcharacterpropertiestype.baseline
