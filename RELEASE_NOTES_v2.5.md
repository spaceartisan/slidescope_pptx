# SlideScope v2.5 — Embedded previews and image transparency

## Changes

- Resolve stored embedded-object previews through direct picture relationships or matching VML shape IDs. Supported legacy equation WMFs become SVG text, with positioned glyphs, fonts, baseline alignment, character advances, and common Symbol-to-Unicode mappings. The slide model exposes the object identity and its preview asset.
- Apply VML gain and black-level settings with browser SVG filters, preserving transparent pixels. This restores the two white equation previews on additional-deck slide 31.
- Apply DrawingML exact-color replacement to picture pixels and shape image fills, including alpha matching and useA behavior. This removes opaque backgrounds from the equations on slide 30, the scatter plot on slide 33, and affected logos. Near-matching colors are deliberately unchanged.
- Use only preview media; never execute embedded programs or read embedded OLE binaries. Unsupported WMF drawing/state records produce a warning rather than an incomplete preview.

No new runtime packages, build steps, or services. Pillow is already a python-pptx dependency. These are format-based fixes, with no slide-specific rules. Slides remain browser objects; equation previews are SVG assets containing text, rather than complete-slide screenshots. The previews are not editable mathematical formulas.

## Full-frame comparison against v2.4

All 71 PowerPoint-exported references were graded at their native resolution, including footers and logos. The baseline is the retained v2.4 final capture set. Both versions use Chromium 153.0.8010.0, device scale 1, the same font substitutions, and the same capture script. Canvas position and dimensions are checked before each capture.

Mean RGB absolute error, 0–255; lower is better:

| Test set | v2.4 | v2.5 |
| --- | ---: | ---: |
| Original 12 slides, 1536 × 864 | 9.916 | 9.918 |
| Additional 59 slides, 960 × 720 | 11.964 | 11.536 |
| All 71, equal weight per slide | 11.618 | 11.263 |

42 slides improve, 27 are unchanged within 0.001 MAE, and 2 regress beyond that tolerance. Overall mean error decreases 3.06%; the additional-deck mean decreases 3.57%. These are image-error changes, not PowerPoint compatibility percentages.

| Additional-deck slide | v2.4 | v2.5 | Change, after − before |
| --- | ---: | ---: | ---: |
| 30 | 12.899 | 4.697 | -8.202 |
| 31 | 13.796 | 13.830 | +0.034 |
| 33 | 24.159 | 14.135 | -10.023 |
| 40 | 14.703 | 14.527 | -0.176 |

The two measured regressions are retained, not removed from the aggregate:

- Original-deck slide 2: 7.882946 → 7.909497 (+0.026551). A targeted recapture reverses the result: baseline 7.909497, candidate 7.882946. This indicates capture variability rather than a stable v2.5 regression. No cause is claimed. The original full-run captures remain the reported comparison; recheck data and PNGs are included separately.
- Additional-deck slide 31: 13.795924 → 13.830101 (+0.034177). The previously missing equations are visible, but their glyph size, style, and placement still differ from the reference. Restoring content does not guarantee a lower pixel error when fonts and metrics differ. Both the visual and numerical results are shown.

Every slide has a full-slide before/current/PowerPoint comparison, scores.csv, and 3×3 regions.csv. Full-precision scores are in results.json. Selected native-resolution PNG triples are in details/; the report does not crop away footers or logos.

## Verification

91 Python tests passed on Linux/Python 3.12.14, including eight new preview/color-replacement cases. JavaScript syntax validation passes. All 71 slides and thumbnails render without uncaught browser errors. Native-object selection and keyboard navigation pass for both decks; the previous two-headed connector check still passes.

Eight browser text-layout checks pass in each deck session, covering wrapping, script offsets, script glyph size, leading, and vertical text. Additional browser checks confirm both recovered equation previews load, retain SVG text (including Δ), receive their color filters, and can be selected as slide objects. The scatter-image background pixel has alpha zero. Raw capture metadata, UI evidence, and the Python test log are included.

## Remaining limitations

This is still an approximate PowerPoint renderer. The portable WMF reader is intentionally limited to supported text records; it is not a complete WMF implementation. Other previews may remain unsupported. The trajectory diagram on slide 40 still has missing mathematical notation and geometry differences. Slide 33 still has an approximate cube, incorrect rounded-corner geometry, and text/spacing differences. Charts, complex image effects, EMF details, group transforms, and earlier text-layout limitations remain.

The QA host lacks Microsoft fonts. Carlito substitutes for Calibri and Arimo for Arial identically in both versions, while other fonts use system fallbacks. These QA-only fonts and browser tools are not packaged. Equation font matching remains approximate. Windows, native GDI, and the VS Code integrated browser still need workstation validation; the reported browser tests ran in headless Chromium on Linux.

## Run

Extract the source ZIP, open the SlideScope folder, run `python server.py`, and open http://127.0.0.1:8000 in the VS Code integrated browser. Upload a presentation as before. Existing Python 3.12 / official public-PyPI wheel requirements are unchanged. See README.md for the installation command.

## Implementation references

- OLE shape ID and stored VML previews: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/4a70eb6c-75a7-4927-b786-f1dd7b84c3db
- Color replacement and useA: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/bda9479c-a102-4807-9008-14549fb40e2d
- WMF positioned text: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wmf/7d07c44a-a828-4b82-9af0-e0a81cced5a8
- VML image adjustments: https://learn.microsoft.com/en-us/windows/win32/vml/web-workshop---how-to-use-vml-on-web-pages-----image--element
