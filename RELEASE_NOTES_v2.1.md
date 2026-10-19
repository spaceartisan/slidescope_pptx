# SlideScope v2.1 — Portable EMF compatibility pass

This release adds a **best-effort, pure-Python GDI EMF-to-SVG compatibility fallback** for older PowerPoint diagrams and spreadsheet/table graphics. It uses the existing Pillow dependency and Python standard library. There is no server framework, native binary, external converter, Node.js, or new pip dependency.

## Fixed

- EMF drawing commands that formerly caused an all-or-nothing preview failure now render common filled rectangles, polygons, lines, Bézier paths, text, fonts, palette brushes, strokes, bitmap DIB records, basic clipping, and GDI coordinate transformations.
- The `DSTCOPY`/`0x00AA0029` no-op raster operation is correctly ignored. Treating it as a filled rectangle previously painted over almost the entire vehicle-sizing table.
- Embedded Excel-derived graphics on slides 14 and 15 of the 59-slide SLS deck now show their original table cells, numbers, colors, and chart bars instead of missing graphics.
- The two embedded simulation diagrams on slide 35 show much more of their original drawing primitives.
- Windows native GDI remains ahead of the compatibility fallback so supported EMFs retain higher-fidelity Windows rendering.

## Visual QA (all frames, unchanged capture setup)

| Test deck | v2.0 MAE | v2.1 MAE | Slides | Regressions |
|---|---:|---:|---:|---:|
| 59-slide SLS deck | 15.225 | 13.011 | 59 | 0 |
| 12-slide causation deck | 10.023 | 10.023 | 12 | 0 |

The three changed slide errors were: slide 14 **78.878 → 17.346**, slide 15 **79.186 → 17.495**, slide 35 **34.238 → 26.833**. Scores are mean absolute RGB pixel differences versus the user's PNG exports: lower is better. These QA results come from Linux Chromium at the reference image dimensions. Cross-platform typography can produce somewhat different scores.

## Verification and limitations

- **50 Python tests** executed; **44 passed, 6 conditionally skipped**.
- All 71 slides rendered in Chromium without JavaScript errors.
- Python 3.12 on Windows is the target; testing here ran under Linux/Python 3.13.
- Some EMF+ effects, advanced GDI clipping/rotation, bitmap transparency, and text metrics remain approximate. This is not a full EMF standard implementation. The quality improvement is verified against the supplied decks, not promised for arbitrary EMF documents.
- No extra packages or source-code downloads are required. Only `python-pptx`, `lxml`, and their public-PyPI wheels as already specified.

## Run

From the extracted `SlideScope` directory:

```powershell
python server.py
```

Then open `http://127.0.0.1:8000` in VS Code's integrated browser.
