# SlideScope v2.0 — Second-deck compatibility and full-fidelity regression

## New benchmark

A second user-supplied PowerPoint was added to our **development-only** regression suite. Its PPTX contains **59 slides** in 4:3 aspect ratio and the supplied export contains 59 PNGs at **960×720**. The original **12-slide** 16:9 deck and its **1536×864** PNG references remain in the suite. Source presentations are **not included** in the runnable ZIP.

## Corrections

1. **Master-picture brightness and contrast**: Read native `a:lum` image effects. Apply browser-native filter approximations to the bitmap itself, rather than to adjacent shapes or text. This restores the dark Earth treatment in the new deck. The filter is a visual approximation, not a formal PowerPoint color-management implementation.
2. **Images embedded as ordinary shape fills**: Read `a:blipFill` from `p:spPr`, follow image relationships from the owning slide, layout or master, and respect source cropping. This restores portraits and other picture-filled rectangles. Shape pictures remain inspectable individual objects.
3. **Actual hollow frame geometry**: Interpret the `frame` preset with its adjustment value as a ring instead of filling the entire rectangle. This restores the underlying diagram screenshots and tables that yellow frame shapes previously concealed.
4. **Native chart DOM bug**: Fix an exception caused by chaining `.lastChild` onto `Element.append()`, which returns `undefined`. Chart-containing slides no longer cause the browser render to fail.
5. **Transparent native charts**: Respect chart-space transparency instead of forcing a white SVG backing. Detect dark master-image layouts for legible light axis labels. This is still an approximate browser chart renderer, not an Office chart engine.
6. **Group-fill inheritance**: Honor `a:grpFill` on custom paths under their parent group. This restores details such as the white ASDL corner logo.
7. **Slide-number fields**: Resolve `a:fld type="slidenum"` per slide and carry over basic number formatting from the master instead of showing the cached `‹#›` field string.

## Grading (same Chromium capture pipeline for before and after)

| Reference deck | Slides | v1.9 before | v2.0 after | Result |
|---|---:|---:|---:|---|
| New 4:3 reference | 59 | 55.214 | 15.225 | 57 improved, 2 regressed |
| Original 16:9 reference | 12 | 10.023 | 10.023 | 12 unchanged |

Mean absolute RGB error on the native complete slide image, 0–255 (lower is better). The 59-slide *before* screenshots apply only the native chart-label crash hotfix so a complete v1.9 baseline can be captured; the unmodified v1.9 crashes on chart slides. All changes otherwise used the unchanged v1.9 code as baseline.

### Remaining limitations

- **Two new slides (14 and 15)** regress in Linux image grading. Their central tables are stored as sophisticated EMF graphics containing records that our conservative cross-platform EMF-to-SVG interpreter cannot render. The existing Windows GDI fallback is retained, but its behavior still needs a Windows/Python 3.12 local test. The displayed dark Earth background is now correct, but without the white EMF table the numerical pixel error can rise. The renderer does **not** replace them with the exported PNG references, as that would defeat its purpose.
- Several other complex EMFs and native charts remain approximated or unrendered on Linux, particularly on slides 33–35.
- Chart plotting, font metrics, SmartArt and Office-specific effects are still not completely PowerPoint-identical.
- The new benchmark is extensive, but it does not establish compatibility with arbitrary corporate presentations.

## Verification

- All **59 new** and **12 original** slides rendered through the actual parser and JavaScript rendering functions in headless Chromium, with **no uncaught JavaScript page errors**.
- Full viewer smoke test: all 59 thumbnails created, slide navigation to 44 succeeded, object DOM present, no uncaught browser errors.
- `python -m unittest discover -s tests -q`: **47 tests successful, 6 conditional skips** (fixtures not bundled). Seven new self-contained cases cover the new features.
- No additional runtime packages, no Node/npm, no office application, no external binaries, and no network requests necessary.
- **Python 3.12 / Windows remains the target**; tests here ran on Linux/Python 3.13. Windows GDI conversion is unverified here.

## Local use

Unzip the project, open its `SlideScope` directory and run:

```powershell
python -m pip install --index-url https://pypi.org/simple --only-binary=:all: -r requirements.txt
python server.py
```

Open `http://127.0.0.1:8000` in the VS Code integrated browser, choose **Open PPTX**, and optionally load **Open reference ZIP**. The provided `more_examples.zip` contains the PNG exports inside its nested folder; the reference loader searches by `SlideN.PNG` regardless of its enclosing folder. The uploaded ZIP size is below the server's 64 MB limit, but a workplace may have its own file-loading limits.
