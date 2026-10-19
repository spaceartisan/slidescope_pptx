# SlideScope — offline PowerPoint browser harness, prototype v2.5

## Latest: v2.5 embedded equation previews and image transparency

Stored legacy equation previews now render as SVG text where their WMF records are supported. Embedded programs are never executed. DrawingML image color replacement now preserves transparent backgrounds in pictures and shape image fills, fixing opaque panels in the supplied diagrams and logos. See `RELEASE_NOTES_v2.5.md` for all 71 visual comparisons, two small measured regressions, and remaining limits. **No runtime dependency changes.**

## Previous: v2.4 text direction, scripts, and wrapping

Text runs now honor non-wrapping frames, vertical labels retain their direction, and subscript/superscript runs keep their baseline offsets without increasing paragraph spacing. The inspectable model retains original run text and declared font sizes. See `RELEASE_NOTES_v2.4.md` for all 71 visual comparisons, test coverage, and remaining limits. **No runtime dependency changes.**

## Previous: v2.3 theme shape styles and connector arrowheads

Shape fills and line colors now resolve PowerPoint theme style references, with direct formatting taking precedence. Straight connectors preserve arrowheads, dash styles, and zero-height/zero-width geometry as inspectable SVG. All 71 slides were compared against the v2.2 captures: both decks improve on average, with two tiny score regressions. See `RELEASE_NOTES_v2.3.md` for measurements and remaining limitations. **No runtime dependency changes.**

## Previous: v2.2 table styles and bullet fidelity

Embedded table styles now retain headers, banding, text colors, margins, and centered border rules. Common Wingdings square bullets work across platforms, and DrawingML tint blending is corrected. All 71 slides were regraded; the original 12-slide deck is unchanged. See `RELEASE_NOTES_v2.2.md` for measured gains, the four small regressions, and font limitations. **No runtime dependency changes.**


## Previous: v2.1 portable EMF compatibility

The cross-platform EMF fallback has been expanded to recover legacy Office
spreadsheets/diagrams as browser-renderable SVG: common text, filled shapes,
DIB bitmaps, clip regions, transforms and brushes are supported. Windows GDI
continues to take priority for complex EMF. This adds **no new pip dependencies**
or required software installations. See `RELEASE_NOTES_v2.1.md` for exact
71-slide reference results and remaining limitations.


Browser-inspectable `.pptx` rendering for a VS Code AI agent with its existing browser/Playwright integration. This is a fidelity-development prototype, **not** an exact Office rendering engine. Python **3.12** target; **no MCP, Docker, npm, database, CDNs, FastAPI or cloud services**.

## Run locally

```powershell
python -m pip install --index-url https://pypi.org/simple --only-binary=:all: -r requirements.txt
python server.py
```

Open **http://127.0.0.1:8000** in the VS Code integrated browser, click **Open PPTX**, and select a `.pptx` file. Optional: `python server.py "C:\path\to\deck.pptx"` starts with a deck and watches that file for changes. Use `--port 8123` if port 8000 is occupied.

## NEW: v2.0 59-slide second-deck compatibility

This version expands visual grading from the original 12 slides to a **second set of 59
slides**, covering **71 original PowerPoint exports** at their native image dimensions.
The renderer now handles image-filled shapes, PowerPoint image luminance, hollow
frame geometry, group-fill inheritance, transparent charts, native chart label
creation and slide-number fields. The full details, *including limitations and
regressions*, are in `RELEASE_NOTES_v2.0.md`.

**No additional Python packages** were introduced. Workstation compatibility is
still Python 3.12, with wheels from official public PyPI and no external binaries.

## Previous: v1.9 full-deck fidelity, wrapping and numbering

v1.9 adjusts inherited leading in the first paragraph of non-autofitted content placeholders, avoiding a common vertical baseline mismatch without moving subsequent text. Text frames with PowerPoint `wrap=none` now stay unwrapped, and auto-numbered bullets render Arabic, alphabetic, or Roman numbering from OOXML. **No new dependencies**. Full details and apples-to-apples 12-slide grading are in `RELEASE_NOTES_v1.9.md`.

## Previous: v1.8 paragraph metrics and browser font-pair behavior

This release was tested against **all 12 slides** in the user-provided original
PowerPoint PNG export ZIP, at **1536 × 864**, including the footer and logo.

- **PowerPoint-style glyph spacing:** Ordinary slide text now disables the
  browser's default optional kerning. OOXML `a:rPr/@kern` remains respected:
  a positive value is interpreted as the font-size threshold (hundredths of
  a point) above which kerning is enabled. Cached SmartArt retains its
  independent rendering behavior. This is a carefully validated browser
  approximation rather than a guarantee of identical font rasterization.
- **Absolute paragraph line spacing:** When a paragraph specifies `a:spcPts`
  instead of a percentage or proportional multiplier, the parser now passes
  an explicit pixel height to the renderer. Prior builds silently ignored
  that line-height because it was a PowerPoint length, not a CSS ratio.
- **Regression coverage:** Added tests for point-valued and ratio-valued line
  spacing, explicit kerning behavior, and cached SmartArt exclusion. The
  before/after images and per-region scores are provided as a separate QA ZIP.

**Measured result (full slide):** **10.593 (v1.7) → 10.319 (v1.8)** mean RGB
absolute error over 12 PowerPoint-exported slide references. Nine slides
improved and three did not change; none regressed on this deck. Largest gains:
slides 3 (7.841 → 7.102), 8 (13.287 → 12.596), 10 (13.171 → 12.728),
12 (12.984 → 12.418). The `a:spcPts` improvement is exercised in synthetic
regression tests but does not produce a further change in this particular deck.

**Verification:** 54 Python tests passed, all 12 slides rendered in isolated
Chromium without JavaScript errors. Tested on Linux/Python 3.13; Windows and
Python 3.12 remain the production target and need local validation. The existing
read-only viewer, CLI, Playwright inspection interface, and dependency rules are
unchanged. No new pip packages, external binaries, server frameworks, build
steps, or hosted services are required.

## Previous: v1.7 DrawingML material and visual-QA refinement

This release continues **full-frame grading on all 12 slides** of the supplied
PowerPoint export ZIP at **1536 × 864**, including logos and footers. It does not
attempt to guess or hardcode any particular missing detail in this deck.

- **Cached SmartArt stroke materials:** Diagram-card outlines can use the same
  semi-transparent, tinted OOXML material as their body fills. The browser
  previously rendered the outlines as over-saturated yellow, even when the
  cards themselves were a soft beige. The renderer now applies the same
  material-color resolution to those strokes. This generalizes to cached
  SmartArt shapes with similar DrawingML styling.
- **Broader DrawingML color semantics:** `color_node` now handles hue and
  saturation modifiers, HSL luminance modifiers, and alpha modifiers in the
  order supplied by the XML. Previously some modifiers were ignored and
  luminance changes were approximated by multiplying RGB channels directly.
  The supplied regression deck does not exercise every transform, so those
  additions have dedicated synthetic tests.
- **Spatial visual grading:** `tools/compare_versions.py` now produces
  `regions.csv`, containing before/after error in a 3×3 grid per slide.
  This supplements the full-slide `scores.csv` to help detect local regressions
  that might otherwise hide inside an improving average.

**Measured outcome:** full-slide mean RGB absolute error **10.619 (v1.6) →
10.593 (v1.7)**. Slide 7 improves **13.025 → 12.711**; the remaining eleven
slides are unchanged in the verified Chromium captures. That is a **small**
benchmark improvement, not PowerPoint parity. All 12 slides were rendered and
compared, with no footer crop. Color transform support is a more substantial
engineering improvement for future PPTX files but has no further measurable
impact on this particular 12-slide example.

**Verification:** 51 automated Python tests passed against the real PPTX and
synthetic cases; 12 browser-rendered screenshots completed without JavaScript
errors. Test environment: Linux/Python 3.13/Chromium. Windows/Python 3.12 is
still the deployment target and needs your local verification. **No new pip
packages, binary dependencies, web frameworks, network services, or build
steps** are required. Launch with `python server.py` as before.

## NEW: v1.6 transparency and inherited text effects

This pass uses the **entire 12-slide exported PNG reference set**, at the
original 1536 × 864 pixel dimensions, including the Georgia Tech footer.

- **Semi-transparent SmartArt fills:** The cached diagram body panels contain
  a DrawingML `a:alpha val="90000"` property. Prior builds approximated the
  final beige color with opaque fills, hiding the underlying ASDL watermark.
  The renderer now preserves the 90%-opaque fill and composites it over the
  actual browser slide background. The source material color is reconstructed
  from a neutral theme approximation and the declared slide backdrop; this is
  still not a complete PowerPoint SmartArt material renderer.
- **Text shadows inherited from slide masters:** Titles can inherit an
  `a:outerShdw` even when their own slide objects have empty `p:spPr` blocks.
  SlideScope now resolves the explicit shadow and exposes it to the browser as
  CSS text shadow, with the OOXML color, direction, distance, and blur.

**Visual grading caveat:** Slide 7 improves **13.095 → 13.025**, but the
**12-slide mean changes 10.598 → 10.619** (a tiny regression of 0.021 RGB MAE).
Native Office and Chromium rasterize shadowed text differently; the shadow is
present in the original PowerPoint master and visually evident in the exported
reference, but it increases strict pixel differences slightly. Do not interpret
this release as a better global pixel score. The supplied reference PPTX and
PNG export ZIP are *not* bundled. All 48 Python tests pass with the test deck.

No extra packages or binaries. Work target remains Python 3.12 on Windows;
automated rendering here was Linux/Chromium/Python 3.13.

## NEW: v1.5 text-flow and kerning refinement

This pass fixes three independent rendering details, tested against the same 12
PowerPoint-exported PNGs as v1.4 at **1536 × 864**, including the footer:

- Subtitle placeholders with no explicit line-spacing setting now use a
  PowerPoint-like 1.44 line-height rather than the normal body-text line box.
  This corrects the author/revision blocks on both title slides.
- Reduced-font normal-autofit text with deliberately empty paragraphs now
  distributes the inherited percentage spacing *between* visible paragraphs
  rather than above the first. This corrects vertical rhythm on the experiment
  slide, without changing explicitly set paragraph spacing.
- Direct `a:rPr kern="0"` properties are passed from the Python parser to
  browser spans, respecting presentations that disable font pair kerning.

**Full-slide average absolute RGB difference: 10.882 (v1.4) → 10.598 (v1.5).**
Slides 1, 2, 6, 9 improved and the other eight did not change measurably.
All 12 slides were graded (no footer excluded). The reference deck and image
exports are user test materials; they are not bundled in the runnable ZIP.

This is still a browser approximation of PowerPoint, and differences remain
in detailed font rasterization, SmartArt text, and Office-specific effects.
The app continues to run with `python server.py`, offline, with no new packages.

## Previous: v1.4 native background tiling and soft-line breaks

This pass compares **all 12 full-frame screenshots** against the PowerPoint
PNG exports in the user's ZIP at the source 1536 × 864 resolution, including
the bottom-right footer. The v1.3 baseline was re-rendered using the same
Chromium session and scoring script as v1.4. No slides were excluded.

- **OOXML background image tiling:** Read `a:blipFill/a:tile` instead of
  stretching every slide background image to the slide's dimensions. The
  tile dimensions are derived from the image's pixel dimensions, embedded
  resolution (DPI), and DrawingML `sx`/`sy`. The slide canvas uses CSS
  background repeat and tile size rather than changing the underlying asset.
  This corrects distorted building photography on both section/title slides.
  Tiles with unsupported transforms or exotic coordinate mappings may still
  differ; this is not a complete DrawingML bitmap fill implementation.
- **Explicit PowerPoint line breaks:** Preserve `a:br` nodes within a
  paragraph in their original sequence, including the adjacent runs' distinct
  font sizes. Previously `python-pptx` returned the breaks as vertical tabs
  in the paragraph text, but the browser renderer only received its `.runs`
  and collapsed the title into one line. This restores slide 9's two-line
  heading without hardcoding the title text.

**Full-frame average RGB error: 11.638 (v1.3) → 10.882 (v1.4)**, about
**6.5% lower**. Slide 1 improves **12.016 → 8.531** and slide 9 improves
**14.594 → 9.017**. The other ten slides are unchanged in this regression deck.
All 12 PowerPoint exports and the matching v1.3/v1.4 screenshots are included
in the separate offline QA report. Tests: 43 passed with the regression deck.
Linux/Chromium/Python 3.13 was used for screenshot QA; the work target remains
Windows/Python 3.12 and requires a local run. No new dependencies, binary
installers, network/CDN requests, or additional server infrastructure.

## NEW: v1.3 unscaled normAutofit paragraph-leading correction

The PowerPoint reference ZIP is the grading baseline for **all 12 slides**, at
1536 × 864, **including** Georgia Tech footer artwork. No slides were omitted
from validation and the improvement is not based on a hand-picked crop.

**Bug addressed:** Some PowerPoint content placeholders use normal autofit but
contain explicitly sized text runs significantly smaller than the slide-master
paragraph style. In v1.2, the CSS line-height calculation combined autofit
`lnSpcReduction` with browser font metrics in a way that **collapsed the gaps
between nested bullets**, making slides 4 and 5 substantially more compact than
PowerPoint. Other text frames were already correct and must not be disturbed.

The browser now preserves about 0.2 em of additional text leading **only** when
all these conditions hold:

- `a:normAutofit`/`TEXT_TO_FIT_SHAPE` is active at full font scale,
- there is a local text run smaller than 85% of the inherited paragraph size,
- the frame specifies positive line-spacing reduction.

This is a targeted, explicitly documented typography approximation, **not an
OOXML-complete line-layout algorithm**. The correction is based on font-size
inheritance and run properties, not slide numbers or authored text strings.
Dense scaled-autofit columns and ordinary content placeholders retain v1.2's
handling. We rejected candidates that worsened slide 2 and tested multiple
line-height values against every reference PNG.

**Full-frame mean RGB absolute error:** **12.618 (v1.2) → 11.638 (v1.3)**,
about **7.8% lower**. Slide 4: **14.012 → 8.492**; Slide 5:
**14.840 → 8.590**. The remaining 10 slides show no measurable change.
Every slide is included in the attached offline QA report; screenshots show
previous output, v1.3 output, and the exported PowerPoint reference. Browser
captures on Linux used Chromium with Python 3.13 and contain all slide imagery,
including the EMF footer. Windows/Python 3.12 still requires local validation.

**No new runtime dependencies:** Python 3.12 target, `python-pptx`, `lxml`,
Python standard-library HTTP, and native HTML/SVG/JS. No Node.js, FastAPI,
MCP, custom registry, GitHub-only dependency, or compilation needed.

## Historical: v1.2 PowerPoint text tracking, shape outlines, and footer alignment

This iteration is graded against **every pixel of all 12 PowerPoint-exported PNGs**
(1536 × 864); it is not a selected-slide or footer-cropped comparison.

- **OOXML stroke width:** SVG shapes with `vector-effect: non-scaling-stroke`
  now take their DrawingML stroke width in output pixels, instead of dividing
  it by the object width. This restores black borders around yellow callouts.
- **`a:normAutofit` character advances:** Tightly autofitted text uses a
  conservative tracking correction based on `fontScale`. This addresses
  a browser-vs-PowerPoint glyph-spacing difference in dense text frames.
  It is a visual approximation, **not** a complete Office typography engine;
  ordinary text frames are deliberately unaffected.
- **Master slide-number alignment:** `p:txBody/a:lstStyle` on the master
  slide-number placeholder takes precedence over generic `p:otherStyle`
  alignment. This restores centered slide numbering without per-slide offsets.

The full-frame average RGB error declined **13.117 → 12.618** (approximately
3.8%); slide 6 declined **16.039 → 11.875**. All 12 slides were checked and
none worsened in this test environment. The accompanying offline QA ZIP
contains before/after/PowerPoint images, a browser-readable report and CSV.
These results were obtained in Linux Chromium under Python 3.13; Python 3.12
on Windows remains the target and needs local verification. No runtime
packages, external binary downloads, or server frameworks were added.

Reproduce with the project tool:

```powershell
python tools/compare_versions.py "PowerPointExports.zip" "captures-v11" "captures-v12" --output qa
```

## Historical: v1.1 paragraph metrics and full-deck QA

This pass uses all 12 PowerPoint-exported PNGs from the user-provided regression
ZIP (1536 × 864), including the footer. It improves **empty-paragraph font
metrics**, resolving extra vertical gaps on slides 10–12, and adjusts line
leading for highly autofitted text on slide 6. DrawingML `a:endParaRPr/@sz`
is used instead of a larger inherited master font for empty paragraphs; this
is a general OOXML behavior correction, not a hardcoded rule for the deck.
Text frame content may extend beyond its own bounds, as PowerPoint sometimes
permits; it remains clipped by the overall slide canvas.

Average absolute RGB pixel error, full slide: **14.551 → 13.117** (9.9% reduction).
The largest improvements were slides 10 (**17.47 → 13.14**), 11 (**20.24 →
14.65**), and 12 (**19.64 → 12.96**); slide 6 improved **16.77 → 16.04**.
Slide 4 regressed very slightly (**14.23 → 14.36**) while honoring an explicitly
sized blank paragraph; the other seven slides were essentially unchanged.
The deck's PowerPoint export ZIP is not part of the source package.

Run `python -m unittest discover -s tests` to exercise dependency-free parser
regressions. Set `SLIDESCOPE_REGRESSION_PPTX` to your own PPTX path to enable the
additional real-deck checks. No new PyPI dependencies, server framework, npm,
custom package repository, or graphics applications were introduced. Linux
Chromium/Python 3.13 visual grading passed; Windows/Python 3.12 remains a
local compatibility check. This is **not yet pixel-identical to PowerPoint**.

## NEW: v1.0 cross-platform EMF vector fidelity

**The footer artwork now renders in the Chromium screenshots on Linux as well
as Windows**. This corrects a visible omission in our own full-deck QA images.
SlideScope's server converts supported Enhanced Metafile (EMF) GDI vector
commands directly to browser-native SVG using only Python's standard library.
It preserves the original logo's paths and colors, instead of baking in the
PowerPoint reference export or requiring Inkscape, ImageMagick, a native binary,
or any additional PyPI packages. Individual paths remain a browser-rendered
vector image, while the containing picture object remains selectable in the DOM.

- Supports standard EMF 16-bit polygons, line and cubic Bézier segments,
  pens, brushes, fill paths and stroke paths. It rejects unsupported drawing
  commands rather than silently dropping artwork. On Windows, the existing
  native GDI converter remains as a fallback for unsupported EMFs.
- **Grading now includes every pixel of every slide, including the footer.**
  Against the actual 12-slide PowerPoint PNG ZIP at 1536 × 864, in the same
  Chromium/Linux environment, the mean RGB error decreased **15.569 → 14.551**.
  All 12 slides improved; the change is concentrated in EMF images rather than
  claiming that text layout suddenly reached parity.
- For example, on slide 7 the bottom-right 546 × 79-pixel footer region's
  error decreased **31.81 → 7.71**. Remaining discrepancies include platform
  font rasterization, text metrics and some SmartArt material effects.
- The source ZIP contains no proprietary slideshow or customer assets. The
  PowerPoint reference images are used *only* for external visual grading,
  never to reconstruct rendered artwork.

Run the comparison on your own 12-slide exports with:

```powershell
python tools/compare_versions.py "PowerPointExports.zip" "old-captures" "new-captures" --output qa-report
```

Do not use `--crop-bottom` for release grading. Browser QA in this development
container used Python 3.13 and Chromium/Linux; the application still targets
Python **3.12** on Windows, which requires your local smoke test.

## NEW: Load real PowerPoint reference images

1. Open the `.pptx` normally in SlideScope.
2. Export all slides to PNG in PowerPoint, which produces `Slide1.PNG`, `Slide2.PNG`, ... .
3. ZIP the exported slide files or their containing folder (do not ZIP the original `.pptx`).
4. Click **Open reference ZIP** in SlideScope and select that ZIP.
5. The canvas enters **SPLIT** mode: its **left half** is SlideScope and its **right half** is the PowerPoint-exported PNG. The middle line separates them. Both halves use identical dimensions and slide coordinates. Click **COMPARE: SPLIT** to cycle: **REFERENCE** (full original), **ONION** (50% transparency), **OFF**, **SPLIT**. Navigation moves both images together.

The reference ZIP stays on the local Python server in memory. No CDN or external upload. Reference images are **not** baked into the native scene graph and don't interfere with object selection; they can be hidden for Playwright object inspection. If the browser refreshes, both presentation and reference image archives persist until the server is restarted.

## Historical: v0.9 full-deck visual fidelity pass

SlideScope v0.9 was tested on **all 12 slides** of the PowerPoint export ZIP
at 1536×864 pixels, not just example slides. It fixes master/layout text-frame
inset inheritance (which python-pptx otherwise reports as a generic inset) and
negative hanging indents on paragraphs with `a:buNone`. This resolves systematic
misalignment and manually numbered text without slide-specific placements.
Both fixes preserve editable DOM slide objects and do not add any packages.

Average RGB absolute error on slide area above the bottom 89 pixels:
**18.143 → 15.744** (lower is better, approximately **13.2%** improvement).
In v0.9 this cropped diagnostic was necessary because Linux could not render
EMF footer artwork. **This limitation is resolved for supported vector EMFs in
v1.0**; use full-frame metrics for current release grading.
Slides 1 and 9 (centered cover layouts) are unchanged; the other ten improve.
This is meaningful progress but not Office-level parity.

| Slide | v0.8 | v0.9 | Improvement |
|---:|---:|---:|---:|
| 1 | 13.23 | 13.23 | -0.01 |
| 2 | 11.58 | 10.23 | +1.34 |
| 3 | 11.64 | 9.09 | +2.56 |
| 4 | 17.90 | 15.37 | +2.53 |
| 5 | 19.58 | 16.21 | +3.36 |
| 6 | 19.54 | 18.33 | +1.21 |
| 7 | 16.75 | 14.74 | +2.01 |
| 8 | 20.63 | 14.46 | +6.16 |
| 9 | 16.11 | 16.11 | -0.00 |
| 10 | 24.21 | 19.05 | +5.17 |
| 11 | 23.39 | 22.17 | +1.22 |
| 12 | 23.16 | 19.93 | +3.23 |

The accompanying **12-slide QA report** (`index.html`, `scores.csv`, and
`Slide1.jpg`–`Slide12.jpg`) shows old render / new render / PowerPoint reference
for every slide. Reproduce the comparison using the existing offline utility:

```powershell
python tools/compare_versions.py "references.zip" "screenshots-v08" "screenshots-v09" --crop-bottom 89 --output qa-report
```

The reference case is not bundled in the source ZIP; keep your local PPTX and
PowerPoint-exported images in a private test directory.

## Fidelity pass in v0.8 (graded against exported PowerPoint PNGs)

The 12 exported PowerPoint PNGs are the acceptance reference, at their **native 1536 × 864** dimensions. This pass:

- Removes a built-in 0.5% inset on SVG rectangle and ellipse geometry (actual object bounds are used).
- Improves muted SmartArt fill colors by using a warm, low-saturation material approximation rather than a saturated translucent yellow.
- Improves vertical rhythm of cached SmartArt, including spacing retained by empty paragraphs.
- Uses more faithful baseline leading in plain HTML text blocks when no explicit OOXML line spacing is specified. The OOXML-specified spacing continues to take priority.
- Adds `tools/compare_versions.py`, a fully offline QA report generator comparing previous/current screenshots with the original PowerPoint PNG export ZIP. It produces per-slide scores, a CSV, and visual three-way comparisons. Requires only Pillow (a python-pptx dependency).

This is a **renderer-wide pass**, not hardcoded slide positions. No new runtime or QA package installations.

### Repeat the visual grade

Capture *unscaled* screenshots from both versions with the existing browser agent, each named `Slide1.png`, `Slide2.png`, etc. Then:

```powershell
python tools/compare_versions.py "C:\work\PowerPointSlides.zip" "C:\work\screenshots-v07" "C:\work\screenshots-v08" --output qa-report
```

Open `qa-report/index.html` locally. The tool also prints mean absolute RGB differences (smaller is better). Optional `--crop-bottom 89` isolates the slide body when comparing **Linux** browser renders that cannot show Windows GDI EMF footers; report images always show the entire slides. **Never treat cropped scores as an overall presentation-fidelity claim**.

## Earlier development passes

## SmartArt improvements in v0.7

- Resolves paragraph styles embedded in *ordinary text boxes* (`a:lstStyle`), not just placeholders. This restores slide 6's right-hand bullets, indent levels and correct inherited type sizes.
- Restores more faithful paragraph spacing, with separate metrics for autofitted text boxes and standard content placeholders. Blank cached SmartArt paragraphs retain additional baseline spacing.
- Applies a limited **neutral-pastel approximation** for cached SmartArt panel fills that specify both tint and transparency, using a neutral color from the deck's theme. This improves the example deck substantially, but is a renderer heuristic, not a complete implementation of PowerPoint's SmartArt material pipeline. Other SmartArt styles may still be inaccurate.
- All changes are inside the renderer. The read-only viewer and existing Playwright integration are unchanged. No new runtime packages.

## SmartArt

The parser now recognizes cached Office SmartArt drawing files (`ppt/diagrams/drawingN.xml`), expands native shapes with text and coordinates into the browser scene graph, and preserves their stacking order. This supports diagrams that include the cached drawing, without recreating SmartArt's complex automatic layout engine. Drawings without a cache remain unsupported; this is not complete SmartArt parity.

In the 12-slide regression deck, **slide 7's previously entirely absent three-column SmartArt** now appears. Cached SmartArt surfaces, text metrics and theme material effects are still approximations. Rendering fidelity is still incomplete. The demonstration deck and reference images used during development are **not included in this ZIP**; use your own originals locally.

## Playwright-friendly browser API

```js
await page.goto('http://127.0.0.1:8000');
await page.locator('#filePicker').setInputFiles('C:/work/presentation.pptx');
await page.locator('#referencePicker').setInputFiles('C:/work/PowerPointSlides.zip');
await page.evaluate(() => window.PPTX_HARNESS.goToSlide(7));
await page.evaluate(() => window.PPTX_HARNESS.setCompare('split'));
await page.getByTestId('slide-canvas').screenshot({path:'slide7-split.png'});
await page.evaluate(() => window.PPTX_HARNESS.setCompare('off'));
const objects = await page.evaluate(() => window.PPTX_HARNESS.getObjects());
const warnings = await page.evaluate(() => window.PPTX_HARNESS.getWarnings());
```

Public browser API: `getDeck`, `getObjects`, `getWarnings`, `getState`, `goToSlide`, `selectObject`, `inspect`, `setCompare`, `getReferenceState`. SmartArt shapes are selectable like other objects; `data-testid="slide-object"` is stable. Press **B** for shape boundaries, arrow keys to change slide.

## Tests and optional pixel-level metrics

```powershell
python -m unittest discover -s tests -v
```

If you have a folder of matching **unscaled** SlideScope screenshots named `Slide1.png` etc., you can compare them with PowerPoint exports using Pillow (already installed transitively with python-pptx):

```powershell
python tools/visual_diff.py "C:\path\to\references.zip" "C:\path\to\screenshots"
```

A lower RGB MAE score means closer pixel agreement, but is sensitive to fonts, scaling, antialiasing and unsupported EMF graphics. For a *diagnostic* comparison that excludes the last 89px of the 864px slides, pass `--crop-bottom 89`. This should not be used to hide actual rendering failures.

## Dependencies and limitations

Runtime remains `python-pptx` and its public PyPI dependencies (`lxml`, `Pillow`, etc.), all fetched through official public PyPI with prebuilt wheels. The webserver uses the Python standard library. No additional runtime package was introduced in v1.0; local QA used an existing Chromium/Playwright installation, not included in the project. Python 3.12 on Windows is the *target*, but automated tests in this development container used Python 3.13/Linux; your Windows test remains essential.

The viewer handles many text, shape, table, image, grouped-shape, theme, gradient and chart scenarios, but text wrapping and rich Office styles are not pixel perfect. Most simple vector EMF files now have a cross-platform native SVG conversion
path; unsupported metafile commands still require the experimental Windows GDI fallback. Full SmartArt, sophisticated shape effects, text metrics, advanced charts, animation and OLE objects need more work. This version is **read-only** for PPTX. Server defaults to `127.0.0.1` so other computers cannot connect unless explicitly configured; uploads stay local.
