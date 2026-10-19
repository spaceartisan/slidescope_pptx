# SlideScope v1.9 — first-paragraph fidelity and numbered bullets

This release maintains the Python 3.12/Windows target and the original `python server.py` offline workflow. There are **no new Python packages or runtime dependencies**.

## Changes

1. **Inherited first-paragraph leading:** non-autofitted PowerPoint content placeholders can carry inherited paragraph spacing that Chromium applies above their first line. SlideScope now keeps the same overall paragraph flow height but positions the first line closer to Office's output by placing that leading after the first paragraph. Other paragraphs are not moved. This is an empirically verified browser/Office compatibility approximation, not a claim of perfect OOXML layout parity.
2. **No-wrap text:** respects PowerPoint text frames explicitly marked not to wrap (`bodyPr@wrap=none`) rather than forcing CSS soft wrapping.
3. **Automatic numbered bullets:** reads `a:buAutoNum` type and start from direct paragraphs and inherited text styles, rendering common Arabic, alphabetic, and Roman numeral forms as browser-inspectable text.

## Full-deck visual QA

PowerPoint reference: all 12 slides exported as 1536×864 PNG images (including the Georgia Tech footer). Comparisons use isolated Chromium rendering without internet requests.

| Metric | v1.8 remeasured | v1.9 |
|---|---:|---:|
| Average RGB MAE (0–255) | 10.929 | 10.510 |
| Slides improved | — | 5 of 12 |
| Slides unchanged | — | 7 of 12 |
| Slides regressed | — | 0 |

Slides 3, 8, 10, 11, and 12 improved. The other seven held steady. This session's test harness captures directly from the browser slide canvas; its **remeasured v1.8 baseline differs from the historic 10.319 score**, so comparisons between these two numbers from different capture sessions would be misleading. See the accompanying QA report for exact same-harness before/after scores and screenshots.

Target verification still required on the user's Windows/Python 3.12 environment. Development/test Chromium and Python versions do not establish Windows parity.
