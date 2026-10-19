# SlideScope v1.8 release notes

Target: Python 3.12, Windows; run `python server.py` from the `SlideScope` folder.

- Browser text pairs: optional kerning disabled for normal slide objects by default, while explicit OOXML run kerning remains honored and cached SmartArt keeps separate metrics.
- Fixed-size text line spacing from `<a:lnSpc><a:spcPts ...>` is interpreted as point units and emitted as CSS pixels.
- Same no-new-install runtime dependencies: `python-pptx` and `lxml`, plus Python standard library and native browser APIs. Both are available from official PyPI with wheels; Python 3.12 support is targeted.
- Full-deck image QA: 10.593 -> 10.319 mean RGB absolute error on 12 slides; no slide regressions.
- 54 Python tests passed. Browser screenshots and tests ran on Linux/Python 3.13; Windows/Python 3.12 still needs local verification.

The original presentation and reference export ZIP are not included in the runtime package. The QA report is distributed separately.
