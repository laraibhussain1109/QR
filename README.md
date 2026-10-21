# Manufacturing QR scanner

Desktop application using **PyQt6**, **OpenCV** and **openpyxl**. It displays the
camera feed, outlines decoded QR labels, overlays the part name and bin number,
and automatically writes each new Part ID to an Excel workbook.

## Install and run

Use Python 3.10+ on a computer with a desktop session and a webcam.

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux instead:
source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

Choose the Excel destination and press **Start camera**. The default is camera 0
and `manufacturing.xlsx` in the current directory. Try camera 1 if another camera
is selected. Windows users can also choose DirectShow in the backend dropdown.
Grant camera permission in your OS settings. Stop releases the camera; it can be
started again without closing the application.

Optional arguments:

```bash
python app.py --camera 1 --output inventory.xlsx
```

The Excel sheet is named `manufacturing`, with the same columns as the supplied
example:

| Part Name | Part ID | BIN Number | TimeStamp |
| --- | --- | --- | --- |
| Steel Bracket | P-001 | B-12 | Local ISO timestamp including UTC offset |

The workbook is created on Start, even before the first scan. Each successful
scan is saved immediately. Keep it **closed in Excel while scanning**; if writing
fails, the UI reports the error and retries while the label remains visible.

## QR label contents

Encode this JSON as text in each QR code. All three values must be strings, so
IDs such as `001` retain their leading zeroes.

```json
{"part_name":"Steel Bracket","part_id":"P-001","bin_number":"B-12"}
```

A single CSV row is also supported: `Steel Bracket,P-001,B-12`. Quote fields
containing commas. Plain IDs, URLs and arbitrary labels are rejected; no OCR or
object recognition is performed. Fields must be nonempty, without control
characters, and at most 256 characters each.

Duplicate prevention uses the **trimmed, case-sensitive Part ID across the entire
workbook**, even after restarts. A previously recorded part is never appended
again, even if its bin changes. Conflicting details are reported and the original
Excel row is retained. To record stock movements repeatedly, the uniqueness rule
would need to change. Existing workbooks must have the exact four headers shown
above in the `manufacturing` sheet. Other sheets are retained; use a dedicated
workbook because openpyxl does not preserve every advanced Excel feature.

Writes use a sidecar lock between app instances and an atomic file replacement.
Excel and other third-party editors do not honor the app's lock; do not edit the
file concurrently. Failed saves preserve the previous workbook. QR text is
stored as text cells rather than executable Excel formulas.

## Try it without a camera

Generate a printable sample label, then scan the same image twice:

```bash
python app.py --make-sample sample-qr.png
python app.py --scan-image sample-qr.png --output demo.xlsx
python app.py --scan-image sample-qr.png --output demo.xlsx
```

The first run saves a row; the second reports a duplicate. The image mode needs
no graphical display. Print the sample or show it on another screen to exercise
the live camera. Multiple readable QR labels in one frame are supported.

## Development checks

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

Tests cover real QR-image decoding into Excel, restart and concurrent duplicate
prevention, invalid labels, safe save failures, and an offscreen PyQt camera
workflow using a simulated capture source. A real webcam and desktop display
still need validation on the target computer. Cloud tasks generally have neither.
On Linux, PyQt's desktop platform plugin may require your distribution's Qt/X11
runtime libraries; `QT_QPA_PLATFORM=offscreen` is for tests, not normal camera use.

The cloud environment is already isolated; use the existing `/workspace/QR`
checkout rather than creating another worktree.
