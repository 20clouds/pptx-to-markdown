# PPTX to Markdown

Drag any PowerPoint file (`.pptx`, `.ppt`, `.pptm`, `.ppsx`, `.potx`, …) onto **PPTX to Markdown** on the Desktop.
You can drop several at once. When it finishes you get a notification, and the output folder opens.

```
Lecture5/                         created next to Lecture5.pptx (" (2)", " (3)"… if it already exists)
  Lecture5.md
  images/
    slide3_img01.png              picture first seen on slide 3
    slide3,7,12_img01.png         same picture on slides 3, 7, 12, saved once
    slide3_full.png               full-slide snapshot, only for slides with visuals (see below)
```

A slide gets a snapshot only if it has a picture, chart, SmartArt diagram, embedded object, equation, ink, or a diagram drawn from 3+ shapes/connectors (flowcharts).
- Text-only and table-only slides don't get one, because the markdown already captures them.
- Icons don't count as pictures. An icon is either under 1.5% of the slide, or under 10% and repeated on 3+ slides (logos, the "question" badge).
- The snapshot's alt text says why it was taken, e.g. `(chart, picture)`.
- `--all-snapshots` / `--no-snapshots` on the command line override this.

- Every PNG has a strip above it naming its slide(s). Lists longer than 10 slides collapse into ranges, e.g. "Slides 1–84".
- `Lecture5.md` has one `## Slide N: Title` section per slide, in this order:
  - the snapshot, when it has one
  - the slide's text in reading order, keeping bullets, bold/italic, links and sub/superscripts
  - tables, plus chart data as tables, plus SmartArt text
  - the images
  - `### Speaker Notes`
- Hidden slides are included and tagged `(hidden)`.
- A footer text box repeated near the bottom of many slides is left out. The front matter lists it.
- Pictures that come from the slide master or layout (logos, backgrounds) are not extracted.

## Setup

Requires macOS, Python 3, and Microsoft PowerPoint for the snapshots and `.ppt` support.

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
./build_app.sh
```

`build_app.sh` clones an existing Automator droplet, `~/Desktop/Convert to Markdown.app`, as the app shell. Any Automator "Application" containing a single **Run Shell Script** action (shell `/bin/zsh`, input "as arguments") will do.

## Files

| File | Purpose |
|---|---|
| `pptx2md.py` | The converter. Also works from the terminal: `.venv/bin/python pptx2md.py [--no-snapshots \| --all-snapshots] [--no-open] [--no-notify] FILE...` |
| `build_app.sh` | Rebuilds `~/Desktop/PPTX to Markdown.app` by cloning `Convert to Markdown.app`. Re-run it if you move this folder. |
| `make_icon.py` | Draws the app icon. |
| `.venv/` | Python environment (python-pptx, pillow, lxml, pypdfium2). |

Log: `~/Library/Logs/pptx2md.log`

## Notes

- Snapshots and `.ppt` support drive Microsoft PowerPoint in the background. macOS asks once for permission to control PowerPoint. If you denied it, enable it under System Settings → Privacy & Security → Automation.
- PowerPoint is sandboxed and can only read files the converter stages in `/tmp`. Don't change the work folder to `$TMPDIR`: under `/var/folders`, PowerPoint stops and waits on a hidden "Grant File Access" dialog.
- If PowerPoint fails, the conversion still finishes without snapshots. The `.md` front matter and the notification say why.
