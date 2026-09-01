# Test-Driven Development — a short slide deck

A 10 slide presentation on Test-Driven Development, written for Unit 8.2 of
Software Engineering (CSC364, Tribhuvan University BSc CSIT, Semester VI).
Content follows Sommerville, *Software Engineering*, 10th edition, chapter 8.

**View it:** https://abhishekbarali.github.io/tdd-presentation/

Everything is in one file, `index.html`. No build step, no dependencies.
Open it in any browser or double click it. The diagrams are inline SVG, drawn in
the file itself, so there are no image files to carry around.

## Controls

| Key | Action |
|---|---|
| Right arrow, space, Page Down | Next slide |
| Left arrow, Page Up | Previous slide |
| Home / End | First / last slide |
| F | Full screen |

The slide area is a fixed 16:9 canvas that scales to the window, so the layout
holds at any size. `Ctrl+P` prints one slide per page if you want a PDF.

## PowerPoint version

`tdd-presentation.pptx` is the same ten slides as a native PowerPoint file, for
when the room has a projector laptop and no browser you trust. It is not a set
of screenshots: every slide is real text boxes and real shapes, so you can edit
it in PowerPoint, Keynote or LibreOffice Impress.

It is generated from the design of `index.html` by `tools/build_pptx.py`, so
`index.html` stays the source of truth. To rebuild it after editing the slides:

```
pip install python-pptx
python tools/build_pptx.py
```

Two differences from the web version, both deliberate. The fonts fall back to
Georgia, Segoe UI and Consolas, which ship everywhere, in place of Instrument
Serif, Inter and JetBrains Mono. The slide transitions and the progress bar
animation are static, since a PowerPoint file has no place for them.

## What it covers

1. Title
2. The idea: write the test first, then the code
3. The cycle, in five steps
4. Red, green, refactor
5. A small worked example
6. Why a new test has to fail once
7. What you get out of it
8. Where it does not help much
9. Testing at the end vs. testing first
10. Summary

## License

MIT for the code. The slide content is coursework notes drawn from the textbook
listed above.
