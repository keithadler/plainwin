**[Download Plain-for-Windows-1.0.1-x64.exe](https://github.com/keithadler/plainwin/releases/download/v1.0.1/Plain-for-Windows-1.0.1-x64.exe)** for ordinary Intel and AMD PCs, or **[Plain-for-Windows-1.0.1-arm64.exe](https://github.com/keithadler/plainwin/releases/download/v1.0.1/Plain-for-Windows-1.0.1-arm64.exe)** for Windows on ARM.

Windows 10 or 11. One exe, no installer, no runtime, no administrator. Put it anywhere and double-click. Windows will warn about an unknown publisher: **More info**, then **Run anyway**.

## Typing into a new document and saving no longer loses the words

This is the whole reason for the release, and it was a bad one.

The editors hand their text over when they lose the caret, which is what keeps a long document from rebuilding on every keystroke. Nothing forced that handover before a save. A file with a single box in it, which is exactly what a new blank document is, never loses the caret on its own. So you could make a new document, type a page, press Ctrl+S, and save an empty file.

Every path that reads the file now takes the caret away first: Save, Save a copy, Print, Save as PDF, and the half-minute keeper that guards against a machine stopping. An existing file with several paragraphs was much less likely to hit this, because clicking from one paragraph to another already committed the last one, but the same fix covers it.

## The interactive tests were measuring an empty window

The test harness waited for the window to appear and then began typing. Opening the file happens after the frame is on screen, and on a slow machine that is seconds later, so the keystrokes went nowhere. Four checks were failing for that reason, and several others were passing without proving anything: "undo puts the old value back" passed because nothing had ever changed. The harness now waits for the title to name the file before it touches the keyboard, and the suite has a check that follows the exact path this bug was found on.

Both the console suite and the interactive suite pass in full against this build, and the released 1.0.0 fails the new check, which is how the fix was confirmed to do something.

## Also

The pictures in the README are taken from a fuller demo file now: a quarterly sheet with real formatting, three sheets and working formulas, a review document with headings, a table and a list, and a six-slide deck. The README says what the download actually weighs, under 60 MB rather than the 57 MB it claimed.

Nothing about the file format changed. A file written by 1.0.0 and a file written by 1.0.1 are the same bytes.
