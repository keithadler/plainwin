# Changelog

## 1.1.0, 2026-09-09

**Sort rows.** By any column, up or down, numbers by value rather than as text and blanks always last. Plain refuses whenever sorting would make a formula mean something else: one inside the block, one reading only part of it, or a lookup whose answer depends on the order. A total underneath a table reads every row of it, and shuffling rows does not change a total, so that case goes ahead.

**Column widths.** Drag the edge of a column heading to resize it; double-click the edge to fit it to its longest value. Widths are written where the file keeps them, so Excel opens the sheet the same way.

**Keep the header on screen.** Freeze rows at the top so they stay while the rest scrolls. Read from and written to the sheet's own pane settings, so a sheet that already had a frozen header keeps it, and one you freeze here is frozen in Excel too.

**Page setup for printing and PDF.** Paper (A4, Letter, Legal, A3, A5), sideways or not, how much white at the edges, and a line at the top and bottom where `{page}` and `{pages}` are filled in.

**Open from Explorer.** An optional "Edit in Plain" line on the right-click menu for Office files. Written under your own account only, so it never asks to be an administrator, and it does not make itself the default for anything. Turning it off takes back exactly what it added, and nothing else.

**A daily check for a new version.** One request to GitHub's releases API, a line in the status bar if there is a newer version, and nothing else: no identifier, nothing about your files, nothing downloaded, nothing installed. Off in Reading and settings; the console twin never checks at all. Plain shipped saying it had no network code, so the README, Help, PRIVACY and the rest now say what the one request is and how to stop it.

**Fixed: bullets and curly quotes in a PDF.** A bullet came out as `€42`. Characters above 126 were written as an octal escape of their Unicode number, and a reader takes the first three digits as one character and prints the rest as text. They are now written where the built-in fonts keep them. This had been wrong since 1.0.0.

**Console twin.** `plain sort`, `plain width`, `plain freeze`, `plain explorer on|off|status`, and `--paper`, `--landscape`, `--margin`, `--header`, `--footer` for `plain pdf`.

## 1.0.1, 2026-09-08

**Typing into a new document and saving no longer loses the words.** The editors hand their text over when they lose the caret, which is what keeps a long document from rebuilding on every keystroke. Nothing forced that handover before a save, and a file with a single box in it, which is exactly what a new blank document is, never loses the caret on its own. So you could make a new document, type a page, press Ctrl+S, and save an empty file. Every path that reads the file now takes the caret away first: save, save a copy, print, export to PDF, and the half-minute keeper that guards against a machine stopping.

**The interactive test suite was measuring an empty window.** It waited for the window to appear and then began typing, but opening the file happens after the frame is on screen, and on a slow machine that is seconds later. Four checks were failing for that reason and several others were passing without proving anything. The harness now waits for the title to name the file before it touches the keyboard.

## 1.0.0, 2026-09-08

First public version.

**The promise.** Plain opens a Word, Excel or PowerPoint file as the package of parts it really is, holds every part as the exact bytes it occupies, and writes those same bytes back for every part it did not edit. Open a file, save it without changing anything, and you get the same file back byte for byte. Checked on every build against real documents, and checkable yourself with `plain roundtrip`.

**Spreadsheets.** Cells, text and formulas across every sheet, with the column widths and number formats the file stores. Insert and delete rows and columns, with every formula in the workbook rewritten so it still means what it meant; one that pointed only at a deleted row becomes `#REF!` rather than a quietly wrong number. Totals are worked out as you go, covering arithmetic, comparisons, text, lookups, conditional sums and dates, and refusing anything not fully understood. Number formats, bold and italic, fill down, block select, copy and paste as tab separated text.

**Documents.** The body text with headings and tables. Bold, italic, heading levels, bullets and numbered lists, using definitions the document already carries. Word count.

**Presentations.** The text on each slide, from a rail of slides.

**All three.** New empty files, built part by part from what each format requires. Find and replace across the whole file. Save, Save a copy, Print, Save as PDF. Comments and tracked changes read, accepted or turned down one at a time. Document properties viewed, set and stripped. The pictures inside a file, shown and saveable. Multi-step undo that survives a save. Unsaved work kept every half minute against a machine that stops. Reading settings for the typeface, paper colour, line spacing and line width. Bigger text, and F6 to move between the parts of the window.

**Console twin.** `plain.exe` does all of it for scripts, including a batch mode that applies many changes in one pass.

**Robustness.** Both the transitional and the strict ISO ways of writing a file are read. A damaged file gets a sentence, not a crash: the readers are fuzzed against truncated, zeroed, overwritten and bit-flipped copies. A save writes beside your file and then replaces its contents, so its permissions and creation date survive and an interrupted save leaves the original whole. A file changed by something else while Plain had it open is noticed and asked about.

641 checks in the engine, 36 end to end, 21 driving the real window, on Windows on ARM and on x64.
