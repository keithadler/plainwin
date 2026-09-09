# Changelog

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
