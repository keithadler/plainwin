# Changelog

## 1.0.0, unreleased

First working version.

The package layer holds every part of a Word, Excel or PowerPoint file as the exact bytes it occupies, compression included, so a part nobody edited is written back byte for byte and an unchanged save returns the identical file. Verified against a corpus of real Office documents.

Excel: cells, text, formulas, shared strings, number formats and the column widths the file stores. Plain does not evaluate formulas, so instead it reads which cells each formula depends on and clears the cached value of exactly those formulas an edit made untrue, following the chain across sheets. Every other total keeps its number. Word: the body text, with headings, lists and tables. PowerPoint: the text on each slide.

The window: one command row of six controls, a formula bar for spreadsheets, a panel listing every part the app is preserving with its size and what it is, and a status bar that counts the parts on every save. Follows the Windows light or dark setting. Multi-step undo. Drag a file in to open it.

Files written with the strict ISO namespaces, which Excel offers as "Strict Open XML", are read as well as the usual transitional ones, and an edit to one keeps it strict. A workbook whose sheets Plain cannot find, or a document whose text it cannot find, is now refused with a sentence instead of opening to an empty window.

A second round of user panels, ten more kinds of person, and the ten things they most needed: PDF export, tracked changes one at a time, unsaved work kept when the machine stops, many changes in one console pass, the pictures inside a file, bullets and numbered lists, CSV in and out, the lookup and conditional and date functions, reading settings for the font and paper and spacing, and a word count.

The preserved panel says what is in the file rather than listing its innards: several parts of a kind are one line with a count, and bookkeeping is a single closing sentence. On a set of 39 real documents that is 847 rows down to 86, and the busiest file goes from 44 rows to 5.

New makes an empty workbook, document or deck, built part by part from the smallest set the format requires. Two blank files of a kind are the same bytes, which is what lets a test check them. LibreOffice opens all three and reads back what Plain writes into them, including a formula it then computes.

Blocks of cells can be selected with Shift, then copied, cut and pasted as tab separated text, the format every spreadsheet uses on the clipboard. A paste is one step of undo however many cells it fills.

Every sheet of a workbook is reachable from a strip along the bottom, and each sheet keeps its own grid and selection once you have visited it. Find (Ctrl+F) runs over the whole file and, in a spreadsheet, over the formulas as well as the values. Save a copy writes the file under a new name and leaves the original alone.

`plain.exe`, the console twin: `info`, `parts`, `text`, `cells`, `get`, `set`, `roundtrip`, `selftest`.

A damaged file produces a message rather than a crash: the readers are fuzzed against truncated, zeroed, overwritten and bit-flipped copies of every fixture, and no part may unpack to more than half a gigabyte. Saving replaces the file's contents rather than swapping a new file into its name, so its permissions and creation date survive; a file changed by something else while Plain had it open is noticed and asked about; a read-only file is refused with an explanation.

Rows are indexed rather than scanned, so a sheet of twenty thousand rows draws any screen in under a millisecond and searches in 81 ms, where both used to walk the whole sheet for every cell.

Set `PLAIN_SOFTWARE_RENDER=1` if the window comes up blank on a virtual machine or a remote desktop.
