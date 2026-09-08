# Changelog

## 1.0.0, unreleased

First working version.

The package layer holds every part of a Word, Excel or PowerPoint file as the exact bytes it occupies, compression included, so a part nobody edited is written back byte for byte and an unchanged save returns the identical file. Verified against a corpus of real Office documents.

Excel: cells, text, formulas, shared strings, number formats and the column widths the file stores. Plain does not evaluate formulas, so instead it reads which cells each formula depends on and clears the cached value of exactly those formulas an edit made untrue, following the chain across sheets. Every other total keeps its number. Word: the body text, with headings, lists and tables. PowerPoint: the text on each slide.

The window: one command row of six controls, a formula bar for spreadsheets, a panel listing every part the app is preserving with its size and what it is, and a status bar that counts the parts on every save. Follows the Windows light or dark setting. Multi-step undo. Drag a file in to open it.

Every sheet of a workbook is reachable from a strip along the bottom, and each sheet keeps its own grid and selection once you have visited it. Find (Ctrl+F) runs over the whole file and, in a spreadsheet, over the formulas as well as the values. Save a copy writes the file under a new name and leaves the original alone.

`plain.exe`, the console twin: `info`, `parts`, `text`, `cells`, `get`, `set`, `roundtrip`, `selftest`.

Rows are indexed rather than scanned, so a sheet of twenty thousand rows draws any screen in under a millisecond and searches in 81 ms, where both used to walk the whole sheet for every cell.

Set `PLAIN_SOFTWARE_RENDER=1` if the window comes up blank on a virtual machine or a remote desktop.
