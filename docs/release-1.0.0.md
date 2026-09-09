**[Download Plain-for-Windows-1.0.0-x64.exe](https://github.com/keithadler/plainwin/releases/latest/download/Plain-for-Windows-1.0.0-x64.exe)** for ordinary Intel and AMD PCs, or **[Plain-for-Windows-1.0.0-arm64.exe](https://github.com/keithadler/plainwin/releases/latest/download/Plain-for-Windows-1.0.0-arm64.exe)** for Windows on ARM.

Windows 10 or 11. One exe, no installer, no runtime, no administrator. Put it anywhere and double-click. Windows will warn about an unknown publisher: **More info**, then **Run anyway**.

The console twin, for scripts: [plain-1.0.0-x64.exe](https://github.com/keithadler/plainwin/releases/latest/download/plain-1.0.0-x64.exe) and [plain-1.0.0-arm64.exe](https://github.com/keithadler/plainwin/releases/latest/download/plain-1.0.0-arm64.exe).

---

Opens Word, Excel and PowerPoint files, edits the basics, and never damages what it doesn't understand.

Every other editor that opens a `.docx` reads it into its own idea of a document and writes that idea back. Whatever the idea has no room for is quietly gone: the chart, the macro, the tracked changes, the header nobody looked at. Plain opens the file as the package of parts it really is and writes back every part it did not edit as the exact bytes it found.

**Open a file, save it without changing anything, and you get the same file back, byte for byte.** That is a claim a test can check, and it is checked against real documents on every build. Check it yourself with `plain roundtrip *.docx *.xlsx *.pptx`.

**What it does.** Cells, formulas and every sheet of a workbook, with rows and columns that can be inserted and deleted and formulas that follow. The body text of a document with headings, tables and lists. The text on each slide. Find and replace across the whole file, print, save as PDF, comments and tracked changes accepted or turned down one at a time, document properties stripped before a file leaves the building, and the pictures inside it shown so you can see what you are sending.

**What it keeps but does not show.** Charts, pivot tables, macros, SmartArt, embedded objects, layouts, masters, themes, headers, footers. All of it is in the saved file unchanged, and the panel on the right tells you what is there.

**Honest limits.** No page layout, so printing and PDF give you what Plain shows rather than Word's pages. No fonts, colours or alignment. English only, left to right. And it has never been opened in Microsoft Office: everything here is checked against LibreOffice, against macOS's PDF engine, and against its own byte-for-byte round trip, which are good proxies that have caught real bugs, but are not the same thing.

Free, MIT, no account, no cloud, and no network code in it at all.
