# Changelog

## 1.4.2, 2026-09-09

Ten things that are not about editing documents.

**An icon.** The exe wore the generic Windows icon in Explorer, the taskbar and Alt+Tab.

**Releases are built by a public job now, with checksums.** Every release before this one was built on one laptop and uploaded by hand. Plain is not signed, so the only thing anyone could check was nothing at all. Now the executables are built on a clean runner from a commit anyone can read, and a `SHA256SUMS.txt` is published beside them by the same job. That does not make the download trusted; it makes it traceable, which is the honest version.

**Something going wrong no longer takes your work with it.** An unexpected error used to end in the .NET crash box. Now a copy of anything unsaved is kept first, what happened is written to a file beside the settings, and you are told in a sentence what happened, where the note is, and that the work was kept. Nothing is sent anywhere.

**Document type definitions are refused.** A .docx is XML, and XML has a way of saying "fetch this and paste it in here". .NET happened not to fetch it, but it accepted the file and quietly dropped what the entity stood for, so a document could lose text with nobody told. Now it is refused, said out loud in the code rather than left to a default, and there are checks that a file naming a path or a web address is turned away and that entities which multiply cannot chew up the machine.

**The grid can be read by a screen reader.** It is drawn rather than built from controls, so a screen reader found one blank surface. It now says where you are and what is in the cell as you move. Sixteen other controls got names too.

**F1 lists every keystroke**, and there is a help page inside the exe that works with no network, in English or Spanish.

**Portable mode.** Put an empty file called `plain-portable` beside the exe and everything lives in a folder beside it: settings, recent files, copies of unsaved work. Plain then leaves nothing on the machine, which is what you want on a USB stick or somebody else's computer.

**A SECURITY.md** that says what Plain is exposed to, what it refuses to do, that it is not signed and why, and what would worry me.

**Startup was measured rather than assumed**: the window appears in 87 ms and the console twin answers in 122 ms, so nothing needed doing. Recorded so it can be watched.

1,020 checks, up from 1,014.

## 1.4.1, 2026-09-09

**A switch's value could end up in your cell.** `plain set book.xlsx B5 25000 --sheet Summary` wrote the text `25000 Summary` into B5: a number quietly became words, and a column of words does not add up. Every switch that takes a value was affected, on every verb that takes both a switch and words of its own, since switches with values were added in 1.3. The console twin now knows which switches are followed by a value and never reads one as something the verb was given.

Found by measuring what each kind of edit rewrites, not by a test, which is why there are now checks for it.

Nothing else changed. The window was never affected.

## 1.4.0, 2026-09-09

**Lines round cells.** Thin, medium, thick, dotted, dashed or double, on all four sides or one at a time.

**Taking out rows that say the same thing, and splitting a column.** The two jobs everyone does to a list and nobody enjoys. Both refuse where a formula would be made to mean something else, and splitting refuses rather than writing over the column beside it.

**What a cell reads, and what reads it.** The question behind most spreadsheet mistakes: a total looks wrong and you cannot see what it is adding up. Right-click a cell and Plain says, both ways.

**The choices a cell will accept.** A sheet someone else built often has cells that only take certain values. Plain does not draw the little arrow, but it reads the rule and says what it wants, which beats a cell that silently refuses what you type.

**Speaker notes, read and changed.** What the presenter was going to say travels with every copy of the deck and nobody sees it on the slide.

**The words that describe a picture.** The thing everyone is asked for and skips. Plain lists every picture, says which have nothing describing them, and writes the description where both Word and PowerPoint will read it.

**Links you can make, not only read.** Put a link on a paragraph, or take one off. Only ordinary web and mail addresses: Plain will not write into a document a link it would refuse to open.

**A picture put into a document.** Nothing is scaled or re-encoded, so what goes in is the file you chose.

**What is already in the column, offered as you type.** A list of clients typed slightly differently each time is the most common way a spreadsheet quietly goes wrong.

**Console twin.** `plain border`, `plain tidy`, `plain reads`, `plain choices`, `plain notes`, `plain describe`, `plain link`, `plain picture`.

1,014 checks, up from 939.

## 1.3.0, 2026-09-09

**Sheets can be added, renamed, moved and taken out**, which slides could already do and sheets could not. Renaming rewrites every formula in the workbook that named the sheet, and says how many it changed; taking one out is refused while any formula still reads it.

**Rows have a height**, as columns have a width, and **columns can be held still** as well as rows.

**Cells can be aligned, made to wrap, and given colour** behind them and for their words. Right-click a selection for "How these cells look". Each change builds on the format the cell already has, so setting a colour does not throw away its number format.

**See the pages before printing them**, laid out on the paper and margins chosen in settings, using the same laying out that printing does.

**Headers and footers.** The lines along the top and bottom of every page: what they say, and a way to change the words. They are the part of a file people forget is there.

**Where a paragraph sits.** Ctrl+L, Ctrl+E, Ctrl+R and Ctrl+J, and the document shows it.

**Links, and where they actually go.** A link's words and its destination are two different things. Plain lists both side by side, and will only open ordinary web and mail addresses: anything else is shown in red and left alone.

**Compare with another file.** What changed between two versions: which parts of the file differ at all, and which lines of text arrived or went. It compares what is there rather than guessing how one became the other.

**Console twin.** `plain sheet`, `plain height`, `plain align`, `plain colour`, `plain band`, `plain links`, `plain compare`.

939 checks, up from 843.

## 1.2.0, 2026-09-09

**Sort was only half of it: rows can now be filtered too.** Show only the rows whose cell in a column contains what you type. Nothing is written to the file; it is a way of looking at the sheet, hiding a row never deletes anything, and clearing it puts everything back.

**What the selection adds up to, in the status bar.** How many numbers, the sum, the average, the lowest and the highest. This is the question a spreadsheet is usually opened to answer, and answering it there means not typing a formula into an empty cell and deleting it afterwards.

**Redo.** Ctrl+Y and Ctrl+Shift+Z. Every step now carries how to undo it and, where it can be worked out, how to do it again. A step that cannot say makes everything after it unrepeatable, so the pile waiting to be redone is thrown away rather than left to put things back in the wrong order.

**Getting about a large sheet.** Ctrl with an arrow goes to the end of the run of filled cells, or across a gap to the next thing there is. Ctrl+Home and Ctrl+End go to the corners. Ctrl+G asks which cell and goes there.

**Before you send it.** One list of what travels with a file that you may not have meant to send: who wrote it, comments, tracked changes, hidden sheets and rows, speaker notes. It takes out only what you tick, it leaves alone the things that are somebody's working rather than an accident, and it never claims to have made a file safe, because it can only find what it knows to look for.

**Slides can be added, taken out and moved.** Moving one only reorders a list. Removing one takes it out of the running order and leaves its part in the file unused, because deleting a part something still points at is how a deck gets broken.

**Rows can be added to and taken out of a table in a document.** A new row is a copy of a neighbour with the words removed, so it keeps the borders, shading and widths of the table it joins.

**Joined cells are drawn as one cell**, rather than as several empty ones with lines through them, and clicking any part of a block selects the corner that holds the value.

**Find in a folder.** Which Office files in a folder hold the words, and where. It only ever reads: you open the ones that matter and change them yourself, one at a time.

**A file with a password on it says so.** It used to report "no ZIP end record found", which sends someone looking for damage that is not there. An older .doc, .xls or .ppt is now named as an older file too, with what to do about it.

**Under all that:** the package can add a part, which a deck needs in order to gain a slide. It is deliberately separate from writing an existing part so that adding one can never be an accident, and the checks prove that afterwards every part that was already there still comes back byte for byte.

**Console twin.** `plain hidden`, `plain find`, `plain slide`, `plain tablerow`.

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
