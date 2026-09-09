**[Download Plain-for-Windows-1.4.1-x64.exe](https://github.com/keithadler/plainwin/releases/download/v1.4.1/Plain-for-Windows-1.4.1-x64.exe)** for ordinary Intel and AMD PCs, or **[Plain-for-Windows-1.4.1-arm64.exe](https://github.com/keithadler/plainwin/releases/download/v1.4.1/Plain-for-Windows-1.4.1-arm64.exe)** for Windows on ARM.

## A switch's value could end up in your cell

In the console twin, `plain set book.xlsx B5 25000 --sheet Summary` wrote the text `25000 Summary` into B5. The number became words, and a column of words does not add up.

Every switch that takes a value was affected, on every verb that takes both a switch and words of its own, from 1.3 onwards. The console twin now knows which switches are followed by a value, and never reads one as something the verb was given. There are checks for it now.

**The window was never affected**, and neither was any file already saved correctly. If you have used `plain set`, `plain colour`, `plain align`, `plain border`, `plain tidy` or `plain height` with a switch, look at the cells you changed: a number that came out as text will be sitting on the left of its cell rather than the right.

This was found by measuring what each kind of edit rewrites in a file, rather than by a test. Nothing else changed.
