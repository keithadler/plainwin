**[Download Plain-for-Windows-1.4.3-x64.exe](https://github.com/keithadler/plainwin/releases/download/v1.4.3/Plain-for-Windows-1.4.3-x64.exe)** for ordinary Intel and AMD PCs, or **[Plain-for-Windows-1.4.3-arm64.exe](https://github.com/keithadler/plainwin/releases/download/v1.4.3/Plain-for-Windows-1.4.3-arm64.exe)** for Windows on ARM.

Windows 10 or 11. One exe, no installer, no runtime, no administrator. Put it anywhere and double-click. Windows will warn about an unknown publisher: **More info**, then **Run anyway**.

`SHA256SUMS.txt` is published beside them, written by the same job that built them.

## Headings written by LibreOffice were shown as bullets

A document exported from LibreOffice, or from anything else that writes a heading the same way, lost its entire outline in Plain: every heading appeared as a list item, and none was styled as a heading.

Two things were read wrongly.

A paragraph can carry numbering whose id is `0`, and that means "no numbering at all". Treating any numbering element as a list turned those paragraphs into list items.

And a heading can declare its level outright with `outlineLvl` rather than by being styled "Heading 1". That is what LibreOffice writes, and reading only the style name missed every one of them.

Both are fixed, with four checks built on a document written the way LibreOffice writes one.

## How it was found

By building Plain for Mac. It shows the same blocks the same way, because it is the same engine, and a document of headings and bullets came out as a document of nothing but bullets. A bug found on one is fixed for both.

Nothing else changed. 1,024 checks.
