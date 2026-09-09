**[Download Plain-for-Windows-1.4.4-x64.exe](https://github.com/keithadler/plainwin/releases/download/v1.4.4/Plain-for-Windows-1.4.4-x64.exe)** — Windows 10 or 11. One exe, no installer, no administrator. ARM64 and the console twin are below.

---

A document can gain and lose a paragraph.

The engine learned how in 1.4.3 and nothing could ask it to, so a document — a new one especially, which starts with exactly one paragraph — stayed the length it was. The More menu now offers **Add a paragraph after this one** and **Take this paragraph out** while a document is open, acting on the paragraph the caret is in, and the console twin has the same:

```
plain paragraph report.docx add 3
plain paragraph report.docx remove 4
```

A paragraph in a table is refused, with the reason: add a row to the table instead.

There is also a new test that opens the real window, opens the real menu with real keystrokes, clicks the item and looks at the file afterwards. What a menu offers is the one thing a unit test cannot see, and that is exactly the gap this release closes.

**Plain for Mac is out today too.** [Plain for Mac 1.0.0](https://github.com/keithadler/plainmac) is the same program with a Mac window, running the same engine compiled native, so the two agree about what a file is by construction rather than by intention.
