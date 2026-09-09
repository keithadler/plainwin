# Security

## What Plain is exposed to

Plain opens files other people send you. That is the whole point of it, and it is also the entire attack surface:
a `.docx`, `.xlsx` or `.pptx` is a ZIP full of XML, and both of those are formats where a malformed file has
historically been a way into a program.

So the things worth knowing:

- **Plain never runs anything in a file.** There is no macro engine. A macro survives a round trip because its
  bytes are copied through untouched, and it is never executed, inspected or understood.
- **XML is parsed with entity resolution off**, so a file cannot make Plain fetch a URL or read a file off your
  disk through a crafted document type. There is no `XmlResolver` anywhere in it.
- **A part is read into memory with a cap** (512 MB), and a package that claims a part is larger is refused rather
  than allowed to exhaust memory.
- **Nothing is written outside the file you saved**, except the settings folder, the copies of unsaved work, and a
  temporary file beside your document while saving.
- **Links are checked before they are opened or written.** Only ordinary web and mail addresses; anything that
  could start a program is refused, in both directions.
- **The only network request Plain makes** is a daily check with GitHub for a newer version, which sends no
  identifier and nothing about your files, and which can be turned off. See [PRIVACY.md](PRIVACY.md).

The checks include a fuzzing suite that feeds deliberately damaged packages to every reader and requires a clear
refusal rather than a crash. CI runs it with several thousand cases on every push.

## What Plain is not

**It is not signed.** A code signing certificate costs money this project does not make. Windows will warn you
about an unknown publisher, and you have to choose **More info** and then **Run anyway**. Anyone telling you a
signature makes software safe is selling something, but the honest position is that you are trusting the download
rather than a certificate authority.

What you can check instead: from version 1.4.2, the released executables are built by a GitHub Actions job on a
clean runner, from a commit you can read, and a `SHA256SUMS.txt` is published beside them by the same job. Compare
what you downloaded:

```powershell
Get-FileHash "Plain-for-Windows-1.4.2-x64.exe" -Algorithm SHA256
```

That tells you the file you have is the file that job produced. It does not tell you the job produced something
trustworthy; for that, the source is right here and it is short enough to read.

## Reporting something

Open a private security advisory on the repository, or email keith.adler@icloud.com.

Please include the file that causes it if you can share one, and what you expected instead. A crash on a malformed
file is worth reporting even though Plain is not a service and has no users to protect but yourself: refusing
damaged input cleanly is a promise this program makes.

There is no bounty. This is one person's free program, and the honest answer about response time is "when I see
it", though anything that damages a file people are editing goes to the front of the queue.

## What would worry me

If you find any of these, they are serious rather than cosmetic:

- A file that, opened and saved without editing, comes back different. That breaks the promise the whole program
  rests on, and there are checks that are supposed to make it impossible.
- Any path where Plain writes outside the file you asked it to save.
- Anything that causes Plain to execute content from a document, or to reach the network because of one.
