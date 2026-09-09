# Contributing

The rule this project is built on: **never damage a part of a file you do not understand.** Anything that risks that is not worth the feature it buys.

## Before a change

```
dotnet run --project src/Plain.Selftest
```

Everything must be green. Point it at your own documents to make the check mean more:

```
PLAIN_CORPUS=/path/to/real/office/files dotnet run --project src/Plain.Selftest
```

Every one of them must round trip byte for byte. If one does not, that is the bug, whatever else you were working on.

## What a good change looks like

- **The engine first.** `Plain.Core` has no dependency and no Windows in it, so it can be tested anywhere. Put logic there and let the window be thin.
- **A test that would have failed before.** The suites live beside the code in `src/Plain.Core/Tests`. Name a check as a sentence about behaviour, not about implementation.
- **Refuse rather than guess.** The formula evaluator computes what it fully understands and returns nothing for the rest. A wrong number is worse than no number, and this applies well beyond formulas.
- **Say the limit out loud.** If a change cannot do something people will expect, the README's honest limits section is where that goes, in the same commit.

## Testing the window

`tests/ui.ps1` drives the running app with real keystrokes and checks the files on disk afterwards. It must run in an interactive Windows session. Three real bugs got past the engine tests, the console tests and the screenshots, and only this caught them; if you change how editing or saving works, run it.

## Fixtures

`tests/fixtures` holds small files made with LibreOffice, and they are committed with `-text` so no line-ending conversion can touch them. Never commit a real document from anybody's work.

## Style

British spelling. Comments explain why, not what. If a comment restates the line below it, delete the comment.
