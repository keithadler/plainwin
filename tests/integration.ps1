# End to end through the console twin, on copies so a fixture is never written over.
# Run on Windows:  pwsh tests/integration.ps1 [path\to\plain.exe]
param([string]$Exe = "dist\win-arm64-cli\plain.exe")
$ErrorActionPreference = "Continue"
if (-not (Test-Path $Exe)) { $Exe = "dist\win-x64-cli\plain.exe" }
if (-not (Test-Path $Exe)) { Write-Host "no plain.exe; run scripts/publish.sh first"; exit 2 }

$fx = "tests\fixtures"
$root = Join-Path ([IO.Path]::GetTempPath()) ("plain-it-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force $root | Out-Null
Copy-Item "$fx\sheet.xlsx" "$root\book.xlsx"; Copy-Item "$fx\doc.docx" "$root\doc.docx"; Copy-Item "$fx\deck.pptx" "$root\deck.pptx"

$fail = 0
function Check($name, [scriptblock]$test) {
  $ok = $false
  try { $ok = & $test } catch { $ok = $false }
  if ($ok) { Write-Host "ok    $name" } else { Write-Host "FAIL  $name"; $script:fail = 1 }
}

Check "info names the workbook"        { (& $Exe info "$root\book.xlsx" | Out-String) -match "Excel workbook" }
Check "info names the document"        { (& $Exe info "$root\doc.docx" | Out-String) -match "Word document" }
Check "info names the deck"            { (& $Exe info "$root\deck.pptx" | Out-String) -match "PowerPoint deck" }
Check "a save changes nothing"         { & $Exe roundtrip "$root\book.xlsx" "$root\doc.docx" "$root\deck.pptx" *> $null; $LASTEXITCODE -eq 0 }
Check "parts are all accounted for"    { ((& $Exe parts "$root\book.xlsx") | Measure-Object).Count -eq 10 }
Check "parts as json"                  { (& $Exe parts "$root\book.xlsx" --json | ConvertFrom-Json).Count -eq 10 }
Check "a cell reads"                   { (& $Exe get "$root\book.xlsx" A1) -eq "Region" }
Check "a missing cell exits 1"         { & $Exe get "$root\book.xlsx" H40 *> $null; $LASTEXITCODE -eq 1 }
Check "a bad reference exits 64"       { & $Exe get "$root\book.xlsx" "not-a-cell" *> $null; $LASTEXITCODE -eq 64 }
Check "a missing file exits 2"         { & $Exe info "$root\nothing.xlsx" *> $null; $LASTEXITCODE -eq 2 }
Check "a file that is not Office exits 2" { "hello" | Set-Content "$root\x.xlsx"; & $Exe info "$root\x.xlsx" *> $null; $LASTEXITCODE -eq 2 }

Check "setting text saves"             { & $Exe set "$root\book.xlsx" A2 "Northern Europe" *> $null; (& $Exe get "$root\book.xlsx" A2) -eq "Northern Europe" }
Check "setting a number saves"         { & $Exe set "$root\book.xlsx" B2 500000 *> $null; (& $Exe get "$root\book.xlsx" B2) -match "500" }
Check "setting a formula saves"        { & $Exe set "$root\book.xlsx" F2 "=B2*2" *> $null; (& $Exe get "$root\book.xlsx" F2) -eq "=B2*2" }
Check "an untouched cell is unchanged" { (& $Exe get "$root\book.xlsx" A4) -eq "APAC" }
Check "the save reports what it rewrote" { (& $Exe set "$root\book.xlsx" A6 "note" | Out-String) -match "kept byte for byte" }
Check "the edited file still opens"    { & $Exe info "$root\book.xlsx" *> $null; $LASTEXITCODE -eq 0 }
Check "no formula keeps a stale value" {
  $cells = & $Exe cells "$root\book.xlsx" | Out-String
  # B5 is =SUM(B2:B4); after editing B2 it must come back as the formula, not the old total
  $cells -match "B5\s+=SUM"
}

Check "document text reads"            { (& $Exe text "$root\doc.docx" | Out-String) -match "Master Services Agreement" }
Check "document blocks are numbered"   { (& $Exe text "$root\doc.docx" --numbered | Out-String) -match "^\d+\t" }
Check "a document block edits" {
  $line = & $Exe text "$root\doc.docx" --numbered | Select-String "thirty"
  $n = $line.ToString().Split([char]9)[0]
  & $Exe set "$root\doc.docx" $n "The Client shall pay within forty-five (45) days." *> $null
  (& $Exe get "$root\doc.docx" $n) -eq "The Client shall pay within forty-five (45) days."
}
Check "the rest of the document is unchanged" { (& $Exe text "$root\doc.docx" | Out-String) -match "Discovery and audit" }

Check "deck text reads"                { (& $Exe text "$root\deck.pptx" | Out-String) -match "Where we are today" }
Check "deck slides are listed"         { (& $Exe info "$root\deck.pptx" | Out-String) -match "slide 2" }

Check "new makes a workbook"           { & $Exe new "$root\new.xlsx" *> $null; (Test-Path "$root\new.xlsx") -and $LASTEXITCODE -eq 0 }
Check "new makes a document"           { & $Exe new "$root\new.docx" *> $null; Test-Path "$root\new.docx" }
Check "new makes a deck"               { & $Exe new "$root\new.pptx" *> $null; Test-Path "$root\new.pptx" }
Check "a new file is the right kind"   { (& $Exe info "$root\new.xlsx" | Out-String) -match "Excel workbook" }
Check "a new file survives a save"     { & $Exe roundtrip "$root\new.xlsx" "$root\new.docx" "$root\new.pptx" *> $null; $LASTEXITCODE -eq 0 }
Check "a new workbook takes an edit"   { & $Exe set "$root\new.xlsx" A1 "typed" *> $null; (& $Exe get "$root\new.xlsx" A1) -eq "typed" }
Check "a new workbook takes a formula" { & $Exe set "$root\new.xlsx" A2 "=1+1" *> $null; (& $Exe get "$root\new.xlsx" A2) -eq "=1+1" }
Check "a new deck has a slide"         { (& $Exe info "$root\new.pptx" | Out-String) -match "slide 1" }
Check "new will not overwrite"         { & $Exe new "$root\new.xlsx" *> $null; $LASTEXITCODE -eq 2 }
Check "new refuses an unknown kind"    { & $Exe new "$root\new.txt" *> $null; $LASTEXITCODE -eq 2 }
Check "two new files are identical"    {
  & $Exe new "$root\a.xlsx" *> $null; & $Exe new "$root\b.xlsx" *> $null
  (Get-FileHash "$root\a.xlsx").Hash -eq (Get-FileHash "$root\b.xlsx").Hash
}
Check "selftest passes"                { (& $Exe selftest | Select-Object -Last 1) -match "0 failed" }

# ---------- sort, widths, freezing and the right-click menu ----------

Check "sort puts the rows in order" {
  # A fresh copy: earlier checks have written formulas into book.xlsx, and sorting those is refused on purpose.
  Copy-Item "$fx\sheet.xlsx" "$root\s.xlsx" -Force
  & $Exe sort "$root\s.xlsx" A2:B4 A | Out-Null
  (& $Exe get "$root\s.xlsx" A2) -le (& $Exe get "$root\s.xlsx" A3)
}
Check "sort refuses to move a formula" {
  Copy-Item "$fx\sheet.xlsx" "$root\f.xlsx" -Force
  & $Exe set "$root\f.xlsx" C2 "=A2" | Out-Null
  $said = (& $Exe sort "$root\f.xlsx" A2:C4 A 2>&1 | Out-String)
  $said -match "formula"
}
Check "a width can be set and read back" {
  & $Exe width "$root\book.xlsx" B 33 | Out-Null
  (& $Exe width "$root\book.xlsx" B 33 | Out-String) -match "33"
}
Check "a column can be fitted to its contents" { (& $Exe width "$root\book.xlsx" A fit | Out-String) -match "characters wide" }
Check "rows can be frozen"                     { (& $Exe freeze "$root\book.xlsx" 1 | Out-String) -match "stay on screen" }
Check "and unfrozen"                           { (& $Exe freeze "$root\book.xlsx" 0 | Out-String) -match "scrolls freely" }
Check "the file still round trips after all that" { (& $Exe roundtrip "$root\book.xlsx" | Out-String) -match "identical" }

Check "a pdf can be given paper and a footer" {
  & $Exe pdf "$root\doc.docx" "$root\p.pdf" --paper Letter --landscape --footer "Page {page} of {pages}" | Out-Null
  $bytes = [IO.File]::ReadAllBytes("$root\p.pdf")
  $text = [Text.Encoding]::ASCII.GetString($bytes)
  $text -match "MediaBox \[0 0 792" -and $text -match "Page 1 of"
}

# The right-click line has to point at the window, so the console twin looks for it beside itself. Where only the
# twin has been published, as on a build machine, there is nothing to point at and Plain must say so rather than
# registering a command that opens nothing. Both are worth checking, so this runs whichever applies.
$window = Join-Path (Split-Path -Parent (Resolve-Path $Exe)) "Plain for Windows.exe"
if (Test-Path $window) {
  Check "the right-click menu goes on and comes off cleanly" {
    & $Exe explorer on | Out-Null
    $on = (& $Exe explorer status | Out-String) -match "^on"
    & $Exe explorer off | Out-Null
    $off = (& $Exe explorer status | Out-String) -match "^off"
    $left = Test-Path "HKCU:\Software\Classes\Plain.Document"
    $on -and $off -and (-not $left)
  }
} else {
  Check "with no window beside it, the right-click menu refuses and says why" {
    $said = (& $Exe explorer on | Out-String)
    $said -match "Plain for Windows.exe"
  }
  Check "and it does not claim to be on" {
    (& $Exe explorer status | Out-String) -match "^off"
  }
  Check "and it leaves nothing behind in the registry" {
    -not (Test-Path "HKCU:\Software\Classes\Plain.Document")
  }
}

# A switch that takes a value must not be read as one of the words the verb was given. This went wrong once:
# "set book.xlsx B5 25000 --sheet Summary" wrote the text "25000 Summary" into the cell, turning a number into
# words, which is exactly the kind of quiet damage this program exists to prevent.
# get reports the cell as it is shown, so 25000 comes back as 25,000 where the cell has a thousands separator.
Copy-Item "$fx\sheet.xlsx" "$root\flags.xlsx" -Force
$sheetName = ((& $Exe info "$root\flags.xlsx" | Select-String "^  sheet") -replace "^\s*sheet\s+", "" -replace ",.*$", "").Trim()
& $Exe set "$root\flags.xlsx" B2 25000 --sheet $sheetName | Out-Null
Check "a value set with a --sheet switch is still a number" {
  ((& $Exe get "$root\flags.xlsx" B2) -replace "[,\s]", "") -eq "25000"
}
Check "and the switch's value is nowhere in the cell" {
  -not ((& $Exe get "$root\flags.xlsx" B2) -match [regex]::Escape($sheetName))
}
& $Exe colour "$root\flags.xlsx" A1 --fill FFE7A1 --sheet $sheetName | Out-Null
Check "two switches with values still leave the cell alone" {
  ((& $Exe get "$root\flags.xlsx" B2) -replace "[,\s]", "") -eq "25000"
}
Check "and the file still opens afterwards" {
  (& $Exe info "$root\flags.xlsx" | Out-String) -match "Excel workbook"
}

Remove-Item -Recurse -Force $root -ErrorAction SilentlyContinue

if ($fail -eq 0) { Write-Host "integration: all passed" } else { Write-Host "integration: FAILURES" }
exit $fail
