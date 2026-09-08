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

Check "selftest passes"                { (& $Exe selftest | Select-Object -Last 1) -match "0 failed" }

Remove-Item -Recurse -Force $root -ErrorAction SilentlyContinue
if ($fail -eq 0) { Write-Host "integration: all passed" } else { Write-Host "integration: FAILURES" }
exit $fail
