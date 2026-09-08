#!/bin/zsh
# Copy the ARM64 exes into the Windows VM and run the checks there.
set -euo pipefail
cd "$(dirname "$0")/.."
VM=~/Downloads/win11arm/vmssh.sh
VER=$(sed -n 's|.*<Version>\(.*\)</Version>.*|\1|p' src/Plain/Plain.csproj)
$VM ssh 'Stop-Process -Name "Plain for Windows",plain -Force -ErrorAction SilentlyContinue; Start-Sleep 3; New-Item -ItemType Directory -Force C:\Plain | Out-Null'
$VM scp "dist/Plain-for-Windows-$VER-arm64.exe" 'keith@VM:C:/Plain/Plain for Windows.exe'
$VM scp "dist/plain-$VER-arm64.exe" 'keith@VM:C:/Plain/plain.exe'
# Copy fixtures one by one: scp of a directory onto an existing one nests it inside itself.
$VM ssh 'Remove-Item C:\Plain\fixtures -Recurse -Force -ErrorAction SilentlyContinue; New-Item -ItemType Directory -Force C:\Plain\fixtures | Out-Null'
for f in tests/fixtures/*; do $VM scp "$f" "keith@VM:C:/Plain/fixtures/$(basename $f)"; done
$VM scp tests/integration.ps1 'keith@VM:C:/Plain/integration.ps1'
$VM ssh '$env:PLAIN_FIXTURES="C:\Plain\fixtures"; C:\Plain\plain.exe selftest | Select-Object -Last 1; C:\Plain\plain.exe version'
