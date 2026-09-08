#!/bin/zsh
# Publish Plain for Windows as single self-contained exes for ARM64 and x64 into dist/.
# Works on macOS (EnableWindowsTargeting) and on Windows. Usage: scripts/publish.sh [arm64|x64|all]
set -euo pipefail
cd "$(dirname "$0")/.."
DOTNET=${DOTNET:-$HOME/.dotnet9/dotnet}
VER=$(sed -n 's|.*<Version>\(.*\)</Version>.*|\1|p' src/Plain/Plain.csproj)
want=${1:-all}
for arch in arm64 x64; do
  [[ "$want" == "all" || "$want" == "$arch" ]] || continue
  rid=win-$arch
  "$DOTNET" publish src/Plain/Plain.csproj -c Release -r $rid --nologo -v q -o dist/$rid
  "$DOTNET" publish src/Plain/Plain.csproj -c Release -r $rid --nologo -v q -o dist/$rid-cli -p:CliBuild=true
  cp "dist/$rid/Plain for Windows.exe" "dist/Plain-for-Windows-$VER-$arch.exe"
  cp dist/$rid-cli/plain.exe "dist/plain-$VER-$arch.exe"
  ls -lh "dist/Plain-for-Windows-$VER-$arch.exe" "dist/plain-$VER-$arch.exe"
done
