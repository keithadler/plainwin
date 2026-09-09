#!/usr/bin/env bash
# The README makes checkable claims: which files to download, and how many checks there are. Both went stale
# without anyone noticing. The console twin's download link said 1.0.0 for four releases, so anyone who clicked it
# got a file that does not exist.
#
# So the claims are checked here, the same way the program's own claims are. Runs anywhere; no Windows needed.
set -u
cd "$(dirname "$0")/.."
fail=0

check() {
  if [ "$2" = "1" ]; then echo "ok    $1"; else echo "FAIL  $1${3:+ - $3}"; fail=1; fi
}

version=$(grep -o '<Version>[^<]*</Version>' src/Plain/Plain.csproj | head -1 | sed 's/<[^>]*>//g')
check "the csproj has a version" "$([ -n "$version" ] && echo 1 || echo 0)"

cli=$(grep -o 'public const string Version = "[^"]*"' src/Plain/Cli.cs | sed 's/.*"\(.*\)"/\1/')
check "the console twin agrees with the csproj" "$([ "$cli" = "$version" ] && echo 1 || echo 0)" "csproj $version, Cli.cs $cli"

# Every download link in the README has to name the version being released.
stale=$(grep -oE '(Plain-for-Windows|plain)-[0-9]+\.[0-9]+\.[0-9]+-(x64|arm64)\.exe' README.md \
        | grep -v -- "-$version-" | sort -u)
check "every download link in the README names $version" "$([ -z "$stale" ] && echo 1 || echo 0)" "$(echo $stale)"

for doc in docs/release-$version.md CHANGELOG.md SECURITY.md PRIVACY.md; do
  check "$doc exists" "$([ -f "$doc" ] && echo 1 || echo 0)"
done

# The release notes name their own files, and those have to be right too.
if [ -f "docs/release-$version.md" ]; then
  stale=$(grep -oE '(Plain-for-Windows|plain)-[0-9]+\.[0-9]+\.[0-9]+-(x64|arm64)\.exe' "docs/release-$version.md" \
          | grep -v -- "-$version-" | sort -u)
  check "every link in the release notes names $version" "$([ -z "$stale" ] && echo 1 || echo 0)" "$(echo $stale)"
fi

# The count the README quotes has to be the count the checks actually produce.
dot=""
[ -x "$HOME/.dotnet9/dotnet" ] && dot="$HOME/.dotnet9/dotnet"
[ -z "$dot" ] && command -v dotnet >/dev/null 2>&1 && dot="$(command -v dotnet)"
if [ -n "$dot" ]; then
  actual=$("$dot" run --project src/Plain.Selftest -c Release 2>&1 | grep -oE '[0-9]+ passed' | tail -1 | grep -oE '[0-9]+')
  # The count is on the same line as its label, so read that line and nothing after it.
  claimed=$(grep 'Checks in the engine' README.md | grep -oE '[0-9][0-9,]*' | tail -1 | tr -d ',')
  check "the README's count of engine checks is right" \
        "$([ -n "$actual" ] && [ "$actual" = "$claimed" ] && echo 1 || echo 0)" \
        "README says ${claimed:-nothing}, there are ${actual:-unknown}"
fi

# Every verb the program answers to should be findable in the README, so nothing ships undocumented.
verbs=$(sed -n '/private static readonly HashSet<string> Verbs/,/};/p' src/Plain/Cli.cs \
        | grep -oE '"[a-z-]+"' | tr -d '"' | grep -vE '^(help|--help|-h|--version)$')
missing=""
for v in $verbs; do
  grep -q "plain $v" README.md || missing="$missing $v"
done
check "every verb is in the README" "$([ -z "$missing" ] && echo 1 || echo 0)" "missing:$missing"

[ $fail -eq 0 ] && echo "docs: all passed" || echo "docs: FAILURES"
exit $fail
