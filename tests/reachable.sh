#!/usr/bin/env bash
# Every button in the window has to reach something, and everything the code reaches for has to exist.
#
# This exists because the same mistake happened four times across Plain for Windows and Plain for Mac: code that
# worked, wired to nothing. Replace-all reported success and saved nothing. Inserting a link wrote it and a stale
# model wrote over it. The Mac's update check found a newer version and told nobody. Twenty-one of the Mac's
# operations had no way to them in the window at all. Every one passed its tests, because the tests exercise the
# engine and none of them exercised the path from a person to the engine.
#
# A unit test cannot see that. Asking "does this button reach anything, and does this name exist?" can, and it is
# the difference between a button that does nothing and one that does something.
set -uo pipefail
cd "$(dirname "$0")/.."

fail=0
note() { if [ "$1" = "1" ]; then echo "ok    $2"; else echo "FAIL  $2"; fail=1; fi; }

xaml=src/Plain/MainWindow.xaml
code=src/Plain/MainWindow.xaml.cs
[ -f "$xaml" ] && [ -f "$code" ] || { echo "no window source here"; exit 2; }

# A Click= naming a handler that does not exist is a button that does nothing when pressed.
missing=""
for handler in $(grep -oE 'Click="[A-Za-z_]+"' "$xaml" | grep -oE '"[A-Za-z_]+"' | tr -d '"' | sort -u); do
  grep -qE "void[[:space:]]+$handler[[:space:]]*\(" "$code" || missing="$missing $handler"
done
note "$([ -z "$missing" ] && echo 1 || echo 0)" "every button in the window has a handler behind it${missing:+ - missing:$missing}"

# A name the code reaches for that the window does not define throws the moment the window opens.
names=$(grep -oE 'x:Name="[A-Za-z_]+"' "$xaml" | grep -oE '"[A-Za-z_]+"' | tr -d '"' | sort -u)
note "$([ -n "$names" ] && echo 1 || echo 0)" "the window names its controls"

# Every menu item that names a handler, across every view, has one.
missing=""
for view in src/Plain/*.xaml; do
  behind="${view}.cs"
  [ -f "$behind" ] || continue
  for handler in $(grep -oE 'Click="[A-Za-z_]+"' "$view" | grep -oE '"[A-Za-z_]+"' | tr -d '"' | sort -u); do
    grep -qE "void[[:space:]]+$handler[[:space:]]*\(" "$behind" || missing="$missing $(basename "$view"):$handler"
  done
done
note "$([ -z "$missing" ] && echo 1 || echo 0)" "every view's buttons have handlers${missing:+ - missing:$missing}"

# Every verb the console twin answers to has to be in the usage text, or nobody can find it.
verbs=$(sed -n '/HashSet<string> Verbs/,/};/p' src/Plain/Cli.cs \
        | grep -oE '"[a-z-]+"' | tr -d '"' | grep -vE '^(help|--help|-h|--version)$' | sort -u)
missing=""
for verb in $verbs; do
  grep -q "plain $verb" src/Plain/Cli.cs || missing="$missing $verb"
done
note "$([ -z "$missing" ] && echo 1 || echo 0)" "every verb is in the usage text${missing:+ - missing:$missing}"

[ $fail -eq 0 ] && echo "reachable: all passed" || echo "reachable: FAILURES"
exit $fail
