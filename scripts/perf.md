# Measuring

`tests/fixtures/large.xlsx` (5,000 rows) is committed and the `scale` suite guards against the code going quadratic.

To measure on something bigger, generate one with LibreOffice and point the harness at it:

```bash
python3 scripts/make-big-sheet.py 20000 > /tmp/big.fods
soffice --headless --convert-to xlsx --outdir /tmp /tmp/big.fods
PLAIN_BIG=/tmp/big.xlsx dotnet run --project scripts/perf
```

The numbers in the README came from 20,000 rows with a formula in every one.
