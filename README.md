# spreadsheet-review

a desktop tool that turns a catalog-transfer spreadsheet into a clean, deduplicated claimant table. drop a report on it and the table is on your clipboard and saved as csv.

## the problem

a transfer report can repeat the same claimant across many catalog rows. for each report, someone has to decide whether the transfer is full or partial, then copy out only the claimants on rows marked yes, without duplicates. doing that by hand is slow and error-prone.

i'm on a three-person team that processes publishing catalog transfers. i built this for my own daily work; i estimate it took me from about 5 to about 20 transfers a day when needed (my own estimate, not a measured benchmark).

## what it does

- reads the first worksheet of an `.xlsx` report.
- classifies the whole report from the answer column: all yes → **true full**, any no → **true partial**, blanks ignored, anything unrecognized → **needs review** (with row numbers).
- accepts `YES`, `Yes`, `Y`, and `(Y)`, ignoring case and spaces.
- extracts four columns from yes rows (**admin, admin p#, publisher, publisher p#**) and keeps the first copy of each distinct row.
- saves a csv next to the report, *then* copies a tab-separated table to the clipboard, so a failed save never overwrites your clipboard.

## run it

**windows** (desktop excel required): extract the download and double-click `windows/Install Claimant Drop.cmd`. then right-click a report → send to → claimant drop. on windows 11, choose show more options first. a popup shows the classification and row counts.

**macos:**

```sh
swiftc src/claimants.swift -o claimants
./claimants /path/to/report.xlsx
```

the macos version reads the workbook xml directly with foundationxml (no excel needed). the windows version drives excel through powershell com automation.

## different report layouts

column positions are configurable. copy [`claimant-columns.example.json`](claimant-columns.example.json) to `claimant-columns.json` next to the tool (or point `CLAIMANTS_COLUMNS` at it) and change the letters and headers:

```json
{
  "answerColumn": "E",
  "fields": [
    { "header": "Admin", "column": "R" },
    { "header": "Admin P#", "column": "Q" }
  ]
}
```

without a settings file, the defaults above (answers in e; r, q, n, m extracted) are used.

## tests

the tests build invented workbooks and check filtering, classification, deduplication, csv escaping, column settings, and error cases. no real report data is in this repository.

```sh
# macos (also runs in github actions)
swiftc src/claimants.swift -o /tmp/claimants
CLAIMANTS_BINARY=/tmp/claimants python3 -m unittest discover -s tests -v
```

```bat
rem windows, with desktop excel
set CLAIMANTS_BINARY=%CD%\tests\run_windows.cmd
python -m unittest discover -s tests -v
```

## limits

- the windows popup and send to shortcut are checked by hand, not by the automated tests.
- formulas use the values excel last saved, with no recalculation.
- not yet hardened for very large files, untrusted workbooks, or spreadsheet formula injection in the csv output.

## how it was built

i defined the workflow, rules, and output format, and confirmed the rules with the workflow owner. the code was written with ai assistance (codex), and i iterated on the behavior.

© jojo. all rights reserved; see [license](LICENSE).
