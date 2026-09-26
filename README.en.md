# Pseudonymise-Before-Upload (NCCU ORD)

*[中文說明](README.md)*

Replaces personal names in Excel rosters with stable codes, so the file can be
handed to an external AI service. **The lookup table never leaves your computer.**

![version](https://img.shields.io/badge/version-v2.9-0f766e)
![licence](https://img.shields.io/badge/license-MIT-blue)
![platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![python](https://img.shields.io/badge/python-3.9%2B-3776ab)

---

## What makes this different

**Codes are consistent across files and across years.**

A code is `SHA-256(local salt + name)`, so the same person gets the same code in
every file you ever process on that machine — this year, next year, whether the
files are processed together or one at a time. Joins and longitudinal tracking
survive pseudonymisation intact.

Most comparable tools mint a fresh random code per document, which breaks exactly
the analysis this tool exists to protect. See
[How this compares](#how-this-compares).

Two more things follow from that design:

- **Name reconciliation.** When the same person is typed two ways
  (`Wei Xiaobao` vs `WeiXiaobao`, a stray space in a Chinese name), the tool
  flags it and lets the operator declare them the same person, merging them onto
  one code.
- **Projects.** One folder per line of business, each with its own salt, so
  unrelated datasets cannot be joined against each other.

---

## Intended user

A university administrator who needs to run an analysis, does not write code, and
should not be asked to open a command prompt. Install is a double-click; the
workflow is four steps in a window.

That focus also sets the limits: **the built-in detection is tuned for Traditional
Chinese rosters** (Taiwanese column headings, 2–3 character Chinese names).

For an English or mixed-language workbook, copy
[`field_settings_example_EN.txt`](field_settings_example_EN.txt) to `欄位設定.txt`
(that Chinese filename is required) and list your own column headings. Matching
ignores case and spacing, and your entries are *added* to the built-in list, so a
workbook with both Chinese and English headings works without extra setup.

What this does **not** fix: the "a name appears inside a sentence" heuristic is
built around 2–3 character Chinese names and does not apply to Western names. On
English rosters, rely on the column lists and always use *① Preview* first.

---

## Install

**Windows, no Python required knowledge:** download the zip from
[Releases](../../releases), move the folder to a local drive (not OneDrive,
Dropbox or any sync folder — the tool checks and warns), then double-click
`1_第一次執行_安裝環境.bat`. It installs Python and dependencies and creates a
desktop shortcut.

**From source:**

```bash
git clone https://github.com/<account>/<repo>.git
cd <repo>
pip install -r requirements.txt

python anonymize_gui.py                      # windowed interface
python anonymize.py                          # CLI, processes the current project
python anonymize.py --dry-run                # preview only, writes nothing
python anonymize.py --decode T-XXXXXXXXXX    # look up a code
```

`tkinterdnd2` is optional; it only enables drag-and-drop onto the window.

---

## How it works

```
Step 0   Pick or create a project     one folder per line of business
Step 1   Put files in input/          drag, browse, or drop on the window
Step 2   Preview, then run            confirm the detected columns first
Step 3   Review                       opens output/ and _private/ side by side
```

```
tool/
└── 專案/  (projects)
    └── <your project>/
        ├── input/      source files — never upload these
        ├── output/     pseudonymised — these are safe to upload
        └── _private/   salt.txt, mapping.csv — never let these leave
```

---

## Three rules that matter

1. **Only `output/` may be uploaded.** `input/` is the original; `_private/` is
   the lookup table.
2. **`_private/salt.txt` must stay the same forever.** With the salt unchanged
   you can re-run, process in batches, and add files later — codes stay stable.
   Delete `_private/` and every code is reshuffled; this year stops matching last
   year.
3. **A project is a *line of business*, not a *year*.** Reuse the same project
   every year. A new project means a new salt.

---

## What it does not do

Stated up front, because knowing the edges matters more than the feature list.

- Only the first header row of a worksheet is recognised when one sheet holds two
  stacked tables
- Three-character Chinese names are replaced even inside running prose, which can
  mangle sentences that happen to contain one
- Two-character names are not auto-replaced inside prose; the tool flags them
- Vertical forms and complex merged cells are flagged, not handled

These are deliberate trade-offs: replace rather than blank (replacement is
reversible via `mapping.csv`; blanking is not), and over-warn rather than fail
silently.

---

## Verification

After v2.4 shipped, the tool was put through an independent vulnerability and
security review, producing v2.6 and v2.7:

| Metric (13,000 randomly generated sheets) | v2.4 | v2.6 |
|---|---:|---:|
| Residual real names | 409,897 | 15,176 (−96.3%) |
| Name columns correctly coded | 31.0% | 93.4% |
| Sheets reported "done" that still had residue | 11,262 | 36 (−99.7%) |
| Sheets worse than v2.4 | — | **0** |
| Regression suite | — | 669 / 673 |

The 4 failures are known limitations, deliberately left failing so that a real
regression is distinguishable from a documented gap.

Full record (17 findings, 126 changes, each with file:line and actual output) is
in [`docs/`](docs/).

---

## How this compares

| Tool | Form | Codes consistent across files? | Best for |
|---|---|---|---|
| **This tool** | Windows desktop, one-click install | **Yes**, and across years | Administrators running recurring analyses |
| [data-deidentification](https://github.com/dean9703111/data-deidentification) | Browser, client-side | No — random, scoped to one document | PDF / Word / multi-format work |
| [Microsoft Presidio](https://github.com/microsoft/presidio) | Python library | Implement it yourself | Developers building a pipeline |
| [ARX](https://arx.deidentifier.org/) | Java desktop | N/A — irreversible anonymisation | Researchers doing k-anonymity work |

**Use something else if:** you need PDF or Word support (use
data-deidentification — far broader format coverage and a human-review preview);
you need to call it from code (use Presidio); you need irreversible anonymisation
with disclosure-risk scoring (use ARX); or your rosters are not in Traditional
Chinese.

**Use this if:** you repeat the same analysis every year, you need several Excel
rosters joined on names, next year's results must line up with this year's, and
the person doing the work does not write code.

---

## Demo data

`專案/示範專案/` contains four **fictional** files; the names are wuxia novel
characters. A run should report:

```
處理檔案　　：4       (files processed)
替換儲存格　：69      (cells replaced)
本次新增姓名：17      (names coded)
```

An orange warning about a suspected duplicate person appears on purpose, to
demonstrate name reconciliation. **This repository contains no real personal
data.**

---

## Contributing

- General issues and suggestions: open an [Issue](../../issues)
- Security issues: read [SECURITY.md](SECURITY.md) first — **do not open a public
  issue**
- Before a PR, run the demo project and confirm it still reports 4 / 69 / 17

Before sending anything, check that your own `_private/` and `output/` are not
included. `.gitignore` blocks them, but check anyway.

The interface and documentation are in Traditional Chinese, matching the intended
users. English documentation is limited to this file.

---

## Licence

MIT · Copyright (c) 2026 Davy HSU

Developed at the Office of Research and Development, National Chengchi
University, Taiwan.
