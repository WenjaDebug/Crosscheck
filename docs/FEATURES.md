# Crosscheck: everything in detail

The short version with a picture is in the [README](../README.md).

Gegenprobe ("cross-check") recalculates your Excel workbook itself and tells you what does not add up. It works without AI, without internet and without any key, and the data stays on your computer. That makes it a fit for sensitive data from hospitals, labs or accounting. On a Mac it sits as a side pane right inside Excel.



It reads the name of a figure ("Umsatz Ost", "Durchschnitt Umsatz Nord") or the headers of a cross table and works out which rows belong to it. When a formula is wrong, it names the part of the formula that causes it:
- a condition too many or a wrong value;
- a range that stops one row early;
- a cell counted twice or a wrong divisor;
- a condition taken from the wrong header cell;
- a number or date stored as text.

It also gives the corrected formula, recalculated beforehand. It also checks the whole calculation chain, ranges that slide when copied down, rates and averages, financial models, data used for analyses, and differences to an earlier version of the workbook.

*Deutsch: [FUNKTIONEN.md](FUNKTIONEN.md)*

> The side panel in Excel speaks German or English (it follows Excel, switchable under **i**). The findings themselves, the start page and the launch windows are still in German.

The project used to be called VoiceAgent and started as a voice assistant. Nothing of that is left; only the launch files are still called **VoiceAgent.command**, **VoiceAgent.bat**, **ExcelAddin.command** and **ExcelAddin.bat**, and the Python package is `voiceagent`.

## Installation (macOS)

**Easiest:** as in the README [in 5 steps](../README.md#install-in-5-steps). **ExcelAddin.command** sets everything up the first time and then the side panel in Excel; **VoiceAgent.command** sets everything up the first time and starts Gegenprobe in the browser. Both first look for a Python 3.10 or newer that is already there (from python.org, from Homebrew or otherwise on the PATH). They only ask macOS's own Python (`/usr/bin/python3`, usually 3.9) when the Command Line Tools are installed, so macOS does not offer to install them. Without a suitable Python they install Python 3.13 with Homebrew if it is there, otherwise they open python.org and say what to do. Then they create a private environment in the `.venv` folder and download the three packages (FastAPI, Uvicorn, pypdf). After that it just starts. Without internet on the first start they say so. The GitHub ZIP keeps the execute permission of the .command files (checked with `git archive`, which is how GitHub packs it). Gegenprobe does not get around macOS's block on files from the internet: you confirm it yourself once per file (see step 3).

**In Terminal:** Gegenprobe needs Python 3.10 or newer, 3.13 is recommended.

```bash
brew install python@3.13
git clone https://github.com/WenjaDebug/Crosscheck.git Gegenprobe
cd Gegenprobe
python3.13 -m venv .venv
source .venv/bin/activate
pip install -e .
```

On Linux use `sudo apt install python3-venv` instead of `brew`.

## Installation (Windows)

Windows 10 or 11, as in the README [in 5 steps](../README.md#install-in-5-steps): download the ZIP, unpack it with **Extract All…** and double-click **ExcelAddin.bat** (side panel in Excel) or **VoiceAgent.bat** (in the browser) in Explorer. The first start looks for a Python 3.10 or newer that is already there. A `python.exe` under `WindowsApps` does not count; it is only the shortcut to the Microsoft Store. Without one it installs Python 3.13 through winget, just for you and without admin rights. Without winget (older Windows 10), it opens python.org and says what to do. Then it sets up the packages; after that it just starts. If Windows says "Windows protected your PC", click **More info → Run anyway**. The folder may have spaces and umlauts in its name (`C:\Users\Jürg Müller\Crosscheck-main`). If it is still inside the ZIP, the launch files say to unpack it first.

### Updates

If you got Gegenprobe with `git clone`, updates come by themselves: **VoiceAgent.command** and **VoiceAgent.bat** fetch the newest version before every start and install new packages. Packages Gegenprobe no longer needs (such as the former speech recognition) are removed from `.venv` on the way. Without internet the current version simply starts. If you changed something in the folder yourself or are testing another branch, Gegenprobe leaves everything as it is and says so in Terminal. Switch it off with `VOICEAGENT_NO_UPDATE=1`.

If you downloaded a ZIP, Gegenprobe never downloads code. At most once a day it asks GitHub at start, without an account, when the newest version was made, and says so if there is a newer one (in the window and under **Zustand**). GitHub writes where the ZIP comes from into `src/voiceagent/stand.txt` when it packs it. You download a new version as a ZIP again; [Help](HELP.md#new-version) says how to put it in place of the old one. Offline or with `VOICEAGENT_NO_UPDATE=1` it does not ask.

## Using it

Double-click **VoiceAgent.command** (Windows: **VoiceAgent.bat**), or run `source .venv/bin/activate && voiceagent`. The browser opens `http://127.0.0.1:8765`. The start page says whether the panel in Excel is set up, takes an Excel file to check (drag it in or pick it, with a report) and has the practice workbook. Below, under **Mehr** (more), are the Excel helper and the status page.

### Excel checker (no AI)

Open **Excel-Prüfer im Browser** on the start page and work in Excel as usual. Gegenprobe reads the open workbook directly (on a Mac, macOS asks once whether Terminal may control Excel: allow it) and shows every few seconds what does not add up: error values like #DIV/0!, a SUM that skips a row of the table or disagrees with the numbers, rounded values that no longer add up, a typed number instead of a formula, numbers copied into formulas instead of cell references, and the same constant in many formulas. It also checks the whole calculation chain with its own formula engine: formulas that break the pattern of their column (with the formula that belongs there), circular references, subtotals counted twice, references to empty cells, values Excel shows that differ from what the formula gives (manual calculation?), outliers, duplicate names and shares that do not add up to 100 %. "Wovon deine Ergebnisse abhängen" shows for every result how much it changes when an input or a constant in the formulas rises by 1 % (elasticity). Each finding comes with what usually helps. It also reads the name of a figure ("Umsatz alle Nord", "Durchschnitt Umsatz Nord", "Umsatz März 2026") and works out which rows belong to it: a range that misses a row or takes one too many, a formula dividing by a fixed count that no longer fits, or SUMIF finding too little because of a space or a typo. It says exactly what is wrong (which row is missing, which one is extra and why) and names the part of the formula that causes it: a condition too many, a wrong value, a range that stops too early, a wrong divisor or a cell counted twice, also two such errors in one formula. In a cross table (regions down, products across) it takes the name from row and column and finds a condition taken from the wrong header cell. Numbers and dates stored as text (often after an import) are found together with the figures they make wrong. It suggests only the corrected formula it has recalculated itself. It also finds ranges that slide when copied down without $, an average of per-row rates instead of the rate over all rows, averages of groups of different size, and growth divided by the new instead of the old value. No AI, no key and no internet; the data stays on the computer.

### The checker as a side panel in Excel (Mac)

Instead of the browser, the checker can sit on the right inside Excel, like the comments pane. Set it up once: double-click **ExcelAddin.command** in Finder (or run `./ExcelAddin.command`). If Gegenprobe is not installed yet, it installs it first, like VoiceAgent.command. It creates a certificate for this Mac only, because Excel opens add-ins only over a secure connection; the Mac asks for your password once. It then registers the add-in with Excel. For that, Terminal has to write into Excel's folder once; if macOS refuses, it opens the **Full Disk Access** setting for you: switch Terminal on there, quit Terminal completely (Cmd-Q) and run ExcelAddin.command again. Afterwards you can switch Full Disk Access off again; later runs leave Excel's folder alone as long as the add-in is unchanged.

Finally it sets up the checker in the background (macOS may report a new background item "python"): macOS keeps port 8766 open itself and starts the checker as soon as Excel opens the panel, also after a restart. Until then it uses no memory, and it never needs Terminal or a browser. ExcelAddin.command tries this right away: it loads the checker the way macOS does at login and then knocks like Excel. If it does not start by itself, data/pruefer-dienst.log says what macOS reports about it. From then on: open Excel (the first time, quit it completely and reopen), click **Add-ins** at the right of the **Home** tab and pick **Gegenprobe Prüfer**; from then on **Home** has a **Prüfer** button that opens the panel on the right. Without the background service, VoiceAgent.command serves the panel while it runs. After an update through VoiceAgent.command the background service restarts with the new version, and an open panel reloads itself. The panel reads the workbook directly through Excel, checks again after every change and selects a cell when you click its hint. When a hint has a recalculated fix, **Einsetzen** (insert) writes it into the cells (it shows first which formula goes where) and **Rückgängig** (undo) takes it out again. This also works for numbers stored as text (they become real numbers in their column's format), ranges that stop too early and VLOOKUP without FALSE. When several hints have a fix, **Alle einsetzen** (insert all) writes them together after Gegenprobe has recalculated the sheet the way it will compute afterwards; fixes that get in each other's way are left out. Otherwise nothing is changed. The cells only go to Gegenprobe on this computer; the only thing Excel loads from the internet is Microsoft's office.js. Remove it (also from the background) with `./ExcelAddin.command --entfernen`.

The panel recalculates the whole workbook, including references to other sheets (Tarife!B3), VLOOKUP, HLOOKUP, XLOOKUP, INDEX and MATCH, text and date functions. It reports typical lookup mistakes on their own: an approximate search (no FALSE at the end) in an unsorted table, a text key that is missing and silently takes the next smaller entry, keys that IFERROR or IFNA silently cover with a fallback (with the reason when it almost matches: spaces at the edge or a number stored as text), duplicate keys, and a table without $ that slides down when the formula is copied. "Auf den anderen Blättern" lists where else there are mistakes; a click opens that sheet.

The **Herleitung** tab shows, for the cell you select in Excel, where its number comes from: the formula (also in words), the value from Excel and whether it recalculates the same, and below it, step by step, the cells it comes from, also on other sheets. For VLOOKUP, XLOOKUP and INDEX it shows the cell that was found. Large ranges are summarised (how many numbers, their sum) and only the cells that are wrong, or build on something wrong, are opened. That way you can follow a wrong total down to its cause. When the formula calculates from the table, one sentence says which rows it really takes (for example the 9 rows with Region "Nord" except D11), so you can see at a glance whether a formula does what it should, even without a hint. "Fließt in" lists the formulas that read the cell. Clicking a cell in the tree selects it in Excel.

The **Mappe** tab explains the whole workbook on one page, from the data to the results: where the data is, which sheet reads from which, whether every formula recalculates the same as in Excel, and where there are hints. Below are the results (formulas and calculated columns nothing else uses), each with what it is built from over all steps, also on other sheets, with fixed numbers and TODAY() on the way and a number on the way that does not add up. For each sheet it lists every calculated column with its formula in words ("=Betrag * Satz", the same in all 7 rows) and every single formula with the sentence saying which rows it takes. A reference without $ that moves when filled down is pointed out ("E3 takes D10"). Clicking selects the cell in Excel. That way you can understand a workbook someone else built. The report with values has the same section.

Under **Eigene Regeln je Spalte** (own rules per column) you decide what a column may contain: select a cell in the column, click **Regel für die markierte Spalte** and choose required, kind (number, whole number, date, text), from and to, not before or after a day, not in the future, allowed values ("m, w, d"), every value only once, a pattern ("PAT-####": # a digit, @ a letter) or a comparison with another column (discharge not before admission). The form shows what is in the column now. The rule applies in every workbook with a column of that name (or on one sheet only) and stays on the computer. Every violation is an error with its cell and reason, including values that are only written differently ("M" instead of "m", a trailing space), a number stored as text or a date without a date format. Clicking selects all the cells, **Regel ändern** opens the rule.

For data analyses (cases, patients, measurements) it also checks: missing values per column and placeholders like "k.A.", codes like 999 or -99 for "unknown" that AVERAGE silently includes, impossible values (age over 120, negative length of stay, shares over 100 %, pulse, BMI, temperature), a discharge before the admission or a birth date in the future, duplicate case numbers and identical rows, two units in one column (1.72 next to 172, mmol/l next to mg/dl), AVERAGE, MEDIAN or COUNT over a range that leaves out the last rows, a denominator that slides when copied (C10 instead of C$10), and SUM or AVERAGE over a filtered table that still counts the hidden rows (use SUBTOTAL). It checks whether tests and intervals are set up right: T.TEST paired over two groups or unpaired over the same cases before and after, with equal variances where they are far apart, or one-sided without saying so; CONFIDENCE with 0.95 instead of 0.05, with the standard error instead of the standard deviation, or the normal distribution for few values; a hand-made confidence interval without the square root of n or with the wrong z. It gives the right p-value and says whether significance changes. The difference of two percentages under "Veränderung %" is percentage points, and it says so. An age computed by hand (days divided by 365 or 365.25, YEAR minus YEAR) is a year off around the birthday; it says in which rows and inserts DATEDIF. TODAY() in a table of cases with admission or discharge changes every day, and usually the age at admission or the stay until discharge is meant. Days "inclusive" without +1, "nights" with +1 and working days that count weekends are found too. It compares pivot tables with their data: one that was not refreshed after a change (with the numbers that are now wrong), a fixed source with more data below it, and "Anzahl von …" (count) where a blank cell made Excel count instead of sum. In charts it notices a series that stops before the end of the table, takes the total row along or is shifted against its labels. In the panel this works with "Aktualisieren", "Summe nehmen" and "Einsetzen". To try it out there is the practice workbook [pruefer-uebung.xlsx](../src/voiceagent/web/demo/pruefer-uebung.xlsx) with invented data and planted mistakes (also linked at the bottom of the panel).

**Large workbooks** stay usable: the panel reads large sheets in chunks and shows how far it is. A sheet that has not changed is not sent again, and Gegenprobe checks other large sheets in the background while you change nothing ("wird noch geprüft", still being checked). It reads up to 500,000 cells per sheet and 1.5 million in the whole workbook, the active sheet first. If a sheet is larger, it reads whole rows up to that point and says up to which row. About formulas that calculate with the rows below, it then says nothing rather than report a false difference to Excel.

Under **Prüfbericht** in the panel a report on all sheets opens, to print or save as PDF (in Chrome or Safari, Print → Save as PDF). **Mit Werten** (with values) contains the full hints with the numbers and names from the workbook. **Ohne Daten** (without data) lists only level, cell and the kind of hint, the sheets are called Blatt 1, Blatt 2, and the workbook name is left out, so you can pass it on even when the workbook holds patient data. Reports are kept in the data folder under `berichte`, only the last 20 stay.

In the **Datei** (file) tab you drag an earlier version of the workbook into the panel (or click and pick the .xlsx file). Gegenprobe compares it with the open workbook and first shows the results that are different now, with the reason: which input changed before it, which formula is different, whether new rows are counted or rows are missing, what changed on another sheet, or that a number was not recalculated in Excel. Below that: formulas that are now fixed numbers, changed formulas (grouped per column), changed inputs, new and deleted rows and columns, renamed sheets. Inserted and re-sorted rows are recognised by row names and content, so moved formulas do not count as changed. **Prüfen** (check) checks the file itself without opening it in Excel, all sheets, with a report too (without data or with values). If the file is no version of the open workbook (no sheet in common), it is checked right away. The file stays in memory.

### The checker as a side panel in Excel (Windows)

On Windows, double-click **ExcelAddin.bat** in Explorer; no admin rights and no extra packages. If the installation is missing, it runs it first like VoiceAgent.bat. Then: (1) a certificate for this PC only, made by Python itself, without openssl; (2) Windows asks once in a security warning whether to trust the authority "VoiceAgent lokal (nur dieser Rechner)": click **Yes** (it goes into your user's certificates only); (3) it registers the add-in as a developer add-in with Excel (registry `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`, the manifest lives in `%USERPROFILE%\.voiceagent\addin`); (4) the checker starts in the background at every login without a window (`HKCU\…\CurrentVersion\Run`, pythonw) and once right away, and ExcelAddin.bat checks that it answers. Then close Excel completely and reopen it: the **Home** tab has a **Prüfer** button. If it is missing, pick **Gegenprobe Prüfer** under developer add-ins in **Home → Add-ins** (or **Insert → My Add-ins**). Nothing needs to change in the Trust Center. While VoiceAgent.bat runs it serves the panel itself; after an update through VoiceAgent.bat the background checker restarts with the new version. Its messages go to `data\pruefer-dienst.log`. Remove it with `ExcelAddin.bat --entfernen` in a command prompt in the folder; the certificate stays in Windows so that setting up again does not ask again; remove it under **Manage user certificates**.

According to Microsoft, Excel 2016 and 2019 show add-ins with Internet Explorer, which cannot run the panel; there the checker stays in the browser. Excel 2021, 2024 and Microsoft 365 use WebView2 (Edge). Older Microsoft 365 builds with legacy Edge (EdgeHTML) need a localhost exemption (`CheckNetIsolation LoopbackExempt`). The panel has not been tried on a real Windows PC yet: registry, certutil and process start are simulated in the tests; the certificate is checked with a real, strict TLS connection.

### Excel helper: pivot summaries on request

Open **Excel-Helfer** on the start page under **Mehr** or in the Excel checker. A narrow window opens that stays next to Excel. Type what you want, for example "Summe der Sitze pro Partei", "Durchschnitt Alter pro Station und Geschlecht", "Anzahl pro Kanton" or "Summe Umsatz pro Region, Jahre in den Spalten". Gegenprobe takes the active sheet, finds the columns (also roughly: "Parteien" for "Partei") and first shows a preview with the resulting numbers. Only after **In Excel schreiben** does it add a new sheet; it never changes existing cells.

The summary consists of real formulas (SUMIFS, COUNTIFS, AVERAGEIFS, MINIFS, MAXIFS) on your data, with a total row, optionally as a crosstab and sorted ("größte zuerst"). It keeps up when the data changes, and every number can be traced in Excel. A grand total below the data is left out. After writing, it reads the new sheet back and compares every number with its own calculation. It can also try a real Excel PivotTable (COM on Windows, AppleScript on a Mac); if that fails, the formula summary is written instead.

Gegenprobe understands the sentence without AI, on your computer. Undo does not work for the new sheet: right-click the sheet name and delete it.

### Excel helper: tables from files

In the Excel helper, under "Oder eine Tabelle aus einer Datei", pick a file or drag it into the window: CSV or text (semicolon, comma, tab or |), Excel (.xlsx, .xlsm, every sheet), Word (.docx, every table) or PDF (tables with real text, no scans). If the file has several tables, pick one. Gegenprobe shows the table first and only writes it into a new sheet of the open workbook after **In Excel schreiben**, with a bold header row.

Numbers become real numbers, also German (1.234,50), Swiss (1'234.50) or English (1,234.50), with percent (12,5 %) and currency (CHF 80.–). Dates (1.3.2024, 2024-03-01) become real dates. What only looks like a number stays text so nothing is lost: postcodes and IDs with a leading zero (08001), phone numbers, IBANs and long ID numbers. After writing, it reads the sheet back and compares every value with the file. The file stays on the computer. You can then ask for a summary of the new sheet right away.

### Status and error report

**Zustand** (status, on the start page under Mehr) shows whether everything is ready: version, connection to Excel and data folder, with what to do when something is missing. If something does not work, download the **Fehlerbericht** (error report) there and send it along when you ask for help. It holds the status, package versions and the latest warnings and errors. Cell contents are never in it, folder names under your user become `~`, and e-mail, IBAN and the like are masked. The page shows exactly what is in it first.

## Privacy

- Everything stays on your computer. Gegenprobe needs no AI, no account and no internet; only updates come from GitHub. From a ZIP it asks GitHub at most once a day whether there is a newer version, without sending anything from the folder or your workbooks.
- The server listens on `127.0.0.1` only and cannot be reached from the network. A web page open in your browser cannot reach your data either: Gegenprobe answers only pages on this computer (`127.0.0.1` and `localhost`).
- The panel in Excel sends cells only to Gegenprobe on this computer; the only thing Excel loads from the internet is Microsoft's office.js.
- A file you check or compare stays in memory and is not saved.
- The report **without data** holds only level, cell and kind of hint, so you can pass it on even when the workbook holds patient data.
- Gegenprobe stores no content of your workbooks. The log (`data/pruefer-dienst.log`) holds no cells, names or values, not even in the error report.
- Only three things are kept, all triggered by you: **Passt so** is stored in the workbook itself, as a checksum without content. **Your own column rules** live in `data/spaltenregeln.json`. A **report** sits in `data/berichte` so the browser can open it, and is deleted after one day; to keep it, save it as PDF from the browser.

## Limits and liability

Gegenprobe recalculates and looks for typical mistakes. It can miss mistakes and can flag something that is correct. It does not replace your own review: whoever uses the workbook stays responsible for its numbers and decisions. Every fix shows what it writes where before it is applied, and **Rückgängig** takes it back out. Gegenprobe is free and open source and provided without any warranty (see [License](../LICENSE)). The panel shows this notice when it first opens and any time via **i**.

## Light on memory

Gegenprobe needs no model and no large packages: Python with FastAPI, Uvicorn and pypdf, about 60 MB on disk together. On a Mac the background checker only starts when Excel opens the panel.

If you used an older version with speech recognition, you can delete its model (about 1.6 GB). The folders `data/sessions` (sessions with screenshots) and `data/anweisungen` also come from the former apprentice; Gegenprobe no longer reads them, so you can delete them.

```bash
rm -rf ~/.cache/huggingface/hub/models--mobiuslabsgmbh--faster-whisper-large-v3-turbo
```

## Settings

| Environment variable | Default | Meaning |
|---|---|---|
| `VOICEAGENT_DATA_DIR` | `data` | Where reports, own rules and the log are kept |
| `VOICEAGENT_NO_UPDATE` | empty | `1`: do not look for a new version before starting |
| `VOICEAGENT_PYTHON` | empty | Use this Python for the installation instead of looking for one |
| `VOICEAGENT_NO_EXCEL` | empty | `1`: do not connect to the open Excel (for tests, for example) |

## Development

```bash
pip install -e ".[dev]"
pytest
```

The code is in `src/voiceagent`. The checks live in `analyse.py` (entry point), `formel.py` (formula engine), `workbook.py` (finding tables), `kopf.py` (headers against formulas), `finanz.py` and `fallen.py` (finance and statistics traps); `pruefer.py` collects the hints, `app.py` serves pages and endpoints, and `web/` holds the start page, the checker and the side panel.

## License

MIT, see [LICENSE](../LICENSE).
