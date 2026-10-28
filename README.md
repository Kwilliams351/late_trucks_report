# Trucks Report (late trucks) builder

Turns the Tableau **Trucks Report_KW** export into:
- a **daily workbook** where every number is an Excel formula you can trace back to one **Data** tab (the cleaned export), with a **Dashboard**, a coordinator **Worklist** for the last 7 days, and a Report Card tab ready to paste;
- a **history workbook** with monthly and weekly trends;
- a **history store** (CSV) of every truck loaded, for Power BI or other reports;
- a JSON feed for the Coordinator Hub and Trucking Pulse.
- a private CSV snapshot on each daily run and, once an earlier snapshot exists, a **Shift Variance** tab.

## Run it
- **Daily:** drag the export CSV onto **`Run_Late_Trucks.bat`**. It opens `Output\` when it's done.
- **Backfill history (one time):** drag old exports, or a folder of them, onto **`Build_History.bat`**.

```
py -3 late_trucks.py export.csv                          # report window = last 7 days + tomorrow (default)
py -3 late_trucks.py export.csv --as-of "2026-10-19 02:45"
py -3 late_trucks.py export.csv --from 2026-10-01 --to 2026-10-18 --origin JFK,ORD --vendor "FLY US"
py -3 late_trucks.py export.csv --data-days 91           # trucks on the Data tab (default 13 weeks)
py -3 late_trucks.py export.csv --report-card none       # skip the Truck Report Card cross-check
py -3 late_trucks.py "Old exports" --history-only        # backfill the history store
py -3 late_trucks.py export.csv --vs prior-day           # compare with a snapshot at least 24 h earlier
py -3 late_trucks.py export.csv --vs "Output\Trucks Report <prior>.xlsx"
```
- **Requirements:** Python 3.10+, pandas, numpy, xlsxwriter, openpyxl and tzdata (Windows time zones).
  Install with `py -3 -m pip install -r requirements.txt`. Keep `formula_book.py`, `report_card_pending.py`
  and `shift_compare.py` next to `late_trucks.py`. Optional: airportsdata for additional station time zones.
- **Run time:** about 45 seconds for 206K trucks (6 years), including both workbooks.
- **As-of time:** comes from the file's time stamp. If the export has departures recorded more than a day after that time (a copied file, or shifted/anonymized dates), the latest departure is used instead and a note is printed. Pass `--as-of` to override.

## Daily workbook: `Output\Trucks Report <as-of>.xlsx`
How the numbers flow:

**Data** (export values + row formulas) → **Summary / Trends / Lanes / Vendors / Stations** (COUNTIFS, SUMIFS, MEDIAN over Data) → **Dashboard** (tiles and charts that point at those tables) and **Worklist / Report Card** (INDEX into Data + your entries)

| Tab | What it is |
|---|---|
| **Dashboard** | KPI tiles for the window, Worklist progress by issue and by status, and charts: daily trucks / late departure % / on-time arrival % (35 days), weekly trend (13 weeks), daily follow-ups (No ATD, No ATA, Check date), and top vendors, origins and destinations. |
| **Worklist** | One row per truck in the window with an open issue, worst first (see below). |
| **Report Card** | Late, pending and check-record trucks in the window. A:W retain the 23-column *Report Card 2026* order. Filter **Card status (X)** to **Ready to log**, and **Already on Report Card** to blank, then copy visible **A:W** using **Paste Special > Values**. Pending/check rows are not for pasting. Times use `1305+1`. Reason, Fault, Summary and Agent come from Worklist entries. Already-logged ready rows are grey. |
| **Summary** | Network KPIs, then results by vendor, origin, destination, lane (top 40), repeat late trucks (3+ days) and day. Every cell is a formula, e.g. `=SUMIFS(Data[Late Dep],Data[Vendor],$A20,Data[In Window],1)`. |
| **Trends** | Daily (last 35 full days) and weekly tables that feed the charts, plus a vendor × week late-% heat map. |
| **Data** | One row per TLM for the last 13 weeks. **Blue headers** are values from the export. **Green headers** are formulas such as `Dep Delay`, `Dep Status`, `Arr Status`, the 1/0 flags, `Proj ATA`, `In Window` and `Issue`. Click any cell to see how it was derived; filter to see the trucks behind any number. |
| **Lanes** | Clean trips and median transit per lane (array MEDIAN over Data). These feed `Proj ATA`. |
| **Vendors** | Vendor email, phone and notes (**yellow, yours**) plus window performance. The email drives the Worklist's *Email vendor* link. |
| **Stations** | Offline stations (from the training guide, with their notes), time zone, and clock offset vs the as-of time. Also has station email and phone (yours). |
| **Settings** | Named cells used by every formula: `AsOf`, `Threshold`, `Grace`, `EarlyCheck`, `LateCheck`, `MinTransit`, `MinLaneTrips`, `WinFrom`, `WinTo`. Also the truck-type, issue, step, fault and status lists. Change a yellow cell and the whole workbook re-scores. |
| **About** | The audit trail and every Data column's formula, written out. |
| **Shift Variance** | Present when an earlier snapshot/report is available: prior/current totals, transition counts and every changed truck with prior/current times and statuses side by side. |

### Pending trucks and safe Report Card pasting
**X: Card status** is linked to a formula on Data, so editing `AsOf`, `Threshold` or actual times in Excel
updates the displayed classification:

| Card status | Meaning |
|---|---|
| **Ready to log** | Late, with actual departure and arrival recorded (or arrival not tracked). Verify the delay before pasting. |
| **Pending - arrival overdue** | Departed, no ATA, and the station-local arrival time plus grace has passed. |
| **Pending - projected late** | Departed, no ATA, not yet overdue, and projected arrival delay exceeds Threshold. |
| **Pending - in transit** | Departed and waiting for ATA, without an overdue/projected-late flag. |
| **Pending - no departure** | No ATD yet, including scheduled future departures and possible no-shows. Confirm before logging; this is not a confirmed cancellation. |
| **Check record** | A departure or arrival failed the existing date sanity checks; verify before logging. |

Check-record takes precedence over pending, which takes precedence over ready (a late departure with no ATA is
still pending). Ready rows sort first at build time; pending rows are amber and check rows red.
Row 3 contains formula-driven group counts. **AF: Proj ATA** and **AG: Proj Arr Delay (min)** are linked to Data.
`Late on` says **Not late yet** when neither actual is late. The Worklist vendor email link uses the Vendor
and Email subject/body headers, not fixed column positions.

**Row membership and order are fixed when the file is built.** Excel updates formulas in existing rows but does
not insert newly eligible rows or re-sort them. Rebuild to change the date window or refresh the row list.
Re-filter X immediately before copying; never copy pending/check rows, X or the helper columns.

### Shift snapshots and comparison
Each successful daily build saves **every Data-tab truck**, not just the Worklist, in
`snapshots\snapshot_<as-of>_<run-time>.csv`. It records TLM, STD/ATD/STA/ATA, projected ATA, statuses,
the report-window flag, as-of time and scoring rules. A rerun never overwrites an earlier snapshot.
Snapshots reflect the build-time export, not subsequent Excel edits; save a recalculated workbook as a snapshot
with the standalone command below to capture those edits. History-only backfills do not save shift snapshots.

- **Default `--vs prior-shift`:** latest snapshot strictly earlier than the current as-of time.
- **`--vs prior-day`:** latest snapshot at or before current as-of minus 24 hours.
- **`--vs "<prior>.xlsx"` or `"<snapshot>.csv"`:** use a specific earlier report/snapshot.
- **`--snapshots-dir "<folder>"`:** use a different local snapshot directory.

Without a prior snapshot the first run saves the baseline, prints that fact and omits Shift Variance.
The next shift adds the tab automatically. Same-as-of reruns cannot compare against themselves.
Reports used as comparison inputs must have cached Data/Settings values: open, recalculate and save in Excel
first if needed. Older formula-linked reports without Card status are supported.

The comparison shows totals for each report's own window. Transition counts use only TLMs in **both** windows
to avoid counting trucks that merely entered/left a sliding window. The detail lists all added, removed or changed
Data-tab trucks, including rows outside the report window; window flags are shown alongside each row.
Transitions include **became late**, **no longer late (time regained)**, **pending -> ready**, and **now overdue**.
Time regained requires valid actual statuses, not a missing or invalid record.
**Clock only** means times, projection, station offsets and scoring rules are unchanged and only advancing the
as-of clock changed a status. It can overlap another transition; transition totals are not additive.
Comparison does not match by truck number alone: TLM must be present and unique.

Standalone tools (source workbooks are never rewritten, preserving their charts):
```
py -3 shift_compare.py snapshot "Output\Trucks Report <as-of>.xlsx"
py -3 shift_compare.py compare "Output\Trucks Report <as-of>.xlsx" --vs prior-day
py -3 shift_compare.py compare "Output\Trucks Report <as-of>.xlsx" --vs "Output\Trucks Report <prior>.xlsx"
```
`compare` writes a separate `... - Shift Variance.xlsx`; it does not save a new snapshot.
For nonstandard Output locations, pass `--snapshots` to the standalone tool.

### Public repository / private operational files
Reports, raw CSVs, history, contacts, coordinator notes, generated feeds, snapshots and Python caches
are excluded by `.gitignore`. Keep them local; never add them with `git add -f`.
Any included SQL must use placeholder schema/table names; configure actual database identifiers privately.
The public checkout does not contain a live Report Card path: pass `--report-card "<local workbook>"`.
Ignoring a file does not erase it from older Git history.

### Worklist (for the coordinator)
- **Which trucks:** each truck in the window with an issue, sorted by priority and then newest first. Each truck gets one issue, the highest that applies:

  | Priority | Issue |
  |---|---|
  | 1 | No ATD (possible no-show) |
  | 1 | No ATA (arrival overdue) |
  | 2 | Late arrival |
  | 2 | Projected late |
  | 3 | Late departure |
  | 4 | Fix record (date keyed wrong) |

- **Green columns (live from Data):**
  - **iCargo key:** `TRK7841/22MAY`, ready for TRK001.
  - **Truck details:** lane, vendor, truck type, whether the destination is offline, STD/ATD/STA/ATA, delays, projected ATA.
  - **How to check:** the steps for this situation, taken from the Trucking Training Guide. It covers TRK001/FLT002, the aa.trucking inbox, emailing the vendor, and arriving the truck (Station = destination, Role CTRKCRD, Reopen > Save > Approve). Offline stations get the FSU/RCF steps in OPR352 and LTE003. Mail trucks: enter STA as the ATA. No-shows: CAP142 > OPR352 > CAP018.
  - **Email vendor / Email station:** links that open a pre-written email asking for the ETA and delay reason, or for the ATA.
  - **Inbox search:** a ready search string for the inbox.
- **Yellow columns (yours):** Status (dropdown), Vendor ETA, Delay Reason (dropdown; Fault fills itself), Updated in iCargo, Notes, Worked by, Worked on.
- **Carry-forward:** your entries carry into the next report by TLM, and are saved in `Output\worklist_notes.csv`. Vendor and station emails carry the same way, in `Output\contacts.csv`. Rows marked Updated / Logged / Closed turn grey.
- **Sorting:** you can sort or filter the Worklist freely. Every row looks up its own truck by TLM.

## History workbook: `History\Trucks History Summary.xlsx`
- **Data** holds the last 24 months, one row per TLM. Statuses here are stored as values to keep the file fast. They use the same rules as the daily Data formulas.
- **Monthly** (all months and the last 26 weeks) and the vendor / origin / destination × month heat maps are formulas on that Data tab.
- **Dashboard:** the last full month vs the month before, plus monthly and weekly trend charts.

`History\late_trucks_history.csv` (every truck, all years) and `late_trucks_daily.csv` (rollup by day, lane and vendor) are for Power BI. On-time arrival % = 1 − `SUM(Late Arrival)` / `SUM(Arrival Measured)`.

## Rules
- **Late:** actual − scheduled > 59 min, counted in whole minutes like Tableau's `DATEDIFF('minute')`. This matches Tableau's RED/GREEN on 100% of comparable trucks in the anonymized export.
- **Check date:** the actual was probably keyed on the wrong day or month. Shown amber and kept out of late and on-time counts. A record gets Check date when:
  - the departure is > 2 days late or > 12 h early;
  - the arrival is before the departure or within 10 min of it;
  - the arrival is > 2 days late.
- **Early arrivals:** arrivals 2+ days early are fine. Friday and Saturday trucks have Monday ETAs.
- **No actual:** nothing recorded 2 h (`Grace`) after the scheduled time, compared in the station's local time. Anything earlier is Pending.
- **Truck types** (first 3 letters of AC Type, on Settings):
  - **Transfer (XFR) and RFS** never have an arrival, so their arrival shows *Not tracked*. The training guide says transfer trucks never have arrival times, and RFS has no arrival in six years of data.
  - **Mail (MAL)** arrivals are keyed as the STA, so they aren't scored; a missing one is still chased.
- **Projected ATA:** ATD + the lane's median transit, when the lane has 3+ clean trips; otherwise the scheduled transit.

## Checks
Run `py -3 -m unittest discover -s tests -v`. Tests cover statuses, the export layout, station time zones,
the Worklist, carry-forward of notes/contacts, Report Card columns and cached values, all six Card statuses,
snapshot round-trips, prior selection, clock-only transitions, additions/removals, empty snapshots and chart preservation.
Formula text and cached values are checked separately; this is not a substitute for recalculation in Excel.
The counts quoted by the other tool (449 rows, 277 ready, etc.) refer to its specific input/as-of/rules,
not fixed acceptance values for other exports or clocks.

## Query and Tableau suggestions
- **Teradata query:** pull a rolling 120 days instead of everything since 2020. The history store keeps the old data.
- **Deduplication:** use `QUALIFY ROW_NUMBER()` instead of `DISTINCT` + `MAX() OVER` + `WHERE I=M`, in both queries.
- **One join:** do the Databricks orders → Teradata legs join in one place, not as a Tableau cross-database join. See `sql\`.
