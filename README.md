# Trucks Report (late trucks) builder

Replaces the download step of the Tableau workbook **Trucks Report_KW**. It builds:
- a daily Excel report that opens on an **Action List** of trucks to act on now, keeps the red/green colors, and has rows ready to paste into the Truck Report Card;
- a JSON feed for the Coordinator Hub and Trucking Pulse;
- a **history store** of every truck ever loaded (one row per TLM) to connect to Power BI or other reports.

## Run it
**Daily:** drag the export CSV onto **`Run_Late_Trucks.bat`**. It builds the last 7 days, adds the whole file to the history store, copies the feed to `hub\`, and opens `Output\`.

**Backfill history (one time):** drag your old exports (several files, or a folder of them) onto **`Build_History.bat`**. It merges them into `History\` and does not build a daily report. Run it once with all the historical files you have; after that the daily run keeps the history up to date.

Accepted inputs (columns are matched by name):
- the custom SQL data source export (`AC Type, Actual Arrival Color, Actual Arrival Date, ... I, M, ... MKT_FLIGHT_NBR, Placeholder`);
- the sheet's **Download > Data** CSV (`Sheet 1_data.csv`, 3 rows per TLM);
- the result of the queries in `sql\` saved as CSV or XLSX;
- `History\late_trucks_history.csv` itself.

Command line options:
```
py -3 late_trucks.py export.csv                                  # last 7 days (default)
py -3 late_trucks.py export.csv --from 2026-10-01 --to 2026-10-18 --origin JFK,ORD --vendor "FLY US"
py -3 late_trucks.py export.csv --arr-basis team                 # score arrivals against the Report Card Schedule tab
py -3 late_trucks.py export.csv --report-card none               # skip the Report Card cross-check
py -3 late_trucks.py "Old exports" 2025.csv --history-only       # backfill the history store
py -3 late_trucks.py export.csv --no-history                     # do not read or write History\
py -3 late_trucks.py export.csv --lookback 24 --grace 90         # Action List hours; minutes before "No actual"
```
Six years of data (about 320K trucks) builds in about 25 seconds. Requires Python 3.9+, pandas, xlsxwriter and openpyxl. Optional: `pyarrow` (also writes a .parquet copy of the history) and `airportsdata` (time zones for stations not in the built-in list).

## Daily report: `Output\Trucks Report <as-of>.xlsx`
| Sheet | Contents |
|---|---|
| **Action List** | Open items from the last 48 h, worst first, each with a next step: **Not departed** (no departure recorded 2 h+ after schedule), **Arrival overdue** (departed, no arrival recorded 2 h+ after the scheduled arrival), **Projected late arrival** (in transit; actual departure + the lane's usual transit lands more than 59 min after the scheduled arrival), **Departed late** (still en route), **Fix record** (actual keyed on the wrong date). Counts across the top. |
| Trucks Report | One row per TLM with the same columns as Tableau. Actuals are colored **red** (late) or **green** as real cell formatting. Adds delay minutes, status, transit hours and an Issue column. *On Report Card* and *Team Sched Arr* appear when the Report Card was read. |
| Late - Report Card | Late trucks in the exact 23-column order of *Report Card 2026*. Times are written like `2030` and `1305+1`, and Source = TABLEAU. Reason/Fault/Y-N dropdowns; Fault fills from Reason. Trucks already logged are grey and sorted last. Paste columns A:W as values. |
| Summary | KPIs, then results by vendor, origin, lane (top 40), **repeat late trucks** (late on 3+ days) and day. |
| Data Quality | Every record to fix, with the problem in words: departure 1+ day off, arrival before departure, arrival keyed same as departure, no departure, arrival not recorded. |
| About | Definitions, sources, counts, and the check against Tableau's own RED/GREEN colors. |

`Output\late_trucks_feed.json` and `.js` (`window.LATE_TRUCKS_FEED`) contain KPIs, results by station, vendor and lane, repeat late trucks, and truck-level rows for the last 3 days plus open Action List items and late trucks from the last 7 days not yet on the Report Card. Each truck row now also has `issue`, `action`, `priority`, `proj_arr` and `proj_arr_delay`. Existing fields are unchanged, so the Hub keeps working.

## History: `History\`
| File | Use |
|---|---|
| `late_trucks_history.csv` | One row per TLM for every date loaded. Clean timestamps (`YYYY-MM-DD HH:MM`), delay minutes, statuses, 0/1 columns (`Late`, `Late Departure`, `Late Arrival`, `Arrival Measured`, `Data Issue`) that sum straight into on-time %, plus Lane, Year, Month, Week Start, Weekday and Source File. Point Power BI at this file. |
| `late_trucks_daily.csv` | Rollup by date, lane and vendor: trucks, late departures and arrivals, arrivals measured, on-time %, median minutes late. Small enough for Excel pivots. |
| `Trucks History Summary.xlsx` | Monthly network trend, plus on-time arrival % by vendor and destination, and late departure % by vendor and origin, for the last 24 months (heat-mapped). |

How the store stays clean:
- Each run reads the store plus the new file(s), keeps one row per TLM, and re-scores everything with the current rules. If you change a rule, the whole history changes with it.
- When the same TLM appears twice, the copy with the newer record version (`I`) wins; without `I`, the newer file wins.
- Rows where `I` is not equal to `M` (superseded versions) are dropped.
- If one TLM matched more than one flight leg (the Tableau join has no date key), the leg closest to the order date (`trkorddat`) is kept.
- TLMs with no matching flight leg are dropped, the same as Tableau's *Exclude Nulls* filter.
- Close `Trucks History Summary.xlsx` before a run, or it is skipped with a message.

## Definitions
- **Late** means actual minus scheduled is more than 59 minutes, counted in whole minutes like Tableau's `DATEDIFF('minute')`. This is the same rule as Tableau's RED calc. Change it with `--threshold`.
- **Check date** means the actual was probably keyed on the wrong day or month. These are shown amber and are not counted as late or on time. A record gets Check date when:
  - the departure is more than 2 days late or more than 12 h early;
  - the arrival is before the departure, or within 10 minutes of it;
  - the arrival is more than 2 days late.
- Arrivals 2+ days **early** are **not** Check date. Friday and Saturday trucks are scheduled to arrive Monday, so a Saturday arrival is on time. The old rule flagged these.
- **No actual** means nothing was recorded 2 h after the scheduled time (`--grace`). These appear on the Report Card tab as `NO SHOW?` to confirm. **Pending** means the time has not come yet.
- Times on the flight record are station local. "Now" is converted from this PC's clock (`--local-tz`, default America/Chicago) to each station's time zone, so a 10:00 LHR truck counts as overdue at 06:00 Chicago time (12:00 UK).
- **Projected arrival** is the actual departure plus the lane's median actual transit over the last 120 days of clean trips. If a lane has fewer than 3 clean trips, the scheduled transit is used instead.

## Findings
From the 6/1–10/18/2026 export (22,109 trucks):
- The Tableau download has 3 rows per TLM (departure, spacer and arrival panes). The red/green color is not exported. This tool collapses the rows and rebuilds the color.
- About 7,300 trucks flagged late by the >59 min rule are **not** on the Report Card. Only about 15–19% of flagged trucks get logged.
- **Scheduled arrival:**
  - for 1,937 trucks, the flight record's `SCHD_ARVL_DTTM` is 1 hour or more later than the team Schedule tab ETA. This hides late arrivals.
  - With `--arr-basis team`, late arrivals go from 4,396 to 5,124, and the share of team-logged trucks shown as late goes from 57% to 72%.
  - Confirm with the data owner which field is the published ETA.
- The arrival actual uses `ACTL_ARVL_TM`, while the departure actual uses `ACTL_LOCAL_DEP_TM`. Verify that arrival is in local time.

From the 10/12–10/18 week, re-scored with the new rules (1,038 trucks):
- 14 arrivals that were Check date are now On time. They were weekend trucks with Monday ETAs that arrived early.
- 22 arrivals are now Check date:
  - the arrival is the same minute as the departure (SUA-MIA 7083/7084 every day, SNN-DUB 8504);
  - or the arrival is before the departure.
- 12 departures that counted as On time are now Check date because they were 13–30 h early. One example is YYZ-JFK at 23:59 the day before.
- Late arrivals go from 197 to 191. On-time arrival goes from 77.3% to 77.8%.
- 91 trucks departed but have no arrival recorded more than 2 h after their scheduled arrival. The 68 from the last 48 h are on the Action List as *Arrival overdue*. All 91 are on Data Quality.

## Query improvements (`sql\`)
| Today | Change | Why |
|---|---|---|
| Teradata pulls `SCHD_DEP_DT BETWEEN '2020-01-01' AND CURRENT_DATE` | Rolling `CURRENT_DATE-120 .. +1` | Scans and sorts about 4 months instead of 6+ years. Now that the history store keeps the old data, the daily pull can be short. |
| `DISTINCT` + `MAX() OVER` in a derived table + `WHERE I=M` | `QUALIFY ROW_NUMBER() ... = 1` | Uses one sort instead of two, and ties can no longer duplicate rows |
| Databricks orders have no `WHERE` and use `SELECT DISTINCT` | `trkorddat >= current_date-120` + `QUALIFY ROW_NUMBER() BY trkordnum` | Pulls a fraction of the rows and gives a true 1 row per TLM |
| Tableau cross-database join (Databricks to Teradata) | Land one table next to the other and join in one query | The cross-database join is run in Tableau's engine on every refresh |

## Tableau speed and export fixes (if you keep the workbook)
1. **Extract:**
   - Add a rolling-date extract filter, for example the last 120 days. The history store keeps everything older.
   - Turn on incremental refresh on `SCHD_DEP_DT`.
   - Move the *Exclude Nulls* date filter into a **data source filter**.
2. **Filters:**
   - Set quick filters to *All values in database*, not *Only relevant values*. Relevant-values filters re-query on every change.
   - Add **Apply** buttons to the date and multi-select filters.
   - Default the date filter to the last 3–7 days.
3. **Layout:**
   - The sheet puts 10 discrete fields on Rows over about 22K rows, and copies a placeholder measure on Columns to fake the panes. This is slow to draw and is why the download has 3 rows per truck.
   - Rebuild it as a plain text table with `Dep Status` / `Arr Status` text fields and delay-minute fields colored by status.

## How it fits the existing tools
- **Coordinator Hub > Ops Day-of board:** `Run_Late_Trucks.bat` copies `late_trucks_feed.js` next to `hub\Coordinator Hub.html`; reload the Hub to pick it up. To publish to the team, set `HUB=` in the .bat to the shared Hub folder. The feed's new `action`/`priority` fields can drive an "open items" panel later.
- **Power BI / Dashboard v10:**
  - Connect to `History\late_trucks_history.csv`.
  - On-time arrival % = `SUM(Late Arrival) / SUM(Arrival Measured)`, subtracted from 1.
  - Late departure % = `SUM(Late Departure) / COUNT(TLM)`.
  - Use `Data Issue = 0` to exclude bad records.
  - The Vendors scorecard can use *all* trucks, not only the logged ones.
- **Trucking Pulse:**
  - Match truck numbers found in emails to `trucks[]` in the feed.
  - Surface `action` next to the email.
  - Prefill the T03 detention template with the real scheduled and actual times.
- **Truck Report Card:** paste from the *Late - Report Card* tab. The *Already on Report Card* column prevents duplicate entries.

## Tests
`py -3 -m unittest discover -s tests -v`
