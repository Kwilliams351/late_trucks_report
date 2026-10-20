-- Trucks Report - Databricks truck orders (replaces the Tableau custom SQL on <catalog>.truck_orders_table)
--
-- What changed vs the original and why it is faster:
--   1. Date window: the original had no WHERE clause, so every refresh pulled the full order history and
--      Tableau then joined it across databases to Teradata in memory. Now a rolling 120 days on trkorddat.
--   2. Duplicates: the original used SELECT DISTINCT and commented out autdat ("generating duplicates").
--      QUALIFY ROW_NUMBER() keeps the latest version of each order instead, so one row per TLM is guaranteed.
--      If trkorddat is not the departure date, widen the window (e.g. 130 days) so no order is dropped.

SELECT
    trkordnum,
    fltcaridr,
    CAST(fltnum    AS INT) AS fltnum,
    CAST(fltseqnum AS INT) AS fltseqnum,
    trkcmpcod,
    trkcmpnam,
    crtarpcod,
    trkorddat
FROM <catalog>.truck_orders_table
WHERE trkorddat >= date_sub(current_date(), 120)
QUALIFY ROW_NUMBER() OVER (PARTITION BY trkordnum ORDER BY autdat DESC) = 1;

-- Join used by the workbook (Tableau relationship / join clause):
--   trucks.fltnum = legs.MKT_FLIGHT_NBR AND trucks.fltseqnum = legs.ITIN_FLIGHT_LEG_SEQ_NBR
-- The join has no date key. fltseqnum is unique per date in the sample, but if both tables are ever in the same
-- database, add AND legs.SCHD_DEP_DT BETWEEN trucks.trkorddat - 2 AND trucks.trkorddat + 2 as a safety check.
-- Best option: land the Teradata legs in Databricks (or the orders in Teradata) and do the join in ONE query, so
-- Tableau does not have to run a cross-database join on every refresh.
