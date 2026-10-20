-- Trucks Report - Teradata flight legs (replaces the Tableau custom SQL on <schema>.FLIGHT_LEG_TABLE)
--
-- What changed vs the original and why it is faster:
--   1. Date window: was SCHD_DEP_DT BETWEEN '2020-01-01' AND CURRENT_DATE (6+ years scanned and sorted on
--      every refresh). Now a rolling 120 days (+1 so tomorrow's scheduled trucks show). Change :days_back as needed.
--   2. Latest-version pick: was SELECT DISTINCT + MAX() OVER (...) in a derived table + outer WHERE I = M
--      (two sorts, and ties on the timestamp still returned duplicate rows). Now one QUALIFY ROW_NUMBER() = 1,
--      which guarantees exactly one row per leg.
--      NOTE: the original window also partitioned by SEG_SEQ_NBR. Dropping it returns one row per leg instead of
--      one per segment; the report uses leg-level columns only. Add SEG_SEQ_NBR back if counts differ.
--   3. Only the columns the report uses are returned; no DISTINCT needed.
--
-- TO VERIFY with the data owner:
--   * Departure actual uses ACTL_LOCAL_DEP_TM but arrival actual uses ACTL_ARVL_TM. If a local arrival column
--     exists (e.g. ACTL_LOCAL_ARVL_TM) use it, otherwise arrival delay mixes local and non-local time.
--   * SCHD_ARVL_DT/TM is often 1-3 days later than the team's published schedule (Report Card 'Schedule' tab),
--     which hides late arrivals. Confirm which field holds the published ETA.

SELECT
    A.MKT_FLIGHT_NBR,
    A.ITIN_FLIGHT_LEG_SEQ_NBR,
    A.FLIGHT_LEG_SEQ_NBR,
    A.SUBFLEET_CD,
    A.DEP_AIRPRT_IATA_CD,
    A.ARVL_AIRPRT_IATA_CD,
    A.SCHD_DEP_DT,
    A.SCHD_DEP_TM,
    CAST(A.SCHD_DEP_DT  AS TIMESTAMP(0)) + (A.SCHD_DEP_TM       - TIME '00:00:00' HOUR TO SECOND) AS SCHD_DEP_DTTM,
    CAST(A.ACTL_DEP_DT  AS TIMESTAMP(0)) + (A.ACTL_LOCAL_DEP_TM - TIME '00:00:00' HOUR TO SECOND) AS ACTL_DEP_DTTM,
    CAST(A.SCHD_ARVL_DT AS TIMESTAMP(0)) + (A.SCHD_ARVL_TM      - TIME '00:00:00' HOUR TO SECOND) AS SCHD_ARVL_DTTM,
    CAST(A.ACTL_ARVL_DT AS TIMESTAMP(0)) + (A.ACTL_ARVL_TM      - TIME '00:00:00' HOUR TO SECOND) AS ACTL_ARVL_DTTM
FROM <schema>.FLIGHT_LEG_TABLE A
WHERE A.ROW_ACTV_IND = 'Y'
  AND A.FLIGHT_LEG_IND = 'Y'
  AND A.SUBFLEET_CD IN ('COND','CONI','DRYD','DRYI','MALD','MALI','MEGA','REFD','REFI','REFR','RFS','ROLD','ROLI','XFRD')
  AND A.SCHD_DEP_DT BETWEEN CURRENT_DATE - 120 AND CURRENT_DATE + 1      -- :days_back
QUALIFY ROW_NUMBER() OVER (
        PARTITION BY A.MKT_FLIGHT_NBR, A.ITIN_FLIGHT_LEG_SEQ_NBR, A.FLIGHT_LEG_SEQ_NBR
        ORDER BY A.SOURCE_DATA_LAST_UPDT_GMT_TMS DESC) = 1;
