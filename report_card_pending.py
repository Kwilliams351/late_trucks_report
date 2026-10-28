"""Shared Report Card classification for Python caches and live Excel formulas."""
import pandas as pd

CARD_STATUSES = [
    "Ready to log",
    "Pending - arrival overdue",
    "Pending - projected late",
    "Pending - in transit",
    "Pending - no departure",
    "Check record",
]


def card_status(df: pd.DataFrame, threshold: int) -> pd.Series:
    status = pd.Series("", index=df.index, dtype=object)
    departed = df["dep_status"].isin(["Late", "On time"])
    awaiting = departed & df["act_arr"].isna() & (df["arr_rule"] != "None")
    status[df["late"]] = CARD_STATUSES[0]
    status[awaiting] = CARD_STATUSES[3]
    status[awaiting & (df["proj_arr_delay"] > threshold).fillna(False)] = CARD_STATUSES[2]
    status[awaiting & (df["arr_status"] == "No actual")] = CARD_STATUSES[1]
    status[df["act_dep"].isna()] = CARD_STATUSES[4]
    status[(df["dep_status"] == "Check date") | (df["arr_status"] == "Check date")] = CARD_STATUSES[5]
    return status


def card_status_formula(ref) -> str:
    """ref resolves a Data header to a this-row structured reference."""
    return (
        f'=IF(OR({ref("Dep Status")}="Check date",{ref("Arr Status")}="Check date"),"Check record",'
        f'IF({ref("ATD")}="","Pending - no departure",'
        f'IF(AND({ref("ATA")}="",{ref("Arr Rule")}<>"None",'
        f'OR({ref("Dep Status")}="Late",{ref("Dep Status")}="On time")),'
        f'IF({ref("Arr Status")}="No actual","Pending - arrival overdue",'
        f'IF(N({ref("Proj Arr Delay (min)")})>Threshold,"Pending - projected late","Pending - in transit")),'
        f'IF({ref("Late Any")}=1,"Ready to log",""))))'
    )
