#!/usr/bin/env python3
"""
planner_to_notion.py
Migrates a Microsoft Planner Excel export to a Notion-ready import CSV.

Outputs (written alongside this script):
  notion_import.csv       — UTF-8-BOM, comma-delimited, ready for Notion CSV import
  verification_report.md  — issues grouped by type with row numbers and task names

Usage:
  pip install pandas openpyxl
  python planner_to_notion.py
"""

import csv
import sys
from datetime import date, datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas is required.  pip install pandas openpyxl")
    sys.exit(1)

# ── Configuration ─────────────────────────────────────────────────────────────

SOURCE_FILE  = r"C:\Users\guybaranyay\Downloads\Weekly_Plan.xlsx"
SOURCE_SHEET = "Tasks"
OUTPUT_DIR   = Path(__file__).parent          # same folder as this script
OUTPUT_CSV    = OUTPUT_DIR / "notion_import.csv"
OUTPUT_REPORT = OUTPUT_DIR / "verification_report.md"

# Team members in display order
ALL_SIX = ["Guy", "Raymond", "Victor", "Edward", "Mark C", "JVT"]
PERSON_LABELS = set(ALL_SIX)

# Planner "Assigned To" full-name → short label
NAME_MAP = {
    "Guy Baranyay":        "Guy",
    "Raymond Muscat":      "Raymond",
    "Mark Catania":        "Mark C",
    "Jasmine Vella Turner":"JVT",
    "Victor Vural":        "Victor",
    "Edward Deguara":      "Edward",
}

# Bucket → Work type
QUOTATION_BUCKETS   = {"Quotes/Tenders", "AQNs to Start"}
MAINTENANCE_BUCKETS = {"Breakdown Maint Works"}
ADMIN_BUCKETS       = {"Long Term - TO DO", "Orders", "Small Jobs"}

# Status label precedence (checked in order; first match wins)
STATUS_LABEL_PRECEDENCE = [
    ("Closing",                    "Closing"),
    ("Awaiting Client Response",   "Awaiting client"),
    ("Paused",                     "Paused"),
    ("In Progress",                "In progress"),
]

# Notion CSV column order
FIELDNAMES = [
    "Task name", "Project", "Priority", "Notes",
    "Created date", "Due date", "Checklist",
    "Status", "Assigned to", "This week",
    "Work type", "Quotation status",
]


# ── Helpers ───────────────────────────────────────────────────────────────────

def parse_labels(value) -> set:
    """Split a semicolon-or-comma separated labels string into a set."""
    if pd.isna(value) or not str(value).strip():
        return set()
    raw = str(value)
    sep = ";" if ";" in raw else ","
    return {p.strip() for p in raw.split(sep) if p.strip()}


def parse_names(value) -> list:
    """Split an Assigned To string into a list of names."""
    if pd.isna(value) or not str(value).strip():
        return []
    raw = str(value)
    sep = ";" if ";" in raw else ","
    return [p.strip() for p in raw.split(sep) if p.strip()]


def progress_to_status(raw) -> str:
    """Convert a Progress field value to a Notion status string."""
    if pd.isna(raw):
        return ""
    s = str(raw).strip()
    mapping = {
        "Completed":   "Completed",
        "In progress": "In progress",
        "Not started": "Not started",
    }
    if s in mapping:
        return mapping[s]
    # Numeric % completion (0 / 50 / 100)
    try:
        pct = int(float(s))
        if pct == 100:
            return "Completed"
        if pct > 0:
            return "In progress"
        return "Not started"
    except ValueError:
        return ""


def resolve_status(labels: set, progress_raw) -> str:
    """Apply label precedence first, then fall back to progress field."""
    for label, status in STATUS_LABEL_PRECEDENCE:
        if label in labels:
            return status
    return progress_to_status(progress_raw) or "Not started"


def resolve_assigned(labels: set, assigned_raw) -> list:
    """Merge person labels + Assigned To column; apply Whole Team / default rules later."""
    if "Whole Team" in labels:
        return ALL_SIX[:]

    from_labels   = {p for p in PERSON_LABELS if p in labels}
    from_assigned = {NAME_MAP[n] for n in parse_names(assigned_raw) if n in NAME_MAP}

    combined = from_labels | from_assigned
    return sorted(combined, key=lambda x: ALL_SIX.index(x) if x in ALL_SIX else 99), from_labels, from_assigned


def infer_work_type(bucket: str) -> str:
    if bucket in QUOTATION_BUCKETS:
        return "Quotation"
    if bucket in MAINTENANCE_BUCKETS:
        return "Maintenance"
    if bucket in ADMIN_BUCKETS:
        return "Admin"
    return "Project"


def resolve_quotation_status(labels: set, work_type: str) -> str:
    if work_type != "Quotation":
        return ""
    if "Backlog" in labels:
        return "Backlog"
    if "Update Required from Assignee" in labels:
        return "Update required"
    if "Quotation" in labels:
        return "Quotation"
    return ""


def fmt_date(val) -> str:
    """Return YYYY-MM-DD string or empty string."""
    if pd.isna(val):
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d") if isinstance(val, datetime) else val.isoformat()
    s = str(val).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return s  # unparseable — pass through as-is


def find_col(df, candidates):
    """Case-insensitive column lookup; returns first match or None."""
    lower = {c.lower(): c for c in df.columns}
    for name in candidates:
        if name.lower() in lower:
            return lower[name.lower()]
    return None


def safe_str(val, default="") -> str:
    if pd.isna(val):
        return default
    return str(val).strip()


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print(f"Reading: {SOURCE_FILE}")
    try:
        df = pd.read_excel(SOURCE_FILE, sheet_name=SOURCE_SHEET, dtype=str)
    except FileNotFoundError:
        print(f"ERROR: File not found — {SOURCE_FILE}")
        sys.exit(1)
    except Exception as exc:
        print(f"ERROR reading file: {exc}")
        sys.exit(1)

    print(f"  {len(df)} rows loaded.")
    print(f"  Columns detected: {list(df.columns)}\n")

    # Flexible column detection
    C = {
        "task":     find_col(df, ["Task Name", "Title"]),
        "bucket":   find_col(df, ["Bucket Name", "Bucket"]),
        "priority": find_col(df, ["Priority"]),
        "desc":     find_col(df, ["Description", "Notes"]),
        "created":  find_col(df, ["Created Date", "Created", "Start date"]),
        "due":      find_col(df, ["Due Date", "Due date"]),
        "checklist":find_col(df, ["Checklist Items", "Checklist"]),
        "labels":   find_col(df, ["Labels", "Label"]),
        "progress": find_col(df, ["Progress", "% Completion", "% Complete"]),
        "assigned": find_col(df, ["Assigned To", "Assigned to", "AssignedTo"]),
    }
    for key, val in C.items():
        if val is None:
            print(f"  WARNING: column '{key}' not found — will be blank in output.")

    # Verification report buckets
    rpt = {
        "conflicts":   [],   # label status vs progress field
        "merges":      [],   # Labels and Assigned To disagreed
        "defaults":    [],   # Guy assigned as default
        "checklists":  [],   # tasks with checklist items
        "overdue":     [],   # due in past, not completed
    }

    today = date.today()
    output_rows = []

    for idx, row in df.iterrows():
        row_num = idx + 2  # Excel row (1-based header + 1)

        def g(key):
            c = C[key]
            return safe_str(row[c]) if c is not None else ""

        task_name = g("task")
        bucket    = g("bucket")
        priority  = g("priority")
        desc      = g("desc")
        checklist = g("checklist")
        labels    = parse_labels(row[C["labels"]]   if C["labels"]   else None)
        prog_raw  = row[C["progress"]] if C["progress"] else None
        asgn_raw  = row[C["assigned"]] if C["assigned"] else None

        # ── Status ────────────────────────────────────────────────────────────
        status = resolve_status(labels, prog_raw)

        # Conflict check: label drove status but progress field says something different
        label_drove = any(lbl in labels for lbl, _ in STATUS_LABEL_PRECEDENCE)
        prog_status = progress_to_status(prog_raw)
        if label_drove and prog_status and prog_status != status:
            rpt["conflicts"].append({
                "row": row_num, "task": task_name,
                "label_status": status, "progress_status": prog_status,
            })

        # ── Assigned to ───────────────────────────────────────────────────────
        if "Whole Team" in labels:
            assigned = ALL_SIX[:]
            from_labels = from_assigned = set()
        else:
            from_labels   = {p for p in PERSON_LABELS if p in labels}
            from_assigned = {NAME_MAP[n] for n in parse_names(asgn_raw) if n in NAME_MAP}
            combined      = from_labels | from_assigned
            assigned      = sorted(combined, key=lambda x: ALL_SIX.index(x) if x in ALL_SIX else 99)

            if from_labels and from_assigned and from_labels != from_assigned:
                rpt["merges"].append({
                    "row": row_num, "task": task_name,
                    "labels": sorted(from_labels),
                    "assigned": sorted(from_assigned),
                })

        if not assigned:
            if status != "Completed":
                assigned = ["Guy"]
                rpt["defaults"].append({"row": row_num, "task": task_name})
            # else leave blank for completed tasks

        # ── Checklist ─────────────────────────────────────────────────────────
        if checklist:
            sep    = ";" if ";" in checklist else "\n"
            count  = len([x for x in checklist.split(sep) if x.strip()])
            rpt["checklists"].append({"row": row_num, "task": task_name, "count": count})

        # ── Dates ─────────────────────────────────────────────────────────────
        created_fmt = fmt_date(row[C["created"]]) if C["created"] else ""
        due_fmt     = fmt_date(row[C["due"]])     if C["due"]     else ""

        # Overdue check
        if due_fmt and status != "Completed":
            try:
                if datetime.strptime(due_fmt, "%Y-%m-%d").date() < today:
                    rpt["overdue"].append({"row": row_num, "task": task_name,
                                           "due": due_fmt, "status": status})
            except ValueError:
                pass

        # ── Work type & Quotation status ──────────────────────────────────────
        work_type        = infer_work_type(bucket)
        quotation_status = resolve_quotation_status(labels, work_type)

        # ── This week ─────────────────────────────────────────────────────────
        this_week = "TRUE" if "ThisWeek" in labels else "FALSE"

        output_rows.append({
            "Task name":        task_name,
            "Project":          bucket,
            "Priority":         priority,
            "Notes":            desc,
            "Created date":     created_fmt,
            "Due date":         due_fmt,
            "Checklist":        checklist,
            "Status":           status,
            "Assigned to":      ", ".join(assigned),
            "This week":        this_week,
            "Work type":        work_type,
            "Quotation status": quotation_status,
        })

    # ── Write CSV ─────────────────────────────────────────────────────────────
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(output_rows)
    print(f"Written: {OUTPUT_CSV}  ({len(output_rows)} rows)")

    # ── Write verification report ─────────────────────────────────────────────
    with open(OUTPUT_REPORT, "w", encoding="utf-8") as f:
        w = f.write

        w("# Notion Migration Verification Report\n\n")
        w(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  \n")
        w(f"Source: `{SOURCE_FILE}`  \n")
        w(f"Total rows migrated: **{len(output_rows)}**\n\n")
        w("---\n\n")

        # 1. Status conflicts
        w("## 1. Status Label Conflicts\n\n")
        w("A status label overrode the Progress field value. Verify the intended status.\n\n")
        if rpt["conflicts"]:
            w("| Row | Task Name | Label-derived Status | Progress Field Status |\n")
            w("|-----|-----------|----------------------|-----------------------|\n")
            for r in rpt["conflicts"]:
                w(f"| {r['row']} | {r['task']} | {r['label_status']} | {r['progress_status']} |\n")
        else:
            w("_No conflicts found._\n")
        w("\n")

        # 2. Assignment merges
        w("## 2. Assignment Source Disagreements (Merged)\n\n")
        w("The Labels column and the Assigned To column named different people — both sources were merged into the output.\n\n")
        if rpt["merges"]:
            w("| Row | Task Name | From Labels | From Assigned To |\n")
            w("|-----|-----------|-------------|------------------|\n")
            for r in rpt["merges"]:
                w(f"| {r['row']} | {r['task']} | {', '.join(r['labels'])} | {', '.join(r['assigned'])} |\n")
        else:
            w("_No disagreements found._\n")
        w("\n")

        # 3. Default owner
        w("## 3. Default Owner Applied — Guy\n\n")
        w("Active tasks with no assignee in either Labels or Assigned To were defaulted to Guy.\n\n")
        if rpt["defaults"]:
            w("| Row | Task Name |\n")
            w("|-----|-----------|\n")
            for r in rpt["defaults"]:
                w(f"| {r['row']} | {r['task']} |\n")
        else:
            w("_No default owner assignments._\n")
        w("\n")

        # 4. Checklist items
        w("## 4. Checklist Items — Manual Review Required in Notion\n\n")
        w(f"{len(rpt['checklists'])} task(s) contain checklist items. "
          "These are imported as plain text in the Checklist column — "
          "recreate them as proper Notion checklist items manually after import.\n\n")
        if rpt["checklists"]:
            w("| Row | Task Name | Item Count |\n")
            w("|-----|-----------|------------|\n")
            for r in rpt["checklists"]:
                w(f"| {r['row']} | {r['task']} | {r['count']} |\n")
        else:
            w("_No checklist items found._\n")
        w("\n")

        # 5. Overdue tasks
        w("## 5. Overdue Tasks\n\n")
        w("Due date is in the past and Status ≠ Completed.\n\n")
        if rpt["overdue"]:
            w("| Row | Task Name | Due Date | Status |\n")
            w("|-----|-----------|----------|--------|\n")
            for r in rpt["overdue"]:
                w(f"| {r['row']} | {r['task']} | {r['due']} | {r['status']} |\n")
        else:
            w("_No overdue tasks._\n")
        w("\n")

        # Summary
        w("---\n\n## Summary\n\n")
        w(f"| Check | Count |\n")
        w(f"|-------|-------|\n")
        w(f"| Status conflicts resolved | {len(rpt['conflicts'])} |\n")
        w(f"| Assignment merges | {len(rpt['merges'])} |\n")
        w(f"| Default owner (Guy) applied | {len(rpt['defaults'])} |\n")
        w(f"| Tasks with checklist items | {len(rpt['checklists'])} |\n")
        w(f"| Overdue active tasks | {len(rpt['overdue'])} |\n")

    print(f"Written: {OUTPUT_REPORT}")
    print("\nDone. Review verification_report.md before importing to Notion.")


if __name__ == "__main__":
    main()
