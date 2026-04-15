"""
run.py
======
BMS Attribution Model — Entry Point
Panta Contracting Limited

Usage:
    python run.py                    # process all projects in settings.py
    python run.py J01577             # process one project by job number
    python run.py J01578             # auto-discovers projects/J01578/ if not in settings.py

Workflow for new projects:
    1. Create projects/J01578/ (folder name = job number)
    2. Drop estimate xlsx files into the folder
    3. Run: python run.py J01578
    The project is auto-detected, added to settings.py, and run immediately.

Output filename format:
    BMS_Attribution_<JobNumber>_<ProjectSlug>.xlsx
"""

import sys
import os
from datetime import date

# Force UTF-8 output on Windows (avoids cp1252 UnicodeEncodeError for €, —, etc.)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# Add project root to path so imports work regardless of where run.py is called from
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config.settings import PROJECTS, OUTPUT_FILENAME_TEMPLATE, PROJECTS_DIR
from core.attribution import run_project
from core.discover import discover_and_register, discover_all_folders
from outputs.excel_builder import build_workbook

# Absolute path to settings.py — needed by discover_and_register
_SETTINGS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               "config", "settings.py")


def slugify(text: str) -> str:
    """Convert project name to a safe filename slug."""
    return (text.lower()
               .replace(" ", "_")
               .replace("—", "")
               .replace("/", "_")
               .replace("\\", "_")
               .replace("(", "")
               .replace(")", "")
               .strip("_"))


def main():
    # Optional: filter to a single job number passed as argument
    filter_job = sys.argv[1] if len(sys.argv) > 1 else None

    projects_to_run = [
        p for p in PROJECTS
        if filter_job is None or p["job_number"] == filter_job
    ]

    # Bulk-discover: when no job number is given, scan projects/ for any folders
    # not yet registered and add them to settings.py before running.
    if filter_job is None:
        registered_folders = {p["folder"] for p in PROJECTS}
        new_projects = discover_all_folders(registered_folders, PROJECTS_DIR, _SETTINGS_PATH)
        projects_to_run = projects_to_run + new_projects

    # Auto-discover: if a specific job was requested but isn't in settings.py yet,
    # scan the projects/<job_number> folder and register it automatically.
    if not projects_to_run and filter_job:
        try:
            discovered = discover_and_register(filter_job, PROJECTS_DIR, _SETTINGS_PATH)
            projects_to_run = [discovered]
        except FileNotFoundError as e:
            print(f"\n[ERROR] {e}")
            print(f"Available job numbers: {[p['job_number'] for p in PROJECTS]}")
            sys.exit(1)

    if not projects_to_run:
        print(f"No projects found matching: {filter_job}")
        print(f"Available job numbers: {[p['job_number'] for p in PROJECTS]}")
        sys.exit(1)

    print(f"\n{'='*60}")
    print(f"BMS Attribution Model — v3.1")
    print(f"Panta Contracting Limited")
    print(f"Run date: {date.today().strftime('%d %B %Y')}")
    print(f"Projects to process: {len(projects_to_run)}")
    print(f"{'='*60}")

    for project in projects_to_run:
        print(f"\nProject: {project['project_name']} ({project['job_number']})")
        print(f"Folder:  {os.path.join(PROJECTS_DIR, project['folder'])}")

        # Run the attribution scan
        result = run_project(project)

        if not result["files"]:
            print(f"  [SKIP] No files processed for this project.")
            continue

        # Build the output filename
        slug = slugify(project["project_name"])
        filename = OUTPUT_FILENAME_TEMPLATE.format(
            job_number=project["job_number"],
            project_slug=slug
        )
        output_path = os.path.join(PROJECTS_DIR, project["folder"], filename)

        # Build the Excel workbook
        build_workbook(result, output_path)

        # Print summary
        all_in_scope = [r for r in result["all_clusters"] if not r.get("is_maintenance")]
        total_rev  = sum(
            f["proj_revenue"] for f in result["files"]
            if f["proj_revenue"] is not None
        )
        total_bms  = sum(r["sell_total"]    for r in all_in_scope)
        total_hrs  = sum(r["lab_hrs_total"] for r in all_in_scope)
        total_gp   = sum(r["gross_profit"]  for r in all_in_scope)
        bms_pct    = total_bms / total_rev * 100 if total_rev else 0
        margin     = total_gp  / total_bms * 100 if total_bms else 0

        print(f"\n  -- Summary ------------------------------------------")
        print(f"  BMS clusters found:  {len(all_in_scope)}")
        print(f"  Project revenue:     €{total_rev:>12,.2f}")
        print(f"  BMS revenue:         €{total_bms:>12,.2f}  ({bms_pct:.1f}%)")
        print(f"  BMS labour hours:    {total_hrs:>12,.1f} hrs")
        print(f"  BMS gross profit:    €{total_gp:>12,.2f}  ({margin:.1f}% margin)")
        if project.get("notes"):
            print(f"\n  Notes: {project['notes'][:100]}")

    print(f"\n{'='*60}")
    print("Done.")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
