"""
core/discover.py
================
Auto-discovers a new project folder and registers it in config/settings.py.

Workflow:
    1. User creates projects/J01578/ and drops estimate xlsx files into it
    2. User runs: python run.py J01578
    3. run.py detects J01578 is not in settings.py, calls discover_and_register()
    4. discover_and_register() reads the files, builds the project definition,
       appends it to settings.py, and returns the project dict for immediate use

File type detection (from filename prefix):
    ELM* -> "ELM"
    E*   -> "EL"

Project name: read from ML_Sum sheet, row 1, col 3.
"""

import os
import pandas as pd


_BMS_ML_CODES = {
    "410", "411", "412", "413", "414", "415",
    "416", "417", "418", "419", "420", "421",
}


def _detect_file_type(filepath: str, filename_upper: str) -> str:
    """
    Determine the scanner type for an estimate file.

    Priority:
    1. Has 'TotalCosts (Hrs)' sheet       → BMS_TC  (standalone BMS panel/SJ estimate)
    2. Has 'ML_Sum' + any 410-421 code    → BMS_ML  (standalone BMS project estimate)
    3. Has 'ML_Sum', filename starts ELM  → ELM
    4. Has 'ML_Sum'                       → EL
    """
    try:
        sheet_names = pd.ExcelFile(filepath).sheet_names
    except Exception:
        # Can't open — fall back to filename prefix
        return "ELM" if filename_upper.startswith("ELM") else "EL"

    if "TotalCosts (Hrs)" in sheet_names:
        return "BMS_TC"

    if "ML_Sum" in sheet_names and "Elec_Est" in sheet_names:
        # Check whether any BMS 410-421 codes appear in Elec_Est
        try:
            df = pd.read_excel(filepath, sheet_name="Elec_Est", header=None)
            for i in range(13, len(df)):
                cd = str(df.iloc[i, 7]).strip()
                if cd in _BMS_ML_CODES:
                    return "BMS_ML"
        except Exception:
            pass
        return "ELM" if filename_upper.startswith("ELM") else "EL"

    return "ELM" if filename_upper.startswith("ELM") else "EL"


def scan_project_folder(job_number: str, folder_path: str) -> dict:
    """
    Scan a project folder and return a settings.py-compatible project dict.

    Reads project name from ML_Sum of the first readable Excel file.
    Determines EL/ELM from filename prefix.
    Files are sorted: EL files first, then ELM.
    """
    xlsx_files = sorted(
        f for f in os.listdir(folder_path)
        if f.lower().endswith(".xlsx") and not f.startswith("BMS_Attribution_")
    )

    if not xlsx_files:
        raise FileNotFoundError(f"No xlsx files found in {folder_path}")

    # Sort: EL files before ELM
    def sort_key(fn):
        return (0 if not fn.upper().startswith("ELM") else 1, fn)
    xlsx_files.sort(key=sort_key)

    file_defs = []
    project_name = None

    for filename in xlsx_files:
        filepath = os.path.join(folder_path, filename)
        fn_upper = filename.upper()

        # Detect file type from sheet names, falling back to filename prefix
        file_type = _detect_file_type(filepath, fn_upper)

        # Build label from first two underscore-parts of filename
        # e.g. E25_0062_Quintano... -> E25-0062 (EL)
        #      ELM25_0085_...       -> ELM25-0085 (ELM)
        parts = filename.split("_")
        if len(parts) >= 2:
            label_base = f"{parts[0]}-{parts[1]}"
        else:
            label_base = os.path.splitext(filename)[0][:12]
        label = f"{label_base} ({file_type})"

        # Extract project name from ML_Sum (row 1, col 3)
        if project_name is None:
            try:
                df = pd.read_excel(filepath, sheet_name="ML_Sum", header=None)
                name_val = str(df.iloc[1, 3]).strip()
                if name_val and name_val != "nan":
                    project_name = name_val
            except Exception:
                pass

        file_defs.append({
            "path":  filename,
            "type":  file_type,
            "label": label,
        })

    if not project_name:
        project_name = job_number

    # Use the job number as the folder name (matches the directory they created)
    folder_name = os.path.basename(folder_path)

    return {
        "job_number":      job_number,
        "project_name":    project_name,
        "folder":          folder_name,
        "files":           file_defs,
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    }


def _format_project_entry(project: dict) -> str:
    """Format a project dict as a Python literal for insertion into settings.py."""
    files_lines = []
    for f in project["files"]:
        files_lines.append(
            f'            {{\n'
            f'                "path":  "{f["path"]}",\n'
            f'                "type":  "{f["type"]}",\n'
            f'                "label": "{f["label"]}",\n'
            f'            }},'
        )
    files_block = "\n".join(files_lines)

    return (
        f'    {{\n'
        f'        "job_number":      "{project["job_number"]}",\n'
        f'        "project_name":    "{project["project_name"]}",\n'
        f'        "folder":          "{project["folder"]}",\n'
        f'        "files": [\n'
        f'{files_block}\n'
        f'        ],\n'
        f'        "project_revenue": {{}},\n'
        f'        "manual_flags":    {{}},\n'
        f'        "notes":           "",\n'
        f'    }},\n'
    )


def append_to_settings(project: dict, settings_path: str) -> None:
    """
    Append a new project entry to config/settings.py, before the
    '# ── ADD NEW PROJECTS BELOW THIS LINE' comment block.
    """
    with open(settings_path, encoding="utf-8") as fh:
        content = fh.read()

    marker = "    # ── ADD NEW PROJECTS BELOW THIS LINE"
    if marker not in content:
        raise ValueError(f"Could not find insertion marker in {settings_path}")

    entry = _format_project_entry(project)
    new_content = content.replace(marker, entry + marker)

    with open(settings_path, "w", encoding="utf-8") as fh:
        fh.write(new_content)


def find_project_folder(job_number: str, projects_dir: str) -> str:
    """
    Find the project folder for a job number.

    Accepts both exact matches (projects/J01578) and prefixed names
    (projects/J01578_ST_Micro_Exhaust_Fans). Returns the full folder path.
    Raises FileNotFoundError if no match is found.
    """
    # Exact match first
    exact = os.path.join(projects_dir, job_number)
    if os.path.isdir(exact):
        return exact

    # Prefix match: J05226_anything
    prefix = job_number + "_"
    matches = [
        d for d in os.listdir(projects_dir)
        if d.startswith(prefix) and os.path.isdir(os.path.join(projects_dir, d))
    ]
    if len(matches) == 1:
        return os.path.join(projects_dir, matches[0])
    if len(matches) > 1:
        raise FileNotFoundError(
            f"Multiple folders found for {job_number}: {matches}\n"
            f"Rename them so only one matches."
        )

    raise FileNotFoundError(
        f"No folder found for {job_number} in {projects_dir}\n"
        f"Create projects/{job_number}_Project_Name/ and add your estimate xlsx files, then re-run."
    )


def discover_and_register(job_number: str, projects_dir: str, settings_path: str) -> dict:
    """
    Main entry point called from run.py when a job number is not found in settings.

    Scans the projects/J0XXXX_Name/ folder, builds the project definition,
    appends it to settings.py, and returns the project dict for immediate use.

    Raises FileNotFoundError if no matching folder exists.
    """
    folder_path = find_project_folder(job_number, projects_dir)

    print(f"  [discover] New project detected — scanning {folder_path}")
    project = scan_project_folder(job_number, folder_path)

    print(f"  [discover] Project name:  {project['project_name']}")
    print(f"  [discover] Files found:   "
          + ", ".join(f['label'] for f in project['files']))
    print(f"  [discover] Registering in settings.py...")

    append_to_settings(project, settings_path)

    print(f"  [discover] Done — {job_number} added to settings.py")
    return project


def discover_all_folders(registered_folders: set, projects_dir: str,
                         settings_path: str) -> list:
    """
    Scan every subfolder in projects_dir and register any that are not already
    in settings.py. Returns a list of newly discovered project dicts.

    Called by run.py when no job number argument is given, so that
    `python run.py` always processes every folder in projects/.

    Folders are identified by their directory name. If the folder name starts
    with a J0XXXX_ prefix that job number is used; otherwise the folder name
    itself is used as the job_number label.
    """
    import re
    new_projects = []

    all_folders = sorted(
        d for d in os.listdir(projects_dir)
        if os.path.isdir(os.path.join(projects_dir, d))
    )

    for folder_name in all_folders:
        if folder_name in registered_folders:
            continue  # already registered

        folder_path = os.path.join(projects_dir, folder_name)

        # Check there are xlsx files to process
        xlsx = [f for f in os.listdir(folder_path)
                if f.lower().endswith(".xlsx") and not f.startswith("BMS_Attribution_")]
        if not xlsx:
            continue

        # Derive job_number: use J0XXXX prefix if present, else folder name
        m = re.match(r'^(J\d+)_', folder_name)
        job_number = m.group(1) if m else folder_name

        print(f"\n  [discover] Unregistered folder: {folder_name}")
        try:
            project = scan_project_folder(job_number, folder_path)
            print(f"  [discover] Project name: {project['project_name']}")
            print(f"  [discover] Files: " + ", ".join(f['label'] for f in project['files']))
            append_to_settings(project, settings_path)
            print(f"  [discover] Registered as {job_number}")
            new_projects.append(project)
        except Exception as e:
            print(f"  [discover] WARNING: Could not register {folder_name}: {e}")

    return new_projects
