import os
import pandas as pd
from src.config import INPUT_DIR, PYXIS_PATH, EPIC_ADMIN_PATH, EPIC_FLOWSHEET_PATH

def consolidate_input_files():
    """
    Scans the Input directory for source files, classifies them, merges them,
    deduplicates rows, and saves the consolidated files to the paths expected
    by the reconciliation pipeline.
    """
    print("=" * 60)
    print("CONSOLIDATING INPUT FILES")
    print("=" * 60)

    # Resolve target filenames to exclude them from the source files list
    target_pyxis_name = os.path.basename(PYXIS_PATH).lower()
    target_admin_name = os.path.basename(EPIC_ADMIN_PATH).lower()
    target_flow_name = os.path.basename(EPIC_FLOWSHEET_PATH).lower()

    pyxis_sources = []
    admin_sources = []
    flow_sources = []

    # Scan files in input directory
    if not os.path.exists(INPUT_DIR):
        print(f"ERROR: Input directory does not exist at {INPUT_DIR}")
        return

    for filename in os.listdir(INPUT_DIR):
        filepath = os.path.join(INPUT_DIR, filename)
        if not os.path.isfile(filepath):
            continue

        filename_lower = filename.lower()

        # Skip target files to avoid recursive consolidation
        if filename_lower in [target_pyxis_name, target_admin_name, target_flow_name]:
            continue

        # Classify files based on extensions and keywords
        if filename_lower.endswith('.csv'):
            # Pyxis files are CSVs (e.g. "CS continuous infusions - weekly report...")
            pyxis_sources.append(filepath)
        elif filename_lower.endswith(('.xlsx', '.xls')):
            if any(kw in filename_lower for kw in ['admin', 'action', 'administration', 'administered']):
                admin_sources.append(filepath)
            elif any(kw in filename_lower for kw in ['flowsheet', 'volume', 'infusion']):
                flow_sources.append(filepath)
            else:
                # Fallback: classify based on columns of the file
                try:
                    df_temp = pd.read_excel(filepath, nrows=2)
                    cols = [c.lower() for c in df_temp.columns]
                    if 'flowsheet row name' in cols or 'order med id' in cols:
                        flow_sources.append(filepath)
                    elif 'administrationaction' in cols or 'order id' in cols:
                        admin_sources.append(filepath)
                except Exception as e:
                    print(f"Warning: Could not inspect columns for {filename}: {e}")

    # 1. Consolidate Pyxis
    if pyxis_sources:
        print(f"Consolidating {len(pyxis_sources)} Pyxis source file(s) into {PYXIS_PATH}:")
        pyxis_dfs = []
        for src in pyxis_sources:
            print(f"  - {os.path.basename(src)}")
            try:
                # Using low_memory=False to prevent mixed-type warning for large columns
                pyxis_dfs.append(pd.read_csv(src, low_memory=False))
            except Exception as e:
                print(f"  [ERROR] Failed to read {src}: {e}")

        if pyxis_dfs:
            try:
                df_pyxis = pd.concat(pyxis_dfs, ignore_index=True)
                initial_len = len(df_pyxis)
                df_pyxis = df_pyxis.drop_duplicates()
                print(f"  -> Total rows: {initial_len}, after removing duplicates: {len(df_pyxis)}")
                df_pyxis.to_csv(PYXIS_PATH, index=False)
                print(f"  -> Saved consolidated Pyxis data successfully.")
            except Exception as e:
                print(f"  [ERROR] Failed to concatenate/save Pyxis: {e}")
    else:
        if os.path.exists(PYXIS_PATH):
            print(f"No Pyxis source files found. Using existing file: {PYXIS_PATH}")
        else:
            print(f"WARNING: No Pyxis files found to consolidate and target does not exist.")

    # 2. Consolidate Epic Administrations (Admin)
    if admin_sources:
        print(f"\nConsolidating {len(admin_sources)} Epic Administrations file(s) into {EPIC_ADMIN_PATH}:")
        admin_dfs = []
        for src in admin_sources:
            print(f"  - {os.path.basename(src)}")
            try:
                # Workbook contains no default style warnings are caught by openpyxl internally
                admin_dfs.append(pd.read_excel(src))
            except Exception as e:
                print(f"  [ERROR] Failed to read {src}: {e}")

        if admin_dfs:
            try:
                df_admin = pd.concat(admin_dfs, ignore_index=True)
                initial_len = len(df_admin)
                df_admin = df_admin.drop_duplicates()
                print(f"  -> Total rows: {initial_len}, after removing duplicates: {len(df_admin)}")
                df_admin.to_excel(EPIC_ADMIN_PATH, index=False)
                print(f"  -> Saved consolidated Administrations data successfully.")
            except Exception as e:
                print(f"  [ERROR] Failed to concatenate/save Administrations: {e}")
    else:
        if os.path.exists(EPIC_ADMIN_PATH):
            print(f"No Epic Administrations source files found. Using existing file: {EPIC_ADMIN_PATH}")
        else:
            print(f"WARNING: No Epic Administrations files found to consolidate and target does not exist.")

    # 3. Consolidate Epic Flowsheet
    if flow_sources:
        print(f"\nConsolidating {len(flow_sources)} Epic Flowsheet file(s) into {EPIC_FLOWSHEET_PATH}:")
        flow_dfs = []
        for src in flow_sources:
            print(f"  - {os.path.basename(src)}")
            try:
                flow_dfs.append(pd.read_excel(src))
            except Exception as e:
                print(f"  [ERROR] Failed to read {src}: {e}")

        if flow_dfs:
            try:
                df_flow = pd.concat(flow_dfs, ignore_index=True)
                initial_len = len(df_flow)
                df_flow = df_flow.drop_duplicates()
                print(f"  -> Total rows: {initial_len}, after removing duplicates: {len(df_flow)}")
                df_flow.to_excel(EPIC_FLOWSHEET_PATH, index=False)
                print(f"  -> Saved consolidated Flowsheet data successfully.")
            except Exception as e:
                print(f"  [ERROR] Failed to concatenate/save Flowsheet: {e}")
    else:
        if os.path.exists(EPIC_FLOWSHEET_PATH):
            print(f"No Epic Flowsheet source files found. Using existing file: {EPIC_FLOWSHEET_PATH}")
        else:
            print(f"WARNING: No Epic Flowsheet files found to consolidate and target does not exist.")
    print("=" * 60 + "\n")

if __name__ == "__main__":
    # Test consolidation if run directly
    consolidate_input_files()
