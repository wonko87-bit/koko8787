"""
Maxwell 3D - Export all model objects as individual STEP files

Each object in the active design is exported to a separate .step file.
Output directory is set by OUTPUT_DIR below.
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import os
import re

# ---------------------------------------------------------
# Settings
# ---------------------------------------------------------

# Output folder - change to your preferred path
OUTPUT_DIR = r"C:\Maxwell_STEP_Export"

# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------

def sanitize_filename(name):
    """Replace characters that are invalid in Windows filenames."""
    return re.sub(r'[\\/:*?"<>|]', '_', name)


# ---------------------------------------------------------
# Main
# ---------------------------------------------------------

oProject = oDesktop.GetActiveProject()
if oProject is None:
    raise RuntimeError("No active project found.")

oDesign = oProject.GetActiveDesign()
if oDesign is None:
    raise RuntimeError("No active design found.")

oEditor = oDesign.SetActiveEditor("3D Modeler")

# Get all solid model objects
all_objects = list(oEditor.GetObjectsInGroup("Model"))
print("=" * 60)
print("Maxwell STEP Exporter")
print("=" * 60)
print("Design  : {}".format(oDesign.GetName()))
print("Objects : {}".format(len(all_objects)))
print("Output  : {}".format(OUTPUT_DIR))
print()

if not all_objects:
    print("No model objects found. Exiting.")
else:
    # Create output directory if it does not exist
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print("Created output directory: {}".format(OUTPUT_DIR))

    success = []
    failed  = []

    for obj_name in all_objects:
        safe_name = sanitize_filename(obj_name)
        filepath  = os.path.join(OUTPUT_DIR, safe_name + ".step")

        try:
            oEditor.Export(
                [
                    "NAME:ExportParameters",
                    "FileName:=",      filepath,
                    "SelectionList:=", obj_name,
                ]
            )
            print("  OK  : {} -> {}".format(obj_name, os.path.basename(filepath)))
            success.append(obj_name)
        except Exception as e:
            print("  FAIL: {} | {}".format(obj_name, str(e)))
            failed.append(obj_name)

    print()
    print("=" * 60)
    print("Done.  Success: {}  /  Failed: {}  /  Total: {}".format(
        len(success), len(failed), len(all_objects)))
    if failed:
        print("Failed objects:")
        for f in failed:
            print("  - {}".format(f))
    print("=" * 60)
