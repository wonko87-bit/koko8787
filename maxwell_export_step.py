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

def msg(text):
    oDesktop.AddMessage("", "", 0, str(text))

def sanitize_filename(name):
    return re.sub(r'[\\/:*?"<>|]', '_', name)


# ---------------------------------------------------------
# Main
# ---------------------------------------------------------

oProject = oDesktop.GetActiveProject()
if oProject is None:
    msg("ERROR: No active project found.")
    raise RuntimeError("No active project found.")

oDesign = oProject.GetActiveDesign()
if oDesign is None:
    msg("ERROR: No active design found.")
    raise RuntimeError("No active design found.")

msg("Project : " + oProject.GetName())
msg("Design  : " + oDesign.GetName())
msg("DesignType: " + str(oDesign.GetDesignType()))

oEditor = oDesign.SetActiveEditor("3D Modeler")
msg("Editor  : " + str(oEditor))

# Try multiple methods to find objects
all_objects = []

# Method 1: wildcard
try:
    r1 = list(oEditor.GetMatchedObjectName("*"))
    msg("[Method1] GetMatchedObjectName(*): " + str(len(r1)) + " -> " + str(r1[:5]))
    if r1:
        all_objects = r1
except Exception as e:
    msg("[Method1] FAIL: " + str(e))

# Method 2: group names
for grp in ["Model", "Solids", "Sheets", "Lines", "Unclassified"]:
    try:
        r = list(oEditor.GetObjectsInGroup(grp))
        msg("[Method2] GetObjectsInGroup('{}')  : {}".format(grp, r))
        if r and not all_objects:
            all_objects = r
    except Exception as e:
        msg("[Method2] GetObjectsInGroup('{}') FAIL: {}".format(grp, str(e)))

msg("Total usable objects: " + str(len(all_objects)))

if not all_objects:
    msg("No objects found. Make sure geometry exists in the 3D Modeler before running this script.")
else:
    try:
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
            msg("Created output directory: " + OUTPUT_DIR)
        else:
            msg("Output directory exists: " + OUTPUT_DIR)
    except Exception as e:
        msg("ERROR creating directory: " + str(e))
        raise

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
            msg("  OK  : " + obj_name + " -> " + os.path.basename(filepath))
            success.append(obj_name)
        except Exception as e:
            msg("  FAIL: " + obj_name + " | " + str(e))
            failed.append(obj_name)

    msg("=" * 50)
    msg("Done.  Success: {}  /  Failed: {}  /  Total: {}".format(
        len(success), len(failed), len(all_objects)))
    if failed:
        msg("Failed objects:")
        for f in failed:
            msg("  - " + f)
    msg("=" * 50)
