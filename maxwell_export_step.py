"""
Maxwell 3D - Export all solid model objects as individual STEP files

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

oEditor = oDesign.SetActiveEditor("3D Modeler")

try:
    solid_objects = list(oEditor.GetObjectsInGroup("Solids"))
except Exception as e:
    msg("ERROR getting solids: " + str(e))
    solid_objects = []

msg("Solid objects found: " + str(len(solid_objects)))

if not solid_objects:
    msg("No solid objects found. Exiting.")
else:
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        msg("Created: " + OUTPUT_DIR)

    success = []
    failed  = []

    for obj_name in solid_objects:
        safe_name = sanitize_filename(obj_name)
        filepath  = os.path.join(OUTPUT_DIR, safe_name + ".step")

        try:
            oEditor.Export(
                [
                    "NAME:ExportParameters",
                    "AllowRegionDependentPartSelectionForPMLCreation:=", True,
                    "AllowRegionSelectionForPMLCreation:=",              True,
                    "Selections:=",   obj_name,
                    "File Name:=",    filepath,
                    "Major Version:=", -1,
                    "Minor Version:=", -1,
                ]
            )
            msg("  OK  : " + obj_name)
            success.append(obj_name)
        except Exception as e:
            msg("  FAIL: " + obj_name + " | " + str(e)[:100])
            failed.append(obj_name)

    msg("=" * 50)
    msg("Done.  Success: {}  /  Failed: {}  /  Total: {}".format(
        len(success), len(failed), len(solid_objects)))
    if failed:
        msg("Failed objects:")
        for f in failed:
            msg("  - " + f)
    msg("=" * 50)
