"""
Maxwell 3D - Export solid model objects as STEP file(s)

MODE = "all"        -> all solids in one STEP file (stable)
MODE = "individual" -> one STEP file per solid (may crash on complex models)
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import os
import re
import time

# ---------------------------------------------------------
# Settings
# ---------------------------------------------------------

OUTPUT_DIR  = r"C:\Maxwell_STEP_Export"
MODE        = "all"   # "all" or "individual"
EXPORT_DELAY = 2.0    # seconds between exports (individual mode only)

# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------

def msg(text):
    oDesktop.AddMessage("", "", 0, str(text))

def sanitize_filename(name):
    return re.sub(r'[\\/:*?"<>|]', '_', name)

def do_export(selections_str, filepath):
    oEditor.Export(
        [
            "NAME:ExportParameters",
            "AllowRegionDependentPartSelectionForPMLCreation:=", True,
            "AllowRegionSelectionForPMLCreation:=",              True,
            "Selections:=",    selections_str,
            "File Name:=",     filepath,
            "Major Version:=", -1,
            "Minor Version:=", -1,
        ]
    )

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
msg("Mode: " + MODE)

if not solid_objects:
    msg("No solid objects found. Exiting.")
else:
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        msg("Created: " + OUTPUT_DIR)

    if MODE == "all":
        # Export all solids in one STEP file
        selections_str = ",".join(solid_objects)
        filepath = os.path.join(OUTPUT_DIR, "all_solids.step")
        try:
            do_export(selections_str, filepath)
            msg("OK: all_solids.step  ({} objects)".format(len(solid_objects)))
        except Exception as e:
            msg("FAIL: " + str(e))

    else:  # individual
        success = []
        failed  = []
        for i, obj_name in enumerate(solid_objects):
            safe_name = sanitize_filename(obj_name)
            filepath  = os.path.join(OUTPUT_DIR, safe_name + ".step")
            msg("[{}/{}] {}".format(i + 1, len(solid_objects), obj_name))
            try:
                do_export(obj_name, filepath)
                msg("  OK")
                success.append(obj_name)
            except Exception as e:
                msg("  FAIL: " + str(e)[:100])
                failed.append(obj_name)
            time.sleep(EXPORT_DELAY)

        msg("=" * 50)
        msg("Done.  Success: {}  /  Failed: {}  /  Total: {}".format(
            len(success), len(failed), len(solid_objects)))
        if failed:
            for f in failed:
                msg("  FAIL: " + f)
        msg("=" * 50)
