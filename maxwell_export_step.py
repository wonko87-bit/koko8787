"""
Maxwell 3D - Export all solid model objects as individual STEP files

Output directory is set by OUTPUT_DIR below.
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

OUTPUT_DIR   = r"C:\Maxwell_STEP_Export"
EXPORT_DELAY = 1.0   # seconds between exports

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

    for i, obj_name in enumerate(solid_objects):
        safe_name = sanitize_filename(obj_name)
        filepath  = os.path.join(OUTPUT_DIR, safe_name + ".step")

        msg("[{}/{}] {}".format(i + 1, len(solid_objects), obj_name))

        # Re-acquire COM references before every export to reset internal state
        oProject = oDesktop.GetActiveProject()
        oDesign  = oProject.GetActiveDesign()
        oEditor  = oDesign.SetActiveEditor("3D Modeler")

        # Save project to flush Maxwell internal buffers
        try:
            oProject.Save()
        except Exception:
            pass

        time.sleep(EXPORT_DELAY)

        try:
            oEditor.Export(
                [
                    "NAME:ExportParameters",
                    "AllowRegionDependentPartSelectionForPMLCreation:=", True,
                    "AllowRegionSelectionForPMLCreation:=",              True,
                    "Selections:=",    obj_name,
                    "File Name:=",     filepath,
                    "Major Version:=", -1,
                    "Minor Version:=", -1,
                ]
            )
            msg("  OK")
            success.append(obj_name)
        except Exception as e:
            msg("  FAIL: " + str(e)[:100])
            failed.append(obj_name)

    msg("=" * 50)
    msg("Done.  Success: {}  /  Failed: {}  /  Total: {}".format(
        len(success), len(failed), len(solid_objects)))
    if failed:
        for f in failed:
            msg("  FAIL: " + f)
    msg("=" * 50)
