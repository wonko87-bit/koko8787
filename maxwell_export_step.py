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

# Only solid objects can be exported as STEP
try:
    solid_objects = list(oEditor.GetObjectsInGroup("Solids"))
except Exception as e:
    msg("ERROR getting solids: " + str(e))
    solid_objects = []

msg("Solid objects: " + str(solid_objects))

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

        exported = False
        for params in [
            # format A: list with SelectionList
            ["NAME:ExportParameters",
             "FileName:=",      filepath,
             "SelectionList:=", obj_name],
            # format B: tuple with SelectionList
            ("NAME:ExportParameters",
             "FileName:=",      filepath,
             "SelectionList:=", obj_name),
            # format C: no SelectionList (export after setting selection)
            ["NAME:ExportParameters",
             "FileName:=",      filepath],
        ]:
            try:
                if isinstance(params, list) and "SelectionList:=" not in params:
                    # Pre-select object before exporting
                    oEditor.SetSelectionList(
                        ["NAME:Selections",
                         "Selections:=", obj_name,
                         "NewPartsModelFlag:=", "Model"]
                    )
                oEditor.Export(params)
                msg("  OK  : " + obj_name + " (" + str(type(params).__name__) + ")")
                success.append(obj_name)
                exported = True
                break
            except Exception as e:
                msg("  try failed (" + str(type(params).__name__) + "): " + str(e)[:80])

        if not exported:
            failed.append(obj_name)
            msg("  FAIL: " + obj_name)

    msg("=" * 50)
    msg("Done.  Success: {}  /  Failed: {}  /  Total: {}".format(
        len(success), len(failed), len(solid_objects)))
    msg("=" * 50)
