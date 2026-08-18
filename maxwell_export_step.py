"""
Maxwell 3D - Export all solid model objects as individual STEP files

Output directory is set by OUTPUT_DIR below.
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import os
import re
import System
from System import Array, Object

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

def make_array(*args):
    """Convert args to a .NET Object array for AEDT COM calls."""
    a = Array.CreateInstance(Object, len(args))
    for i, v in enumerate(args):
        a[i] = v
    return a


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

        # Attempt 1: .NET Object array
        try:
            params = make_array(
                "NAME:ExportParameters",
                "FileName:=",      filepath,
                "SelectionList:=", obj_name,
            )
            oEditor.Export(params)
            msg("  OK (NET array): " + obj_name)
            success.append(obj_name)
            exported = True
        except Exception as e:
            msg("  A fail: " + str(e)[:100])

        # Attempt 2: positional args (no wrapping)
        if not exported:
            try:
                oEditor.Export(
                    "NAME:ExportParameters",
                    "FileName:=",      filepath,
                    "SelectionList:=", obj_name,
                )
                msg("  OK (positional): " + obj_name)
                success.append(obj_name)
                exported = True
            except Exception as e:
                msg("  B fail: " + str(e)[:100])

        if not exported:
            failed.append(obj_name)
            msg("  FAIL: " + obj_name)

    msg("=" * 50)
    msg("Done.  Success: {}  /  Failed: {}  /  Total: {}".format(
        len(success), len(failed), len(solid_objects)))
    msg("=" * 50)
    msg("If all failed: use Maxwell script recorder to capture")
    msg("  Modeler -> Export on a selected object, then share the .py")
