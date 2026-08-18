"""
Maxwell 3D - STEP Export Script Generator

Run this script ONCE in Maxwell to generate individual export scripts
(one per solid object) in SCRIPTS_DIR.

Then run each generated script one at a time from Maxwell's script runner.
This avoids the Maxwell STEP exporter crash that occurs on consecutive exports.
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import os
import re

# ---------------------------------------------------------
# Settings
# ---------------------------------------------------------

# Where the exported .step files will be saved (used in generated scripts)
OUTPUT_DIR  = r"C:\Maxwell_STEP_Export"

# Where the generated per-object scripts will be saved
SCRIPTS_DIR = r"C:\Maxwell_STEP_Export\scripts"

# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------

def msg(text):
    oDesktop.AddMessage("", "", 0, str(text))

def sanitize_filename(name):
    return re.sub(r'[\\/:*?"<>|]', '_', name)


SCRIPT_TEMPLATE = '''
import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
import os

OBJ_NAME  = {obj_name_repr}
FILEPATH  = {filepath_repr}
PROJECT   = {project_repr}
DESIGN    = {design_repr}

def msg(text):
    oDesktop.AddMessage("", "", 0, str(text))

oProject = oDesktop.SetActiveProject(PROJECT)
oDesign  = oProject.SetActiveDesign(DESIGN)
oEditor  = oDesign.SetActiveEditor("3D Modeler")

if not os.path.exists(os.path.dirname(FILEPATH)):
    os.makedirs(os.path.dirname(FILEPATH))

try:
    oEditor.Export(
        [
            "NAME:ExportParameters",
            "AllowRegionDependentPartSelectionForPMLCreation:=", True,
            "AllowRegionSelectionForPMLCreation:=",              True,
            "Selections:=",    OBJ_NAME,
            "File Name:=",     FILEPATH,
            "Major Version:=", -1,
            "Minor Version:=", -1,
        ]
    )
    msg("OK: " + OBJ_NAME + " -> " + FILEPATH)
except Exception as e:
    msg("FAIL: " + str(e))
'''

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

project_name = oProject.GetName()
design_name  = oDesign.GetName()

try:
    solid_objects = list(oEditor.GetObjectsInGroup("Solids"))
except Exception as e:
    msg("ERROR getting solids: " + str(e))
    solid_objects = []

msg("Solid objects found: " + str(len(solid_objects)))

if not solid_objects:
    msg("No solid objects found. Exiting.")
else:
    for d in [OUTPUT_DIR, SCRIPTS_DIR]:
        if not os.path.exists(d):
            os.makedirs(d)

    for i, obj_name in enumerate(solid_objects):
        safe_name  = sanitize_filename(obj_name)
        step_path  = os.path.join(OUTPUT_DIR, safe_name + ".step")
        script_path = os.path.join(SCRIPTS_DIR,
                                   "{:03d}_{}.py".format(i + 1, safe_name))

        code = SCRIPT_TEMPLATE.format(
            obj_name_repr = repr(obj_name),
            filepath_repr = repr(step_path),
            project_repr  = repr(project_name),
            design_repr   = repr(design_name),
        )

        with open(script_path, "w") as f:
            f.write(code)

        msg("Generated: " + os.path.basename(script_path))

    msg("=" * 50)
    msg("Generated {} scripts in:".format(len(solid_objects)))
    msg("  " + SCRIPTS_DIR)
    msg("Run each script ONE AT A TIME from Maxwell script runner.")
    msg("=" * 50)
