"""
Maxwell 3D - STEP Export Automation

STEP 1: Run THIS script once inside Maxwell.
        It generates per-object export scripts + run_all.bat in SCRIPTS_DIR.

STEP 2: Close Maxwell (optional but recommended).

STEP 3: Double-click run_all.bat.
        Maxwell starts fresh for each object, exports it, then exits.
        No crash from consecutive exports.

Set MAXWELL_EXE to your actual Maxwell installation path if needed.
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import os
import re

# ---------------------------------------------------------
# Settings
# ---------------------------------------------------------

OUTPUT_DIR  = r"C:\Maxwell_STEP_Export"
SCRIPTS_DIR = r"C:\Maxwell_STEP_Export\scripts"

# Maxwell 2021 R1 default path - adjust if installed elsewhere
MAXWELL_EXE = r"C:\Program Files\AnsysEM\AnsysEM21.1\Win64\ansysedt.exe"

# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------

def msg(text):
    oDesktop.AddMessage("", "", 0, str(text))

def sanitize_filename(name):
    return re.sub(r'[\\/:*?"<>|]', '_', name)


PER_OBJECT_SCRIPT = """import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
import os

PROJECT_FILE = {project_file_repr}
DESIGN_NAME  = {design_repr}
OBJ_NAME     = {obj_name_repr}
FILEPATH     = {filepath_repr}

def msg(text):
    oDesktop.AddMessage("", "", 0, str(text))

oProject = oDesktop.OpenProject(PROJECT_FILE)
if oProject is None:
    raise RuntimeError("Failed to open project: " + PROJECT_FILE)

oDesign = oProject.SetActiveDesign(DESIGN_NAME)
oEditor = oDesign.SetActiveEditor("3D Modeler")

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
    msg("OK: " + OBJ_NAME)
except Exception as e:
    msg("FAIL: " + OBJ_NAME + " | " + str(e))

oProject.Close()
"""

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
project_path = oProject.GetPath()
project_file = os.path.join(project_path, project_name + ".aedt")

msg("Project file: " + project_file)

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

    script_paths = []

    for i, obj_name in enumerate(solid_objects):
        safe_name   = sanitize_filename(obj_name)
        step_path   = os.path.join(OUTPUT_DIR, safe_name + ".step")
        script_name = "{:03d}_{}.py".format(i + 1, safe_name)
        script_path = os.path.join(SCRIPTS_DIR, script_name)

        code = PER_OBJECT_SCRIPT.format(
            project_file_repr = repr(project_file),
            design_repr       = repr(design_name),
            obj_name_repr     = repr(obj_name),
            filepath_repr     = repr(step_path),
        )

        with open(script_path, "w") as f:
            f.write(code)

        script_paths.append(script_path)
        msg("  Generated: " + script_name)

    # Generate run_all.bat
    bat_path = os.path.join(SCRIPTS_DIR, "run_all.bat")
    bat_lines = ["@echo off",
                 "echo Maxwell STEP Batch Export",
                 "echo =========================",
                 "set MAXWELL=\"{}\" ".format(MAXWELL_EXE),
                 ""]
    for i, sp in enumerate(script_paths):
        bat_lines.append("echo [{}/{}] Exporting...".format(i + 1, len(script_paths)))
        bat_lines.append("%MAXWELL% -RunScriptAndExit \"{}\"".format(sp))
        bat_lines.append("")
    bat_lines += ["echo Done.", "pause"]

    with open(bat_path, "w") as f:
        f.write("\r\n".join(bat_lines))

    msg("=" * 50)
    msg("Generated {} scripts + run_all.bat".format(len(solid_objects)))
    msg("Folder: " + SCRIPTS_DIR)
    msg("-> Double-click run_all.bat to export all objects automatically.")
    msg("   (Close Maxwell first, or at minimum save the project.)")
    msg("=" * 50)
