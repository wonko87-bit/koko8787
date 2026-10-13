"""python -m efield2d file.dxf   - list the shapes found in a DXF file."""
import sys

from .dxf_io import import_dxf

if len(sys.argv) < 2:
    raise SystemExit("usage: python -m efield2d file.dxf   (GUI: python -m efield2d.app)")
print(import_dxf(sys.argv[1]).describe())
