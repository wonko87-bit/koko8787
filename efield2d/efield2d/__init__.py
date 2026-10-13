"""efield2d - 2D electric field FEM (planar / axisymmetric), DC & AC."""
from .materials import Material, EPS0
from .geometry import Polygon, rect
from .problem import Problem, Region, Conductor, DirichletLine, Refinement
from .mesh import Mesh, build_mesh
from .fem import Solution, solve
from . import post

__all__ = [
    "Material", "EPS0", "Polygon", "rect", "Problem", "Region", "Conductor",
    "DirichletLine", "Refinement", "Mesh", "build_mesh", "Solution", "solve",
    "post",
]
